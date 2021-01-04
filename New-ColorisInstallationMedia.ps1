<#

.SYNOPSIS
Makes an offline installer with all dependencies (typically on a flash drive)

.DESCRIPTION
Makes an offline installer with all dependencies (typically on a flash drive) from an remote repository.

.EXAMPLE
./New-ColorisInstallationMedia.ps1 -Destination G:
Make an installation media on G: drive from default repository (ftp://maintenance.cosoluce.fr).

.EXAMPLE
./New-ColorisInstallationMedia.ps1 -Destination C:\temp\ -RepoBaseAddress ftp://vm-supernova-qa
Make a local installation source in C:\temp\ from ftp://vm-supernova-qa.

#>
param(
    # Local directory where installation source will be downloaded
    [Parameter(Mandatory = $true, HelpMessage = "Indiquez le répertoire dans lequel la distribution locale sera générée.")]
    [ValidateScript( { Test-Path $_ -PathType Container })]
    [string]
    $Destination,

    # Repository base address
    [ValidateSet("ftp://maintenance.cosoluce.fr", "ftp://beta-maintenance.cosoluce.fr", "ftp://vm-supernova-st", "ftp://vm-supernova-qa")]
    [string]
    $RepoBaseAddress = "ftp://maintenance.cosoluce.fr"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

function DownloadFile {
    param(
        [string]
        $url,
        [string]
        $targetFile
    )

    $uri = New-Object -TypeName System.Uri -ArgumentList $url
    
    # Get file size
    $request = [System.Net.WebRequest]::Create($uri)
    $request.Method = [System.Net.WebRequestMethods+Ftp]::GetFileSize
    $response = $request.GetResponse()
    $fileSizeInBytes = $response.ContentLength

    # Download the file
    $request = [System.Net.WebRequest]::Create($uri)
    $response = $request.GetResponse()

    $responseStream = $response.GetResponseStream()

    $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create

    $buffer = New-Object byte[] 1MB
    $bytesCount = $responseStream.Read($buffer, 0, $buffer.length)

    $downloadedBytes = 0
    $TotalDownloadedBytesCopy = $script:TotalDownloadedBytes

    try {
        while ($bytesCount -gt 0) {
            $targetStream.Write($buffer, 0, $bytesCount)
            $bytesCount = $responseStream.Read($buffer, 0, $buffer.length)
            $downloadedBytes += $bytesCount
            
            $fileName = [System.IO.Path]::GetFileName($url)
            
            # Increment a copy of $script:TotalDownloadedBytes to prevent miscalculation after a retry
            $TotalDownloadedBytesCopy += $bytesCount
            $totalCompletedPercent = [Math]::Floor(($TotalDownloadedBytesCopy / $TotalSizeInBytes) * 100)

            $TotalDownloadedMBytesCopy = [Math]::Floor($TotalDownloadedBytesCopy / 1MB)
            $TotalSizeInMBytes = [Math]::Floor($TotalSizeInBytes / 1MB)
            
            Write-Progress `
                -Id 1 `
                -Activity "Téléchargement global" `
                -Status "$TotalDownloadedMBytesCopy / $TotalSizeInMBytes Mo ($totalCompletedPercent%)" `
                -PercentComplete $totalCompletedPercent
            
            $completedPercent = [Math]::Floor(($downloadedBytes / $fileSizeInBytes) * 100)
            $fileSizeInKBytes = [Math]::Floor($fileSizeInBytes / 1KB)
            $downloadedKBytes = [Math]::Floor($downloadedBytes / 1KB)

            Write-Progress `
                -Id 2 `
                -ParentID 1 `
                -Activity "Téléchargement de $fileName" `
                -Status "$downloadedKBytes / $fileSizeInKBytes Ko ($completedPercent%)" `
                -PercentComplete $completedPercent
        }
    
        Write-Progress -Activity "Téléchargement de $fileName terminé." -Id 2 -ParentId 1 -Completed

        # Set the script-scoped $TotalDownloadedBytes only after the download has completed successfully
        $script:TotalDownloadedBytes = $TotalDownloadedBytesCopy
    }
    catch {
        throw $_
    }
    finally {
        $targetStream.Dispose()
        $responseStream.Dispose()
    }
}

function DownloadFtpDirectory {
    param(
        [string]
        $url, 
        [string]
        $localPath,
        [switch]
        $sizeEstimation
    )

    $listRequest = [Net.WebRequest]::Create($url)
    $listRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails

    $lines = @()

    $listResponse = $listRequest.GetResponse()
    $listStream = $listResponse.GetResponseStream()
    $listReader = New-Object System.IO.StreamReader($listStream)

    while (-Not $listReader.EndOfStream) {
        $line = $listReader.ReadLine()
        $lines += $line
    }

    $listReader.Dispose()
    $listStream.Dispose()
    $listResponse.Dispose()

    $totalSize = 0

    foreach ($line in $lines) {
        # Response on IIS FTP Server for a file is formatted as :
        # "12-02-20  04:24PM             60570002 Ambre_6.00.00.cosopack"
        # And for a directory as :
        # "12-02-20  11:14PM       <DIR>          supernova"
        $tokens = $line.Split(" ", 9, [StringSplitOptions]::RemoveEmptyEntries)
        $name = $tokens[3]
        # in case of a file, 2nd entry will be file size
        $size = $tokens[2]
        # in case of a directory, 2nd entry will be the type ("<DIR>)
        $type = $tokens[2]

        $localFilePath = Join-Path $localPath $name
        $fileUrl = "$url/$name"

        if ($type -eq '<DIR>') {
            if ($sizeEstimation) {
                $totalSize += DownloadFtpDirectory $fileUrl $localFilePath -sizeEstimation
            }
            else {
                New-Item $localFilePath -Type Directory -Force | Out-Null
                DownloadFtpDirectory $fileUrl $localFilePath
            }
        }
        else {
            if ($sizeEstimation) {
                $totalSize += $size
            }
            else {
                $fileDownloaded = $false
                $downloadFailure = 0

                while ($fileDownloaded -eq $false) {
                    try {
                        DownloadFile $fileUrl $localFilePath
                        $fileDownloaded = $true
                        if ($downloadFailure -ge 1) {
                            Write-Information "Nouvelle tentative de téléchargement de $fileUrl réussie."
                        }
                    }
                    catch {
                        $downloadFailure++
                        Write-Warning "Echec du téléchargement de $fileUrl. Nouvelle tentative..."
                        if ($downloadFailure -eq 3) {
                            throw "Impossible de télécharger $fileUrl après 3 tentatives.`r`nDétails de l'erreur :`r`n$_"
                        }
                    }
                }
            }
        }
    }

    if ($sizeEstimation) {
        return $totalSize
    }
}

"Génération d'une installation locale depuis $RepoBaseAddress"

# Check if drive is an USB drive
"Analyse de la destination"
$IsFlashDrive = $false
$FlashDriveLetters = @()

Get-CimInstance -Class Win32_Volume -Filter ("DriveType={0}" -f [int][System.IO.DriveType]::Removable) | `
    ForEach-Object {
    $FlashDriveLetters += $_.DriveLetter
}

# If $Destination length is less than or equal to 3, it should be a drive
if ( $Destination.Length -le 3) {
    # Remove the end slash to compare
    $IsFlashDrive = ($Destination.Substring(0, 2) -in $FlashDriveLetters)
}

if (-Not $IsFlashDrive) {
    if ((Get-Item $Destination).BaseName -ne "CosoluceInstaller") {
        $Destination = New-Item (Join-Path $Destination "CosoluceInstaller") -ItemType Directory -Force
        New-Item $Destination -Type Directory -Force | Out-Null
    }
}

$DirectoriesToDownload = "cosopackages", "tools", "docs"

# Estimate size
$TotalSizeInBytes = 0

$DirectoriesToDownload | `
    ForEach-Object {
    $TotalSizeInBytes += DownloadFtpDirectory "$RepoBaseAddress/$_" (Join-Path $Destination $_) -sizeEstimation
}

"Taille du téléchargement estimée à $([Math]::Floor($TotalSizeInBytes / 1MB))Mo."

Write-Progress -Activity "Début du téléchargement" -Id 1

$DirectoriesToDownload | `
    ForEach-Object {
    $LocalTargetDirectory = (Join-Path $Destination $_)
    Remove-Item $LocalTargetDirectory -Force -Recurse -ErrorAction Ignore
    New-Item $LocalTargetDirectory -Type Directory -Force | Out-Null
    DownloadFtpDirectory "$RepoBaseAddress/$_" $LocalTargetDirectory
}

Write-Progress -Activity "Téléchargement terminé." -Id 1

# Remove downloaded catalog
Remove-Item (Join-Path $Destination "tools/supernova/Bootstrapper/catalog.zip") -Force -ErrorAction Ignore

# Remove repoBaseAddress from catalog
$CatalogMetadataPath = Join-Path $Destination "tools/supernova/Bootstrapper/catalog/files/metadata.json"
$CatalogMetadata = Get-Content $CatalogMetadataPath | ConvertFrom-Json
$CatalogMetadata.repoBaseAddress = ""
ConvertTo-Json $CatalogMetadata | Set-Content $CatalogMetadataPath

# Delete existing catalog
Remove-Item (Join-Path $Destination "catalog.zip") -Force -ErrorAction Ignore

# Generate catalog.zip
Compress-Archive `
    -Path (Join-Path $Destination "tools/supernova/Bootstrapper/catalog/files/*") `
    -DestinationPath (Join-Path $Destination "catalog.zip") `
    -Force

# Delete previous bootstrappers (if any)
Get-ChildItem $Destination -Filter "Supernova.Client.Bootstrapper*.exe" | `
    ForEach-Object {
    Remove-Item -Path $_.FullName -Force
}

# Move the downloaded bootstrappers to root level
Get-ChildItem (Join-Path $Destination "tools/supernova/Bootstrapper/") -Filter "Supernova.Client.Bootstrapper*.exe" | `
    ForEach-Object {
    Move-Item -Path $_.FullName -Destination $Destination -Force
}

if ($IsFlashDrive) {
    # Rename drive
    $FormatedDriveLetter = ($Destination.ToUpper())[0]
    if ((Get-Volume -DriveLetter $FormatedDriveLetter).FileSystem -eq "FAT32") {
        # For FAT32 file system, we can use only 11 characters
        Set-Volume -DriveLetter $FormatedDriveLetter -NewFileSystemLabel "Coloris"
    }
    else {
        Set-Volume -DriveLetter $FormatedDriveLetter -NewFileSystemLabel "Cosoluce Coloris"
    }

    # Add an autorun.inf to set icon and label
    Get-ChildItem $Destination -Filter "Supernova.Client.Bootstrapper*.exe" | Select-Object -First 1 |`
        ForEach-Object {
        "[AutoRun]`r`nicon=$($_.BaseName).exe,0`r`nlabel=Cosoluce Coloris" | `
            Set-Content (Join-Path $Destination "autorun.inf")
    }
}

"Traitement terminé : votre support d'installation est prêt à être utilisé !"
# SIG # Begin signature block
# MIIccQYJKoZIhvcNAQcCoIIcYjCCHF4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU6wlPJ1yO+wNxZMvexVc3Y5IA
# H3igghegMIIFKTCCBBGgAwIBAgIQAufjyjh6OBzpoTtErdGO/jANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMTAwMTAwMDAwMFoXDTIzMTAw
# NjEyMDAwMFowZjELMAkGA1UEBhMCRlIxGzAZBgNVBAgTEk5vdXZlbGxlLUFxdWl0
# YWluZTEMMAoGA1UEBxMDUGF1MRUwEwYDVQQKEwxDb3NvbHVjZSBTQVMxFTATBgNV
# BAMTDENvc29sdWNlIFNBUzCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# ALNt2fNjD0Ufk6if1Jp/9uDTVT42kNEuE+38dqJPAPUh8VgDT1bKRe35i3xxHkDK
# 4G86gWE43M2QzaPZOzd8SgsDmbDs+6lLgU2Ap4PP2X6BpkWUjhl1gslfijcdzkvZ
# L17ueBPS87ZfsnsCEfSstaogofwEUkxS6xXmzNPYHs5BwtN072lBtgXIR+FjBhL5
# 42n+mMC0QfOMTgBUXfUqUq3cTbACMS9r+1aSPDtq6hjb3pa10brLNu3ZTfbC+7aN
# h2YtvNN7UI5jUGzSQpU7lUPyMJBe+qXiehexhu0/PVZFpUAgqofPr2ffczHd0HsY
# 2Iit4mGxRCeTxw/BgMvDK40CAwEAAaOCAcUwggHBMB8GA1UdIwQYMBaAFFrEuXsq
# CqOl6nEDwGD5LfZldQ5YMB0GA1UdDgQWBBQ2fEoz0uigTfWyyz6pY3C/PRvzfDAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwdwYDVR0fBHAwbjA1
# oDOgMYYvaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1jcy1n
# MS5jcmwwNaAzoDGGL2h0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMEwGA1UdIARFMEMwNwYJYIZIAYb9bAMBMCowKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCAYGZ4EMAQQBMIGEBggr
# BgEFBQcBAQR4MHYwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNv
# bTBOBggrBgEFBQcwAoZCaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0U0hBMkFzc3VyZWRJRENvZGVTaWduaW5nQ0EuY3J0MAwGA1UdEwEB/wQCMAAw
# DQYJKoZIhvcNAQELBQADggEBAEkpCp6pma1DJlE/pf6JH+Wt73A/CjAk4P45aQx6
# cKe8aGE2fDJyiuHO/YuH/Mhqni4X2YCwITs2AOHLluvT5EwOaqJzF+omldWoXsyD
# +7fyPKNqkpFrpOFCQ3+HyGP+vKsF2459WYl9nQ1IP+OgBMbDe5E0mX6K6BIZ4jhX
# RiSZuM76MgUjOnqHAK7xogu0lvUSVfeSeCyxDoTv0z9uVZaUXJM0rPId4kQ9xnKm
# DoRXOFRuXOVi3UST2GjJhKo4jWgqmdher3IN5wKt5Bxylgc39yTbryFkmyjv3pgJ
# KYYvHkfFb5q8/SNBwsfnsmfGcuVmHz39+0U6BYWqekM98V0wggUwMIIEGKADAgEC
# AhAECRgbX9W7ZnVTQ7VvlVAIMA0GCSqGSIb3DQEBCwUAMGUxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0xMzEw
# MjIxMjAwMDBaFw0yODEwMjIxMjAwMDBaMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNV
# BAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQD407Mcfw4Rr2d3B9MLMUkZz9D7
# RZmxOttE9X/lqJ3bMtdx6nadBS63j/qSQ8Cl+YnUNxnXtqrwnIal2CWsDnkoOn7p
# 0WfTxvspJ8fTeyOU5JEjlpB3gvmhhCNmElQzUHSxKCa7JGnCwlLyFGeKiUXULaGj
# 6YgsIJWuHEqHCN8M9eJNYBi+qsSyrnAxZjNxPqxwoqvOf+l8y5Kh5TsxHM/q8grk
# V7tKtel05iv+bMt+dDk2DZDv5LVOpKnqagqrhPOsZ061xPeM0SAlI+sIZD5SlsHy
# DxL0xY4PwaLoLFH3c7y9hbFig3NBggfkOItqcyDQD2RzPJ6fpjOp/RnfJZPRAgMB
# AAGjggHNMIIByTASBgNVHRMBAf8ECDAGAQH/AgEAMA4GA1UdDwEB/wQEAwIBhjAT
# BgNVHSUEDDAKBggrBgEFBQcDAzB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDCB
# gQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDBPBgNVHSAESDBGMDgG
# CmCGSAGG/WwAAgQwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQu
# Y29tL0NQUzAKBghghkgBhv1sAzAdBgNVHQ4EFgQUWsS5eyoKo6XqcQPAYPkt9mV1
# DlgwHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEL
# BQADggEBAD7sDVoks/Mi0RXILHwlKXaoHV0cLToaxO8wYdd+C2D9wz0PxK+L/e8q
# 3yBVN7Dh9tGSdQ9RtG6ljlriXiSBThCk7j9xjmMOE0ut119EefM2FAaK95xGTlz/
# kLEbBw6RFfu6r7VRwo0kriTGxycqoSkoGjpxKAI8LpGjwCUR4pwUR6F6aGivm6dc
# IFzZcbEMj7uo+MUSaJ/PQMtARKUT8OZkDCUIQjKyNookAv4vcn4c10lFluhZHen6
# dGRrsutmQ9qzsIzV6Q3d9gEgzpkxYz0IGhizgZtPxpMQBvwHgfqL2vmCSfdibqFT
# +hKUGIUukpHqaGxEMrJmoecYpJpkUe8wggZqMIIFUqADAgECAhADAZoCOv9YsWvW
# 1ermF/BmMA0GCSqGSIb3DQEBBQUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTAeFw0xNDEwMjIwMDAwMDBaFw0yNDEw
# MjIwMDAwMDBaMEcxCzAJBgNVBAYTAlVTMREwDwYDVQQKEwhEaWdpQ2VydDElMCMG
# A1UEAxMcRGlnaUNlcnQgVGltZXN0YW1wIFJlc3BvbmRlcjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBAKNkXfx8s+CCNeDg9sYq5kl1O8xu4FOpnx9kWeZ8
# a39rjJ1V+JLjntVaY1sCSVDZg85vZu7dy4XpX6X51Id0iEQ7Gcnl9ZGfxhQ5rCTq
# qEsskYnMXij0ZLZQt/USs3OWCmejvmGfrvP9Enh1DqZbFP1FI46GRFV9GIYFjFWH
# eUhG98oOjafeTl/iqLYtWQJhiGFyGGi5uHzu5uc0LzF3gTAfuzYBje8n4/ea8Ewx
# ZI3j6/oZh6h+z+yMDDZbesF6uHjHyQYuRhDIjegEYNu8c3T6Ttj+qkDxss5wRoPp
# 2kChWTrZFQlXmVYwk/PJYczQCMxr7GJCkawCwO+k8IkRj3cCAwEAAaOCAzUwggMx
# MA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsG
# AQUFBwMIMIIBvwYDVR0gBIIBtjCCAbIwggGhBglghkgBhv1sBwEwggGSMCgGCCsG
# AQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMIIBZAYIKwYBBQUH
# AgIwggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQBy
# AHQAaQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBj
# AGUAcAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAg
# AEMAUAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQ
# AGEAcgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBt
# AGkAdAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBj
# AG8AcgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBl
# AHIAZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTAfBgNVHSMEGDAWgBQVABIrE5iymQft
# Ht+ivlcNK2cCzTAdBgNVHQ4EFgQUYVpNJLZJMp1KKnkag0v0HonByn0wfQYDVR0f
# BHYwdDA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNz
# dXJlZElEQ0EtMS5jcmwwOKA2oDSGMmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydEFzc3VyZWRJRENBLTEuY3JsMHcGCCsGAQUFBwEBBGswaTAkBggrBgEF
# BQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRw
# Oi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNy
# dDANBgkqhkiG9w0BAQUFAAOCAQEAnSV+GzNNsiaBXJuGziMgD4CH5Yj//7HUaiwx
# 7ToXGXEXzakbvFoWOQCd42yE5FpA+94GAYw3+puxnSR+/iCkV61bt5qwYCbqaVch
# XTQvH3Gwg5QZBWs1kBCge5fH9j/n4hFBpr1i2fAnPTgdKG86Ugnw7HBi02JLsOBz
# ppLA044x2C/jbRcTBu7kA7YUq/OPQ6dxnSHdFMoVXZJB2vkPgdGZdA0mxA5/G7X1
# oPHGdwYoFenYk+VVFvC7Cqsc21xIJ2bIo4sKHOWV2q7ELlmgYd3a822iYemKC23s
# Ehi991VUQAOSK2vCUcIKSK+w1G7g9BQKOhvjjz3Kr2qNe9zYRDCCBs0wggW1oAMC
# AQICEAb9+QOWA63qAArrPye7uhswDQYJKoZIhvcNAQEFBQAwZTELMAkGA1UEBhMC
# VVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0
# LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTA2
# MTExMDAwMDAwMFoXDTIxMTExMDAwMDAwMFowYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgQXNzdXJlZCBJRCBDQS0xMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEA6IItmfnKwkKVpYBzQHDSnlZUXKnE0kEGj8kz/E1FkVyB
# n+0snPgWWd+etSQVwpi5tHdJ3InECtqvy15r7a2wcTHrzzpADEZNk+yLejYIA6sM
# NP4YSYL+x8cxSIB8HqIPkg5QycaH6zY/2DDD/6b3+6LNb3Mj/qxWBZDwMiEWicZw
# iPkFl32jx0PdAug7Pe2xQaPtP77blUjE7h6z8rwMK5nQxl0SQoHhg26Ccz8mSxSQ
# rllmCsSNvtLOBq6thG9IhJtPQLnxTPKvmPv2zkBdXPao8S+v7Iki8msYZbHBc63X
# 8djPHgp0XEK4aH631XcKJ1Z8D2KkPzIUYJX9BwSiCQIDAQABo4IDejCCA3YwDgYD
# VR0PAQH/BAQDAgGGMDsGA1UdJQQ0MDIGCCsGAQUFBwMBBggrBgEFBQcDAgYIKwYB
# BQUHAwMGCCsGAQUFBwMEBggrBgEFBQcDCDCCAdIGA1UdIASCAckwggHFMIIBtAYK
# YIZIAYb9bAABBDCCAaQwOgYIKwYBBQUHAgEWLmh0dHA6Ly93d3cuZGlnaWNlcnQu
# Y29tL3NzbC1jcHMtcmVwb3NpdG9yeS5odG0wggFkBggrBgEFBQcCAjCCAVYeggFS
# AEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABpAGYAaQBj
# AGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBwAHQAYQBu
# AGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQAC8AQwBQ
# AFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQByAHQAeQAg
# AEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0ACAAbABp
# AGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwByAHAAbwBy
# AGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBlAG4AYwBl
# AC4wCwYJYIZIAYb9bAMVMBIGA1UdEwEB/wQIMAYBAf8CAQAweQYIKwYBBQUHAQEE
# bTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYB
# BQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmwzLmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0
# dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5j
# cmwwHQYDVR0OBBYEFBUAEisTmLKZB+0e36K+Vw0rZwLNMB8GA1UdIwQYMBaAFEXr
# oq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqGSIb3DQEBBQUAA4IBAQBGUD7Jtygkpzgd
# tlspr1LPUukxR6tWXHvVDQtBs+/sdR90OPKyXGGinJXDUOSCuSPRujqGcq04eKx1
# XRcXNHJHhZRW0eu7NoR3zCSl8wQZVann4+erYs37iy2QwsDStZS9Xk+xBdIOPRqp
# FFumhjFiqKgz5Js5p8T1zh14dpQlc+Qqq8+cdkvtX8JLFuRLcEwAiR78xXm8TBJX
# /l/hHrwCXaj++wc4Tw3GXZG5D2dFzdaD7eeSDY2xaYxP+1ngIw/Sqq4AfO6cQg7P
# kdcntxbuD8O9fAqg7iwIVYUiuOsYGk38KiGtSTGDR5V3cdyxG0tLHBCcdxTBnU8v
# WpUIKRAmMYIEOzCCBDcCAQEwgYYwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERp
# Z2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMo
# RGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQQIQAufjyjh6
# OBzpoTtErdGO/jAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKA
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU5g2C9eTwvEnD5HT2/7F6af73OLAw
# DQYJKoZIhvcNAQEBBQAEggEAoVtoiseok0ViwNSUblFpEazgN+n29YKfFQHFyrvQ
# NjSjljzQ/bkCgtv3g97jiQ4Iml+GspdL9pP6rk735XD6Ye+I1dBnAE7+oRm6Cf+N
# Ll4H+jZyyU8RQX4/oZfN5XfuCXbflK4rOsfzNZ/weGYnXRTKZ8EYWJ/1Ls65eyQv
# bhxFM8v9IVxjfmAAVH8U9hR24kjSX2grmNZRtmvQcuPaB2hAMmrh2k9R475A3yy/
# sdyWM5KHmEJiUVJ03BinvX2TNwLwWrcuhNYXw2+s/Kwt8ze8pI1/45qek2yGKJ7L
# s+uVvqyfhzhWYbrAzPHF7sYIor0mdXEhn2UbNphwBfLaVqGCAg8wggILBgkqhkiG
# 9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdp
# Q2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERp
# Z2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJBgUrDgMC
# GgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjAxMjExMjMzMzA0WjAjBgkqhkiG9w0BCQQxFgQUsoFXcItnLubm1WEVDuJ6/KPy
# xucwDQYJKoZIhvcNAQEBBQAEggEAEbee4gLBgbay1Gjp0fFwNufGi+KyhVq3/pJU
# 6ludDcvn7/gv4HqglMilojZuOL+0Q5vIlwEh0kl+feGKzCupvAVpGwFiCh5J+7NX
# AEbWAu0AOpMc/s5OcTu88iKe3EkNO3Pluwpb0CeGoxUEj/ODX/LyRHnlFY3uy/14
# xF3vxBvu9IQrVg6jxMIx7IzysvfESzp50KKwSPmQ+rNd32mDlRP0TXdJ0xFeegtw
# /iAshNXWTxftpYkbGrT8U2fL1UzeZSG+8xgFfvxdR6DpCAZjdKELXBK4LICnZCXc
# Xe1P6K82vqPl3dKZC0fpS7BucBlzIllTRX6AsW4KSdICWZfGrw==
# SIG # End signature block
