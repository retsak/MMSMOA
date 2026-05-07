#Requires -Version 7.0
<#
.SYNOPSIS
    Script downloads presentations from the "mms2026atmoa" conference.
.DESCRIPTION
    This PowerShell script downloads presentations from the "mms2026atmoa" conference.
    It retrieves the presentations' URLs from the conference's schedule webpage and saves the files to a user-selected folder.
    The script also supports authenticated sessions if the user provides their sched.com credentials.
.NOTES
    This is written for PowerShell 7
.LINK
    https://github.com/retsak/MMSMOA
.EXAMPLE
    ."Get-AllPresentations_v2026_MMSMOA_PowerShell7Support.ps1"
#>

# Define the conference name and the conference schedule dates
$conferenceName = "mms2026atmoa"
$dates = @(
    '2026-05-03',
    '2026-05-04',
    '2026-05-05',
    '2026-05-06',
    '2026-05-07'
)
# Initialize an array to store the URLs of the presentations
$urls = @()
#Loop through the dates and generate the URLs for the presentations
$dates | ForEach-Object {$urls += "https://$conferenceName.sched.com/$psitem/list/descriptions"}

# Set the maximum number of concurrent threads for downloading files
$MaxThreads = 10

# Resolve destination folder (dialog on supported platforms, prompt fallback otherwise)
$folder = $null
$progressPreference = 'silentlyContinue'
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    # Create and configure a FileBrowser dialog to select a folder
    $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FileBrowser.Description = "Select a folder"
    $FileBrowser.rootfolder = "MyComputer"
    $FileBrowser.SelectedPath = $initialDirectory

    # Show the FileBrowser dialog and store the selected folder path
    if ($FileBrowser.ShowDialog() -eq "OK") {
        $folder = $FileBrowser.SelectedPath
    }
}
catch {
    Write-Output "Folder dialog unavailable. Please enter a destination path."
}

while ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder -PathType Container)) {
    $folder = Read-Host "Enter folder path for downloads"
    if ([string]::IsNullOrWhiteSpace($folder)) {
        continue
    }

    if (-not (Test-Path -LiteralPath $folder)) {
        $createFolder = Read-Host "Path does not exist. Create it? (Y/N)"
        if ($createFolder -match '^(?i:y|yes)$') {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
    }
}

# Confirm the download location with the user
Read-Host "This will download all presentations to: $folder (press enter to continue)"

try {
# Download and load HtmlAgilityPack.dll if not already present in the folder
if (!(Test-Path $folder\HtmlAgilityPack\lib\netstandard2.0\HtmlAgilityPack.dll -ErrorAction SilentlyContinue)) {
    $HtmlAgilityPackversion = (Find-Package -Name HtmlAgilityPack -Source https://www.nuget.org/api/v2).Version
    Invoke-WebRequest "https://www.nuget.org/api/v2/package/HtmlAgilityPack/$HtmlAgilityPackversion" -OutFile $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg
    Expand-Archive $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg -DestinationPath $folder\HtmlAgilityPack
}
Add-Type -LiteralPath $folder\HtmlAgilityPack\lib\netstandard2.0\HtmlAgilityPack.dll

# Request sched.com credentials for authenticated sessions
$schedCreds = Get-Credential -Message "Enter your sched username and password (or cancel for only publically available content)"

# Check if credentials were provided
if ($null -eq $schedCreds.UserName) {
    $schedUserName = 'blank'
} else {
    $schedUserName = $schedCreds.UserName
    $schedPassword = $schedCreds.GetNetworkCredential().Password
}

# Authenticate with the provided credentials if necessary
if ($schedUserName -ne "blank") {
    Invoke-WebRequest -UseBasicParsing -Uri "https://$conferenceName.sched.com/login" -Method "POST" -Headers @{
        "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        "Origin" = "https://$conferenceName.sched.com"
        "Referer" = "https://$conferenceName.sched.com/login"
        "Upgrade-Insecure-Requests" = "1"
    } -ContentType "application/x-www-form-urlencoded" -Body "landing_conf=https%3A%2F%2F$conferenceName.sched.com&username=$schedUserName&password=$schedPassword&login=" -SessionVariable newSession | Out-Null

    Write-Output "Using authenticated session."
}
# Iterate over the URLs of the presentations
$unnamedSessionCounter = 0
$urls | ForEach-Object {
    $url = $_
    # Send requests with or without authenticated sessions
    if ($newSession) {
        $response = Invoke-RestMethod -Uri $url -WebSession $newSession
    }
    else {
        $response = Invoke-RestMethod -Uri $url
    }
    
    # Parse the HTML response using HtmlAgilityPack
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($response)

    # Find the sched-container elements containing the presentation files
    $schedContainerElements = $doc.DocumentNode.SelectNodes('//div[contains(@class, "sched-container")]')

    # Iterate over the sched-container elements
    $schedContainerElements | ForEach-Object {
        $result = $_
        # Check if the element contains presentation files
        if ($result.InnerHtml.Contains("sched-container-inner") -and $result.InnerHtml.Contains("sched-file")) {
            $eventName = $null
            $eventNameNode = $result.SelectSingleNode('.//a[contains(@class, "name")]')
            if ($eventNameNode) {
                $eventName = [HtmlAgilityPack.HtmlEntity]::DeEntitize($eventNameNode.InnerText).Trim()
            }

            if ([string]::IsNullOrWhiteSpace($eventName)) {
                $eventName = ($result.InnerText -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -First 1)
            }

            if ([string]::IsNullOrWhiteSpace($eventName)) {
                $unnamedSessionCounter++
                $eventName = "Session_$unnamedSessionCounter"
            }

            $eventName = $eventName -replace '[\x00-\x1F\x7F<>:"/\\|?*]', '_'
            Write-Output "$($eventName):"
            # Find and iterate over the file links in the sched-file elements
            $files = $result.SelectNodes('.//div[contains(@class, "sched-file")]/a[@href]')
            $files | ForEach-Object -ThrottleLimit $MaxThreads -Parallel {
                $file = $_
                $fileUrl = $file.GetAttributeValue("href", "")
                $fileName = [uri]::UnescapeDataString(($fileUrl -split '/')[-1])
                $fileName = $fileName -replace '[\x00-\x1F\x7F<>:"/\\|?*]', '_'
                
                Write-Output "  $filename"
                # Create a directory for each event if it doesn't exist
                $eventFolderPath = Join-Path $using:folder $using:eventName
                if (!(Test-Path $eventFolderPath)) {
                    New-Item -Type Directory -Path $eventFolderPath
                }
                # Download the file to the destination folder
                $destinationPath = Join-Path $eventFolderPath $fileName
                if (!(Test-Path $destinationPath)) {
                    if ($using:newSession) {
                        Invoke-RestMethod -Uri $fileUrl -OutFile $destinationPath -WebSession $using:newSession -ErrorAction SilentlyContinue
                    }
                    else {
                        Invoke-RestMethod -Uri $fileUrl -OutFile $destinationPath -ErrorAction SilentlyContinue
                    }
                }
            }
        }
    }
}
} catch {
    Write-Error $_
} finally {
    # Clean up downloaded HtmlAgilityPack files
    Remove-Item -Path $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg -ErrorAction SilentlyContinue
    Remove-Item -Path $folder\HtmlAgilityPack\ -Recurse -Force -ErrorAction SilentlyContinue    
}
