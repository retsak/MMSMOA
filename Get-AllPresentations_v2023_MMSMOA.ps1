$conferenceName = "mms2023atmoa"
$dates = @(
    '2023-04-30',
    '2023-05-01',
    '2023-05-02',
    '2023-05-03',
    '2023-05-04'
)
$urls = @()
$dates | ForEach-Object { $urls += "https://$conferenceName.sched.com/$psitem/list/descriptions" }

# Adds the System.Windows.Forms assembly to access the FolderBrowserDialog class
Add-Type -AssemblyName System.Windows.Forms
$progressPreference = 'silentlyContinue'

# Prompts the user to select a folder
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
    Description  = "Select a folder"
    RootFolder   = "MyComputer"
    SelectedPath = $initialDirectory
}

# Sets the selected folder to $folder if the user clicks OK
if ($FileBrowser.ShowDialog() -eq "OK") {
    $folder = $FileBrowser.SelectedPath
}
else {
    break
}

# Confirms folder selection
Read-Host "This will download all presentations to: $folder (press enter to continue)"

# Prompts the user for their login credentials if required to access the files
$schedCreds = Get-Credential -Message "Enter your sched username and password (or cancel for only publically available content)"

# Sets the schedUserName variable to "blank" if the user cancels, or to the provided username otherwise
if ($null -eq $schedCreds.UserName) {
    $schedUserName = 'blank'
}
else {
    $schedUserName = $schedCreds.UserName
    $schedPassword = $schedCreds.GetNetworkCredential().Password
}

# Logs in to the website using the provided credentials and creates a new web session variable
if ($schedUserName -ne "blank") {
    Invoke-WebRequest -UseBasicParsing -Uri "https://$conferenceName.sched.com/login" -Method "POST" -Headers @{
        "Accept"                    = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        "Origin"                    = "https://$conferenceName.sched.com"
        "Referer"                   = "https://$conferenceName.sched.com/login"
        "Upgrade-Insecure-Requests" = "1"
    } -ContentType "application/x-www-form-urlencoded" -Body "landing_conf=https%3A%2F%2F$conferenceName.sched.com&username=$schedUserName&password=$schedPassword&login=" -SessionVariable newSession | Out-Null

    Write-Output "Using authenticated session."
}

$urls | ForEach-Object {
    $url = $_
    $res = Invoke-WebRequest -Uri $url -WebSession $newSession

    $res.ParsedHtml.documentElement.getElementsByClassName('sched-container') | ForEach-Object {
        $result = $_
        if ($result.innerHTML -like "*sched-container-inner*" -and $result.innerHTML -like "*sched-file*") {
            $eventName = (($result.innerText).Split([Environment]::NewLine)[0]).Trim()
            $eventName = $eventName -replace '[\\/:*?"<>|]', '_'

            $result.getElementsByTagName('div') | Where-Object { $_.ClassName -match '\bsched-file\b' } | ForEach-Object {
                $file = $_.innerHTML
                $fileUrl = ($file.Split(" ") | Where-Object { $_ -match "href" }).replace("href=", "").replace('"', '')
                $fileName = [uri]::UnescapeDataString(($fileUrl.Split('/'))[-1])
                $fileName = $fileName -replace '[\\/:*?"<>|]', '_'
                if (($PSVersionTable.PSVersion).Major -ne 7) {
                    $fileName = $fileName.Replace('[', '').Replace(']', '')
                }

                $eventPath = Join-Path $folder $eventName
                if (!(Test-Path $eventPath)) {
                    New-Item -Type Directory -Path $eventPath
                }

                $filePath = Join-Path $eventPath $fileName
                if (!(Test-Path $filePath)) {
                    Write-Output "Downloading $fileName to: $filePath"
                    if ($newSession) {
                        Invoke-WebRequest $fileUrl -OutFile "$folder\$eventName\$fileName" -WebSession $newSession
                    }
                    else {
                        Invoke-WebRequest $fileUrl -OutFile "$folder\$eventName\$fileName"
                    }
                }
            }
        }
    }
}
