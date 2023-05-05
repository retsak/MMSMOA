# Adds the System.Windows.Forms assembly to access the FolderBrowserDialog class
Add-Type -AssemblyName System.Windows.Forms

# Creates a new instance of FolderBrowserDialog
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# Sets the description for the folder selection dialog box
$FileBrowser.Description = "Select a folder"

# Sets the root folder to MyComputer
$FileBrowser.rootfolder = "MyComputer"

# Prompts the user to select a folder
$FileBrowser.SelectedPath = $initialDirectory

# Displays the folder selection dialog box and sets the selected folder to $folder if the user clicks OK
if ($FileBrowser.ShowDialog() -eq "OK") {
    $folder = $FileBrowser.SelectedPath
} else {
    break
}

# Prompts the user with a message that displays the folder where the files will be downloaded
Read-Host "This will download all presentations to: $folder (press enter to continue)"

# Prompts the user for their login credentials if required to access the files
$schedCreds = Get-Credential -Message "Enter your sched username and password (or cancel for only publically available content)"

# Sets the schedUserName variable to "blank" if the user cancels, or to the provided username otherwise
if ($null -eq $schedCreds.UserName) {
    $schedUserName = 'blank'
} else {
    $schedUserName = $schedCreds.UserName
    $schedPassword = $schedCreds.GetNetworkCredential().Password
}

# Logs in to the website using the provided credentials and creates a new web session variable
if ($schedUserName -ne "blank") {
    
    #Password is used to create a new web session variable that is used to download the files
    Invoke-WebRequest -UseBasicParsing -Uri "https://mms2023atmoa.sched.com/login" `
        -Method "POST" `
        -Headers @{
        "Accept"                    = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        "Accept-Encoding"           = "gzip, deflate, br"
        "Accept-Language"           = "en-US,en;q=0.9"
        "Cache-Control"             = "max-age=0"
        "DNT"                       = "1"
        "Origin"                    = "https://mms2023atmoa.sched.com"
        "Referer"                   = "https://mms2023atmoa.sched.com/login"
        "Sec-Fetch-Dest"            = "document"
        "Sec-Fetch-Mode"            = "navigate"
        "Sec-Fetch-Site"            = "same-origin"
        "Sec-Fetch-User"            = "?1"
        "Upgrade-Insecure-Requests" = "1"
        "sec-ch-ua"                 = "`" Not A;Brand`";v=`"99`", `"Chromium`";v=`"100`", `"Microsoft Edge`";v=`"100`""
        "sec-ch-ua-mobile"          = "?0"
        "sec-ch-ua-platform"        = "`"Windows`""
    } `
        -ContentType "application/x-www-form-urlencoded" `
        -Body "landing_conf=https%3A%2F%2Fmms2023atmoa.sched.com&username=$schedUserName&password=$schedPassword&login=" `
        -SessionVariable newSession | Out-Null

        Write-Output "Using authenicated session."
}

#Define an array of URLs to scrape
$urls = @(
    'https://mms2023atmoa.sched.com/2022-04-30/list/descriptions',
    'https://mms2023atmoa.sched.com/2022-05-01/list/descriptions',
    'https://mms2023atmoa.sched.com/2022-05-02/list/descriptions',
    'https://mms2023atmoa.sched.com/2022-05-03/list/descriptions',
    'https://mms2023atmoa.sched.com/2022-05-04/list/descriptions'
)

#Loop through each URL
$urls | ForEach-Object {
    $url = $_
    # If there is a new session, use it in the web request, otherwise, just make a normal web request
    if ($newSession) {
        $res = Invoke-WebRequest -Uri $url -WebSession $newSession
    }
    else {
        $res = Invoke-WebRequest -Uri $url
    }

    # Loop through each sched-container in the HTML and check if it contains a sched-file
    $res.ParsedHtml.documentElement.getElementsByClassName('sched-container') | ForEach-Object {
        $result = $_
        # If the sched-container contains sched-file, get the event name and download the file(s)
        if ($result.innerHTML -like "*sched-container-inner*" -and $result.innerHTML -like "*sched-file*") {

            # Get the event name and replace any invalid characters with underscores
            $eventName = (($result.innerText).Split([Environment]::NewLine)[0]).Trim()
            [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
                if ($_.length -gt 0) {
                    $eventName = $eventName.Replace($_, '_')
                }
            }

            # Get the HTML of all the sched-file elements
            $files = ($result.getElementsByTagName('div') | Where-Object { $_.ClassName -match '\bsched-file\b' }).innerHTML

            # Loop through each file element and download the file
            $files | ForEach-Object {
                $file = $_

                # Get the URL of the file and extract the file name
                $file = ($file.Split(" ") | Where-Object { $_ -match "href" }).replace("href=", "").replace('"', '')
                $fileName = $file.Split('/')
                $fileName = $fileName[$fileName.count - 1]

                # Fix any URL encoding issues with the file name
                $fileName = [uri]::UnescapeDataString($fileName)
                $eventName
                $file
                $fileName

                # Replace any invalid characters in the file name with underscores
                [System.IO.Path]::GetInvalidFileNameChars() | ForEach-Object {
                    if ($_.length -gt 0) {
                        $fileName = $fileName.Replace($_, '_')
                    }
                }
                #Escape Brackets - PowerShell 5.1
                if (($PSVersionTable.PSVersion).Major -ne 7) {
                    $filename = $filename.Replace('[', '').Replace(']', '')
                }

                # Create the event folder if it doesn't exist
                if (!(Test-Path "$folder\$eventName")) {
                    New-Item -Type Directory -Path "$folder\$eventName"
                }

                # Download the file if it doesn't already exist in the event folder
                if (!(Test-Path "$folder\$eventName\$fileName")) {
                    if ($newSession) {
                        Invoke-WebRequest $file -OutFile "$folder\$eventName\$fileName" -Verbose -WebSession $newSession
                    }
                    else {
                        Invoke-WebRequest $file -OutFile "$folder\$eventName\$fileName" -Verbose
                    }
                }
            }
        }
    }
}
