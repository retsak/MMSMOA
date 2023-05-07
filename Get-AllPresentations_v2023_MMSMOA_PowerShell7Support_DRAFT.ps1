<#
    This code is still a work in progress - it should be functional but you may run into issues.
#>
Add-Type -AssemblyName System.Windows.Forms

# Select a folder
$FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FileBrowser.Description = "Select a folder"
$FileBrowser.rootfolder = "MyComputer"
$FileBrowser.SelectedPath = $initialDirectory

if ($FileBrowser.ShowDialog() -eq "OK") {
    $folder = $FileBrowser.SelectedPath
} else {
    break
}

Read-Host "This will download all presentations to: $folder (press enter to continue)"

#Downlaod HtmlAgilityPack.dll
if (!(Test-Path $folder\HtmlAgilityPack\lib\netstandard2.0\HtmlAgilityPack.dll -ErrorAction SilentlyContinue)) {
    $HtmlAgilityPackversion = (Find-Package -Name HtmlAgilityPack -Source https://www.nuget.org/api/v2).Version
    Invoke-WebRequest "https://www.nuget.org/api/v2/package/HtmlAgilityPack/$HtmlAgilityPackversion" -OutFile $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg
    Expand-Archive $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg -DestinationPath $folder\HtmlAgilityPack
}
Add-Type -LiteralPath $folder\HtmlAgilityPack\lib\netstandard2.0\HtmlAgilityPack.dll

$schedCreds = Get-Credential -Message "Enter your sched username and password (or cancel for only publically available content)"

if ($null -eq $schedCreds.UserName) {
    $schedUserName = 'blank'
} else {
    $schedUserName = $schedCreds.UserName
    $schedPassword = $schedCreds.GetNetworkCredential().Password
}

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

    Write-Output "Using authenticated session."
}

$urls = @(
    'https://mms2023atmoa.sched.com/2023-04-30/list/descriptions',
    'https://mms2023atmoa.sched.com/2023-05-01/list/descriptions',
    'https://mms2023atmoa.sched.com/2023-05-02/list/descriptions',
    'https://mms2023atmoa.sched.com/2023-05-03/list/descriptions',
    'https://mms2023atmoa.sched.com/2023-05-04/list/descriptions'
)

$urls | ForEach-Object {
    $url = $_
    if ($newSession) {
        $response = Invoke-RestMethod -Uri $url -WebSession $newSession
    }
    else {
        $response = Invoke-RestMethod -Uri $url
    }
    
    #$doc = New-HtmlDocument -Html $response
    $doc = New-Object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($response)

    $schedContainerElements = $doc.DocumentNode.SelectNodes('//div[contains(@class, "sched-container")]')

    $schedContainerElements | ForEach-Object {
        $result = $_
        if ($result.InnerHtml.Contains("sched-container-inner") -and $result.InnerHtml.Contains("sched-file")) {
            $pattern = '<a.*?class="name".*?>([\s\S]*?)<\/a>'
            $eventName = [regex]::Match($result.InnerHtml, $pattern)
            $eventName = $eventName.Value.Split(">")[1].Trim().Split("<")[0].Trim()
            $eventName = $eventName -replace '[\x00-\x1F\x7F<>:"/\\|?*]', '_'
            Write-Output "$($eventName):"
            $files = $result.SelectNodes('.//div[contains(@class, "sched-file")]/a[@href]')
            $files | ForEach-Object {
                $file = $_
                $fileUrl = $file.GetAttributeValue("href", "")
                $fileName = [uri]::UnescapeDataString(($fileUrl -split '/')[-1])
                $fileName = $fileName -replace '[\x00-\x1F\x7F<>:"/\\|?*]', '_'
                
                Write-Output "  $filename"
                $eventFolderPath = Join-Path $folder $eventName
                if (!(Test-Path $eventFolderPath)) {
                    New-Item -Type Directory -Path $eventFolderPath
                }

                $destinationPath = Join-Path $eventFolderPath $fileName
                if (!(Test-Path $destinationPath)) {
                    if ($newSession) {
                        Invoke-WebRequest $fileUrl -OutFile $destinationPath -Verbose -WebSession $newSession
                    }
                    else {
                        Invoke-WebRequest $fileUrl -OutFile $destinationPath -Verbose
                    }
                }
            }
        }
    }
}
