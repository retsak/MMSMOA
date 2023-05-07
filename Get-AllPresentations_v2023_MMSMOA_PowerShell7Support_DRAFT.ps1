<#
    This code is still a work in progress - it should be functional but you may run into issues.
#>
$conferenceName = "mms2023atmoa"
$dates = @(
    '2023-04-30',
    '2023-05-01',
    '2023-05-02',
    '2023-05-03',
    '2023-05-04'
)
$urls = @()
$dates | ForEach-Object {$urls += "https://$conferenceName.sched.com/$psitem/list/descriptions"}

$MaxThreads = 10
$ThrottleLimit = $MaxThreads

Add-Type -AssemblyName System.Windows.Forms
$progressPreference = 'silentlyContinue'

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

try {

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
    Invoke-WebRequest -UseBasicParsing -Uri "https://$conferenceName.sched.com/login" -Method "POST" -Headers @{
        "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        "Origin" = "https://$conferenceName.sched.com"
        "Referer" = "https://$conferenceName.sched.com/login"
        "Upgrade-Insecure-Requests" = "1"
    } -ContentType "application/x-www-form-urlencoded" -Body "landing_conf=https%3A%2F%2F$conferenceName.sched.com&username=$schedUserName&password=$schedPassword&login=" -SessionVariable newSession | Out-Null

    Write-Output "Using authenticated session."
}

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
            $files | ForEach-Object -ThrottleLimit $MaxThreads -Parallel {
                $file = $_
                $fileUrl = $file.GetAttributeValue("href", "")
                $fileName = [uri]::UnescapeDataString(($fileUrl -split '/')[-1])
                $fileName = $fileName -replace '[\x00-\x1F\x7F<>:"/\\|?*]', '_'
                
                Write-Output "  $filename"
                $eventFolderPath = Join-Path $using:folder $using:eventName
                if (!(Test-Path $eventFolderPath)) {
                    New-Item -Type Directory -Path $eventFolderPath
                }

                $destinationPath = Join-Path $eventFolderPath $fileName
                if (!(Test-Path $destinationPath)) {
                    if ($newSession) {
                        Invoke-RestMethod -Uri $fileUrl -OutFile $destinationPath -WebSession $newSession
                    }
                    else {
                        Invoke-RestMethod -Uri $fileUrl -OutFile $destinationPath
                    }
                }
            }
        }
    }
}
} catch {
    Write-Error $_
} finally {
    Remove-Item -Path $folder\HtmlAgilityPack.$HtmlAgilityPackversion.nupkg -ErrorAction SilentlyContinue
    Remove-Item -Path $folder\HtmlAgilityPack\ -Recurse -Force -ErrorAction SilentlyContinue    
}
