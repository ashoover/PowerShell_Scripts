#Copy-Item "\\path\to\icons\folder" -Destination "C:\" -Recurse

#### Word Icon ####
$WordurlShortcutPath = "$HOME\Desktop\Word.url"
$Wordurl = 'https://word.cloud.microsoft/'
$WordiconPath = 'C:\365_icons\word.ico'

# Create the shortcut
$WordurlShortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($WordurlShortcutPath)
$WordurlShortcut.TargetPath = $Wordurl
$WordurlShortcut.Save()

# Append the icon information.
'IconIndex=0', "IconFile=$WordiconPath" | Add-Content -LiteralPath $WordurlShortcutPath

if (Test-Path -Path $WordurlShortcutPath) {
    Write-Host "Word Icon created."
} else {
    Write-Host "Word Icon not created."
}


#### Excel Icon ####
$ExcelurlShortcutPath = "$HOME\Desktop\Excel.url"
$Excelurl = 'https://word.cloud.microsoft/'
$ExceliconPath = 'C:\365_icons\excel.ico'

$ExcelurlShortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($ExcelurlShortcutPath)
$ExcelurlShortcut.TargetPath = $Excelurl
$ExcelurlShortcut.Save()

'IconIndex=0', "IconFile=$ExceliconPath" | Add-Content -LiteralPath $ExcelurlShortcutPath

if (Test-Path -Path $ExcelurlShortcutPath) {
    Write-Host "Excel Icon created."
} else {
    Write-Host "Excel Icon not created."
}

#### Outlook Icon ####
$OutlookurlShortcutPath = "$HOME\Desktop\Outlook.url"
$Outlookurl = 'https://word.cloud.microsoft/'
$OutlookiconPath = 'C:\365_icons\Outlook.ico'

$OutlookurlShortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($OutlookurlShortcutPath)
$OutlookurlShortcut.TargetPath = $Outlookurl
$OutlookurlShortcut.Save()

'IconIndex=0', "IconFile=$OutlookiconPath" | Add-Content -LiteralPath $OutlookurlShortcutPath

if (Test-Path -Path $OutlookurlShortcutPath) {
    Write-Host "Outlook Icon created."
} else {
    Write-Host "Outlook Icon not created."
}


#### PowerPoint Icon ####
$PowerPointurlShortcutPath = "$HOME\Desktop\PowerPoint.url"
$PowerPointurl = 'https://word.cloud.microsoft/'
$PowerPointiconPath = 'C:\365_icons\powerpoint.ico'

$PowerPointurlShortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($PowerPointurlShortcutPath)
$PowerPointurlShortcut.TargetPath = $PowerPointurl
$PowerPointurlShortcut.Save()

'IconIndex=0', "IconFile=$PowerPointiconPath" | Add-Content -LiteralPath $PowerPointurlShortcutPath

if (Test-Path -Path $PowerPointurlShortcutPath) {
    Write-Host "PowerPoint Icon created."
} else {
    Write-Host "PowerPoint Icon not created."
}
