#############################
#
#   Export Script for Backup 
#   Written by Brad Roberts
#   Backs up user data for 2010/2013, config data - copies to $drFolderPath below
#   Updated 14 January 2014
#
#############################

Function Add-Zip{
	Param([string]$ZipFilename)
	If(-Not (Test-Path($ZipFilename))){
		Set-Content $ZipFilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(Dir $ZipFilename).IsReadOnly = $False	
	}
	$ShellApplication = New-Object -COM Shell.Application
	$ZipPackage = $ShellApplication.NameSpace($ZipFilename)
	ForEach($File in $Input){ 
		$ZipPackage.CopyHere($File.FullName)
		Start-Sleep -Milliseconds 15000
	}
}

### Import Lync Module 
Import-Module "C:\Program Files\Common Files\Microsoft Lync Server 2013\Modules\Lync\Lync.psd1"

### Variables To Set 
$folderPath = "C:\Backup" 
$lengthOfBackup = "-30" 
$drFolderPath = "\\company.com\Lync\Backup" 
$poolFQDN = "lync-2013-pool.company.com"

### Production – Delete Older Than x Days 
get-childitem $folderPath -recurse | where {$_.lastwritetime -lt (get-date).adddays($lengthOfBackup) -and -not $_.psiscontainer} |% {remove-item $_.fullname -force }

### Production – Delete Empty Folders 
$a = Get-ChildItem $folderPath -recurse | Where-Object {$_.PSIsContainer -eq $True} 
$a | Where-Object {$_.GetFiles().Count -eq 0} | Remove-Item

### Production – Get Date and Create Folder 
$currDate = get-date -uformat "%a-%m-%d-%Y-%H-%M" 
New-Item $folderPath\$currDate -Type Directory

### Delete Older Than x Days – DR Side 
get-childitem $drFolderPath -recurse | where {$_.lastwritetime -lt (get-date).adddays($lengthOfBackup) -and -not $_.psiscontainer} |% {remove-item $_.fullname -force }

### Delete Empty Folders – DR Side 
$a = Get-ChildItem $drFolderPath -recurse | Where-Object {$_.PSIsContainer -eq $True} 
$a | Where-Object {$_.GetFiles().Count -eq 0} | Remove-Item

### Message Out 
Write-Host -ForegroundColor Green "Backup to server in progress"
$strTranscript += "<br>Backup to server in progress...<br>"

### Export CMS/XDS and LIS 
Export-CsConfiguration -FileName $folderPath\$currDate\XdsConfig.zip 
Export-CsLisConfiguration -FileName $folderPath\$currDate\LisConfig.zip

### Export Voice Information
Get-CsDialPlan | Export-Clixml -path $folderPath\$currDate\DialPlan.xml
Get-CsVoicePolicy | Export-Clixml -path $folderPath\$currDate\VoicePolicy.xml
Get-CsVoiceRoute | Export-Clixml -path $folderPath\$currDate\VoiceRoute.xml
Get-CsPstnUsage | Export-Clixml -path $folderPath\$currDate\PSTNUsage.xml
Get-CsVoiceConfiguration | Export-Clixml -path $folderPath\$currDate\VoiceConfiguration.xml
Get-CsTrunkConfiguration | Export-Clixml -path $folderPath\$currDate\TrunkConfiguration.xml

### Export RGS Config 
Export-CsRgsConfiguration -Source "service:ApplicationServer:$poolFQDN" -FileName $folderPath\$currDate\RgsConfig.zip
Write-Host -ForegroundColor Green "XDS, LIS and RGS backup to server is completed." 
$strTranscript += "<br>XDS, LIS and RGS backup to server is complete."

### Export User Information 
# Export 2013 data
Export-CsUserData -PoolFqdn $poolFQDN -FileName $folderPath\$currDate\Lync2013UserData.zip
$strTranscript += "<br>Export of Lync 2013 user data complete."
# Export 2010 data
C:\Admin\Software\dbimpexp /hrxmlfile:"$folderPath\$currDate\Lync2010UserData.xml" /sqlserver:sql-01
$strTranscript += "<br>Export of Lync 2010 user data complete."

#Create new zip file
Get-ChildItem $folderPath\$currDate\Lync2010UserData.xml | Add-Zip $folderPath\$currDate\Lync2010UserData.zip
Remove-Item $folderPath\$currDate\Lync2010UserData.xml

### Copy Files to DR Server 
robocopy $folderPath/$currDate $drFolderPath/$currDate /COPY:DATSO /S
$strTranscript += "<br><br>Files copied to $drfolderPath\$currDate"

#Stop-Transcript

### Email Transcript
$strSMTPServer = "smtp.company.com"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($strSMTPServer)

$strHTMLHeader = $strHTMLHeader + "<html>"
$strHTMLHeader = $strHTMLHeader + "<head>"
$strHTMLHeader = $strHTMLHeader + "<style type=""text/css"">"
$strHTMLHeader = $strHTMLHeader + "body {"
$strHTMLHeader = $strHTMLHeader + "margin-bottom: 0;"
$strHTMLHeader = $strHTMLHeader + "margin-left: 0;"
$strHTMLHeader = $strHTMLHeader + "margin-right: 0;"
$strHTMLHeader = $strHTMLHeader + "margin-top: 0;	 "
$strHTMLHeader = $strHTMLHeader + "FONT-FAMILY: Tahoma, Verdana;"
$strHTMLHeader = $strHTMLHeader + "font-size: 12px;"
$strHTMLHeader = $strHTMLHeader + "COLOR: black}"
$strHTMLHeader = $strHTMLHeader + "</style>"
$strHTMLHeader = $strHTMLHeader + "</head>"
$strHTMLHeader = $strHTMLHeader + "<body>"

$strComputerName = gc env:computername
$dtTimeNow = get-date
$strScriptInfo = "<br><br><br><b>Script Info</b><br>Script Name:  " + $MyInvocation.MyCommand.Definition + "<br>Time:  " + $dtTimeNow + "<br>Run From:  " + $strComputerName
$strHTMLFooter = $strHTMLFooter + $strScriptInfo
$strHTMLFooter = $strHTMLFooter + "</body>"
$strHTMLFooter = $strHTMLFooter + "</html>"
#$strTranscript = Get-Content $logpath
$strHTMLBody = $strHTMLHeader + $strTranscript + $strHTMLFooter

$msg = New-Object System.Net.Mail.MailMessage 
$msg.From = "no-reply@company.com" 
$msg.To.Add("lync.admin@company.com") 
$msg.Subject = "Lync Backup Report - " + (Get-Date -format D)
$msg.IsBodyHtml = $true
$msg.body = $strHTMLBody
$smtp.Send($msg)
