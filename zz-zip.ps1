# Usage:
# powershell.exe -File .\zz-zip.ps1 [-m] -f <folder>
param([switch]$m, [string]$f="")
$ver="1.0"
$7z_path="$env:ProgramFiles\7-Zip\7z.exe"
if (-not (Test-Path -Path $7z_path -PathType Leaf)) {
	Write-Host "'$7z_path' not found.`nPlease get 7-Zip from https://www.7-zip.org/" -Foregroundcolor Red
	pause; exit
}

$f_is_folder=$false
$f_is_file=$false
if ($f) {
	Write-Host $f
	if (Test-Path -Path $f -PathType Container) {
		$f_is_folder=$true
		Write-Host "Zip this folder?" -Foregroundcolor DarkYellow
	}
	elseif (Test-Path -Path $f -PathType Leaf) {
	    $f_is_file=$true
		Write-Host "Zip this file?" -Foregroundcolor DarkYellow
	}
	else {
		Write-Host "Not a valid file/folder, or path contains invalid symbols. Exit." -Foregroundcolor Red
		pause; exit
	}
	if ($m) {
		Write-Host "And create an email with the password?" -Foregroundcolor Yellow
	}
	pause
}
else {
	Write-Host "No inputs. Exit." -Foregroundcolor Red
	pause; exit
}

if ($f_is_folder -or $f_is_file) {
	Add-Type -AssemblyName System.Web
	$pw_len=15
	$pw_sp_char=2
	$pw=[System.Web.Security.Membership]::GeneratePassword($pw_len, $pw_sp_char)

	$zip_format="zip"
	$zip_suffix=".zi_"

	if ($f_is_folder) {
		$zip_target=$f+$zip_suffix
		$zip_source="$f\*"
		$pw_file=$f+"_pw.txt"
	}
	elseif ($f_is_file) {
		$zip_target=(Get-Item $f).DirectoryName+"\"+(Get-Item $f).Basename+$zip_suffix
		$zip_source=$f
		$pw_file=(Get-Item $f).DirectoryName+"\"+(Get-Item $f).Basename+"_pw.txt"
	}

	# $work_dir=Split-Path $f -Parent
	# cd $work_dir
	Set-Alias 7z $7z_path
	7z a "-t$zip_format" "-p$pw" $zip_target $zip_source

 # Create a TXT file contains the password.
	Set-Content "$pw_file" $pw -NoNewline

 # Copy the password to clipboard.
	Set-Clipboard -Value $pw
	Write-Host -Foregroundcolor Green `
"
** $pw_len-digit password copied! **
*                             *
*       $pw       *
*                             *
*******************************"

	# Create an email message in Outlook.
	if ($m) {
		$ol=New-Object -comObject Outlook.Application

		$mail=$ol.CreateItem(0)
		$mail.Subject="[PASSWORD]"
		$zip_target_short=Split-Path "$zip_target" -Leaf
		$mail.HTMLBody= `
"[FILE]<br>
$zip_target_short<br>
<br>
[PASSWORD]<br>
$pw<br>"
		$mail.Attachments.Add($pw_file)
		$mail.save()
		
		$inspector=$mail.GetInspector
		$inspector.Display()
	}
}
pause; exit
