# ==============================================================================================
# 
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. It will also append
# a random number to the end of the file name just to avoid name collisions. The script will
# look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under
# DestinationRootPath, e.g. DestinationRootPath\2011\02_February\2011-02-09_21-41-47_680.jpg
# 
# JPG files contain EXIF data which has a DateTaken value. Other media files have a MediaCreated
# date. 
#
# ============================================================================================== 


# These value work for Windows 10:
[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll") 
$MediaCreatedColumn = 208

# For older versions of Windos you may need to use these values:
#[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 
#$MediaCreatedColumn = 191

$SourceRootPath = ".\RecentUploads"
$DestinationRootPath = ".\Years"
$FileTypesToOrganize = @("*.jpg","*.avi","*.mp4", "*.3gp", "*.mov")
$global:ConfirmAll = $false

function GetMediaCreatedDate($File) {
	$Shell = New-Object -ComObject Shell.Application
	$Folder = $Shell.Namespace($File.DirectoryName)
	$CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), $MediaCreatedColumn).Replace([char]8206, ' ').Replace([char]8207, ' ')

	if (($CreatedDate -as [DateTime]) -ne $null) {
		return [DateTime]::Parse($CreatedDate)
	} else {
		return $null
	}
}

function GetCreatedDateFromFilename($File) {
	$Filename = $File.Name.Substring(0, 11).Replace("_", " ") + $File.Name.Substring(11, 8).Replace("-", ":")
	Write-Host $Filename
	Write-Host ($Filename -as [DateTime])
	if (($Filename -as [DateTime]) -ne $null) {
		return [DateTime]::ParseExact($Filename,"yyyy-MM-dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) 
	} else {
		return $null
	}
}

function GetCreatedDateFromFileInfo($File) {
	return $File.CreationTime
}

function ConvertAsciiArrayToString($CharArray) {
	$ReturnVal = ""
	foreach ($Char in $CharArray) {
		$ReturnVal += [char]$Char
	}
	return $ReturnVal
}

function GetDateTakenFromExifData($File) {
	$FileDetail = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $File.Fullname 
	$DateTimePropertyItem = $FileDetail.GetPropertyItem(36867)
	$FileDetail.Dispose()

	$Year = ConvertAsciiArrayToString $DateTimePropertyItem.value[0..3]
	$Month = ConvertAsciiArrayToString $DateTimePropertyItem.value[5..6]
	$Day = ConvertAsciiArrayToString $DateTimePropertyItem.value[8..9]
	$Hour = ConvertAsciiArrayToString $DateTimePropertyItem.value[11..12]
	$Minute = ConvertAsciiArrayToString $DateTimePropertyItem.value[14..15]
	$Second = ConvertAsciiArrayToString $DateTimePropertyItem.value[17..18]
	
	$DateString = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
	
	if (($DateString -as [DateTime]) -ne $null) {
		return [DateTime]::Parse($DateString)
	} else {
		return $null
	}
}

function GetCreationDate($File) {
	switch ($File.Extension) { 
        ".jpg" { $CreationDate = GetDateTakenFromExifData($File) } 
		".3gp" { $CreationDate =  GetCreatedDateFromFilename($File) }
		".mov" { $CreationDate =  GetCreatedDateFromFileInfo($File) }
        default { $CreationDate = GetMediaCreatedDate($File) }
    }
	return $CreationDate
}

function BuildDesinationPath($Path, $Date) {
	return [String]::Format("{0}\{1}\{2}_{3}", $Path, $Date.Year, $Date.ToString("MM"), $Date.ToString("MMMM"))
}

$RandomGenerator = New-Object System.Random
function BuildNewFilePath($Path, $Date, $Extension) {
	return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyy-MM-dd_HH-mm-ss"), $RandomGenerator.Next(100, 1000).ToString(), $Extension)
}

function CreateDirectory($Path){
	if (!(Test-Path $Path)) {
		New-Item $Path -Type Directory
	}
}

function ConfirmContinueProcessing() {
	if ($global:ConfirmAll -eq $false) {
		$Response = Read-Host "Continue? (Y/N/A)"
		if ($Response.Substring(0,1).ToUpper() -eq "A") {
			$global:ConfirmAll = $true
		}
		if ($Response.Substring(0,1).ToUpper() -eq "N") { 
			break 
		}
	}
}

function GetAllSourceFiles() {
	return @(Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize)
}

Write-Host "Begin Organize"
$Files = GetAllSourceFiles
foreach ($File in $Files) {
	$CreationDate = GetCreationDate($File)
	if (($CreationDate -as [DateTime]) -ne $null) {
		$DestinationPath = BuildDesinationPath $DestinationRootPath $CreationDate
		CreateDirectory $DestinationPath
		$NewFilePath = BuildNewFilePath $DestinationPath $CreationDate $File.Extension
		
		Write-Host $File.FullName -> $NewFilePath
		if (!(Test-Path $NewFilePath)) {
			Move-Item $File.FullName $NewFilePath
		} else {
			Write-Host "Unable to rename file. File already exists. "
			ConfirmContinueProcessing
		}
	} else {
		Write-Host "Unable to determine creation date of file. " $File.FullName
		ConfirmContinueProcessing
	}
} 

Write-Host "Done"