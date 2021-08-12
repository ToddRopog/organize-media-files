# ==============================================================================================
# 
# Microsoft PowerShell Source File 
# 
# This script will organize photo and video files by renaming the file based on the date the
# file was created and moving them into folders based on the year and month. The script will
# look in the SourceRootPath (recursing through all subdirectories) for any files matching
# the extensions in FileTypesToOrganize. It will rename the files and move them to folders under
# DestinationRootPath, e.g. DestinationRootPath\2011\02_February\2011-02-09 21.41.47.jpg
# 
# JPG files contain EXIF data which has a DateTaken value. 
# Other media files have a MediaCreated date. 
# ============================================================================================== 
Param(
	[Parameter(Mandatory = $true)]
	[string]$SourceRootPath,
	[Parameter(Mandatory = $true)]
	[string]$DestinationRootPath
)

# Should cancel if the params aren't valid.

# These value work for Windows 10:
[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Drawing.dll") 
$MediaCreatedColumn = 208

# For older versions of Windos you may need to use these values:
#[reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll") 
#$MediaCreatedColumn = 191

$FileTypesToOrganize = @("*.jpg", "*.avi", "*.mp4", "*.3gp", "*.mov")
$global:ConfirmAll = $false

function GetMediaCreatedDate($File) {
	$Shell = New-Object -ComObject Shell.Application
	$Folder = $Shell.Namespace($File.DirectoryName)
	$CreatedDate = $Folder.GetDetailsOf($Folder.Parsename($File.Name), $MediaCreatedColumn).Replace([char]8206, ' ').Replace([char]8207, ' ')

	if ($null -ne ($CreatedDate -as [DateTime])) {
		return [DateTime]::Parse($CreatedDate)
	}
 else {
		return $null
	}
}

function GetCreatedDateFromFilename($File) {
	if ($File.Name.Length -ge 20) {
		$Filename = $File.Name.Substring(0, 11).Replace("_", " ") + $File.Name.Substring(11, 8).Replace("-", ":").Replace(".", ":")
		Write-Host $Filename
		$DateValue = $Filename -as [DateTime]
		Write-Host $DateValue	
		if ($null -ne $DateValue) {
			return [DateTime]::ParseExact($Filename, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture) 
		}
	}
	else {
		if ($File.Name.Length -ge 15) {
			$Filename = $File.Name.Substring(0, 15)
			Write-Host $Filename
			$DateValue = [DateTime]::ParseExact($Filename, "yyyyMMdd HHmmss", [System.Globalization.CultureInfo]::InvariantCulture) 
		}

		if ($null -ne $DateValue) {
			return $DateValue
		}
		else {
			return $null
		}
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
	Try {
		$DateTimePropertyItem = $FileDetail.GetPropertyItem(36867)

		$Year = ConvertAsciiArrayToString $DateTimePropertyItem.value[0..3]
		$Month = ConvertAsciiArrayToString $DateTimePropertyItem.value[5..6]
		$Day = ConvertAsciiArrayToString $DateTimePropertyItem.value[8..9]
		$Hour = ConvertAsciiArrayToString $DateTimePropertyItem.value[11..12]
		$Minute = ConvertAsciiArrayToString $DateTimePropertyItem.value[14..15]
		$Second = ConvertAsciiArrayToString $DateTimePropertyItem.value[17..18]
		
		$DateString = [String]::Format("{0}/{1}/{2} {3}:{4}:{5}", $Year, $Month, $Day, $Hour, $Minute, $Second)
	}
	Catch {
		write-host $_.Exception.Message
	}
	Finally {
		$FileDetail.Dispose()
	}

	if ($null -ne ($DateString -as [DateTime])) {
		return [DateTime]::Parse($DateString)
	}
 else {
		return $null
	}
}

function GetCreationDate($File) {
	switch ($File.Extension) { 
		".jpg" { $CreationDate = GetDateTakenFromExifData($File) } 
		".3gp" { $CreationDate = GetCreatedDateFromFilename($File) }
		".mov" { $CreationDate = GetCreatedDateFromFileInfo($File) }
		default { $CreationDate = GetMediaCreatedDate($File) }
	}
	write-host $CreationDate
	if ($null -eq ($CreationDate -as [DateTime])) {
		$CreationDate = GetCreatedDateFromFilename($File)
	}
	return $CreationDate
}

function BuildDesinationPath($Path, $Date) {
	return [String]::Format("{0}\{1}\{2}_{3}", $Path, $Date.Year, $Date.ToString("MM"), $Date.ToString("MMMM"))
}

function BuildNewFilePath($Path, $Date, $Extension) {
	return [String]::Format("{0}\{1}{2}", $Path, $Date.ToString("yyyy-MM-dd HH.mm.ss"), $Extension)
}

function BuildNewFilePathWithDifferentiator($Path, $Date, $Differentiator, $Extension) {
	return [String]::Format("{0}\{1}_{2}{3}", $Path, $Date.ToString("yyyy-MM-dd HH.mm.ss"), $Differentiator, $Extension)
}

function CreateDirectory($Path) {
	if (!(Test-Path $Path)) {
		New-Item $Path -Type Directory
	}
}

function ConfirmContinueProcessing() {
	if ($global:ConfirmAll -eq $false) {
		$Response = Read-Host "Continue? (Y/N/A)"
		if ($Response.Substring(0, 1).ToUpper() -eq "A") {
			$global:ConfirmAll = $true
		}
		if ($Response.Substring(0, 1).ToUpper() -eq "N") { 
			break 
		}
	}
}

function GetAllSourceFiles() {
	return @(Get-ChildItem $SourceRootPath -Recurse -Include $FileTypesToOrganize)
}

function MakeUniqueFilePath($DestinationPath, $CreationDate, $FileExtension) {
	$ndx = 1
	do {
		$TempFilePath = BuildNewFilePathWithDifferentiator $DestinationPath $CreationDate $ndx $File.Extension
		$ndx++
	} while (Test-Path $TempFilePath)
	return $TempFilePath
}

Write-Host "Begin Organize"
$Files = GetAllSourceFiles
foreach ($File in $Files) {
	write-host $File.Name
	$CreationDate = GetCreationDate($File)
	if ($null -ne ($CreationDate -as [DateTime])) {
		$DestinationPath = BuildDesinationPath $DestinationRootPath $CreationDate
		CreateDirectory $DestinationPath
		$NewFilePath = BuildNewFilePath $DestinationPath $CreationDate $File.Extension
		
		if (Test-Path $NewFilePath) {
			$NewFilePath = MakeUniqueFilePath $DestinationPath $CreationDate $File.Extension
		}

		Write-Host $File.FullName -> $NewFilePath
		Move-Item $File.FullName $NewFilePath
	}
	else {
		Write-Host "Unable to determine creation date of file. " $File.FullName
		ConfirmContinueProcessing
	}
} 

Write-Host "Done"