
$exeFilePath= "D:\Program Files\Thomson Reuters\ASAPJobSubmitter"
$outputFolder = $rootFolder + "\output"
$hostname = $env:COMPUTERNAME
$source = "PAL"
$logFile = "D:\Logs\RiskFactoryIngestion"


function GetLastFolder($root)
{
    $folderName = ""
	$dirs = @(get-childitem $outputFolder | where {$_.mode -match "d" -and $_.name -match "\d{8}" -and $_.name.length  -match 8} | select -property name | sort-object -property name -descending)
    if($dirs.length -ge 0)
	{
       $folderName = $dirs[0].name
    }
    return $folderName
	if ($error[0] -ne "")
	{
		write-eventlog -logname Application -source "PAL" -eventID 0 -entrytype "Error" -message $error[0]  -category 0 	
	}
}

function CheckOutputFolder($dir)
{
	return Test-Path $dir
}

#make sure server has log facility for "PAL" entity source
if ([System.Diagnostics.EventLog]::SourceExists($source) -eq $false) {
    [System.Diagnostics.EventLog]::CreateEventSource($source, "Application")
}
#make sure the log file directory exists and then set the file location
if (!(Test-Path -path $logFile)) {New-Item $logFile -file}
$logFile = $logFile + "\LogFile.txt"

Make sure you are using the correct NAS
$cenv=$env:computername.tostring().substring(0,4)  # Ex. "US1P"
switch ($cenv)
{
"US1I"{ $rootFolder = "\\10.248.58.4\us1i_asap_1\FILEWATCHER\RiskFactoryOutput\" }
"US1S"{ $rootFolder = "\\10.248.30.228\us1s_asap_1\FILEWATCHER\RiskFactoryOutput\" }
"US1P"{ $rootFolder = "\\10.248.162.68\us1p_asap_1\FILEWATCHER\RiskFactoryOutput\" }
"US2P"{ $rootFolder = "\\10.249.29.84\us2p_asap_1\FILEWATCHER\RiskFactoryOutput\" }
}


for($i=1; $i -le 5; $i++)
{
    $latestFolder = GetLastFolder($outputFolder)
    if($latestFolder.length -eq 8)
    {
        write-host $latestFolder
        $flag = CheckOutputFolder($rootFolder)
        write-host "first flag - folder level" +  $flag
        if ($flag)
        {
        	$theDate = get-date ($latestFolder -replace '^(\d{4})(\d{2})(\d{2})', '$1-$2-$3')
        	write-host $theDate
            #copy the latest folder from the output directory to the root folder
            Copy-Item $outputFolder"\"$latestFolder $rootFolder -recurse
            
            #rename the folder within /output to indicate that it has already been copied
            Rename-Item $outputFolder"\"$latestFolder DONE_$latestFolder
        	
        	Push-Location $exeFilePath
        	$command = "./ASAPJobSubmitter.exe RISKENGINESCHEDULER " + $latestFolder + " 142615224 142615208 Y"
        	write-host "invoking command " + $command 
        	invoke-expression $command
			Add-Content $logFile ($command)
			if ($error[0] -ne "")
			{
				write-eventlog -logname Application -source "PAL" -eventID 0 -entrytype "Error" -message $error[0]  -category 0 
			}
			$error.clear()
        	pop-location
        }
    }
    else
    {
        write-host "No more folders to process"
    }
}



