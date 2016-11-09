# ==============================================================================================
# 
# Windows PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2015
# 
# NAME: 
# 
# AUTHOR: Umang Kapadia 
# DATE  : 06/24/2016
# 
# COMMENT: 
# Version: 2.0.0
# ==============================================================================================
param
(
	[switch] $OverrideBusinessHours
)

Function Update-SQL()
{
    $uri = "http://eusserviceapi.joann.com/api/scriptstatus/virtualcapture"

    if($psversiontable.psversion.major -gt 2)
	{
        $body = @{
            ScriptVersion = $Script:version
            Storenumber = $script:Store
            dfree = $script:FreeSpace
            raidhealthy = $script:RAIDHealthy
            imagecaptured = $script:ImageCaptured
            timeoutreached = $script:TimeoutReached
            hourslastcaptured = $script:Hours_Since_Last_Capture
            timestampoldimage = $script:OldImageTimeStamp
            captureruntime = $script:RunTime
            cfree = $script:CFree
            ctotal = $script:CTotal
            dtotal = $script:dTotal
            updatestartdate = $script:StartofScript       
        } | ConvertTo-Json
    
        $pstable = Invoke-RestMethod -URI $uri -Method PUT -Body $body -ContentType "application/json"
    }
    else
    {
        $client = New-Object System.Net.WebClient
        $client.headers.Add("Content-Type","application/json")
        $client.headers.Add("Accept","application/xml")
        $data = $client.UploadString("http://eusserviceapi.joann.com/api/services/status","PUT",$jbody)
        if($data)
        {
            $pstable = $([xml] $data).sp_ServiceState_UpdateEntry_Result | select ComputerName, OverrideBusinessHours
        }
    }
	
return $pstable	
}
Function Get-MD5HashWithProgress($file)
{
	$hash_txt = $null
	$total = 0
	$md5 = [System.Security.Cryptography.MD5]::Create("MD5")
	$fd = [System.IO.File]::OpenRead($file)
	$buf = New-Object byte[] (1024*1024*8) # 8mb buffer
	while (($read_len = $fd.Read($buf,0,$buf.length)) -eq $buf.length)
	{
	    $total += $buf.length
	    $null = $md5.TransformBlock($buf,$offset,$buf.length,$buf,$offset)
	    Write-Progress -Activity "Hashing File" -Status $file -percentComplete ($total/$fd.length * 100)
	}
	
	# finalize the last read
	$null = $md5.TransformFinalBlock($buf,0,$read_len)
	$hash = $md5.Hash
	
	# convert hash bytes to hex formatted string
	$hash | foreach { $hash_txt += $_.ToString("x2") }
	$fd.close()
	
	Write-Progress -Activity "Hashing File" -Status $file -Completed
	
	Return $hash_txt
}
Function Check-RunningLatestSoftware
{
	param
	(
		[Parameter(Mandatory=$true)]
		[string] $FileFullName
	)
    $uri = "http://eusserviceapi.joann.com/api/version/download"
	$canrun = $false
	
	if(Test-Path $FileFullName)
	{
		$CurrentHash = Get-MD5HashWithProgress $FileFullName
		write-host "Local Hash: $CurrentHash"
		
		$filename = $(Get-ChildItem $FileFullName).Name
		
		if($psversiontable.psversion.major -gt 2)
		{
			$body = @{Filename=$filename;Hash=$CurrentHash}
			$result = Invoke-RestMethod -Uri $uri -Method GET -Body $body
		}
		else
		{		
			#This code exists because of non PS3.0 systems
			$client = New-Object System.Net.WebClient
			$client.headers.Add("Accept","application/xml")
			$downloaduri = "$($uri)?Filename=$filename&Hash=$CurrentHash"
			$data = $client.DownloadString($downloaduri)
				
			if($data)
			{	
				$result = $([xml] $data).DownloadRequest
			}
		}
		
		if($result)
		{
			if($result.mustupdate)
			{
				write-host "Repo Hash: $($result.Repohash)"
				write-host "Must Update Software"
				write-host "Updating from $($result.Location)"
				$WebClient = New-Object System.Net.WebClient
				$WebClient.DownloadFile("$($result.Location)", $FileFullName)
				$CurrentHash = Get-MD5HashWithProgress $FileFullName
				if($CurrentHash -eq $result.Repohash)
				{
					write-host "Software Updated Successfully"
					
					if($script:MyInvocation.BoundParameters.keys)
					{
						foreach($item in $script:MyInvocation.BoundParameters.keys)
						{
							$arguments += "-$($item) "
						}
					}
					
					if($script:MyInvocation.UnboundArguments)
					{
						foreach($item in $script:MyInvocation.UnboundArguments)
						{
							$arguments += "$($item) "
						}
					}
					
					Write-Host "Script Arguments: $arguments"
					Write-Host "Starting Script again"
					Write-Host "Script $($script:MyInvocation.InvocationName)"
					powershell $script:MyInvocation.InvocationName $arguments
				}
				else
				{
					write-host "Failed to update software"
				}
			}
			elseif($result.approved)
			{
				$canrun = $true				
			}
			else
			{
				write-host "$($result.status)"
			}
				
		}
	}
	else
	{
		write-host "Failed to Find local file"
	}
	
	return $canrun
}


$running = $true
$script:Store = 0
$script:Date = Get-Date	
$script:FreeSpace = 0
$script:CFree = 0
$script:CTotal = 0
$script:dTotal = 0
$Script:version = 0
$script:ScriptFound = 0
$Script:CurrentTime = Get-Date
$script:RAIDHealthy = 0
$script:RunTime = 0
$script:TimeoutReached = 0
$script:ImageCaptured = 0
$script:OldImageTimeStamp = $null
$script:StartofScript = 1
$script:Hours_Since_Last_Capture = 0
$computername = $env:ComputerName
$Script:Store = $computername.substring(1,4)
$dpsystem = "S$($Store)dist01"
$Required_Hours_Between_Capture = 12
$CurrentTotalHours = $date.TimeOfDay.Totalhours
$StartofBusinessday = 8
$EndofBusinessDay = 23
$NeededFreespace = 20
$MaxRunTimeMins = 180

$FileFullName = $MyInvocation.InvocationName

$null = Update-SQL
$script:StartofScript = 0

Stop-Process -Name disk2vhd -Force -ErrorAction "SilentlyContinue"
Stop-Process -Name disk2vhd-tmp -Force -ErrorAction "SilentlyContinue"

Write-Host "Starting Script for Primary Image Capture"
if(Check-RunningLatestSoftware $FileFullName)
{
    $versionline = Get-Content $FileFullName | where{$_ -match "[#].*Version:.*\d{0,3}[\.]\d{0,3}[\.]\d{0,3}"}
    $temp = $versionline -match "\d{1,3}[.]\d{1,3}[.]\d{1,3}"
    $version = $matches[0]	
    Write-Host "Script Version: $version"

	$scriptblock = {
		
        $C_disk = get-wmiobject win32_logicaldisk | where {$_.deviceid -eq "C:"}
        $D_disk = get-wmiobject win32_logicaldisk | where {$_.deviceid -eq "D:"}
		$D_freeGB = $D_disk.freespace/1024/1024/1024
        $D_TotalGB = $D_disk.Size/1024/1024/1024
        $C_freeGB = $C_disk.freespace/1024/1024/1024
        $C_TotalGB = $C_disk.Size/1024/1024/1024
		
        $row = @{CDriveTotalGB = $C_TotalGB; CDriveFreeGB = $C_freeGB; DDriveTotalGB = $D_TotalGB; DDriveFreeGB = $D_freeGB}
        $pstable = New-Object -TypeName PSObject -Property $row

		return $pstable
	}
		
	$pstable = Invoke-Command -ScriptBlock $scriptblock -ComputerName $dpsystem

    $FreeSpace = $pstable.DDriveFreeGB
    $CFree = $pstable.CDriveFreeGB
    $CTotal = $pstable.CDriveTotalGB
    $dTotal = $pstable.DDriveTotalGB
	
	if(!(Test-Path \\$dpsystem\d$\VHDS\Image.vhd))
	{
		$NeededFreespace += 100
	}
	
	if($FreeSpace -gt $NeededFreespace)
	{
		if(!(Test-Path \\$dpsystem\d$\VHDS))
		{
			mkdir \\$dpsystem\d$\VHDS
		}
		
		Write-Host "Checking Date when the last image was captured"
		if(Test-Path \\$dpsystem\d$\VHDS\Image.vhd)
		{
			$OldImageTimeStamp = [datetime] $(Get-ItemProperty \\$dpsystem\d$\VHDS\Image.VHD).LastWriteTime
			$Hours_Since_Last_Capture = [math]::Round($($date - $OldImageTimeStamp).TotalHours)
		}
		else
		{
			Write-Host "Image doesn't exist"
			$Hours_Since_Last_Capture = 9999
		}
		
		if(($CurrentTotalHours -gt $StartofBusinessday -and $CurrentTotalHours -lt $EndofBusinessDay) -and !($OverrideBusinessHours))
		{
			Write-Host "Script is not authorized to run during business hours"
		}
		else
		{
            if($OverrideBusinessHours)
            {
                 Write-Host "Ignoring Business Hours"
			}
            else
            {
                Write-Host "Current time is outside of business hours"
            }

			if(($Hours_Since_Last_Capture -gt $Required_Hours_Between_Capture) -or $OverrideBusinessHours)
			{
	
				if(Test-Path \\$dpsystem\d$\VHDS\Image.vhd.old)
				{
					#Removing Last BackUp
					Write-Host "Removing last backup image"
					Remove-Item \\$dpsystem\d$\VHDS\Image.vhd.old
				}
				else
				{
					$NeededFreespace += 100
				}
				
				if((Test-Path \\$dpsystem\d$\VHDS\Image.vhd) -and $FreeSpace -gt $NeededFreespace)
				{
					#Renaming Last Image to Old Image file
					Write-Host "Renaming Image file inorder to maintain a backup Image"
					Rename-Item \\$dpsystem\d$\VHDS\Image.vhd \\$dpsystem\d$\VHDS\Image.vhd.old
				}
				else
				{
					Write-Host "System doesn't have enough free space to keep 2 images or Image Files Doesn't exist"
				}
				
				#Set a timeout on a process
				Write-Host "Starting Capture"
				c:\tools\disk2vhd\disk2vhd.exe /accepteula * \\$dpsystem\d$\VHDS\Image.vhd
				
				Start-Sleep 5
				
				while($running -ne $null)
				{
					$currentTime = Get-Date
					$running = Get-Process disk2vhd
					if($Running)
					{
						$runtime += 5
						if($runtime -gt $MaxRunTimeMins)
						{
							Write-Host "Run time has surpassed max allowed run time"
							Write-Host "Stopping all processes"
							$script:TimeoutReached = 1
							Stop-Process -Name disk2vhd -Force
							Stop-Process -Name disk2vhd-tmp -Force
							Remove-Item \\$dpsystem\d$\VHDS\Image.vhd						
						}
						else
						{
							Start-Sleep 300
						}
					}
				}
			}
			else
			{
				Write-Host "Attempting to Capture Image within $($Required_Hours_Between_Capture) hours is prohibited"
			}
		} #End of else for Busines Time Check
	} #End of If check for FreeSpace
	
	if(Test-Path \\$dpsystem\d$\VHDS\Image.vhd)
	{
		$ImageTimeStamp = [datetime] $(Get-ItemProperty \\$dpsystem\d$\VHDS\Image.VHD).LastWriteTime
		if($($(get-date) - $ImageTimeStamp).TotalMinutes -lt 20)
		{ 
			$ImageCaptured = $true
		}
		else
		{
			$ImageCaptured = $false
		}
	}
	else
	{
		$ImageCaptured = $false
	}
	
	$null = Update-SQL
}
else
{
	Write-Host "Script Execution is not Authorized since it could not be Verifed from Repo"
	$null = Update-SQL
}


