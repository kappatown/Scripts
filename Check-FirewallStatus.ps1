# ==============================================================================================
# 
# Windows PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2016
# 
# NAME: 
# 
# AUTHOR: Windows User , 
# DATE  : 7/26/2016
# 
# COMMENT: 
# Version: 1.5.3
# ==============================================================================================

Function write-host($msg)
{
	$msg | Out-Host
	$msg | Out-File c:\tools\logs\FirewallCheck.log -Append -Encoding ASCII
}

Function Update-SQL($ServiceName, $ServiceState)
{
    $uri = "http://eusserviceapi.joann.com/api/services/status"

    if($psversiontable.psversion.major -gt 2)
	{
        $body = @{
            computerName = $script:ComputerName;
            storeNumber  = $Script:StoreNumber;
            running = $script:IsRunning;
            ServiceName = "$ServiceName";
            Version = $script:version;
            NumAttemps = $script:Attempts;
            InHours = $script:inBusinessHours;
            State = $ServiceState;
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

if(!(Test-Path c:\tools\logs))
{
	mkdir c:\tools\logs
}

if(Test-Path c:\tools\logs\FirewallCheck.log)
{
	"$(Get-Date)" | Out-File c:\tools\logs\FirewallCheck.log -Encoding ASCII
}

write-host "Starting Script"
$FilePath = "$($MyInvocation.InvocationName)"
$computername = $env:ComputerName
$Script:StoreNumber = $computername.substring(1,4)
$serviceName = "MpsSvc"
$script:attempts = 0
$RunTime = Get-Date
$bussinessHoursStart = Get-Date "$($runTime.ToShortDateString()) 8:00:00"
$bussinessHoursEnd = Get-Date "$($runTime.ToShortDateString()) 22:00:00"
$script:inBusinessHours = $RunTime -gt $bussinessHoursStart -and $RunTime -lt $bussinessHoursEnd
$override_Business_Hours = $false
$script:IsRunning = $true
$FileFullName = $(Get-ChildItem $FilePath).Fullname
$FileName = $(Get-ChildItem $FileFullName).Name
$updateSQL = $true

write-host "Script Name: $FileName"

$versionline = Get-Content $FileFullName | where{$_ -match "[#].*Version:.*\d{0,3}[\.]\d{0,3}[\.]\d{0,3}"}
$temp = $versionline -match "\d{1,3}[.]\d{1,3}[.]\d{1,3}"
$script:version = $matches[0]

$checkforrule = netsh advfirewall firewall show rule name="SQL EUS"
	 
if($checkforrule | where{$_ -match "Allow"})
{
	write-host	"Firewall rule exists"
}
else
{
	write-host "Firewall rule needs to be added"
	$ruleadd = netsh advfirewall firewall add rule dir=out name="SQL EUS" action=allow protocol=TCP remoteport=1433 remoteip=10.3.12.81
	
	Start-Sleep 2
	
	$checkforrule = netsh advfirewall firewall show rule name="SQL EUS"
	if($checkforrule | where {$_ -match "Allow"})
	{
		write-host "Rule Added but still not showing in Advanced Firewall"
	}
	else
	{
		write-host "Rule Added Sucessfully"
	}
}

write-host "Checking SQL Entry"
$sqlinfo = Update-SQL $serviceName

if($sqlinfo)
{
	if($sqlinfo.overrideBusinessHours -eq $true)
	{
		write-host "Business Hours Override is currently Active"
		$override_Business_Hours = $true
	}

	write-host "Checking Admin mode"
	$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
	$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
	$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
	 
	if (!($myWindowsPrincipal.IsInRole($adminRole)))
	{
        #prevent sql from updating again on return
        $updateSQL = $false

		if($MyInvocation.BoundParameters.keys)
		{
			foreach($item in $MyInvocation.BoundParameters.keys)
			{
				$arguments += "-$($item) "
			}
		}
		
		if($MyInvocation.UnboundArguments)
		{
			foreach($item in $MyInvocation.UnboundArguments)
			{
				$arguments += "$($item) "
			}
		}
		
		$startstring = "$($myInvocation.MyCommand.Definition) $arguments"
		$newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
		$newProcess.Arguments = $myInvocation.MyCommand.Definition
		$newProcess.Verb = "runas"
		
		[System.Diagnostics.Process]::Start($newProcess) | Out-Null	
	}
	else
	{
		$file = $MyInvocation.InvocationName
		$temp = $file.split("\")
		$file = $temp[$temp.count-1]
		
		Write-Host "Checking is Script is current Version of Software"
		
		#check IIS version
		
		if(Check-RunningLatestSoftware $FileFullName)
		{
			write-host "Getting Firewall status"
			$firewall = Get-Service MpsSvc
			$starttype = Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='MpsSvc'"
				
			write-host "Vars retrived"
			if($starttype.StartMode -ne "Disabled")
			{	
				write-host "Firewall not disabled"
				if($firewall.status -eq "Running")
				{
					$status = "Running"
					write-host "Firewall already running"
				}
				
				while($firewall.status -ne "Running" -and $attempts -lt 3)
				{
					write-host "Starting Firewall"
					Start-Service -Name MpsSvc
					Start-Sleep 10
					$firewall = Get-Service MpsSvc
					if($firewall.status -ne "Running")
					{
						write-host "Firewall Still not running"
						Start-Sleep 30
					}
					$attempts++
				}
				
				$firewall = Get-Service MpsSvc
				
				if($attempts -ge 3 -and $firewall.status -ne "Running")
				{
					if($inbusinessHours -or !($override_Business_Hours))
					{
						$status = "InBusinessHours - Needs Disabled"
					}
					else
					{
						#only reboot outside of business hours
						$status = "Disabling, Pending Reboot"
						write-host "Disabling Firewall"
						Set-Service -name MpsSvc -StartupType disabled
						shutdown -r -t 0
						$script:IsRunning = $false
						Update-SQL $serviceName $status
					}
				}
			}
			else
			{
				$status = "Disabled"
				write-host "Firewall already disabled"
			}
			
			$script:IsRunning = $false
		}
		else
		{
			$script:IsRunning = $false
			$status = "Script Could not be Verifed"
			Write-Host "Script Execution is not Authorized since it could not be Verifed from Repo"
		}
	}
}
else
{
	$status = "Failed to Connect to SQL"
}

if($updateSQL)
{
	Update-SQL $serviceName $status
}

exit 0
