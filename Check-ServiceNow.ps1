# ==============================================================================================
# 
# Windows PowerShell Source File -- Created with SAPIEN Technologies PrimalScript 2016
# 
# NAME: 
# 
# AUTHOR:  , 
# DATE  : 8/30/2016
# Version: 1.0.4
# Rocking this out
# ==============================================================================================

param
(
[int] $SqlID, 
[string] $environment
)

Function get-SqlInfo($SqlID)
{
	$sqlserver = "h12dmtsr201.joann.com"
	$username = "ESU"
	$password = "HelloWorld123"
	$database = "EUSTesting"
	$tablename = "ServiceState"
		
	Write-Debug "Updating SQL"
	
	$conn = New-Object System.Data.SqlClient.SqlConnection
	$conn.ConnectionString = "Data Source=$sqlserver;ConnectRetryCount=10;Initial Catalog=$database;User Id =$username; Password=$password"
	$conn.Open()
	if($conn.State -eq "Open")
	{
		$sqlCommand = New-Object System.Data.SqlClient.SqlCommand
		$sqlCommand.Connection = $conn
		$sqlCommand.CommandTimeout = 120
		$sqlCommand.CommandText = "execute [sp_PrinterTonerNotification_GetInfo] '$sqlid'"
	    
		Write-Debug "$($sqlCommand.CommandText)"
		
		$results = $sqlCommand.ExecuteReader()
		if($results.fieldcount -gt 0)
		{
			$pstable = $null
			$table = $null
			$table = New-Object "System.Data.DataTable"
			$table.Load($results)
			$pstable = @($table  | select *)
		}
		$conn.close()
	}
	
	return $pstable	
}

Function UpdateSQL-ServiceNowInfo($SqlID,$SsysID,$Incident,$hasreopened)
{
	$sqlserver = "h12dmtsr201.joann.com"
	$username = "ESU"
	$password = "HelloWorld123"
	$database = "EUSTesting"
	$tablename = "ServiceState"
		
	Write-Debug "Updating SQL"
	
	$conn = New-Object System.Data.SqlClient.SqlConnection
	$conn.ConnectionString = "Data Source=$sqlserver;ConnectRetryCount=10;Initial Catalog=$database;User Id =$username; Password=$password"
	$conn.Open()
	if($conn.State -eq "Open")
	{
		$sqlCommand = New-Object System.Data.SqlClient.SqlCommand
		$sqlCommand.Connection = $conn
		$sqlCommand.CommandTimeout = 120
		$sqlCommand.CommandText = "execute [sp_PrinterTonerNotification_WriteServiceNowInfo] '$sqlid', '$SsysID', '$Incident', '$hasreopened'"
	    
		Write-Debug "$($sqlCommand.CommandText)"
		
		$results = $sqlCommand.ExecuteReader()
		if($results.fieldcount -gt 0)
		{
			$pstable = $null
			$table = $null
			$table = New-Object "System.Data.DataTable"
			$table.Load($results)
			$pstable = @($table  | select *)
		}
		$conn.close()
	}
	
	return $pstable	
}

Write-Host "Start"
$return = $null

if($environment -eq "Prod")
{
	Write-Host "Setting Production Settings"
	$baseuri = "joannstores.service-now.com"
	$username = "ukapadia"
	$password = Get-Content c:\tools\ServiceNowDev | ConvertTo-SecureString
}
elseif($environment -eq "Dev")
{
	Write-Host "Setting Development Settings"
	$baseuri = "joannstoresdev.service-now.com"
	$username = "ukapadia"
	$password = Get-Content c:\tools\ServiceNowDev | ConvertTo-SecureString
}
elseif($environment -eq "SandBox")
{
	Write-Host "Setting SandBox Settings"
	$baseuri = "joannstoressb.service-now.com"
	$username = "ukapadia"
	$password = Get-Content c:\tools\ServiceNowDev | ConvertTo-SecureString
}

if($baseuri)
{
	Write-Host "Building Creds"
	$cred = New-Object System.Management.Automation.PSCredential ($username, $password)
	
	#check sql for if incident exists
	$SQLinfo = get-SqlInfo $SqlID
	if($SQLinfo)
	{
		$storenumber = $SQLinfo.storenumber
		$incidentsysid = $SQLinfo.ServiceNowID
		$incidentnumber = $SQLinfo.ServiceNowIncident
		$SQLReopen = $SQLinfo.ReopenResponse
		
		$body = @{'name'="$($storenumber.tostring('0000'))"}
		$result = Invoke-RestMethod -Method GET -Uri https://$baseuri/api/now/table/sys_user?sysparm_limit=10 -Credential $cred -Body $body
		$callerid = $result.result.sys_id
			
		if($callerid)
		{
			if($incidentsysid -ne [System.DBNull]::Value)
			{
				Write-Host "Using Existing ID"
				$body = @{'sys_id'="$incidentsysid"}
			}
			elseif($incidentnumber -ne [System.DBNull]::Value)
			{
				Write-Host "Using Existing Number"
				$body = @{'number'="$incidentnumber"}
			}
			else
			{
				Write-Host "Trying to find based on description"
				$body = @{
					'short_description'="Low Toner"
					'caller_id'="$callerid"
				}
			}
			
			$result = Invoke-RestMethod -Method Get -Uri https://$baseuri/api/now/table/incident -Credential $cred -ContentType "application/json" -Body $body
			$existing_incident = $result.result
			
			if($existing_incident.sys_id)
			{
				Write-Host "existing Record Found"
				foreach	($record in $existing_incident)
				{
					if(!($latested) -or [datetime] $latested.opened_at -gt [datetime] $existing_incident.opened_at)
					{
						Write-Host "found newest record"
						$latested = $existing_incident
					}
				}
				
				$ticketid = $latested.sys_id
				
				if(($(get-date) - [datetime] $latested.opened_at).days -lt 90)
				{
					Write-Host "INC has been open for less then 90 days."
					$ticketid = $latested.sys_id
					$ticketnumber = $latested.number
					if(!($SQLReopen) -and ($latested.incident_state -eq 20 -or $latested.incident_state -eq 7))
					{
						Write-Host "Send API Command to reopen INC"
						$body = @{'state'="1";'incident_state'="2";'comments'="Store has responded to automated process that they are still missing toner."} | ConvertTo-Json
						$result = Invoke-RestMethod -Method PUT -Uri https://$baseuri/api/now/table/incident/$ticketid -Credential $cred -ContentType "application/json" -Body $body
						
						$latested = $result.result
					}
					
					if($latested.incident_state -eq 2)
					{
						Write-Host "INC is open"
						$hasreopened = 1
					}
				}
			}
			
			if(!($ticketid))
			{
			 	Write-Host "Creating New Incident"
				$body = @{
				        'caller_id' = "$callerid"
				        'short_description' = "Low Toner"
				        'assignment_group' = "cd2f3b580a0a3d28002b542565fc5143"
				        'category' = "Service Request"
				        'subcategory' = "Toner Request"
				        'cmdb_ci' = "0c53b6364f2f42008fb9dd2f0310c7c2"
				    } | ConvertTo-Json
				    			
				$result = Invoke-RestMethod -Method Post -Uri https://$baseuri/api/now/table/incident -Credential $cred -ContentType "application/json" -Body $body
				
				$ticketid = $result.result.sys_id
				$ticketnumber = $result.result.number
				
				$comment = "On $($SQLinfo.DateEmailed.tostring())(Eastern Time) your store received an email for low toner notification. 
				A user had clicked `"yes`" in that email which auto-generated this ticket to fulfill a toner request for your $($Sqlinfo.Model) printer. 
				Your toner request has been processed, please allow 5-10 businesses days for arrival."
				
				$body = @{'state'="20";'incident_state'="20";'comments'="$comment"} | ConvertTo-Json
				$result = Invoke-RestMethod -Method PUT -Uri https://$baseuri/api/now/table/incident/$ticketid -Credential $cred -ContentType "application/json" -Body $body
				
			}
			
			$null = UpdateSQL-ServiceNowInfo $SqlID $ticketid $ticketnumber $hasreopened
		}
		else
		{
			Write-Host "Store $storenumber not found in Service Now"
		}
	}
	else
	{
		Write-Host "Couldn't find Record"
	}
}
else
{
	Write-Host "no Settings have been configured"
}


return $return
