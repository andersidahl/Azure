workflow AutoShutdownSchedule
{
	
    # Settings
	$TimeZone = "W. Europe Standard Time"
    $TAGName = "AutoShutdownSchedule"

    # Get Current Time
	$currentTime = (Get-Date).ToUniversalTime()
	$CurrentTime = Get-CurrentTime -TimeZone $TimeZone
	
	"Runbook started"
	"Current UTC/GMT time [$($currentTime.ToString("dddd, yyyy MMM dd HH:mm:ss"))]"
	"Current $TimeZone time [$($currentTime.ToString("dddd, yyyy MMM dd HH:mm:ss"))]"
	
    # Connecting to Azure
    $connectionName = "AzureRunAsConnection"
    try
    {
        # Get the connection "AzureRunAsConnection "
        $servicePrincipalConnection=Get-AutomationConnection -Name $connectionName         

        "Logging in to Azure..."
        Add-AzureRmAccount `
            -ServicePrincipal `
            -TenantId $servicePrincipalConnection.TenantId `
            -ApplicationId $servicePrincipalConnection.ApplicationId `
            -CertificateThumbprint $servicePrincipalConnection.CertificateThumbprint 
    }
    catch {
        if (!$servicePrincipalConnection)
        {
            $ErrorMessage = "Connection $connectionName not found."
            throw $ErrorMessage
        } else{
            Write-Error -Message $_.Exception
            throw $_.Exception
        }
    }

	# Running Logic
	"Getting ARM based virtual machines with AutoShutdownSchedule tag"
	$VirtualMachines = Get-AzureRmVM | Where {$_.Tags.AutoShutdownSchedule}
	if(!$VirtualMachines){
		Throw "Could not fing any ARM based VMs with AutoShutdownSchedule tag"
	}
	
	$startVMs = @()
	$stopVMs = @()
	
	ForEach ($VM in $VirtualMachines){
        $Tag = ($VM.TagsText | ConvertFrom-Json)
        
        "[$($VM.Name)] $TAGName = $($tag.$TAGName)"
        
        # Get the shutdown time ranges definition tag and extract the value
		$shutdownTag = $tag.$TAGName
        $shutdownTimeRangesDefinition = $shutdownTag
        
        # Parse the ranges in the Tag value. Expects a string of comma-separated time ranges, or a single time range
		$timeRangeList = $shutdownTimeRangesDefinition -split "," | foreach {$_.Trim()}
        
		$checkScheduleEntry = 'not matched'
		$matchedEntry = $null
		foreach($entry in $timeRangeList){
			if($checkScheduleEntry -eq 'not matched'){
	            $checkScheduleEntry = CheckScheduleEntry -TimeRange $entry -currentTime $currentTime
				"[$($VM.Name)] $entry $checkScheduleEntry against $currentTime"
				$matchedEntry = $entry
			}
		}
        
        # Record desired state for virtual machine based on result. If schedule is matched, shut down the VM if it is running. Otherwise start the VM if stopped.
		if($checkScheduleEntry -eq 'matched')
		{
		    "[$($VM.Name)] Current time falls within the range [$matchedEntry]"
            $targetState = "PowerState/Deallocated"
		}
		else
		{
		    "[$($VM.Name)] Current time '$(Get-Date($currenttime) -format "dddd HH:mm")' is outside of all shutdown schedule ranges"
            $targetState = "PowerState/running"
		}
        
        $VMPowerState = (Get-AzureRmVm -Status $VM.ResourceGroupName $VM.Name).Statuses.Code[1]
        
        if ($VMPowerState -eq $targetState){
            "[$($VM.Name)] PowerState '$VMPowerState' equals target state - Do Nothing"
        }
        
        else{
 			if($targetState -eq "PowerState/running"){
               $startVMs += $VM
               "[$($VM.Name)] PowerState '$VMPowerState' does not match targetState '$targetState' - Start virtual machine"
            }
            if ($targetState -eq "PowerState/Deallocated"){
                $stopVMs += $VM
 				"[$($VM.Name)] PowerState '$VMPowerState' does not match targetState '$targetState' - Stop virtual machine"
            }
        }
        
    }
    
    If ($startVMs){
        ForEach -Parallel ($VM in $startVMs){
            "[ACTION] Starting virtual machines [$($VM.Name)] in resource group [$($VM.ResourceGroupName)]"
    		$VM | Start-AzureRmVm -verbose
            ""
	   }
    }
    
    else{
    	"[INFO] No virtual machine to start"
    }
    
    If ($stopVMs){
	   ForEach -Parallel ($VM in $stopVMs){
    	   "[ACTION] Stopping virtual machines [$($VM.Name)] in resource group [$($VM.ResourceGroupName)]"
    	   $VM | Stop-AzureRmVm -Force -verbose
           ""
	   }
    }

    else{
	    "[INFO] No virtual machine to stop"
    }
    
    ""
    "Runbook completed"


    # Functions

    # Function - Get Current Time
	Function Get-CurrentTime{
    	Param(
        	[String]$TimeZone = $TimeZone
    	)

    	$TimeZones = $null
    	$timeZones = [System.TimeZoneInfo]::GetSystemTimeZones() | Where-Object StandardName -Match $TimeZone
    	$timeZones = $timeZones | Add-Member -MemberType ScriptProperty -Name "CurrentTime" -Value { [TimeZoneInfo]::ConvertTime([DateTIme]::Now, $this) } -PassThru  -Force
    	$timeZones = $timeZones | Add-Member -MemberType ScriptProperty -Name "IsDayLightSavingTime" -Value { $this.IsDaylightSavingTime([DateTime]::Now) } -PassThru  -Force
    	$timeZones | Add-Member -MemberType ScriptProperty -Name "HoursApart" -Value { $this.CurrentTime - [DateTime]::Now } -PassThru  -Force | out-null
		    
    	Return $timeZones.CurrentTime
	}
	
	# Fnction - Check Schedule
	Function CheckScheduleEntry{
	    Param(
	        $TimeRange,
            $currentTime
	       )
	
	    try
		{
		    # Parse as range if contains '->'
		    if($TimeRange -like "*->*")
		    {
		        $timeRangeComponents = $TimeRange -split "->" | foreach {$_.Trim()}
		        if($timeRangeComponents.Count -eq 2)
		        {
		            $rangeStart = Get-Date $timeRangeComponents[0]
		            $rangeEnd = Get-Date $timeRangeComponents[1]
					$midnight = $currentTime.AddDays(1).Date

		            # Check for crossing midnight
		            if($rangeStart -gt $rangeEnd)
		            {
	                    # If current time is between the start of range and midnight tonight, interpret start time as earlier today and end time as tomorrow
	                    if($currentTime -ge $rangeStart -and $currentTime -lt $midnight)
	                    {
	                        $rangeEnd = $rangeEnd.AddDays(1)
	                    }
	                    # Otherwise interpret start time as yesterday and end time as today   
	                    else
	                    {
	                        $rangeStart = $rangeStart.AddDays(-1)
	                    }
		            }
		        }
		        else
		        {
		            Write-Error "`tWARNING: Invalid time range format. Expects valid .Net DateTime-formatted start time and end time separated by '->'" 
		        }
		    }
		    # Otherwise attempt to parse as a full day entry, e.g. 'Monday' or 'December 25' 
		    else{
		        # If specified as day of week, check if today
		        if([System.DayOfWeek].GetEnumValues() -contains $TimeRange){
					$Today = (Get-Date).DayOfWeek
					
					if($TimeRange -eq $Today){
		                $parsedDay = Get-Date "00:00"
		            }
                    
		            else{
		                # Skip detected day of week that isn't today
		            }
                    
		        }
		        # Otherwise attempt to parse as a date, e.g. 'December 25'
		        else{
   		            $parsedDay = Get-Date $TimeRange
		        }
		    
		        if($parsedDay -ne $null){
		            $rangeStart = $parsedDay # Defaults to midnight
		            $rangeEnd = $parsedDay.AddHours(23).AddMinutes(59).AddSeconds(59) # End of the same day
		        }
		    }
	    }
	    catch{
		    # Record any errors and return false by default
		    Write-Error "`tWARNING: Exception encountered while parsing time range. Details: $($_.Exception.Message). Check the syntax of entry, e.g. '<StartTime> -> <EndTime>', or days/dates like 'Sunday' and 'December 25'"   
		    return $false
	    }
		
	    # Check if current time falls within range
	    if($currentTime -ge $rangeStart -and $currentTime -le $rangeEnd){
			return 'matched'
		}
	    else{
		    return 'not matched'
	    }
	}
}