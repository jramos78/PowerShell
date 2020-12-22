<#
.SYNOPSIS
  Create or update and existing TXT record on Route 53.
  
.DESCRIPTION
  
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have permissions to view and update DNS records on Route 53.
  -Setup the AWS PowerShell module as described on https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  
.PARAMETER
  None
  
.INPUTS
  The function will prompt the user to enter the data it needs.
  
.OUTPUTS
  A new or updated TXT record on a ROute 53 hosted zone.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		Dec. 22, 2020
  Purpose/Change: 	Working version
  
.EXAMPLE
  TXTRecord
#>

function TXTRecord {
	#Get the names of every Route 53 hosted zone
	$zones = (Get-R53HostedZoneList).Name | Sort
    Write-Host "`n=========== Select a Route 53 hosted zone ===========`n"
	$count = 0
	forEach ($i in $zones) {
		$count++
		Write-Host "`tEnter ""$count"" for $i"
	}
    Write-Host "`tEnter ""Q"" to quit."
	Write-Host "`n====================================================="
	$selection = Read-Host "`nPlease make a selection"
	#List every Route 53 hosted zone
	for ($i = 1;$i -le $zones.Length;$i++){
		switch ($selection){
			$i {$r53Zone = $zones[$i-1];Break}
		}
	}
	#Exit the function if the user enters "q", "Q" or an invalid number. 
	if ($selection -eq "q"){
		Write-Warning "The script was terminated by the user!"
		Break
	} else {
		$selection = $selection -As [int]
		if (($selection -lt 1) -Or ($selection -gt $zones.Length)){Write-Warning "The script has been terminated because an invalid value was entered!";Break}
	}
	#Define the subdomain that will be updated or created
	$subDomainName = Read-Host "`nEnter the name of the subdomain"
	#Remove empty spaces from the subdomain
	$subDomainName = $subDomainName -Replace '\W',''
	#Set the name of the DNS record that we will be working with
	$recordName = $subDomainName + "." + $r53Zone
	#Define the value of the TXT record
	$txtRecordValue = Read-Host "`nEnter the value of the TXT record"
	#Remove empty spaces from the TXT record value
	$txtRecordValue = $txtRecordValue -Replace '\W',''
	#Set the DNS record type that we will be working on
	$recordType = "TXT"
	#Set the values for the DNS record 
	$r53Change = New-Object -TypeName Amazon.Route53.Model.Change
	$r53Change.ResourceRecordSet = New-Object -TypeName Amazon.Route53.Model.ResourceRecordSet
	$r53Change.ResourceRecordSet.Name = $recordName
	$r53Change.ResourceRecordSet.Type = $recordType
	$r53Change.ResourceRecordSet.TTL = 60
	$r53Change.ResourceRecordSet.ResourceRecords.Add(@{Value = "`"$txtRecordValue`""})
	#Get the ID of the selected Route53 zone
	$r53ZoneId = ((Get-R53HostedZones | Where Name -eq $r53Zone).Id).TrimStart("/hostedzone/") 
	#Update the DNS record if it currently exists or create it if it doesn't 
	if (Test-R53DNSAnswer -RecordName $subDomain -HostedZoneId $r53ZoneId -RecordType TXT){ 
		$r53Change.Action = "UPSERT" 
		$action = "updated" 
	} else { 
		$r53Change.Action = "CREATE" 
		$action = "created" 
	}
	#Get the current time
	$now = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	#Edit the DNS record
	$r53Update = Edit-R53ResourceRecordSet -HostedZoneId $r53ZoneId -ChangeBatch_Change $r53Change -ChangeBatch_Comment "This record was $action on $now" 
	#Wait until the record has been created or updated 
	$count = 0 
	do { 
		if ($count -eq 0){Write-Host "`nWaiting for the DNS record to be updated" -NoNewLine;$count++} 
		Write-Host "." -NoNewLine 
		Start-Sleep 1 
	} While ((Get-R53Change -Id $r53Update.Id).Status -eq 'PENDING') 
	Write-Host "`n`nThe following DNS record has been $action`n`n`tName:`t$recordName`n`tType:`t$recordType`n`tValue:`t$txtRecordValue`n"
}
TXTRecord
