<#
.SYNOPSIS
  Create or update and existing CNAME or TXT record on Route 53.
  
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
  A new or updated CNAME or TXT record on a Route 53 hosted zone.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		Feb. 23, 2021
  Purpose/Change: 	Modified script to support CNAME records 
  
.EXAMPLE
  R53Record
#>

#Create or update a TXT record
function R53Record {
    #Get the names of every Route 53 hosted zone
	$zones = (Get-R53HostedZoneList).Name | Sort
	for ($i = 0;$i -lt $zones.Length;$i++){Write-Host "`tEnter $i for" $zones[$i]}
	Write-Host "`tEnter ""Q"" to quit."
	Write-Host "`n====================================================="
	$selection = Read-Host "`nPlease make a selection"
	#List every Route 53 hosted zone
	for ($i = 0;$i -le $zones.Length;$i++){
		switch ($selection){$i {$r53Zone = $zones[$i];Break}}
	}
    #Exit the function if the user enters "q", "Q" or an invalid number.
    if ($selection -eq "q"){
        Write-Warning "The script was terminated by the user!";Break
    } else {
        $selection = $selection -As [int]
        if (($selection -lt 0) -Or ($selection -ge $zones.Length)){Write-Warning "Invalid entry. The script has been terminated!`n";Break}
    }
	#Get the type of record that will be created/updated
	$types = ("CNAME","TXT")
    Write-Host "`n=========== Select a DNS record type ===========`n"
	for ($i = 0;$i -lt $types.Length;$i++){Write-Host "`tEnter $i for" $types[$i]}
    Write-Host "`tEnter ""Q"" to quit."
    Write-Host "`n====================================================="
    $selection = Read-Host "`nPlease make a selection"
	#Set the DNS record's type based on the user's selection
	if ($selection -eq 0){
		$recordType = $types[0]
	} elseif ($selection -eq 1){
		$recordType = $types[1]
	} elseif (($selection -lt 0) -Or ($selection -gt $type.Length)){
		Write-Warning "Invalid entry. The script has been terminated!`n";Break
	}
	#Exit the function if the user enters "q", "Q" or an invalid number.
    if ($selection -eq "q"){
        Write-Warning "The script was terminated by the user!";Break
    } else {
        $selection = $selection -As [int]
        if (($selection -lt 0) -Or ($selection -gt $zones.Length)){Write-Warning "Invalid entry. The script has been terminated!`n";Break}
	}
    #Define the subdomain that will be updated or created
    $subDomainName = Read-Host "`nEnter the DNS record`'s subdomain (name)"
    #Remove empty spaces from the subdomain
    $subDomainName = $subDomainName -Replace " ",""
    #Set the DNS record's value
    $recordValue = Read-Host "`nEnter the DNS record`'s value"
    #Remove empty spaces from the TXT record value
    $recordValue = $recordValue -Replace " ",""
    #Set the DNS record's FQDN
    $recordName = $subDomainName + "." + $r53Zone
    #Set the values for the DNS record
    $r53Change = New-Object -TypeName Amazon.Route53.Model.Change
    $r53Change.ResourceRecordSet = New-Object -TypeName Amazon.Route53.Model.ResourceRecordSet
    $r53Change.ResourceRecordSet.Name = $recordName
    $r53Change.ResourceRecordSet.Type = $recordType
    $r53Change.ResourceRecordSet.TTL = 60
	#Prepend and append double quotes to the DNS record's value if its a TXT record
	if($recordType -eq "TXT"){$recordValue = "`"$recordValue`""}
    $r53Change.ResourceRecordSet.ResourceRecords.Add(@{Value = $recordValue})
    #Get the ID of the selected Route53 zone  
    $r53ZoneId = ((Get-R53HostedZones | Where Name -eq $r53Zone).Id).TrimStart("/hostedzone/")
    #Update the DNS record if it currently exists or create it if it doesn't   
    if (Test-R53DNSAnswer -RecordName $recordName -HostedZoneId $r53ZoneId -RecordType $recordType){
        $r53Change.Action = "UPSERT"
        $action = "updated"
    } else {
        $r53Change.Action = "CREATE"
        $action = "created"
    }
    #Edit the DNS record
    $now = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
    $r53Update = Edit-R53ResourceRecordSet -HostedZoneId $r53ZoneId -ChangeBatch_Change $r53Change -ChangeBatch_Comment "This record was $action on $now"
    #Wait until the record has been created or updated
    $count = 0
    do {
        if ($count -eq 0){Write-Host "`nWaiting for the DNS record to be updated" -NoNewLine -ForeGroundColor Cyan;$count++}
        Write-Host "." -NoNewLine -ForeGroundColor Cyan   
        Start-Sleep 1   
    } While ((Get-R53Change -Id $r53Update.Id).Status -eq 'PENDING')
    ####Write-Host "`n`nThe DNS record has been updated!`n`n"
    Write-Host "`n`nThe following DNS record has been $action`n`n`tName:`t$recordName`n`tType:`t$recordType`n`tValue:`t$recordValue`n" -ForeGroundColor Cyan
}  
R53Record
