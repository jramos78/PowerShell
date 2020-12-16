<#
.SYNOPSIS
  Creates or updates a DNS record on Route 53 and points it to an existing Application Load-Balancer.
  
.DESCRIPTION
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have the appropriate Route 53 permissions.
  -Setup the AWS PowerShell module as described on https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  
.PARAMETER
  None
  
.INPUTS
  None
  
.OUTPUTS
  An Excel spreadsheet.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78)
  Updated on:  		Dec. 16, 2020
  Purpose/Change: 	Working version
  
.EXAMPLE
  None, just make sure to assing values to the following variables:
	$domainName
	$subDomainName
	$albName
#>

##### Requires the AWS PowerShell module

#Enter the name of the domain where the record will be created or updated 
$domainName = "<domain name>" 
#Enter the name of the subdomain that will be created or updated 
$subDomainName = "<subdomain name>" 
$subDomain = $subDomainName + "." + $domainName 
#Enter the name of the Application Load Balancer that the DNS record will point to 
$albName = "<load balancer name>" 
$alb = (Get-ELB2LoadBalancer -Name $albName).DNSName 
$albCanonicalId = (Get-ELB2LoadBalancer -Name $albName).CanonicalHostedZoneId 
#Set the values for the DNS record 
$r53Change = New-Object -TypeName Amazon.Route53.Model.Change 
$r53Change.ResourceRecordSet = New-Object Amazon.Route53.Model.ResourceRecordSet 
$r53Change.ResourceRecordSet.Name = $subDomain + "." 
$r53Change.ResourceRecordSet.Type = "A" 
$r53Change.ResourceRecordSet.AliasTarget = New-Object Amazon.Route53.Model.AliasTarget 
$r53Change.ResourceRecordSet.AliasTarget.HostedZoneId = $albCanonicalId 
$r53Change.ResourceRecordSet.AliasTarget.DNSName = $alb + "." 
$r53Change.ResourceRecordSet.AliasTarget.EvaluateTargetHealth = $False 
#Get the Zone Id of the Route 53 hosted zone 
$r53ZoneId = ((Get-R53HostedZones | Where Name -like "$domainName.").Id).TrimStart("/hostedzone/") 
#Update the DNS record if it currently exists or create it if it doesn't 
if (Test-R53DNSAnswer -RecordName $subDomain -HostedZoneId $r53ZoneId -RecordType A){ 
    $r53Change.Action = "UPSERT" 
    $action = "updated" 
} else { 
    $r53Change.Action = "CREATE" 
    $action = "created" 
}
#Update Route53 
$now = Get-Date -Format "MM/dd/yyyy HH:mm:ss" 
$r53Update = Edit-R53ResourceRecordSet -HostedZoneId $r53ZoneId -ChangeBatch_Change $r53Change -ChangeBatch_Comment "This record was $action on $now" 
#Wait until the record has been created our updated 
$count = 0 
do { 
    if ($count -eq 0){Write-Host "`nWaiting for the DNS record to be $action" -NoNewLine;$count++} 
    Write-Host "." -NoNewLine 
    Start-Sleep 1 
} While ((Get-R53Change -Id $r53Update.Id).Status -eq 'PENDING') 
Write-Host "`n`nThe DNS record has been $action!`n`n" 
; 
