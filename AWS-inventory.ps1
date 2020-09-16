<#
.SYNOPSIS
  Generates an AWS resource inventory on an Excel spreadsheet. This script 
  
.DESCRIPTION
  
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have permissions to query AWS services.
  -Setup the AWS PowerShell module as described on https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  -Excel installed on the computer where the script is run.
  
.PARAMETER
  None
  
.INPUTS
  None
  
.OUTPUTS
  An Excel spreadsheet.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		  Sept. 15, 2020
  Purpose/Change: 	Working version
  
.EXAMPLE
  AWS-inventory
#>

##### Requires the AWS PowerShell module

#Prompt the user to choose which US AWS region to inventory
function Select-EC2Region {
	$regions = @(Get-EC2Region | Where RegionName -like "us-*").RegionName
    Write-Host "`n========== Select an AWS region ==========`n"
    #for($i -eq 1; $i -lt $regions.Count; i++) {
	$count = 0
	forEach ($i in $regions) {
		$count++
		Write-Host "`tPress ""$count"" for $i"
	}
    Write-Host "`tPress ""Q"" to quit."
	Write-Host "`n--> The default option is ""1""."
	Write-Host "`n=========================================="
	$global:selection = Read-Host "`nPlease make a selection"
	for ($i = 1;$i -le $regions.Length;$i++){
		switch ($selection){
			$i {$global:region = $regions[$i-1];Break}
		}
	}
	#Exit the function if the user enters "q" or "Q". Set the region to the first on the list if an invalid option is entered.
	if ($selection -eq "q"){
		Write-Warning "The script was terminated by the user"
		Break
	} elseif (!(($selection -ge 1)) -or (!($selection -le $global:regions.Length))){
		$global:region = $regions[0]
		Write-Warning "The region has been defaulted to $global:region due to an invalid entry!"
	}
}

#Get an inventory of EC2 instances
function EC2-inventory{ 
	Write-Host "`nGenerating EC2 inventory.`n" -ForegroundColor Green

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "EC2"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True

	#Define the column headers
	$headers = ("Name tag","State","Operating System","Domain tag","Instance ID","Instance size","VPC Id","Subnet Id","Private IP address","Public IP address","Description tag")
	$column = 1
	
	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}
	
	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Length) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1
		
	#Get the EC2 instances
	$instances = ((Get-EC2Instance).Instances)
	forEach ($i in $instances){
		$spreadsheet.Cells.Item($row,$column++) = ((Get-EC2Tag | Where ResourceID -eq $i.instanceID ) | Where Key -eq "Name").Value #Value of "Name" tag
		$spreadsheet.Cells.Item($row,$column++) = $i.state.name.value #Instance state
		$spreadsheet.Cells.Item($row,$column++) = (Get-SSMInstanceInformation | Where InstanceId -eq $i.instanceID).PlatformName #Operating System
		$spreadsheet.Cells.Item($row,$column++) = ((Get-EC2Tag | Where ResourceID -eq $i.instanceID ) | Where Key -eq "Domain").Value #Value of "Domain" tag
		$spreadsheet.Cells.Item($row,$column++) = $i.instanceID #Instance ID
		$spreadsheet.Cells.Item($row,$column++) = $i.Instancetype.Value #Instance size
		$spreadsheet.Cells.Item($row,$column++) = $i.VpcId #VPC ID
		$spreadsheet.Cells.Item($row,$column++) = $i.SubnetId #Subnet ID
		$spreadsheet.Cells.Item($row,$column++) = $i.PrivateIpAddress #Private IP address
		$spreadsheet.Cells.Item($row,$column++) = $i.PublicIpAddress #Public IP address
		$spreadsheet.Cells.Item($row,$column++) = ((Get-EC2Tag | Where ResourceID -eq $i.instanceID ) | Where Key -eq "Description").Value #Value of "Description" tag
			
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function VPC-inventory {

	Write-Host "`nGenerating VPC inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "VPC"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("VPC ID","CIDR bock","State","Default?","DCHP option")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all VPCs
	$vpc = (Get-EC2Vpc).VpcId
	forEach ($i in $vpc){
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).VpcId
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).CidrBlock
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).State.Value
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).IsDefault
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).DhcpOptionsId
		
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}

	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function Subnets-inventory {
	Write-Host "`nGenerating Subnets inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "Subnets"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Subnet ID","Subnet VPC","Availability Zone","Availability Zone Id","Default for AZ?","State","CIDR block","Available IP addresses","Public IP on launch?")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all Subnets
	$subnets = (Get-EC2Subnet).SubnetId
	forEach ($i in $subnets){
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).SubnetId
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).VpcId
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).AvailabilityZone
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).AvailabilityZoneId
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).DefaultForAz
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).State.Value
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).CidrBlock
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).AvailableIpAddressCount
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Subnet $i).MapPublicIpOnLaunch
			
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function S3-inventory {

	Write-Host "`nGenerating S3 inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "S3"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Name","Description")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all S3 buckets
	$buckets = ((Get-S3Bucket).BucketName)
	forEach ($i in $buckets){
		$spreadsheet.Cells.Item($row,$column++) = $i #Bucket name
		$spreadsheet.Cells.Item($row,$column++) = ((Get-S3BucketTagging $i) | Where Key -eq "Description").Value #Value of "Description" tag
		
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}

	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function RDS-inventory {
	Write-Host "`nGenerating RDS inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "RDS"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("ARN","Engine","Name","Size","Cluster","Availability Zone","Multi AZ?","Description")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all RDS instances
	$rds = ((Get-RDSDBInstance).DBInstanceArn)
	forEach ($i in $rds){
		$spreadsheet.Cells.Item($row,$column++) = $i #ARN
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).Engine #Database engine
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).DBName #DBName
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).DBInstanceClass #Instance size
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).DBClusterIdentifier #Cluster 
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).AvailabilityZone #Availability Zone
		$spreadsheet.Cells.Item($row,$column++) = (Get-RDSDBInstance $i).MultiAZ #Hosted on multiple availabilty zones?
		$spreadsheet.Cells.Item($row,$column++) = ((Get-RDSTagForResource $i) | Where Key -eq "Description").Value #Value of "Description" tag
			
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function IAM-inventory {
	Write-Host "`nGenerating IAM inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "IAM"	
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Username","User ID","Created","Password last used")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all IAM users
	$iamUsers = Get-IAMUserList
	forEach ($i in $iamUsers){
		$spreadsheet.Cells.Item($row,$column++) = $i.Username # Username
		$spreadsheet.Cells.Item($row,$column++) = $i.UserId # User ID
		$spreadsheet.Cells.Item($row,$column++) = $i.CreateDate #When the account was created
		$spreadsheet.Cells.Item($row,$column++) = $i.PasswordLastUsed #Last time the password was used
		
		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function ELB-inventory {
	Write-Host "`nGenerating ELB inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "ELB"	
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Name","DNS name","ARN","Scheme","Type","Description tag","Availability Zones","IP address(es)","Target Group","Target Group instance and port")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all ELBs instances
	$elbs = Get-ELB2LoadBalancer
	forEach ($i in $elbs){
		$spreadsheet.Cells.Item($row,$column++) = $i.LoadBalancerName
		$spreadsheet.Cells.Item($row,$column++) = $i.DNSName
		$spreadsheet.Cells.Item($row,$column++) = $i.LoadBalancerArn 
		$spreadsheet.Cells.Item($row,$column++) = $i.Scheme.Value 
		$spreadsheet.Cells.Item($row,$column++) = ($i).Type.Value
		$spreadsheet.Cells.Item($row,$column++) = ((Get-ELB2Tag -ResourceArn ($i.LoadBalancerArn)).Tags | Where Key -eq "Description").Value #Value of "Description" tag

		#Get the ELB's availability zones
		$values = ($i).AvailabilityZones.ZoneName
		$AZs = @()
		forEach ($j in $values){$AZs += "$j"}
		$spreadsheet.Cells.Item($row,$column++) = $AZs -Join ", " #ELB availability zone(s)
		
		#Get the ELB's IP address(es)
		$values = (Resolve-DnsName $i.DnsName) | Select IPAddress -ExpandProperty IPAddress
		$IPs = @()
		forEach ($k in $values){$IPs += "$k"}
		$spreadsheet.Cells.Item($row,$column++) = $IPs -Join ", " #IP addresses assigned to the ELB

		$spreadsheet.Cells.Item($row,$column++) = (Get-ELB2TargetGroup $i.LoadBalancerArn).TargetGroupName #Target Group name

		#Get the IDs of the EC2 instances attached to the target group
		try {
			$tgARN = (Get-ELB2TargetGroup | Where LoadBalancerArns -eq $i.LoadBalancerARN).TargetGroupArn
			$targets = (Get-ELB2TargetHealth $tgARN).Target | Select Id,Port | forEach {$_.Id + " (" + $_.Port + ")"}
			$targets = $targets -Join ", "
			$spreadsheet.Cells.Item($row,$column++) = $targets #Target Group EC2 instance ID and port
			$targets = $null 
		} catch {}

		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function DS-inventory {
	Write-Host "`nGenerating Directory Service inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "Directory Service"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Name","Directory Id","DNS servers","Access URL")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all Directory Service instances
	$directoryService = Get-DSDirectory
	
	forEach ($i in $directoryService){
		$spreadsheet.Cells.Item($row,$column++) = $i.Name
		$spreadsheet.Cells.Item($row,$column++) = $i.DirectoryId
		$spreadsheet.Cells.Item($row,$column++) = $i.DnsIpAddrs[0] + "," + $i.DnsIpAddrs[1]
		$spreadsheet.Cells.Item($row,$column++) = $i.AccessUrl

		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function WS-inventory {
	Write-Host "`nGenerating Workspaces inventory.`n" -ForegroundColor Green 

	#Create a spreadsheet and name it
	if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $global:excel.Worksheets.Add()}
	$spreadsheet.Name = "Workspaces"
	
	#Freeze the top row
	$global:excel.Rows.Item("2:2").Select() | Out-Null
	$global:excel.ActiveWindow.FreezePanes = $True
	
	#Define the column headers
	$headers = ("Hostname","Assigned user","Domain","IP address","Network Interface ID","Public IP address")
	$column = 1

	#Write the headers on the top row in bold text
	forEach($i in $headers) {
		$spreadsheet.Cells.Item(1,$column) = $i
		$spreadsheet.Cells.Item(1,$column).Font.Bold = $True
		$column++
	}

	#Add an auto-filter to each column header
	$spreadsheet.Cells.Autofilter(1,$headers.Count) | Out-Null
	 
	#Set the starting column and row in the spreadsheet to write data 
	$row = 2
	$column = 1

	#Get all Workspaces
	$workSpaces = Get-WKSWorkspace
	
	forEach ($i in $workSpaces){
		$spreadsheet.Cells.Item($row,$column++) = $i.ComputerName
		$spreadsheet.Cells.Item($row,$column++) = $i.UserName
		$spreadsheet.Cells.Item($row,$column++) = (Get-DSDirectory $i.DirectoryId).Name
		$spreadsheet.Cells.Item($row,$column++) = $i.IpAddress
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2NetworkInterface | Where PrivateIpAddress -eq $i.IpAddress).NetworkInterfaceId
		$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Address | where PrivateIpAddress -eq $i.IpAddress).PublicIp

		#Start the next row at column 1
		$column = 1
		#Go to the next row
		$row++
	}
	#Auto fit the column width
	$global:excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
	
	#Format active cells into a table
	$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
	$ListObject.Name = "TableData"
	$ListObject.TableStyle = "TableStyleMedium9"
}

function AWS-inventory {
	#Define a new Excel object as a global variable and create a new workbook
	$global:excel = New-Object -ComObject Excel.Application
	#create a new Excel workbook
	$global:workbook = $global:excel.Workbooks.Add()
	#Create a new spreadsheet
	$spreadsheet = $global:workbook.Worksheets.Item(1)

	#Check if the AWS PowerShell module has been installed
	$modules = (Get-Module -ListAvailable).Name
	if (!($modules.Contains("AWSPowerShell"))) {
		Write-Warning "`nThis script will not continue because the AWS PowerShell module has not been installed.`nVisit https://docs.aws.amazon.com/powershell/latest/userguide/pstools-getting-set-up-windows.html for instructions on how to download and install it."
	} else {
		#Import the AWS PowerShell module
		Import-Module AWSPowerShell

		#Call the functions
		Select-EC2Region
		Write-Host "`nGenerating AWS inventory. Excel will automatically open when all of the data has been collected.`n" -ForegroundColor Green
		if ($region) {
			S3-inventory
			RDS-inventory
			IAM-inventory
			ELB-inventory
			WS-inventory
			DS-inventory
			Subnets-inventory
			VPC-inventory
			EC2-inventory
			#Open the spreadsheet
			$global:excel.Visible = $True
		} else {Clear;Write-Warning "The script has exited because an AWS EC2 region was not selected."}
	}
}
AWS-inventory
