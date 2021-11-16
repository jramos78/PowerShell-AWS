<#
.SYNOPSIS
  Generates an AWS resource inventory on an Excel spreadsheet.
  
.DESCRIPTION
 This script generates an inventory of the AWS services/resources listed below and writes it to an Excel spreadsheet. 
 -S3 buckets
 -RDS instances
 -IAM users
 -Elastic Load Balancers (ELB)
 -Workspaces
 -Directory Services
 -VPC
 -VPC subnets
 -EC2 instances
  
.PARAMETER Region
  This parameter is optional and determines which AWS region will be queried. If not entered, the user will be prompted to select an AWS region from among those in the US. 
  
.INPUTS
  None
  
.OUTPUTS
  An Excel spreadsheet.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		Nov. 16, 2021
  Purpose/Change: 	Fixed syntax errors
  
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have permissions to query AWS services.
  -The desktop version of Excel.
  -The AWS PowerShell module has been configured as described at https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  
.EXAMPLE
  Get-AwsInventory
  Get-AwsInventory -Region us-east-1
#>
function Get-AwsInventory {
	param (
		#Allow the $Region parameter to be empty
		[AllowNull()]
		[AllowEmptyString()]
		[AllowEmptyCollection()]
		[String]$Region
	)
	#Create a new Excel file
	$excel = New-Object -ComObject Excel.Application
	#create a new Excel workbook
	$workbook = $excel.Workbooks.Add()
	#Create a new spreadsheet
	$spreadsheet = $workbook.Worksheets.Item(1)
	#Prompt the user to choose which US AWS region to inventory
	function Select-AwsRegion {
		$regions = @(Get-EC2Region | Where RegionName -like "us-*").RegionName
		Write-Host "`n========== Select an AWS region ==========`n"
		$count = 0
		forEach ($i in $regions) {
			$count++
			Write-Host "`tPress ""$count"" for $i"
		}
		Write-Host "`tPress ""Q"" to quit."
		Write-Host "`n--> The default option is ""1""."
		Write-Host "`n=========================================="
		$selection = Read-Host "`nPlease make a selection"	
		for ($i = 1;$i -le $regions.Length;$i++){
			switch ($selection){
				$i {$region = $regions[$i-1];Break}
			}
		}
		#Exit the function if the user enters "q" or "Q". Set the region to the first on the list if an invalid option is entered.
		if ($selection -eq "q"){
			Write-Warning "The script was terminated by the user!"
			Break
		} elseIf (($selection -ge 1) -And ($selection -le $regions.Length)){
			Return $regions[$selection-1]
		} else {
			Write-Host "The region has been defaulted to" $regions[0] "due to an invalid entry!" -ForeGroundColor Yellow
			Return $regions[0]
		}
	}
	#Get an inventory of EC2 instances
	function Get-Ec2Inventory{ 
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "EC2"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $instances){
			[int]$progress = ($count / $instances.count) * 100
			Write-Progress -Activity "Gathering EC2 instance data" -Status "$progress% complete:" -PercentComplete $progress
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
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-VpcInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "VPC"
		
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
		
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
		$count = 1
		forEach ($i in $vpc){
			[int]$progress = ($count / $vpc.count) * 100
			Write-Progress -Activity "Gathering VPC data" -Status "$progress% complete:" -PercentComplete $progress
			$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).VpcId
			$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).CidrBlock
			$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).State.Value
			$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).IsDefault
			$spreadsheet.Cells.Item($row,$column++) = (Get-EC2Vpc $i).DhcpOptionsId
			#Start the next row at column 1
			$column = 1
			#Go to the next row
			$row++
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-SubnetsInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "Subnets"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $subnets){
			[int]$progress = ($count / $subnets.count) * 100
			Write-Progress -Activity "Gathering data on subnets" -Status "$progress% complete:" -PercentComplete $progress
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
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-S3Inventory{
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "S3"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $buckets){
			[int]$progress = ($count / $buckets.count) * 100
			Write-Progress -Activity "Gathering S3 bucket data" -Status "$progress% complete:" -PercentComplete $progress
			$spreadsheet.Cells.Item($row,$column++) = $i #Bucket name
			$spreadsheet.Cells.Item($row,$column++) = ((Get-S3BucketTagging $i) | Where Key -eq "Description").Value #Value of "Description" tag
			#Start the next row at column 1
			$column = 1
			#Go to the next row
			$row++
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-RdsInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "RDS"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $rds){
			[int]$progress = ($count / $rds.count) * 100
			Write-Progress -Activity "Gathering RDS instance data" -Status "$progress% complete:" -PercentComplete $progress
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
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-IamInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "IAM"	
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $iamUsers){
			[int]$progress = ($count / $iamUsers.count) * 100
			Write-Progress -Activity "Gathering data on IAM users" -Status "$progress% complete:" -PercentComplete $progress
			$spreadsheet.Cells.Item($row,$column++) = $i.Username # Username
			$spreadsheet.Cells.Item($row,$column++) = $i.UserId # User ID
			$spreadsheet.Cells.Item($row,$column++) = $i.CreateDate #When the account was created
			$spreadsheet.Cells.Item($row,$column++) = $i.PasswordLastUsed #Last time the password was used
			#Start the next row at column 1
			$column = 1
			#Go to the next row
			$row++
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-ElbInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "ELB"	
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $elbs){
			[int]$progress = ($count / $elbs.count) * 100
			Write-Progress -Activity "Gathering data on Elastic Load Balancers" -Status "$progress% complete:" -PercentComplete $progress
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
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-DsInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "Directory Service"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $directoryService){
			[int]$progress = ($count / $directoryService.count) * 100
			Write-Progress -Activity "Gathering data on Directory Services" -Status "$progress% complete:" -PercentComplete $progress
			$spreadsheet.Cells.Item($row,$column++) = $i.Name
			$spreadsheet.Cells.Item($row,$column++) = $i.DirectoryId
			$spreadsheet.Cells.Item($row,$column++) = $i.DnsIpAddrs[0] + "," + $i.DnsIpAddrs[1]
			$spreadsheet.Cells.Item($row,$column++) = $i.AccessUrl
			#Start the next row at column 1
			$column = 1
			#Go to the next row
			$row++
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	function Get-WorkspaceInventory {
		#Create a spreadsheet and assign in a name
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "Workspaces"
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
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
		$count = 1
		forEach ($i in $workSpaces){
			[int]$progress = ($count / $workSpaces.count) * 100
			Write-Progress -Activity "Gathering data on AWS Workspaces" -Status "$progress% complete:" -PercentComplete $progress
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
			$count++
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
	}
	#Check if the AWS PowerShell module has been installed
	$modules = (Get-Module -ListAvailable).Name
	if (!($modules.Contains("AWSPowerShell"))) {
		Write-Host "`nThis script will not continue because the AWS PowerShell module has not been installed.`nVisit https://docs.aws.amazon.com/powershell/latest/userguide/pstools-getting-set-up-windows.html for instructions on how to download and install it.`n" -ForeGroundColor Yellow
	} else {
		#Import the AWS PowerShell module
		Import-Module AWSPowerShell
		#Call a function if a variable has not been declared
		if (!($Region)){$Region = Select-AwsRegion}
		#Call the functions
		Write-Host "`nGenerating AWS inventory, Excel will automatically open when all of the data has been collected.`n" -ForegroundColor Green
		if ($Region) {
			Get-S3Inventory
			Get-RdsInventory
			Get-IamInventory
			Get-ElbInventory
			Get-WorkspaceInventory
			Get-DsInventory
			Get-SubnetsInventory
			Get-VpcInventory
			Get-Ec2Inventory
			#Open the spreadsheet
			$excel.Visible = $True
			Write-Host "`nDone!`n" -ForeGroundColor Green
		} else {Clear;Write-Warning "The script has exited because an AWS EC2 region was not selected."}
	}
}
