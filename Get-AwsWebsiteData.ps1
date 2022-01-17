<#
.SYNOPSIS
  Collects data regarding registered websites and writes them to an Excel spreadsheet
  
.DESCRIPTION
  
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have permissions to query AWS services.
  -Setup the AWS PowerShell module as described on https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  -The user account that runs the script must have a valid Office 365 account.
  
.PARAMETER 
  Region - This parameter is optional and determines which AWS region will be queried. If not entered, the user will be prompted to select an AWS region from among those in the US. 
  Domains - This parameter is a string array that supports different domains.
  
.INPUTS
  None
  
.OUTPUTS
  An Excel spreadsheet.

.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		Jan. 17, 2022
  Purpose/Change: 	Updated logic to get the SSL policy on ELBs

.EXAMPLE
	Get-AwsWebsiteData
	Get-AwsWebsiteData -Region us-east-1
#>

function Get-AwsWebsiteData {
	param (
		#Allow the $Region parameter to be empty
		[AllowNull()]
		[AllowEmptyString()]
		[AllowEmptyCollection()]
		[String]$Region,
		[parameter(Mandatory=$true)]
        [string[]]$Domains
	)
	#Ask the user to choose which US AWS region to query
	function Select-EC2Region {
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
	function Get-WebsiteData {
		#Create a spreadsheet and name it
		if ($spreadsheet.Name -ne "Sheet1"){$spreadsheet = $excel.Worksheets.Add()}
		$spreadsheet.Name = "Website data"	
		#Freeze the top row
		$excel.Rows.Item("2:2").Select() | Out-Null
		$excel.ActiveWindow.FreezePanes = $True
		#Define the column headers
		$headers = ("Website URL","Load balancer name","Load balancer public IP addresses","Load balancer scheme","Load balancer type","Target Group Name","Security policy","Target Group EC2 instance data","Target Group ports")
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
		$domainIds = @()
		#Get each domains' Id
		forEach ($i in $Domains){$domainIds += (Get-R53HostedZoneList | Where Name -eq "$i.").Id}
		#Get data from every DNS record
		forEach ($i in $domainIds){
			#Get every "A" and "CNAME" record in the specified DNS zones
			$records = ((Get-R53ResourceRecordSet -HostedZoneId $i -MaxItem 300).ResourceRecordSets | Where {($_.Type -eq "A") -or ($_.Type -eq "CNAME")}) | Sort Name
			$data = @()
			$count = 1
			forEach ($j in $records){
				[int]$progress = ($count / $records.count) * 100
				Write-Progress -Activity "Search in Progress" -Status "$progress% complete:" -PercentComplete $progress
				$name = ($j.Name).TrimEnd(".")
				$elbDnsName = $j.AliasTarget.DnsName
				#Get records that point to an ELB
				###if (($elbDnsName -like "*elb.amazonaws.com.") -or ($elbDnsName -like "*elb.us-east-1.amazonaws.com.")){
				if ($elbDnsName -like "*.amazonaws.com."){
					$elbDnsName = ($elbDnsName -Replace ("dualstack.", "")).TrimEnd(".")
					#$ipAddress = (Resolve-DnsName $elbDnsName -ErrorAction Ignore).IPAddress #| Select IPAddress -ExpandProperty IPAddress
					$ipAddress = (Resolve-DnsName $elbDnsName -ErrorAction Ignore).IPAddress -Join ", "
					$elb = Get-ELB2LoadBalancer | Where DNSName -eq $elbDnsName
					$elbArn = $elb.LoadBalancerArn
					$elbName = $elb.LoadBalancerName
					$elbScheme = $elb.Scheme.Value
					$elbType = ($elb).Type.Value
					$targetGroup = (Get-ELB2TargetGroup | Where LoadBalancerArns -eq $elbArn)
					if($targetGroup){
						$tgArn = $targetGroup.TargetGroupArn
						$tgName = $targetGroup.TargetGroupName
						$instanceIds = (Get-ELB2TargetHealth $tgArn).Target.Id
						$instance = try {(Get-EC2Instance -InstanceId $instanceIds).Instances[0]} catch {}
						if ($instance){
							$instanceName = ($instance.Tags | ? {$_.Key -eq "Name"} | Select -expand Value)
							$instanceIp = ((Get-EC2Instance).Instances | Where InstanceId -eq $instanceIds).PrivateIpAddress
							$instanceData = "$instanceName ($instanceIp)"
						} else {$instanceName = "The instance was not found!"}
						$ports = (Get-ELB2TargetHealth $tgArn).Target.Port
						$listeners = Get-ELB2Listener -LoadBalancerArn $elbArn
						$sslPolicy = [string]$listeners.SslPolicy | Where SslPolicy -ne ""
						if (!$sslPolicy){$sslPolicy = "N/A"}
						#Write the data to the spreadsheet
						$spreadsheet.Cells.Item($row,$column++) = $name
						$spreadsheet.Cells.Item($row,$column++) = $elbName
						$spreadsheet.Cells.Item($row,$column++) = $ipAddress
						$spreadsheet.Cells.Item($row,$column++) = $elbScheme
						$spreadsheet.Cells.Item($row,$column++) = $elbType
						$spreadsheet.Cells.Item($row,$column++) = $tgName
						$spreadsheet.Cells.Item($row,$column++) = $sslPolicy
						$spreadsheet.Cells.Item($row,$column++) = $instanceData
						$spreadsheet.Cells.Item($row,$column++) = $ports
						#Start the next row at column 1
						$column = 1
						#Go to the next row
						$row++
					}
				}
				$count++
			}
		}
		#Auto fit the column width
		$excel.ActiveSheet.UsedRange.EntireColumn.AutoFit() | Out-Null
		#Format active cells into a table
		$ListObject = $excel.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $excel.ActiveCell.CurrentRegion, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
		$ListObject.Name = "TableData"
		$ListObject.TableStyle = "TableStyleMedium9"
		Write-Host "`nDone!" -ForegroundColor Green
	}
	#Define a new Excel object as a global variable and create a new workbook
	$excel = New-Object -ComObject Excel.Application
	#create a new Excel workbook
	$workbook = $excel.Workbooks.Add()
	#Create a new spreadsheet
	$spreadsheet = $workbook.Worksheets.Item(1)
	#Check if the AWS PowerShell module has been installed
	$modules = (Get-Module -ListAvailable).Name
	if (!($modules.Contains("AWSPowerShell"))) {
		Write-Warning "`nThis script will not continue because the AWS PowerShell module has not been installed.`nVisit https://docs.aws.amazon.com/powershell/latest/userguide/pstools-getting-set-up-windows.html for instructions on how to download and install it."
	} else {
		#Import the AWS PowerShell module
		Import-Module AWSPowerShell
		#Call a function if a variable has not been declared
		if (!($Region)){$Region = Select-EC2Region}
		Write-Host "`nGenerating infrastructure data for every registered website. Excel will automatically open when all of the data has been collected.`n" -ForegroundColor Green
		#Call the other function only if a valid region was selected
		if ($Region) {
			Get-WebsiteData
			#Open the spreadsheet
			$excel.Visible = $True
		} else {Clear;Write-Warning "The script has exited because an AWS EC2 region was not selected."}
	}
}
Get-AwsWebsiteData
