<#
.SYNOPSIS
  Creates or updates a DNS TXT record on Route 53.
  
.DESCRIPTION
 This script creates or updates a DNS TXT record on Route 53. I usually use it to submit ACME challenges when renewing Let's Encrypt SSL certificates.
  
.PARAMETER 
	TxtRecordName - The name of the TXT record.
	TxtRecordValue - The record's value.
  
.NOTES
  Author:         	Jorge Ramos (https://github.com/jramos78/PowerShell)
  Updated on:  		Dec. 13, 2021
  Purpose/Change: 	Initial upload.
  
  This script requires the following:
  -The AWS PowerShell module (this can be installed directly from PowerShell by running the following command: Install-Module -Name AWSPowerShell -Force)
  -IAM credentials that have permissions to query AWS services.
  -The AWS PowerShell module has been configured as described at https://docs.aws.amazon.com/powershell/latest/userguide/specifying-your-aws-credentials.html
  
.EXAMPLE
  Set-R53TxtRecord
  Set-R53TxtRecord -TextRecordName _acme-challenge.mydomain.com -TextRecordValue lTE8BkGu_Mp8BxM6Xcy78p2f_9QcfgL6CxuzrmI6TEI 
#>
function Set-R53TxtRecord {  
	param(  
		[Parameter(Mandatory,ValueFromPipelinebyPropertyName,HelpMessage="Enter the TXT record's name.")]  
		[String]$TextRecordName,  
		[Parameter(Mandatory,ValueFromPipelinebyPropertyName,HelpMessage="Enter the TXT record value.")]  
		[String]$TextRecordValue  
	)  
	#Extract the domain name from the value of the $TextRecordName variable    
	$domainName = ($TextRecordName.Split(".",3) | Select -Index 2,2) -Join "."  
	#Set the values for the DNS record    
	$r53Change = New-Object -TypeName Amazon.Route53.Model.Change   
	$r53Change.Action = "UPSERT"   
	$r53Change.ResourceRecordSet = New-Object -TypeName Amazon.Route53.Model.ResourceRecordSet   
	$r53Change.ResourceRecordSet.Name = $TextRecordName + "."   
	$r53Change.ResourceRecordSet.Type = "TXT"   
	$r53Change.ResourceRecordSet.TTL = 60   
	$r53Change.ResourceRecordSet.ResourceRecords.Add(@{Value = "`"$TextRecordValue`""})   
	#Get the ID of the Route53 zone   
	$r53ZoneId = ((Get-R53HostedZones | Where Name -like "$domainName.").Id).TrimStart("/hostedzone/")  
	#Update the DNS record if it currently exists or create it if it doesn't  
	if (Test-R53DNSAnswer -RecordName $TextRecordName -HostedZoneId $r53ZoneId -RecordType $r53Change.ResourceRecordSet.Type){  
		$r53Change.Action = "UPSERT"  
		$action = "updated"  
	} else {  
		$r53Change.Action = "CREATE"  
		$action = "created"
	}  
	#Edit the DNS record  
	$now = Get-Date -Format "M/d/yyyy HH:MM" 
	$r53Update = Edit-R53ResourceRecordSet -HostedZoneId $r53ZoneId -ChangeBatch_Change $r53Change -ChangeBatch_Comment "This record was $action on $now"    
	#Wait until the record has been created our updated    
	$count = 0    
	do {    
		if ($count -eq 0){Write-Host "`nWaiting for the DNS record to be updated" -NoNewLine;$count++}    
		Write-Host "." -NoNewLine    
		Start-Sleep 1    
	} While ((Get-R53Change -Id $r53Update.Id).Status -eq 'PENDING')
	#Compare the current value of the TXT record with what was entered and display a message if the update was successful 
	$value = ((Test-R53DNSAnswer -RecordName $TextRecordName -HostedZoneId $r53ZoneId -RecordType $r53Change.ResourceRecordSet.Type).RecordData).Replace('"','') 
	$type = $r53Change.ResourceRecordSet.Type.Value 
	if ($value -eq $TextRecordValue){Write-Host "`n`nThe DNS record has been updated as follows:`n`tName:`t$TextRecordName`n`tType:`t$type`n`tValue:`t$TextRecordValue`n" -ForeGroundColor Cyan} 
} 
