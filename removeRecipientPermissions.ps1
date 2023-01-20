Param (
	[Parameter(Mandatory, ParameterSetName = 'Trustee')]
	[string]$trustee
)

$mb = @{ }
foreach ($m in $mb)
{
	$r = Get-RecipientPermission -Identity $m.Name | where { $_.Trustee -eq $trustee }
	
	if ($r)
	{
		foreach ($c in $r)
		{
			$mb["$($c.Identity)"] = $c
		}
	}
	
}
if ($mb.count -eq 0)
{
	write-host "No recipient permissions for $($trustee) - nothing to do."
}
else
{
	Write-Host
	write-host "Trustee has recipient permissions on these mailboxes:"
	$mb.keys
	$confirmation = Read-Host "Do you want to continue?"
	if ($confirmation -eq 'y')
	{
		foreach ($m in $mb.keys)
		{
			$accessrights = $mb[$m].AccessRights
			$name = $mb[$m].Identity
			write-host "Removing $($accessrights) from $($name) for $($trustee)"
			Remove-RecipientPermission -Identity $name -Trustee $trustee -AccessRights $accessrights -Confirm
		}
	}
}<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2022 v5.8.213
	 Created on:   	12/01/2023 17:29
	 Created by:   	rbutterworth
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>



