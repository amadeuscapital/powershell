Param(
    [Parameter(Mandatory=$true,HelpMessage='email address of the calendar to which you are removing write access')]
    [string]$mailBox,
    [Parameter(Mandatory=$true,HelpMessage='email address of the delgate to whom you are removing access for')]
    [string]$delegate
)
function write_logFile
{
	Param ([string]$log_file,
		$log_message)
	$time_stamp = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
	$log_message = "[$($time_stamp)]: " + $log_message
	Add-content $log_file -value $log_message
}
Write-Output 'Loading O365 Powershell support'

& C:\bin\justlogon.ps1

if (!(Get-Recipient $mailBox -ErrorAction SilentlyContinue)) {
      Write-Output "ERROR: no such mailbox ($mailBox)"
      & C:\bin\cleanup.ps1
      Exit
}
if (!(Get-Recipient $delegate -ErrorAction SilentlyContinue)) {
      Write-Output "ERROR: no such delegate ($delegate)"
      & C:\bin\cleanup.ps1
      Exit
}


$title = "Confirm Delegation Removal"
$message = "Remove $delegate write access from the calendar for $mailBox ?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Remove write access to $delegate."

$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Cancel this request."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

switch ($result)
    {
        0 {
            Remove-MailboxFolderPermission -Identity $mailBox":\Deleted Items" -user $delegate
            Remove-MailboxFolderPermission -Identity $mailBox":\Calendar" -user $delegate
            Write-Output "Delegation enabled"
            Remove-MailboxFolderPermission -Identity $mailBox":\Sent Items" -user $delegate -ErrorAction SilentlyContinue
            Remove-RecipientPermission -Trustee $delegate -Identity $mailBox -AccessRights SendAs  -ErrorAction SilentlyContinue
            Write-Output "SendAs permissions added"
            
        }
        1 {"Delegation cancelled."
        & C:\bin\cleanup.ps1
        exit
        }
    }



& C:\bin\cleanup.ps1