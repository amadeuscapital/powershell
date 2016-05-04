Param(
    [Parameter(Mandatory=$true,HelpMessage='email address of the calendar to which you are granting write access')]
    [string]$mailBox,
    [Parameter(Mandatory=$true,HelpMessage='email address of the delgate to whom you are giving access')]
    [string]$delegate,  
    [Parameter(Mandatory=$true,HelpMessage='Should the delegate be allowed to send on behalf of the mailbox user?')]
    [bool]$sendAs,
    [Parameter(Mandatory=$false,HelpMessage='Should the script load the EOL module?')]
    [bool]$loadmod

)

if ($loadmod) {
    Write-Output 'Loading O365 Powershell support'
    & C:\bin\justlogon.ps1
}

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


$title = "Confirm Delegation"
$message = "Give $delegate write access to the calendar for $mailBox ?"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
    "Grant write access to $delegate."

$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
    "Cancel this request."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

switch ($result)
    {
        0 {
            Add-MailboxFolderPermission -Identity $mailBox":\Deleted Items" -user $delegate -AccessRights NonEditingAuthor
            Add-MailboxFolderPermission -Identity $mailBox":\Calendar" -user $delegate -AccessRights Editor
            Write-Output "Delegation enabled"
            if ($sendAs) {
                  Add-MailboxFolderPermission -Identity $mailBox":\Sent Items" -user $delegate -AccessRights NonEditingAuthor
                  Add-RecipientPermission -Trustee $delegate -Identity $mailBox -AccessRights SendAs
                  Write-Output "SendAs permissions added"
            }
        }
        1 {"Delegation cancelled."
        exit
        }
    }


if ($loadmod) {
    & C:\bin\cleanup.ps1
}