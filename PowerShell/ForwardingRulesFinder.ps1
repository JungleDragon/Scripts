Set-ExecutionPolicy RemoteSigned -Force

Write-Host "Checking if ExchangeOnlineManagement Module is already installed"
try
    {Write-Host "Importing ExchangeOnlineManagement Module"
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "ExchangeOnlineManagement Module Imported"}
catch
    {Write-Host "Need to install ExchangeOnlineManagement Module"
    Install-Module -Name ExchangeOnlineManagement -Force
    Write-Host "ExchangeOnlineManagement Module Installed"
    Write-Host "Importing ExchangeOnlineManagement Module"
    Import-Module ExchangeOnlineManagement
    Write-Host "ExchangeOnlineManagement Module Imported"}

Write-Host "Connecting to Exchange Online"
Connect-ExchangeOnline -ShowProgress $true
Write-Host "Connected to Exchange Online"

Write-Host "Obtaining Accepted Domain(s)"
$domains = Get-AcceptedDomain
Write-Host "Obtained Accepted Domain(s)"

Write-Host "Obtaining Mailboxes"
$mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties DisplayName, Identity, PrimarySmtpAddress, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward |
    Select-Object DisplayName, Identity, PrimarySmtpAddress, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward
Write-Host "Obtained Mailboxes"

Write-Host "Obtaining Transport Rules"
$TransportRules = Get-TransportRule -ResultSize Unlimited |
    Select-Object Identity, State, Priority, Description, AddToRecipients, CopyTo, BlindCopyTo, RedirectMessageTo, ActivationDate, ExpiryDate, CreatedBy, LastModifiedBy, Comments |
    Where-Object {($_.AddToRecipients -ne $null) -or ($_.CopyTo -ne $null) -or ($_.BlindCopyTo -ne $null) -or ($_.RedirectMessageTo -ne $null)}
Write-Host "Obtained Transport Rules"

Write-Host "Checking for Transport Rules with Forwarding to External Addresses and Exporting to Downloads"
foreach ($TransportRule in $TransportRules) {
 
    Write-Host "Checking rules for $($TransportRule.Identity)" -foregroundColor Green
 
        $recipients = @()
        $recipients = $TransportRule.AddToRecipients | Where-Object {$_ -ne $null}
        $recipients += $TransportRule.BlindCopyTo | Where-Object {$_ -ne $null}
        $recipients += $TransportRule.CopyTo | Where-Object {$_ -ne $null}
        $recipients += $TransportRule.RedirectMessageTo | Where-Object {$_ -ne $null}
     
        $externalRecipients = @()
 
        foreach ($recipient in $recipients) {
            $email = ($recipient.Trim("{","}"))
            $domain = ($email -split "@")[1]
 
            if ($domains.DomainName -notcontains $domain) {
                $externalRecipients += $email
            }    
        }
 
        if ($externalRecipients) {
            $extRecString = $externalRecipients -join ", "
            Write-Host "$($TransportRule.Name) forwards to $extRecString" -ForegroundColor Yellow
 
            $ruleHash = $null
            $ruleHash = [ordered]@{
                Identity           = $TransportRule.Identity
                State              = $TransportRule.State
                Description        = $TransportRule.Description
                ExternalRecipients = $extRecString
                ActivationDate     = $TransportRule.ActivationDate
                ExpiryDate         = $TransportRule.ExpiryDate
                CreatedBy          = $TransportRule.CreatedBy
                LastModifiedBy     = $TransportRule.LastModifiedBy
                Comments           = $TransportRule.Comments
            }
            $ruleObject = New-Object PSObject -Property $ruleHash
            $ruleObject | Export-Csv -Path "$ENV:userprofile\Downloads\Transport_Rules.csv" -NoTypeInformation -Append
        }
}
Write-Host "Obtained Transport Rules with Forwarding to External Addresses and Exported to Downloads"

Write-Host "Obtaining and Exporting Account Forwarding Rules for Users to Downloads"
foreach ($mailbox in $mailboxes) {
  
    $forwardingSMTPAddress = $null
    $ForwardingAddress = $null
    Write-Host "Checking SMTP Forwarding for $($mailbox.DisplayName) - $($mailbox.PrimarySmtpAddress)"
    $forwardingSMTPAddress = $mailbox.ForwardingSMTPAddress
    $ForwardingAddress = $mailbox.ForwardingAddress
    $externalRecipient = $null
    if ($forwardingSMTPAddress) {
        $email = ($forwardingSMTPAddress -split "SMTP:")[1]
        $domain = ($email -split "@")[1]
        if ($domains.DomainName -notcontains $domain) {
            $externalRecipient = $email
        }
  
        if ($externalRecipient) {
            Write-Host "$($mailbox.DisplayName) - $($mailbox.PrimarySmtpAddress) forwards to $externalRecipient" -ForegroundColor Yellow
  
            $forwardHash = $null
            $forwardHash = [ordered]@{
                DisplayName                = $mailbox.DisplayName
                PrimarySmtpAddress         = $mailbox.PrimarySmtpAddress
                ExternalRecipient          = $externalRecipient
                DeliverToMailboxandForward = $mailbox.DeliverToMailboxandForward
            }
            $ruleObject = New-Object PSObject -Property $forwardHash
            $ruleObject | Export-Csv -Path "$ENV:userprofile\Downloads\Account_Forwarding.csv" -NoTypeInformation -Append
        }
    }
    if ($ForwardingAddress) {
        $ForwardingAddressEmail = Get-EXORecipient -Identity $mailbox.ForwardingAddress| Select-Object PrimarySmtpAddress
        $email = ($ForwardingAddressEmail -split "@{PrimarySmtpAddress=")[1].Trim("}")
        $domain = ($email -split "@")[1]
        if ($domains.DomainName -notcontains $domain) {
            $externalRecipient = $email
        }
  
        if ($externalRecipient) {
            Write-Host "$($mailbox.DisplayName) - $($mailbox.PrimarySmtpAddress) forwards to $externalRecipient" -ForegroundColor Yellow
  
            $forwardHash = $null
            $forwardHash = [ordered]@{
                DisplayName                = $mailbox.DisplayName
                PrimarySmtpAddress         = $mailbox.PrimarySmtpAddress
                ExternalRecipient          = $externalRecipient
                DeliverToMailboxandForward = $mailbox.DeliverToMailboxandForward
            }
            $ruleObject = New-Object PSObject -Property $forwardHash
            $ruleObject | Export-Csv -Path "$ENV:userprofile\Downloads\Account_Forwarding.csv" -NoTypeInformation -Append
        }
    }
}
Write-Host "Obtained and Exported Account Forwarding Rules for Users to Downloads"

Write-Host "Obtaining and Exporting Inbox Forwarding Rules for Users to Downloads"
foreach ($mailbox in $mailboxes) {
 
    $forwardingRules = $null
    Write-Host "Checking rules for $($mailbox.DisplayName) - $($mailbox.PrimarySmtpAddress)" -foregroundColor Green
    $rules = Get-InboxRule -Mailbox $mailbox.PrimarySmtpAddress -IncludeHidden
     
    $forwardingRules = $rules | Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null)}
 
    foreach ($rule in $forwardingRules) {
        $recipients = @()
        $recipients = $rule.ForwardTo | Where-Object {$_ -ne $null}
        $recipients += $rule.ForwardAsAttachmentTo | Where-Object {$_ -ne $null}
        $recipients += $rule.RedirectTo | Where-Object {$_ -ne $null}
     
        $externalRecipients = @()

        foreach ($recipient in $recipients) {
            
            if ($recipient -match "SMTP") {
                $email = ($recipient -split "SMTP:")[1].Trim("]")}
            else {
                $mailcontactname = @()
                $mailcontactemail = @()
                $mailcontactname = ($recipient -split "`"")[1].Trim()
                $mailcontactemail = Get-EXORecipient -Identity $mailcontactname | Select-Object PrimarySmtpAddress
                $email = ($mailcontactemail -split "@{PrimarySmtpAddress=")[1].Trim("}")}
            
            $domain = ($email -split "@")[1]
            if ($domains.DomainName -notcontains $domain) {
                $externalRecipients += $email
            }    
        }
 
        if ($externalRecipients) {
            $extRecString = $externalRecipients -join ", "
            Write-Host "$($rule.Name) forwards to $extRecString" -ForegroundColor Yellow
 
            $ruleHash = $null
            $ruleHash = [ordered]@{
                DisplayName        = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                RuleId             = $rule.Identity
                RuleName           = $rule.Name
                RuleStatus         = $rule.Enabled
                RuleDescription    = $rule.Description
                ExternalRecipients = $extRecString
            }
            $ruleObject = New-Object PSObject -Property $ruleHash
            $ruleObject | Export-Csv -Path "$ENV:userprofile\Downloads\Inbox_Rule_Forwarding.csv" -NoTypeInformation -Append
        }
    }
}
Write-Host "Obtained and Exported Inbox Forwarding Rules for Users to Downloads"

Write-Host "Disconnecting from Exchange Online"
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
