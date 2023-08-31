#Connect to Exchange Online
Connect-ExchangeOnline -ShowProgress $true

#Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

# Get information on Mailbox
Get-ExoMailbox -Identity "name" -Properties DisplayName, RecipientTypeDetails, ForwardingSMTPAddress, ForwardingAddress, DeliverToMailboxandForward

# Get Mailbox Statistics
Get-MailboxStatistics -Identity "names" | Select DisplayName, ItemCount, TotalItemSize, DeletedItemCount, TotalDeletedItemSize

# Check SMTP Forwarding for a Particular User
Get-ExoMailbox -Identity "name" -Properties DisplayName, RecipientTypeDetails, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward |
    select DisplayName, RecipientTypeDetails, ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxandForward |
    where {($_.ForwardingSMTPAddress -ne $null) -or ($_.ForwardingAddress -ne $null)}

# Find Inbox Rules for user
Get-InboxRule -Mailbox "name" -IncludeHidden  | 
    Select Identity, Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage |
    Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null)}

# Find Transport Rules with Forwarding
Get-TransportRule -ResultSize Unlimited |
    select Identity, State, Priority, Description, AddToRecipients, CopyTo, BlindCopyTo, RedirectMessageTo, ActivationDate, ExpiryDate, CreatedBy, LastModifiedBy, Comments |
    where {($_.AddToRecipients -ne $null) -or ($_.CopyTo -ne $null) -or ($_.BlindCopyTo -ne $null) -or ($_.RedirectMessageTo -ne $null)}

# Remove SMTP Auto-Forwarding on an Account
Set-Mailbox "name" -ForwardingAddress $NULL -ForwardingSmtpAddress $NULL -DeliverToMailboxAndForward $true

# Remove Individual Inbox Rule
Remove-InboxRule -Mailbox "name" -Identity "name of rule" -Confirm:$false

# Remove All Inbox Rules for a User
Get-InboxRule -Mailbox "name" |
    Remove-InboxRule -Confirm:$false

# Remove All Inbox Rules with Auto-forwarding for a User
Get-InboxRule -Mailbox "name" -IncludeHidden | 
    Select Identity, Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage |
    Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectTo -ne $null)} |
    Remove-InboxRule -Confirm:$false

# Get Roles Assigned to User
Get-ManagementRoleAssignment -RoleAssignee "name" -Delegating $false |
    Format-Table -Auto Role,RoleAssigneeName,RoleAssigneeType

# Get Roles Required for a Command
Get-ManagementRole -Cmdlet "command"

# Get All Role Groups
Get-RoleGroup

# Get Details for a Particular Role Group
Get-RoleGroup "name" | Format-List

# Get Group Members for a Role
Get-RoleGroupMember "name"

# Get Information on Mail Contact
Get-EXORecipient -Identity "name" -ResultSize Unlimited

#Create Transport Rule
New-TransportRule -Name "REQ0000000" -SentTo employee@domain.com -From email -CopyTo externalemail -Comments "REQ0000000"
