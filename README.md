# Exchange-Administration-Tool

his script is a summary of several PowerShell scripts that are useful for BULK tasks in Exchange Online.

## Prerequisites CSV

Create Multiple Shared Mailboxes
Create Multiple Equipment Mailboxes
Create Multiple Room Mailboxes

"Name","Alias","NewName","NewAlias"

--------------------------------------------

Set Shared Mailbox permissions

"Mailbox","UPN","Permission","AssignedTo","MailboxType"

--------------------------------------------

Create Multiple PST exports

"UserAccount","PSTName"

--------------------------------------------

Set External SMTP forwardings for multiple Users

"Name","SMTPold","SMTPnew"

--------------------------------------------

Set OOF Message for multiple Users

"UPN","ExternalMessage","InternalMessage","StartDate","EndDate"
For StartTime and EndTime use the following format: 7/15/2018 17:00:00

--------------------------------------------

Disable OWA, ActiveSync and MAPI for multiple Users

"Name","UPN"

--------------------------------------------

Convert Mail User to Remote Mailbox

This script requires the Active Directory and ExchangePowerShell modules to be installed on the machine running the script.
The modules it self will be loaded when the script is executed.

"Samaccountname","Mail"

## Author
Drago Petrovic
MSB365 Blog (https://www.msb365.blog)
