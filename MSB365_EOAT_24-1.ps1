#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2024/02/16 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           MSB365_EOAT_24-1.ps1
       Current version:    V1.0     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       Exchange online Administration Tool
       This script is a summary of several PowerShell scripts that are useful for BULK tasks in Exchange Online.


       .NOTES
       This script can be executed without prior customisation.


       .EXAMPLE
       .\MSB365_EOAT_24-1.ps1 
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2024/02/16 - DrPe - Initial version

             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "MSB365 - Exchange online Administration Tool"
$RKEY = "MSB365_EOAT_24-1"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2024 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
write-host ""
###############################################################################

Write-Host "Would you like to connect to Microsoft Exchange online using this Script?? "
$selection3 =  Read-Host "[Y] for yes / [N] for no or already connected." 
switch ($selection3)
       { 'Y' {
            # Load Microsoft Teams PowerShell Module
            write-host "Connectig Exchange online" -ForegroundColor Magenta
            Start-Sleep -s 5
                if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                    Write-Host "Exchange online Module Already Installed" -ForegroundColor Green
                    start-sleep -s 2
                    Write-Host "Checking for Module update..." -ForegroundColor cyan
                    Update-Module ExchangeOnlineManagement
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                } 
            else {
                    Write-Host "Exchange online Module Not Installed. Installing........." -ForegroundColor Red
                    Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
                    Write-Host "Exchange online Module Installed" -ForegroundColor Green
                    start-sleep -s 2
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                }
            Import-Module ExchangeOnline
            Connect-ExchangeOnline
            Start-Sleep -s 5
     } 'N' {
         
     }
     
     }
##############################################################################################################
# Define the menu options
$MenuOptions = @(
    "0. Script information"
    "1. Create Multiple Equipment Mailboxes"
    "2. Create Multiple Room Mailboxes"
    "3. Create Multiple PST exports"
    "4. Create Multiple Shared Mailboxes"
    "5. Set External SMTP forwardings for multiple Users"
    "6. Set OOF Message for multiple Users"
    "7. Set Shared Mailbox permissions"
    "8. Disable OWA, ActiveSync and MAPI for multiple Users"
	"9. Convert Mail User to Remote Mailbox"
    "99. Exit the menu"
)

# Define the menu function
function Show-Menu {
    # Clear the screen
    Clear-Host
    # Display the menu title
    write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
write-host ""
write-host ""
write-host ""
    Write-Host "Welcome to the powershell script menu"
    Write-Host "Please select an option from the list below:"
    Write-Host "-------------------------------------------"

    foreach ($option in $MenuOptions) {
        Write-Host $option
    }
}

# Define the script actions
function Invoke-ScriptAction {
    # Get the user input
    $UserInput = Read-Host "Enter your choice"
    # Switch based on the user input
    switch ($UserInput) {
        # Script information
        0 {
            Write-Host "Thank you for using this PowerShell tool for the administration of BULK Exchange online tasks." -ForegroundColor Cyan
Write-Host "A CSV file is required for each of the options listed in the menu. These look like this depending on the application:" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Create Multiple Shared Mailboxes" -ForegroundColor Yellow
Write-Host "Create Multiple Equipment Mailboxes " -ForegroundColor Yellow
Write-Host "Create Multiple Room Mailboxes" -ForegroundColor Yellow
Write-Host ""
Write-Host '"Name","Alias","NewName","NewAlias"' -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Set Shared Mailbox permissions" -ForegroundColor Yellow
Write-Host ""
Write-Host '"Mailbox","UPN","Permission","AssignedTo","MailboxType"' -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Create Multiple PST exports" -ForegroundColor Yellow
Write-Host ""
Write-Host '"UserAccount","PSTName"' -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Set External SMTP forwardings for multiple Users" -ForegroundColor Yellow
Write-Host ""
Write-Host '"Name","SMTPold","SMTPnew"' -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Set OOF Message for multiple Users" -ForegroundColor Yellow
Write-Host ""
Write-Host '"UPN","ExternalMessage","InternalMessage","StartDate","EndDate"' -ForegroundColor Gray
Write-Host "For StartTime and EndTime use the following format: 7/15/2018 17:00:00" -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Disable OWA, ActiveSync and MAPI for multiple Users" -ForegroundColor Yellow
Write-Host ""
Write-Host '"Name","UPN"' -ForegroundColor Gray
Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host "Convert Mail User to Remote Mailbox" -ForegroundColor Yellow
Write-Host ""
Write-Host "This script requires the Active Directory and ExchangePowerShell modules to be installed on the machine running the script." -ForegroundColor Yellow
Write-Host "The modules it self will be loaded when the script is executed." -ForegroundColor Yellow
Write-Host ""
Write-Host '"samaccountname","Mail"' -ForegroundColor Gray
Write-Host ""
            pause
        }
        # Create Multiple Equipment Mailboxes
        1 {
            ##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"Name","Alias","NewName","NewAlias"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring Mailboxes

start-sleep -s 3			
write-host "Creating the Room Mailboxes..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					New-Mailbox -Name $($user.NewName) -Equipment
					Set-CalendarProcessing $($user.NewName) -AutomateProcessing AutoAccept
					Write-Host "Equipment Mailbox $($user.NewName) created!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not create the Equipment Mailbox $($user.NewName) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Create Multiple Room Mailboxes
        2 {
            ##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"Name","Alias","NewName","NewAlias"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring Mailboxes

start-sleep -s 3			
write-host "Creating the Room Mailboxes..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					New-Mailbox -Name $($user.NewName) -Room
					Set-CalendarProcessing $($user.NewName) -AutomateProcessing AutoAccept
					Write-Host "Room Mailbox $($user.NewName) created!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not create the Room Mailbox $($user.NewName) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Create Multiple PST exports
        3 {
            ###############################################################################



$Server = Read-Host "Enter the Exchange Server Name! eg. SRVEXC01"
$folderPath = "C:\MDM\PSTexport"
if (Test-Path $folderPath) {
	Write-Host< "Folder already exists" -ForegroundColor Green
} else {
	Write-Host "Folder does not exist, creating folder..." -ForegroundColor Cyan
	New-Item -Path $folderPath -ItemType Directory
	Write-Host "Folder path: $folderPath created successfully" -ForegroundColor Green
}

# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"UserAccount","PSTName"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Creating PST Files
$Server = Read-Host "Enter the Exchange Server Name! eg. SRVEXC01"
$folderPath = "C:\PSTexport"
if (Test-Path $folderPath) {
    Write-Host "Folder already exists" -ForegroundColor Green
} else {
    Write-Host "Folder does not exist, creating folder..." -ForegroundColor Cyan
    New-Item -Path $folderPath -ItemType Directory
    Write-Host "Folder path: $folderPath created successfully" -ForegroundColor Green
}

$networkPath = Convert-Path $folderPath
Start-Sleep -s 1
Write-Host "Please note that this will only work if the directory is shared over the network. " -ForegroundColor Yellow
Write-Host "If the directory is not shared, you’ll need to share it first before you can access it using a UNC path. " -ForegroundColor Yellow
Write-Host "Also, the resulting network path will depend on the network configuration of your machine. " -ForegroundColor Yellow
Write-Host "If your machine is not part of a domain, the network path will likely start with \\<MachineName>\.... " -ForegroundColor Yellow
Write-Host "If your machine is part of a domain, the network path will likely start with \\<DomainName>\<MachineName>\...." -ForegroundColor Yellow
Start-Sleep -s 4


start-sleep -s 3			
write-host "Changing PST file for selected users..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					New-MailboxExportRequest -Mailbox $($user.UserAccount) -FilePath "\\$Server\$folderPath\$($user.PSTName).pst" -ErrorAction Stop	
					Write-Host "PST file for $($user.UserAccount) has been created successfully" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Create Multiple Shared Mailboxes
        4 {
            ##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host 'Name","Alias"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring Mailboxes

start-sleep -s 3			
write-host "Creating shared Mailboxes..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					New-Mailbox -Shared -Name $($user.Name) -DisplayName $($user.Name) -Alias $($user.Alias)
					
					Write-Host "Shared Mailbox $($user.NewName) created!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not create the Shared Mailbox $($user.NewName) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks Done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Set External SMTP forwardings for multiple Users
        5 {
            ###############################################################################


$selection3 =  Read-Host "Would you like to connect to Microsoft Exchange online using this Script?? [Y] for yes / [N] for no or already connected." 
switch ($selection3)
       { 'Y' {
            # Load Microsoft Exchange online PowerShell Module
            write-host "Connectig Exchange online" -ForegroundColor Magenta
            Start-Sleep -s 5
                if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                    Write-Host "Exchange online Module Already Installed" -ForegroundColor Green
                    start-sleep -s 2
                    Write-Host "Checking for Module update..." -ForegroundColor cyan
                    Update-Module ExchangeOnlineManagement
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                } 
            else {
                    Write-Host "Exchange online Module Not Installed. Installing........." -ForegroundColor Red
                    Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
                    Write-Host "Exchange online Module Installed" -ForegroundColor Green
                    start-sleep -s 2
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                }
            Import-Module ExchangeOnline
            Connect-ExchangeOnline
            Start-Sleep -s 5
     } 'N' {
         
     }
     
     }
##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"Name","SMTPold","SMTPnew"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring external forwardings

start-sleep -s 3			
write-host "Creating the external forwardings..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					Set-Mailbox -Identity $user.SMTPold -ForwardingSmtpAddress $user.SMTPnew -DeliverToMailboxAndForward $True
					Write-Host "External forwarding for the User $($user.Name) created!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the external forwarding for the User $($user.Name) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Set OOF Message for multiple Users
        6 {
            ###############################################################################


$selection3 =  Read-Host "Would you like to connect to Microsoft Exchange online using this Script?? [Y] for yes / [N] for no or already connected." 
switch ($selection3)
       { 'Y' {
            # Load Microsoft Exchange online PowerShell Module
            write-host "Connectig Exchange online" -ForegroundColor Magenta
            Start-Sleep -s 5
                if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                    Write-Host "Exchange online Module Already Installed" -ForegroundColor Green
                    start-sleep -s 2
                    Write-Host "Checking for Module update..." -ForegroundColor cyan
                    Update-Module ExchangeOnlineManagement
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                } 
            else {
                    Write-Host "Exchange online Module Not Installed. Installing........." -ForegroundColor Red
                    Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
                    Write-Host "Exchange online Module Installed" -ForegroundColor Green
                    start-sleep -s 2
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                }
            Import-Module ExchangeOnline
            Connect-ExchangeOnline
            Start-Sleep -s 2
     } 'N' {
         
     }
     
     }
##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 2
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"User","ExternalMessage","InternalMessage","StartDate","EndDate"' -ForegroundColor Gray -BackgroundColor Black
            Start-Sleep -s 2
            Write-Host "NOTE: For StartTime and EndTime use the following format: " -ForegroundColor Gray -BackgroundColor Black -NoNewline
            Write-Host "7/15/2018 17:00:00" -ForegroundColor Magenta -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 3
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring OOF

start-sleep -s 3			
write-host "Creating Out Of Office entries..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					Set-MailboxAutoReplyConfiguration -Identity $user.UPN -AutoReplyState Scheduled -StartTime $user.StartDate -EndTime $user.EndDate -ExternalMessage $user.ExternalMessage -InternalMessage $user.InternalMessage
					Write-Host "Out Of Office message for the User $($user.UPN) created!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the Out Of Office for the User $($user.UPN) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Set Shared Mailbox permissions
        7 {
            ###############################################################################


$selection3 =  Read-Host "Would you like to connect to Microsoft Exchange online using this Script?? [Y] for yes / [N] for no or already connected." 
switch ($selection3)
       { 'Y' {
            # Load Microsoft Teams PowerShell Module
            write-host "Connectig Exchange online" -ForegroundColor Magenta
            Start-Sleep -s 5
                if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                    Write-Host "Exchange online Module Already Installed" -ForegroundColor Green
                    start-sleep -s 2
                    Write-Host "Checking for Module update..." -ForegroundColor cyan
                    Update-Module ExchangeOnlineManagement
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                } 
            else {
                    Write-Host "Exchange online Module Not Installed. Installing........." -ForegroundColor Red
                    Install-Module -Name ExchangeOnlineManagement -AllowClobber -Force
                    Write-Host "Exchange online Module Installed" -ForegroundColor Green
                    start-sleep -s 2
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                }
            Import-Module ExchangeOnline
            Connect-ExchangeOnline
            Start-Sleep -s 5
     } 'N' {
         
     }
     
     }
##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"Mailbox","UPN","Permission","AssignedTo","MailboxType"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Configuring permissions

start-sleep -s 3			
write-host "Setting the permissions..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					Add-MailboxPermission $($user.Mailbox) -User $($user.AssignedTo) -AccessRights $($user.Permission) -InheritanceType All
					Write-Host "Permission for the Mailbox $PolicyID1 and the users $($user.UserPrincipalName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the permission $($user.Permission) for user $($user.AssignedTo) " + $_.Exception -ForegroundColor Red 
					Write-Host "Trying again, please hold..." -ForegroundColor yellow
					Set-Mailbox $($user.Mailbox) -GrantSendOnBehalfTo @{add=$($user.AssignedTo)}
				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause
        }
        # Disable Services for Mailboxes
        8 {
            ##############################################################################################################


			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"Name","SMTPold"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Disabling Mailbox Services

start-sleep -s 3			
write-host "Disabling Mailbox Services..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
                    Set-CASMailbox -Identity $user.SMTPold -MAPIEnabled $False
                    Write-Host "MAPI for the User $($user.Name) disabled!" -ForegroundColor Green
                    Set-CASMailbox -Identity $user.SMTPold -OWAEnabled $False
                    Write-Host "Outlook Web Service (OWA) for the User $($user.Name) disabled!" -ForegroundColor Green
                    Set-CASMailbox -Identity $user.SMTPold -ActiveSyncEnabled $False
					Write-Host "ActiveSync for the User $($user.Name) disabled!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not disable service for the User $($user.Name) " + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black
pause

        }

		# Convert Mail User to Remote Mailbox
        9 {
            ##############################################################################################################
			Write-Host "Important: This script requires the Active Directory and ExchangePowerShell modules to be installed on the machine running the script." -ForegroundColor Yellow
			Start-Sleep -s 4

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "NOTE!" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "The following information are needed in the CSV file:" -ForegroundColor White -BackgroundColor black
			Write-Host '"samaccountname","Mail"' -ForegroundColor Gray -BackgroundColor Black
			Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep -s 6
			$File = New-Object System.Windows.Forms.OpenFileDialog
			$null = $File.ShowDialog()
			$FilePath = $File.FileName
			$users = Import-Csv $FilePath
			Start-Sleep -s 3
			Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
			$users | ft
			Start-Sleep -s 3
			do
			{
				$selection = Read-Host "Are the data correct? - Choose between [Y] and [N]"
				switch ($selection)
				{
					'y' {
						
					} 'n' {
						$File = New-Object System.Windows.Forms.OpenFileDialog
						$null = $File.ShowDialog()
						$FilePath = $File.FileName
						$users = Import-Csv $FilePath
						Start-Sleep -s 3
						Write-Host "*****************************************************" -ForegroundColor Yellow -BackgroundColor Black
						Write-Host "The following data is imported from CSV:" -ForegroundColor Cyan
						$users | ft
						Start-Sleep -s 3
					}
				}
			}
			until ($selection -eq "y")
			

			start-sleep -s 3
##############################################################################################################
# Check if Active Directory PowerShell module is imported
if (Get-Module -Name activedirectory -ErrorAction SilentlyContinue) {
	Write-Host "Already connected to Active Directory." -ForegroundColor Green
} else {
	Write-Host "Importing Active Directory PowerShell module..." -ForegroundColor Cyan

}
Start-Sleep -s 2
Import-Module activedirectory
# Check if Exchange PowerShell module is imported
if (Get-Module -Name ExchangeManagementShell -ErrorAction SilentlyContinue) {
	Write-Host "Already connected to Exchange." -ForegroundColor Green
} else {
	Write-Host "Importing Exchange PowerShell module..." -ForegroundColor Cyan
	Import-Module ExchangeManagementShell
}
Start-Sleep -s 2

# Converting Mail Users

start-sleep -s 3			
write-host "Converting Mail Users..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					Set-MailUser -Identity $user.samaccountname -ExternalEmailAddress $user.Mail
                    Write-Host "External E-Mail Address for the User $($user.samaccountname) set!" -ForegroundColor Green
                    Start-Sleep -s 3
                    Set-ADuser -Identity $user.samaccountname -Replace @{msExchRecipientTypeDetails="2147483648"}
                    Write-Host "Modification of the Exchange Recipient Type for the User $($user.samaccountname) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Coverting the User $($user.Name) was not working." + $_.Exception -ForegroundColor Red 

				}
				
			}
start-sleep -s 3
Write-Host "Tasks done!" -ForegroundColor Green -BackgroundColor Black

            pause
        }
        # Exit the menu
        99 {
            Write-Host "Thank you for using the Exchange online Administration Tool." -ForegroundColor Cyan -NoNewline; Write-Host " Goodbye!" -ForegroundColor Green
            break
        }
        # Invalid input
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
        }
    }
}

# Loop until the user exits the menu
do {
    # Show the menu
    Show-Menu
    # Invoke the script action
    Invoke-ScriptAction
    # Pause until the user presses a key
    Write-Host "Press any key to continue ..."
    $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
} while ($UserInput -ne 4)
