#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2022/04/13 
       Created by:         Drago Petrovic & Daniel Barrera
       Organization:       MSB365.blog & vordis technologies
       Filename:           Handle-BasicVoicePolicies.ps1
       Current version:    V1.01     

       ==============================================================================
       .DESCRIPTION
       Assign and manage Phone number, voice routing policy and Dial Plan to users (Bulk)            
       
       .NOTES
       This script can be executed without prior customisation.
       This script is used to manage Teams usres and assign the phone number, voice routing policy and Dial Plan (bulk) with PowerShell
       .EXAMPLE
       .\Handle-BasicVoicePolicies.ps1
             
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
             V1.00, 2022/06/14 - DrPe - Initial version
             V1.01, 2022/08/23 - KainDevOps - Assigning Voice Routing Policies, Phone number and Dial Plan policy
			 
--- keep it simple, but significant ---
#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Teams Voice - Enable Basic Policies in Direct Routing mode"
$RKEY = "MSB365_Teams_Voice_PN_VR_DP_Policies"
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
Copyright (c) 2022 Drago Petrovic & Daniel Barrera\par
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
\pard\widctlpar\fi-360\li360\qj 7.\tab A partial script documentation can be found on the msb65 blog, but just a few part :o)\par
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
write-host ""
write-host ""
write-host " ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ " -ForegroundColor Yellow
write-host " ░░┌┐░░░░░░┌───┐░░░░░░░░░┌┐░┌──┐░░░ " -ForegroundColor Yellow
write-host " ░░││░░░░░░└┐┌┐│░░░░░░░░░││░│┌┐│░░░ " -ForegroundColor Yellow
write-host " ░░│└─┬┐░┌┐░│││├──┬─┐┌┬──┤│░│└┘└┐░░ " -ForegroundColor Yellow
write-host " ░░│┌┐││░││░││││┌┐│┌┐┼┤│─┤│░│┌─┐│░░ " -ForegroundColor Yellow
write-host " ░░│└┘│└─┘│┌┘└┘│┌┐││││││─┤└┐│└─┘│░░ " -ForegroundColor Yellow
write-host " ░░└──┴─┐┌┘└───┴┘└┴┘└┴┴──┴─┘└───┘░░ " -ForegroundColor Yellow
write-host " ░░░░░┌─┘│░░░░░░░░░░░░░░░░░░░░░░░░░ " -ForegroundColor Yellow
write-host " ░░░░░└──┘░░░░░░░░░░░░░░░░░░░░░░░░░ " -ForegroundColor Yellow
Start-Sleep -s 1
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
write-host ""
###############################################################################











$selection3 =  Read-Host "Would you like to connect to Microsoft Teams using this Script?? [Y] for yes / [N] for no or already connected." 
switch ($selection3)
       { 'Y' {
            # Load Microsoft Teams PowerShell Module
            write-host "Connectig Microsoft Teams" -ForegroundColor Magenta
            Start-Sleep -s 5
                if (Get-Module -ListAvailable -Name MicrosoftTeams) {
                    Write-Host "Microsoft Teams Module Already Installed" -ForegroundColor Green
                    start-sleep -s 2
                    Write-Host "Checking for Module update..." -ForegroundColor cyan
                    Update-Module MicrosoftTeams
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                } 
            else {
                    Write-Host "MicrosoftTeams Module Not Installed. Installing........." -ForegroundColor Red
                    Install-Module -Name MicrosoftTeams -AllowClobber -Force
                    Write-Host "MicrosoftTeams Module Installed" -ForegroundColor Green
                    start-sleep -s 2
                    write-host " - Please enter the credentials..." -ForegroundColor Yellow 
                }
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams
            Start-Sleep -s 5
     } 'N' {
         
     }
     
     }
##############################################################################################################
# Get Location ID
			write-host "Gettering Tenant location ID..." -ForegroundColor Cyan
			start-sleep -s 2
			$Lid = Get-CsOnlineLisLocation | Sort-Object LocationID | select-object -ExpandProperty LocationID
			write-host "Tenant LocationID is: $Lid" -ForegroundColor White -BackgroundColor Black
			Start-Sleep -s 2

			# Getting CSV Information
			write-host "Please select and import the CSV File from your device:" -ForegroundColor Cyan
			Write-Host ""
			Write-Host ""
			Start-Sleep -s 4
			Write-Host "**************************************************************************************************************************************" -ForegroundColor Yellow -BackgroundColor Black
			Write-Host "* NOTE!                                                                                                                              *" -ForegroundColor Yellow -BackgroundColor Black                                                                          
			Write-Host "* The following information are needed in the CSV file:                                                                              *"
			Write-Host "* "UserPrincipalName","DisplayName","CanonicalPhoneNumber","ExtensionNumber","CallingIDpolicy","VoiceRoutingPolicy“,"DialPlanPolicy" *" -ForegroundColor Gray -BackgroundColor Black
			Write-Host "**************************************************************************************************************************************" -ForegroundColor Yellow -BackgroundColor Black
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
# Configuring Teams Calling ID Policy
#$company = $(Write-Host "Enter the Company name. Example: " -NoNewLine) + $(Write-Host """" -NoNewline) +$(Write-Host "Contoso Ltd" -ForegroundColor Yellow -NoNewline; Read-Host """")
#start-sleep -s 3			
#write-host "Setting the calling ID Policies..." -ForegroundColor cyan 
#foreach($user in $users)
#			{
#				try
#				{
#					Grant-CsCallingLineIdentity -Identity $user.UserPrincipalName -PolicyName $user.PolicyName -ErrorAction Stop ###
#					Write-Host "Voice calling ID policy $($user.PolicyName) for the users $($user.DisplayName) set!" -ForegroundColor Green
#					Start-Sleep -s 1
#				}
#				catch
#				{
#					Write-Host "Could not set the policy $($user.PolicyName) for user $($user.DisplayName) " + $_.Exception -ForegroundColor Red 
#				}
#				
#			}
            start-sleep -s 3


##############################################################################################################
# Configuring Teams Phone Number for Direct Routing
#$company = $(Write-Host "Enter the Company name. Example: " -NoNewLine) + $(Write-Host """" -NoNewline) +$(Write-Host "Contoso Ltd" -ForegroundColor Yellow -NoNewline; Read-Host """")
start-sleep -s 3			
write-host "Setting users phone number..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					# Set-CsPhoneNumberAssignment -Identity francesco.amato@algeco.com -PhoneNumber +390577592093;ext=20407 -PhoneNumberType DirectRouting
                    Set-CsPhoneNumberAssignment -Identity $user.UserPrincipalName -PhoneNumber $user.CanonicalPhoneNumber -PhoneNumberType DirectRouting -ErrorAction Stop ###
					Write-Host "Number $($user.CanonicalPhoneNumber) for the user $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the number $($user.CanonicalPhoneNumber) for user $($user.DisplayName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
            start-sleep -s 3


##############################################################################################################
# Configuring Dial Plan Policy
#$company = $(Write-Host "Enter the Company name. Example: " -NoNewLine) + $(Write-Host """" -NoNewline) +$(Write-Host "Contoso Ltd" -ForegroundColor Yellow -NoNewline; Read-Host """")
start-sleep -s 3			
write-host "Setting Dial Plan policy..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					# Grant-CsTenantDialPlan -Identity francesco.amato@algeco.com -PolicyName IT
                    Grant-CsTenantDialPlan -Identity $user.UserPrincipalName -PolicyName $user.DialPlanPolicy -ErrorAction Stop ###
					Write-Host "Dial Plan Policy $($user.DialPlanPolicy) for the user $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the Dial Plan policy $($user.DialPlanPolicy) for user $($user.DisplayName) " + $_.Exception -ForegroundColor Red 
				}
				
			}
            start-sleep -s 3
            

###############################################################################################################
# Configuring Voice Routing Policy
#$company = $(Write-Host "Enter the Company name. Example: " -NoNewLine) + $(Write-Host """" -NoNewline) +$(Write-Host "Contoso Ltd" -ForegroundColor Yellow -NoNewline; Read-Host """")
start-sleep -s 3			
write-host "Setting Voice Routing policy..." -ForegroundColor cyan 
foreach($user in $users)
			{
				try
				{
					# Grant-CsOnlineVoiceRoutingPolicy -Identity francesco.amato@algeco.com -PolicyName IT-All-International
                    Grant-CsOnlineVoiceRoutingPolicy -Identity $user.UserPrincipalName -PolicyName $user.VoiceRoutingPolicy -ErrorAction Stop ###
					Write-Host "Voice Routing Policy $($user.VoiceRoutingPolicy) for the user $($user.DisplayName) set!" -ForegroundColor Green
					Start-Sleep -s 1
				}
				catch
				{
					Write-Host "Could not set the Voice Rouing policy $($user.VoiceRoutingPolicy) for user $($user.DisplayName) " + $_.Exception -ForegroundColor Red 
				}
				
			}