#requires -version 2
<#
.SYNOPSIS
  Script that will sent an email to users that are about to have an expired password.
.DESCRIPTION
  Script that will sent an email to users that are about to have an expired password.

  Will sent two times an email reminder that the AD password is about to expire.
  Script needs to be scheduled.

.PARAMETER <Parameter_Name>
  <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None>
.NOTES
  Version:        1.0
  Author:         Bart Tacken Client ICT Groep
  Creation Date:  05-2017
  Purpose/Change: Initial script development
.EXAMPLE
    not available
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string
Start-Transcript ('c:\windows\temp\' + $DateStr  + '_PasswordExpirationReminder.log') # Start logging 
Import-Module ActiveDirectory
Import-Module Send-Mail -DisableNameChecking # Import custom Send Mail module and ignore warnings
#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Convertto-HereString {
# Function for importing HTML to here string 
begin {$temp_h_string = ‘@”‘ + “`n”} # For creating here string: start with "@" followed by a new line
process {$temp_h_string += $_ + “`n”} # Process all content and store in string
end {
    $temp_h_string += ‘”@’ # End with "@" and a new line 
    iex $temp_h_string
    }
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
$SMTPserver= ”DC02.client-ict.local”
$SenderMail = "support@client.nl"
$EmailTemplate1 = "C:\Beheer\Scripts\HTMLtemplates\ClientPasswordExpirationMail.html"

$expireindays1 = 5
$expireindays2 = 2
$ErrorActionPreference = 'SilentlyContinue'

# Get AD accounts for all enabled and non password expired users (whose passwords can expire) 
$users = get-aduser -filter * -Properties enabled, passwordneverexpires, passwordexpired, emailaddress, passwordlastset | where { $_.name -like "*" } |
Where-Object {$_.Enabled -eq “True”} | Where-Object { $_.PasswordNeverExpires -eq $false } | Where-Object { $_.passwordexpired -eq $false }

foreach ($user in $users) {
    $Name = $user.name
    $emailaddress = $user.emailaddress
    
    # Get date where password is last changed
    $passwordSetDate = (get-aduser $user -properties passwordlastset | foreach { $_.PasswordLastSet })
    
    # Check for Fine Grained Password
    $PasswordPol = (Get-AduserResultantPasswordPolicy $user) 
    if (($PasswordPol) -ne $null) {
            $maxPasswordAge = ($PasswordPol).MaxPasswordAge
        }

    else {
        $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
    }

    $ExpiresOn = $Passwordsetdate + $maxPasswordAge # Calculate when password expires
    $Today = (get-date)
    $daystoexpire = (New-TimeSpan -Start $Today -End $ExpiresOn).Days    
    Write-Output ”$name, your password will expire in $daystoExpire days” # Test

    If (($daystoexpire -eq $expireindays1) -or ($daystoexpire -eq $expireindays2)) {
        [string]$SubjectStr = "Uw wachtwoord verloopt over $daystoExpire dagen"

        $HTMLtemplate = Get-Content $EmailTemplate1 | convertto-herestring
        ForEach ($word in $EmailTemplate1) { 
                $HTMLtemplate = $HTMLtemplate -replace "%PLACEHOLDER_UserName%", $Name
                $HTMLtemplate = $HTMLtemplate -replace "%PLACEHOLDER_DaysToExpire%", $daystoexpire
        }
        $HTMLtemplate | Out-File -FilePath "C:\Beheer\Scripts\HTMLtemplates\PassWordExpirationMailUser.html" #Test
        Send-Mail -SMTPserverStr $SMTPserver -SenderMailAddStr $SenderMail -RecepientMailAddStr $EmailAddress -SubjectStr $SubjectStr -BodyStr $HTMLtemplate -Attach1Str "C:\Beheer\Scripts\HTMLtemplates\Werkinstructie - Wachtwoord wijzigen Client.docx" # to contact user
    }
} # End ForEach
