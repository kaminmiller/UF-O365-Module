<#
 .Synopsis
  Simplifies Functions needed to manipulate O365 Mailboxes

 .Description
  Contains the following functions:
  Add-MailboxDeligations
  Check-Object

 .Example
  To do - Add examples.  Figure it out on your own until then

  Check-Object -Object <Active Directory Opbject>

  OR

  Check-Object <Active Directory Object>

  Returns the object type of an AD Object, ex: user, group

    Check-Object -Object "kaminm"
    >> user

    Check-Object -Object "wcba-exchange-users"
    >> group

---------------------------------

  Add-MailboxDeligations -AdminCredentials <Credential Object> -O365User <gatorlink> -SendAs <user or group> 
        -SendOnBehalf <user or group> -FullAccess <user or group>

        If no admin credentials are supplied, you will be prompted for them.

        O365User - O365 enabled mailbox gatorlink
        SendAs - User or Group for SendAs permissions
        SendOnBehalf - User or Group for SendOnBehalf permissions
        FullAccess - User or Group for FullAccess permissions

  Add-MailboxDeligations -O365User "kaminm" -SendAs "w-sec-itsp-KaminMiller-SA" -FullAccess "w-sec-itsp-KaminMiller-FA"
        
#>

function dbg{
param(
    $message
)

Write-host -ForegroundColor Green -Object $message

}

function Add-MailboxDelegations{
    [CmdletBinding()]
    param(
        [Parameter()]
        $AdminCredentials,
        [Parameter(Mandatory)]
        $O365User,
        [Parameter()]
        $SendAs,
        [Parameter()]
        $SendOnBehalf,
        [Parameter()]
        $FullAccess
    )

 # dbg -message "Parameters"   

    # Check for Admin Credentials
    # dbg -message "Checking for Admin"
    if (!$AdminCredentials){
        Write-Output "Administrative Credentials Needed"
        Write-Output "Use w-adm-<username>@ufl.edu format"
        try {
            $AdminCredentials = Get-Credential -Message "Administrative Credentials Needed.  Format 'w-adm-<gatorlink>@ufl.edu'."
        } Catch {
            Write-Output "Error! Administrative Credentials Required!"
            throw "No Admin Credentials Supplied or Requested"        
        }
    }
 # dbg -message "O365 User Check"
    # Check if we have a user account to process
    if (!$O365User){
        Write-Output "Error! No User Account to process."
        throw "Error! No User Account to process."
    }
 # dbg -message "Runspace Check"
 # dbg -message (Get-PSSession | where ComputerName -eq "outlook.office365.com")
    # Check if Office 365 Runspace is available and imported
    try {
        if(Check-O365Session){
            # dbg -message "Session Established"
        }
        else {
            try {
                # dbg -message "No session. Starting Session"
                Start-O365Session
                # dbg -message "Started Session"
            }
            catch {
                throw "Unable to make connection to Office 365 for powershell modules"
            }
        }
    }
    catch {
        throw "Unable to create PSSession for Office 365 Management"
    }

    # If send-as user group provided, process the group
    # dbg -message "SendAs Check"
    if($SendAs){
       # dbg -message "SendAs True"         
       try{
            # Check if supplied object is user or group
            if((Check-Object($SendAs)) -eq 'user'){
                # dbg -message "SendAs is User"
                Add-O365RecipientPermission -Identity $O365User -Trustee $SendAs -AccessRights SendAs -Confirm:$false -Verbose 
                # dbg -message $tempvar
            }
            if((Check-Object($SendAs)) -eq 'group'){
                # dbg -message "SendAs is Group"
                $SendAsUsers = Get-ADGroupMember -Identity $SendAs -Credential $AdminCredentials -Recursive
                foreach($user in $SendAsUsers){
                    Add-O365RecipientPermission -Identity $O365User -Trustee $user.SamAccountName -AccessRights SendAs -Confirm:$false -Verbose 
                    # dbg -message $tempvar
                }
            }
       }
       catch
       {
           write-output "Error Adding SENDAS permission to $O365User"
           Write-Output $Error
       }
    }

    # If Send-On-Behalf user group provided, process the group
    # dbg -message "SendOnBehalf Check"
    if($SendOnBehalf){
        # dbg -message "SendOnBehalf is True"
        try{
            if((Check-Object($SendOnBehalf)) -eq 'user'){
             # dbg -message "SendOnBehalf is User"
                Set-O365Mailbox -Identity $O365User -GrantSendOnBehalfTo $SendOnBehalf -Confirm:$false 
                # dbg -message $tempvar
            }
            if((Check-Object($SendOnBehalf)) -eq 'group'){
             # dbg -message "SendOnbehalf is Group"
                $SendOnBehalfUsers = Get-ADGroupMember -Identity $SendOnBehalf -Credential $AdminCredentials -Recursive
                foreach($user in $SendOnBehalfUsers){
                    Set-O365Mailbox -Identity $O365User -GrantSendOnBehalfTo $user.SamAccountName -Confirm:$false 
                    # dbg -message $tempvar
                }
            }
        }
        catch {
            write-output "Error Adding SEND ON BEHALF permission to $O365User"
            Write-Output $Error
        }
    }

    # dbg -message "FullAccess Check"
    # If Full-Access user group provided, process the group
    if($FullAccess){
        # dbg -message "FullAccess True"
        try{
            
            if((Check-Object($FullAccess)) -eq 'user'){
            # dbg -message "FullAccess is User"
                Add-O365MailboxPermission -Identity $O365User -user $FullAccess -AccessRights FullAccess -Confirm:$false 
                # dbg -message $tempvar
            }
            if((Check-Object($FullAccess)) -eq 'group'){
             # dbg -message "FullAccess is Group"
                $FullAccessUsers = Get-ADGroupMember -Identity $FullAccess -Credential $AdminCredentials -Recursive
                foreach($user in $FullAccessUsers){
                    Add-O365MailboxPermission -Identity $O365User -user $user.SamAccountName -AccessRights FullAccess -Confirm:$false 
                    # dbg -message $tempvar
                }
            }
        }
        catch {
            write-output "Error Adding FULL ACCESS permission to $O365User"
            Write-Output $Error
        }

    }

    # dbg -message "Process Done"
}

function Check-O365Session{
    #Returns a [bool] indicating if there is an active session to Office365 

    # dbg -message "Checking Session"
    try {
        if(Get-PSSession | where ComputerName -eq "outlook.office365.com"){
        # dbg -message "Session is TRUE"
            return $true
        }
        else{
        # dbg -message "Session is FALSE"
            return $false
        }
    } 
    catch 
    {
        # Write-Host "Unable to detect Office 365 Session Status" -ForegroundColor Red
        return $false
    }

}

function Start-O365Session{

# dbg -message "STARTING SESSION"
    
    $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $AdminCredentials -Authentication Basic -AllowRedirection
    Import-PSSession $O365Session -DisableNameChecking -Prefix O365

}

function Start-OnPremSession{

# dbg -message "STARTING SESSION"
    
    $OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://exmbxprd01.ad.ufl.edu/Powershell/"
    Import-PSSession $OnPremSession -DisableNameChecking -Prefix OnPrem

}

function Check-Object{
    param(
    [Parameter(Mandatory)]$Object
    )
    try{
    $ObjectType = (Get-ADObject -filter {SamAccountName -eq $Object}).ObjectClass
    }
    catch {
        throw "Error Determining Object Type"
    }

    return $ObjectType
}

Export-ModuleMember -Function Check-Object -Alias *
Export-ModuleMember -Function Add-MailboxDelegations -Alias *