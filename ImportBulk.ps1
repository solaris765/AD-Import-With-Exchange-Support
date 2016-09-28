#Var $CompanyFolder is the OU container all users will be created in.
#Set this variable before running the code.
$COMPANY_FOLDER = "GrowOffice"

#Designates the working Directory and navigates there.
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
cd $dir

#Checks if Script is being run in the Exchange Managment Shell or not.
#Switch "-ErrorAction Stop" is necessary for Try|Catch since Try|Catch doesn't catch non-terminating errors and would therefore fail without designating "Stop"
Try{$PowershellTest = Get-MailBox -ErrorAction Stop}
Catch{}

If (!$PowershellTest)
{
   $Reply = [System.Windows.Forms.MessageBox]::Show("Mailboxs will not be created unless this is run from the Exchange Managment Console" , "Warning" , 1)
   If ($Reply -eq "Cancel")
   {exit}
}

#Imports the Active-Directory Snap-In.
Import-Module ActiveDirectory 

#Imports csv with Delimiter "-Delimiter" from the working directory.
$Users = Import-Csv -Delimiter "," -Path ".\userlist.csv" 

<#
 Imports each user from the Csv into AD and creates their mailboxes
   If the user already exists it is ignored
   If the mailbox already exists it is ignored 
#>
foreach ($User in $Users)  
{   
    <#
    Var $UserName & $SAM are imported directly from the "Username" colomn
    Var $Domain is set by the domain of the local computer
    Var $Password is imported directly from the "password" column
    Var $Detailedname is imported directly from the "Name" column
    Var $UserFirstname is imported directly from the "FirstName" column
    Var $UserLastName is imported directly from the "LastName" column
    Var $Email is imported from the "EmailAddress" column
    Var $Policy is imported from the "EmailPolicy" column
    #>
    $UserName = $User.Username
    $Domain = (Get-WmiObject Win32_ComputerSystem).Domain
    $Password = $User.Password 
    $Detailedname = $User.Name
    $UserFirstname = $User.FirstName 
    $UserLastName = $User.LastName
    $SAM =  $User.username 
    $Email = $User.EmailAddress
    $Policy = $User.EmailPolicy
     <#
    Designates which OU the new users are to be created in.
    Var $pos designates the dilimiter used in the "DC" variables
    Var "OU" is set by Var $CompanyFolder
    Var "DC" is set by the left part of the $Domain variable delimited by "."
    Var "DC" is set by the right part of the $Domain variable delimited by "."
    If this variable is left blank User will be created in "Default $OU = "OU=Users,DC=gre,DC=local"
    #>
    $pos = $Domain.IndexOf(".")
    $OU = "OU=$COMPANY_FOLDER,DC=" + $Domain.Substring(0, $pos) + ",DC=" + $Domain.Substring($pos+1)
    
    #Checks to see if a user already exists with this username
    #Var $UserTest is used to store the output if a user exists
    Try {$UserTest = get-aduser $UserName} 
    Catch {}

    <#
    If user exists throw warning and continue
    Else creates user and outputs a green success message
    #>
    If($UserTest)
    {Write-Warning "A user account with username $Username already exists in Active Directory."}
    Else
    {
        New-ADUser -sAMAccountName $UserName -UserPrincipalName "$UserName@$Domain" -Name $Detailedname -DisplayName $Detailedname -GivenName $UserFirstName -Surname $UserLastName -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true -Path $OU
        Write-Host "User account $UserName, has been created." -foregroundcolor "Green"
    }
    
    If ($PowershellTest)
    {
       #Checks to see if a user mailbox exists for this user
       #Var $MailTest is used to store the output if a user mailbox exists
       #Switch "-ErrorAction Stop" is necessary for Try|Catch since Try|Catch doesn't catch non-terminating errors and would therefore fail without designating "Stop"
       Try{$MailTest = get-mailbox $UserName -ErrorAction Stop}
       Catch{}
       
       <#
       If the mailbox exists throw warning and continue
       Else enables the user mailbox with the default mailbox policy
       #>
       if($MailTest)
       {Write-Warning "User $Username already has a mailbox enabled."}
       Else
       {enable-mailbox $UserName}
       
       #If mailbox policy is set to FALSE sets the users email address to what has been designated in the $Email variable
       If($policy = "FALSE")
       {set-mailbox -Identity $Detailedname -emailaddresses $Email -emailaddresspolicy 0}
    }
    #Resets the $UserTest & $MailTest Vars to NULL for the next user
    $UserTest= $null
    $MailTest= $null

    #Writes a new-line between Users
    Write-Host "`r`n"
}   
