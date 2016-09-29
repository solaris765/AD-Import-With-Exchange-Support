[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

#Designates the working Directory and navigates there.
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
cd $dir

#Checks if Script is being run in the Exchange Managment Shell or not.
#Switch "-ErrorAction Stop" is necessary for Try|Catch since Try|Catch doesn't catch non-terminating errors and would therefore fail without designating "Stop"
Try{$PowershellTest = Get-MailBox -ErrorAction Stop}
Catch{}

#Warning that mailboxes will not be created since the program was not run in Exchange management console.
If (!$PowershellTest)
{
   $Reply = [System.Windows.Forms.MessageBox]::Show("Mailboxs will not be created unless this is run from the Exchange Managment Console.
   Hit `"OK`" to continue or `"Cancel`" to exit." , "Warning" , 1)
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
    
    #If FirstName or OU are blank, throw error and exit.
    If (!$User.Firstname -or !$User.OU)
    {
        [System.Windows.Forms.MessageBox]::Show("There are errors in the user config file ""userlist.csv"", 
        please fix them and run this program again.`r`nRequired items are: Username, Firstname and OU","ERROR", 0)
        exit 1
    }
    #If Var $Detailedname is NULL set it to $UserFirstName + $UserLastName
    If (!$Detailedname){$Detailedname = $UserFirstname; If ($UserLastName){$Detailedname += " $UserLastName"}}
    #If $User.password doesn't exist, use "P@ssword" to initialize.
    If (!$User.Password){$Password = "P@ssword"}
    #If $User.username does not exist create it with Firstname and lastname.
    If (!$UserName){If ($UserLastName){$UserName=$UserFirstname.Substring(0,1) + $UserLastName} Else {$UserName=$UserFirstname}}

     <#
    Designates which OU the new users are to be created in.
    Var $pos designates the dilimiter used in the "DC" variables
    Var "OU" imported directly from OU column.
    Var "DC" is set by the left part of the $Domain variable delimited by "."
    Var "DC" is set by the right part of the $Domain variable delimited by "."
    If this variable is left blank User will be created in "Default $OU = "OU=Users,DC=ts,DC=local"
    #>
    $pos = $Domain.IndexOf(".")
    $OU = "OU=" + $User.OU + ",DC=" + $Domain.Substring(0, $pos) + ",DC=" + $Domain.Substring($pos+1)
    
    #If defined OU does not exist, create OU.
    If (![ADSI]::Exists("LDAP://$OU")){New-ADOrganizationalUnit $User.OU}


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
       If ($Policy = 0 -and !$Email)
       {
          [System.Windows.Forms.MessageBox]::Show("Email address is marked assigned but has been left blank, Please review $Detailedname's email settings.","ERROR",0)
          exit
       }

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
       If($policy = 0)
       {set-mailbox -Identity $Detailedname -emailaddresses $Email -emailaddresspolicy 0}
    }
    #Resets the $UserTest & $MailTest Vars to NULL for the next user
    $UserTest= $null
    $MailTest= $null

    #Writes a new-line between Users
    Write-Host "`r`n"
}   
