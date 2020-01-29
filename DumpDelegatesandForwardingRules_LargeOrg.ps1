function messageBox {
    # Displays pop-up message windows.
    param($boxMessage, $boxTitle, $boxIcon)
    Add-Type -AssemblyName System.Windows.Forms
    [Windows.Forms.MessageBox]::Show($boxMessage, $boxTitle,
        [Windows.Forms.MessageBoxButtons]::OK ,
        [Windows.Forms.MessageBoxIcon]::$boxIcon)
}#end messageBox

function moveToCompletedFolder
{
    #Date and Time
    $Date = (Get-Date -Format MMdd)
    $Time = (get-date -UFormat %R) -replace ':',''

    if (!(Test-Path -Path "$mainFolderPath\_Completed\Results_From_$Date-At-$Time"))
    {
        New-Item -Path "$mainFolderPath\_Completed\Results_From_$Date-At-$Time" -ItemType Directory | Out-Null
    }
    Move-Item -Path "$mainFolderPath\ProcessingData\*" -Destination "$mainFolderPath\_Completed\Results_From_$Date-At-$Time" | Out-Null
    messageBox "Dump of Delegates and Forwarding Rules are completed" "Dump Completed" Information
}

#Report Path
$mainFolderPath = "$env:userprofile\Desktop\Reports\DumpDelegatesandForwardingRules"

#Set File Path
$mailboxSMTPForwardingPath = "$mainFolderPath\ProcessingData\Mailboxsmtpforwarding.csv"
$allUsersPath = "$mainFolderPath\ProcessingData\AllUsers.csv"
$processedUserPath = "$mainFolderPath\ProcessingData\ProcessedUsers.csv"
$userInboxRulesPath = "$mainFolderPath\ProcessingData\UserInboxRules.csv"
$userDelegatesPath = "$mainFolderPath\ProcessingData\UserDelegates.csv"

#Verify if folder exist
if (!(Test-Path -Path $mainFolderPath))
{
    New-Item -Path "$mainFolderPath" -ItemType Directory | Out-Null
    New-Item -Path "$mainFolderPath\_Completed" -ItemType Directory | Out-Null
    New-Item -Path "$mainFolderPath\ProcessingData" -ItemType Directory | Out-Null
}

if ((Get-PSSession | Measure-Object).Count -ge 1) {
    if (Test-Path -Path $allUsersPath)
    {
        if (Test-Path -Path $processedUserPath)
        {
            $allUsers = Import-Csv $allUsersPath
            $processedUsers = Import-Csv $processedUserPath
            $SCRIPT:allUsers = Compare-Object -ReferenceObject $AllUsers -DifferenceObject $processedUsers -Property UserPrincipalName | Where {$_.SideIndicator -eq "<="} | Select UserPrincipalName
            if (($allUsers.Count -eq 0 -or $allUsers -eq $null -or $allUsers -eq "") -and (Test-Path variable:Global:allUsers))
            {
                moveToCompletedFolder
            }
        }
        else
        {
            $SCRIPT:AllUsers = Import-Csv $allUsersPath
        }
    }
    else
    {
        #Get All Users
        $SCRIPT:allUsers = @()
        $allUsers = Get-MsolUser -All -EnabledFilter EnabledOnly | select ObjectID, UserPrincipalName, FirstName, LastName, StrongAuthenticationRequirements, StsRefreshTokensValidFrom, StrongPasswordRequired, LastPasswordChangeTimestamp | Where-Object {($_.UserPrincipalName -notlike "*#EXT#*")}
        $allUsers | Export-Csv $allUsersPath -NoTypeInformation
        #Get All SMTP Forwarders
        $SMTPForwarding = Get-Mailbox -ResultSize Unlimited | select DisplayName,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxandForward | where {$_.ForwardingSMTPAddress -ne $null}
        $SMTPForwarding | Export-Csv $mailboxSMTPForwardingPath -NoTypeInformation
    }

    if (!(Test-Path -Path $processedUserPath) -and ((Get-ChildItem "$mainFolderPath\ProcessingData" | Measure-Object).count -ge 2)) 
    {
        "UserPrincipalName" > $processedUserPath
    }

    $UserInboxRules = @()
    $UserDelegates = @()

    if ($allUsers.count -ge 1) 
    {
        foreach ($User in $allUsers)
        {
            Write-Host "Checking inbox rules and delegates for user: " $User.UserPrincipalName;
            $UserInboxRules = Get-InboxRule -Mailbox $User.UserPrincipalname | Select MailboxOwnerId, RuleIdentity, Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectsTo -ne $null)}

            if ($UserInboxRules) 
            {
                $UserInboxRules | Export-Csv -NoTypeInformation $userInboxRulesPath -Append
            }
    
            #Adding Micro Delays between commands
            Start-Sleep -milliseconds 500

            $UserDelegates = Get-MailboxPermission -Identity $User.UserPrincipalName | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
            if ($UserDelegates)
            {
                $UserDelegates | Export-Csv -NoTypeInformation $userDelegatesPath -Append
            }

            #This will stop the script if you loose the remote connection   
            if ($Error[0] -clike "*Access is denied*")
            {
                messageBox "Connection to Office 365 has been lost" "Lost Connection" Error
                Break
            }
            $User.UserPrincipalName >> $processedUserPath
        }
    }

    if ((Get-ChildItem "$mainFolderPath\ProcessingData" | Measure-Object).count -ne 0 -and !($Error[0] -clike "*Access is denied*")) 
    {
        moveToCompletedFolder
    }
}
else
{
    messageBox "Sorry, it doesn't seem like you are connected to the Office 365 Exchange Server" "No Connection" Error   
}