#$Mailboxes=Get-Mailbox | Where-Object{$_.isshared -eq $true}
$Mailboxes=Get-Mailbox -RecipientTypeDetails SharedMailbox

$OutPath = ".\Output.csv"

foreach($mbx in $Mailboxes){
    $smtp=$mbx.primarysmtpaddress
    $path=$smtp + ":\Calendar" 
    $perms=Get-MailboxFolderPermission -Identity $path
        foreach($perm in $perms){
            if($perm.user -notmatch "Default" -and $perm.user -notmatch "anonymous"){
             $OutputVar=    $smtp + "," + $perm.foldername + "," + $perm.user + "," + $perm.accessrights
             
             Out-File -FilePath $OutPath -InputObject $OutputVar -Append
            }

        }
       
}