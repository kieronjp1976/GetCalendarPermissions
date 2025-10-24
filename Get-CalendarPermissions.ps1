$Mailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox
$Results = @()

foreach ($mbx in $Mailboxes) {
    $smtp = $mbx.PrimarySmtpAddress
    $path = "${smtp}:\Calendar"


    try {
        $perms = Get-MailboxFolderPermission -Identity $path -ErrorAction Stop
        foreach ($perm in $perms) {
            $Results += [pscustomobject]@{
                Mailbox      = $smtp
                Folder       = $perm.FolderName
                User         = $perm.User
                AccessRights = $perm.AccessRights -join ';'
            }
        }
    }
    catch {
        Write-Warning "Failed to get permissions for $smtp ($path): $_"
    }
}

$OutPath = ".\Output.csv"
$Results | Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8
