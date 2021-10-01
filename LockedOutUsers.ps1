try
{
    $LockedOutUsers = Get-ADUser -Filter * -Properties * | Where-Object -Property "LockedOut" -EQ "True" | Select-Object "Name","SamAccountName","Title","Mail" | Sort-Object "Name" | Out-Gridview -Title "Locked Out Users" -PassThru | Select-Object -ExpandProperty SamAccountName
    if ($LockedOutUsers)
    {
        Out-Gridview -Title "Locked Out Users" -PassThru | Select-Object -ExpandProperty SamAccountName ; ForEach ($LockedOutUser in $LockedOutUsers) {Unlock-ADAccount -Identity $LockedOutUser}
    }
    else
    {
        $null = [System.Windows.Forms.MessageBox]::Show("There are no locked out users", "Account Report", "OK", "Information")
    }
}
catch
{
    $ErrMsg = $_.Exception.Message
    $null = [System.Windows.Forms.MessageBox]::Show("Error Occurred: $ErrMsg", "Error", "OK", "Error")
}
