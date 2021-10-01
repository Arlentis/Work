$ADUserFrom = Get-ADUser -Filter * -Properties * | Select-Object -Property "GivenName","Surname","Name","SamAccountName","Title","Mail" | Where-Object -Property "GivenName" -NotLike "" | Where-Object -Property "Surname" -NotLike ""  | Where-Object -Property "mail" -NotLike "" | Sort-Object "GivenName" | Out-Gridview -Title "All Users" -PassThru | Select-Object -ExpandProperty SamAccountName

ForEach ($ADUser in $ADUserFrom) {Get-ADUSer -Identity $ADUser -Properties * | Select-Object -Property "mail" }
