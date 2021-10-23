

$ExcelObj = New-Object -comobject Excel.Application
$ExcelObj.visible=$true

$ExcelWorkBook = $ExcelObj.Workbooks.Open("$PSScriptRoot\add_user.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("herzo_student")

$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

for($i=2;$i -le $rowcount;$i++){

    $class = $ExcelWorkSheet.Columns.Item(3).Rows.Item($i).Text
    $upn = $ExcelWorkSheet.Columns.Item(6).Rows.Item($i).Text

    Try{

        $usr = Get-AzureADUser -ObjectId $upn
        $group = Get-AzureADGroup -SearchString $class
        
        Add-AzureADGroupMember -ObjectId $group.ObjectId -RefObjectId $usr.ObjectId
    }
    Catch{
    
        "#addUsersToGroup upn: "+$upn+" class: "+$class |  Out-File -FilePath .\error.txt -Append

    }

    

}

$ExcelObj.WorkBooks.Close()
$ExcelObj.Quit()