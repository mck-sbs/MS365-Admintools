Import-Module AzureAD

$Credentials = Get-Credential
Connect-AzureAD -Credential $Credentials

$ExcelObj = New-Object -comobject Excel.Application
$ExcelObj.visible=$true

$ExcelWorkBook = $ExcelObj.Workbooks.Open("$PSScriptRoot\add_user.xlsx")
$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("herzo_student")

$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

for($i=2;$i -le $rowcount;$i++){

    $givenname = $ExcelWorkSheet.Columns.Item(4).Rows.Item($i).Text
    $surname = $ExcelWorkSheet.Columns.Item(5).Rows.Item($i).Text
    $upn = $ExcelWorkSheet.Columns.Item(6).Rows.Item($i).Text
    $pw = $ExcelWorkSheet.Columns.Item(7).Rows.Item($i).Text


    $password = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $password.Password = $pw

    Try{

        New-AzureADUser -UserPrincipalName $upn -MailNickName ($givenname + "." + $surname) -DisplayName ($surname + ", " + $givenname) -GivenName $givenname -Surname $surname -PasswordProfile $password -AccountEnabled $true -UsageLocation "DE"


        $planName1 = "OFFICESUBSCRIPTION_STUDENT"
        $planName2 = "STANDARDWOFFPACK_STUDENT"

        $License1 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License1.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName1 -EQ).SkuID

        $License2 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License2.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName2 -EQ).SkuID

        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License1,$License2

        Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $LicensesToAssign
    }
    Catch{
    
        $upn |  Out-File -FilePath .\error.txt -Append

    }

}

$ExcelObj.WorkBooks.Close()
$ExcelObj.Quit()