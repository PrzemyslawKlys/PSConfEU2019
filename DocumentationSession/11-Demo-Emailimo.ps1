Import-Module Emailimo -Force

if ($null -eq $DataSetForest) {
    $DataSetForest = Get-WinADForestInformation -Verbose
}
if ($null -eq $DataSetEvents) {
    $DataSetEvents = Find-Events -Report ADUserChangesDetailed, ADUserChanges, ADUserLockouts, ADUserStatus, ADGroupChanges -Servers 'AD1', 'AD2' -DatesRange Last14days
}

$UserNotify = 'Przemyslaw Klys'

# Email -FilePath "$PSScriptRoot\11-Demo.html" {
Email -WhatIf -FilePath "$PSScriptRoot\11-Demo.html" {
    EmailHeader {
        EmailFrom -Address 'testO365@evotec.xyz'
        EmailTo -Addresses "testO365@evotec.xyz"
        EmailServer -Server 'smtp.office365.com' -Port 587 -UserName 'testO365@evotec.xyz' -Password 'C:\Support\Important\Password-O365Test.txt' -PasswordAsSecure -PasswordFromFile -SSL
        EmailOptions -Priority High -DeliveryNotifications Never
        EmailSubject -Subject 'This is a test email'
    }
    EmailBody -FontFamily 'Calibri' -Size 15 {
        EmailText -Text "Hello ", $UserNotify, "," -Color None, Blue, None -Verbose -LineBreak

        EmailText -Text "Here's some ", 'Forest information:' -Color None, Blue
        EmailTable -DataTable $DataSetForest.ForestInformation -HideFooter

        EmailText -LineBreak

        EmailText -Text 'And here are ', 'User changes', ' detailed:' -Color BlueViolet, Green, Red
        EmailTable -DataTable $DataSetEvents.ADUserChangesDetailed -HideFooter


        EmailTextBox {
            'You can add more tables, and more things. You can even attach self which gives you'
            'couple of nice things... such as different JS/CSS.'
        }
        EmailText -LineBreak
        EmailText -Text 'Keep in mind this is still ', 'work in progress!' `
            -Color None, LightSkyBlue, None, LightSkyBlue -TextDecoration none, underline, none, underline -FontWeight normal, bold, normal, bold
        EmailText -LineBreak
        EmailTextBox {
            'Kind regards,'
            'Evotec IT'
        }
    }
} -AttachSelf
