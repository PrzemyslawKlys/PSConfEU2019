Import-Module Emailimo -Force

if ($null -eq $DataSetForest) {
    $DataSetForest = Get-WinADForestInformation -Verbose
}
if ($null -eq $DataSetEvents) {
    $DataSetEvents = Find-Events -Report ADUserChangesDetailed, ADUserChanges, ADUserLockouts, ADUserStatus, ADGroupChanges -Servers 'AD1', 'AD2' -DatesRange Last14days
}

$UserNotify = 'Przemyslaw Klys'

Email -FilePath "$PSScriptRoot\11-Demo.html" {
    EmailHeader {
        EmailFrom -Address 'testO365@evotec.xyz'
        EmailTo -Addresses "testO365@evotec.xyz"
        EmailServer -Server 'smtp.office365.com' -Port 587 -UserName 'testO365@evotec.xyz' -Password 'C:\Support\Important\Password-O365Test.txt' -PasswordAsSecure -PasswordFromFile -SSL
        EmailOptions -Priority High -DeliveryNotifications Never
        EmailSubject -Subject 'This is a test email'
    }
    EmailBody -FontFamily 'Calibri' -Size 15 {
        EmailTextBox -FontFamily 'Calibri' -Size 17 -TextDecoration underline -Color DarkSalmon -Alignment center {
            'Demonstration'
        }

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

        EmailTextBox -FontSize 15 -Color DarkCyan -FontStyle italic {
            ""
            @"
                This is tricky 😁 but it works like one
                big line... even thou we've split this over few lines
                already this will be one continues line. Get it right?
                Notice how I gave it color and made it font size 15.
"@
        }

        EmailList -FontSize 15 {
            EmailListItem -Text 'First item' -Color Red
            EmailListItem -Text '2nd item' -Color Green
            EmailList {
                EmailListItem -Text '3rd item' -FontStyle italic
                EmailListItem -Text '4th item' -TextDecoration line-through
            }
        }

        EmailText -FontSize 15 -Text 'This is my', 'text' -Color Red, Green -TextDecoration underline -FontFamily 'Arial'
        EmailText -FontSize 15 -Text 'This is my', 'text', ' but ', ' with different formatting.' -Color Blue, Red, Green -TextDecoration underline, none, 'line-through' -FontFamily 'Calibri' -LineBreak

        EmailTextBox {
            'Kind regards,'
            'Evotec IT'
        }
    }
} -AttachSelf
