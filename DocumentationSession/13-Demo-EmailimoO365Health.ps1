Import-Module PSWinDocumentation.O365HealthService -Force
Import-Module Emailimo -Force

$ApplicationID = Get-Content -LiteralPath "C:\Support\Important\O365-ApplicationID.txt"
$ApplicationKey = Get-Content -LiteralPath "C:\Support\Important\O365-ApplicationKey.txt"
$TenantDomain = 'evotec.pl' # CustomDomain (onmicrosoft.com won't work), alternatively you can use DirectoryID

if (-not $O365) {
    $O365 = Get-Office365Health -ApplicationID $ApplicationID -ApplicationKey $ApplicationKey -TenantDomain $TenantDomain -ToLocalTime -Verbose
}

Email -FilePath "$PSScriptRoot\13-Demo.html" {
    EmailHeader {
        EmailFrom -Address 'testO365@evotec.xyz'
        EmailTo -Addresses "testO365@evotec.xyz"
        EmailServer -Server 'smtp.office365.com' -Port 587 -UserName 'testO365@evotec.xyz' -Password 'C:\Support\Important\Password-O365Test.txt' -PasswordAsSecure -PasswordFromFile -SSL
        EmailOptions -Priority High -DeliveryNotifications Never
        EmailSubject -Subject 'Office 365 Health Services Status'
    }
    EmailBody -FontFamily 'Calibri' -Size 15 {
        EmailTextBox -FontFamily 'Calibri' -Size 17 -TextDecoration underline -Color DarkSalmon -Alignment center {
            'Demonstration of Office 365'
        }

        EmailText -TextBlock {
            "Here's a list of things this email covers:"
        }

        EmailList -FontSize 15 {
            EmailListItem -Text 'List of Services' -Color Red
            EmailListItem -Text 'List of Incidents' -Color Green
            EmailList {
                EmailListItem -Text 'Only 2 days' -FontStyle italic
                EmailListItem -Text 'All of it','Last 1 day' -TextDecoration line-through, none
            }
        }


        EmailText -Text 'O365', ' Services Status ❤' -Color BlueViolet, Green
        EmailTable -DataTable $O365.CurrentStatus -HideFooter

        EmailText -LineBreak

        $ListLast2Days = $O365.IncidentsMessages | Where-Object { $_.PublishedDaysAgo -le 2 }

        EmailText -Text 'O365', ' Incidents Messages 😁' -Color BlueViolet, Green
        EmailTable -DataTable $ListLast2Days -HideFooter

        EmailText -LineBreak

        $ListLast1Days = $O365.IncidentsMessages | Where-Object { $_.PublishedDaysAgo -le 1 }

        EmailText -Text 'O365', ' Incidents Messages 🤔', ' - ', 'All', ' Last Day' -Color BlueViolet, Green, None, Red -TextDecoration None, None, None, line-through, None
        EmailTable -DataTable $ListLast1Days -HideFooter

        EmailText -LineBreak

        $Incidents = $O365.MessageCenterInformationExtended | Where-Object { $_.PublishedDaysAgo -le 5 }
        EmailText -Text 'O365', ' Message Center 🤔', ' Showing how InvokeHTMLTags works'
        EmailTable -DataTable $Incidents -HideFooter -InvokeHTMLTags

        EmailTextBox {
            'Kind regards,'
            'Evotec IT'
        }
    }
} -AttachSelf
