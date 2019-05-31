Import-Module PSWinDocumentation -Force

$Document = [ordered]@{
    Configuration = [ordered] @{
        Prettify       = @{
            CompanyName        = 'Evotec'
            UseBuiltinTemplate = $true
            CustomTemplatePath = ''
            Language           = 'en-US'
        }
        Options        = @{
            OpenDocument = $false
            OpenExcel    = $false
        }
        DisplayConsole = @{
            ShowTime   = $true
            LogFile    = "C:\Support\Logs\PSWinDocumentationTesting.log"
            TimeFormat = 'yyyy-MM-dd HH:mm:ss'
        }
        Debug          = @{
            Verbose = $true
        }
    }
    DocumentAD    = [ordered] @{
        Enable        = $true
        ExportWord    = $false
        ExportExcel   = $false
        ExportSql     = $true
        FilePathWord  = "$Env:USERPROFILE\Desktop\PSWinDocumentation-Report.docx"
        FilePathExcel = "$Env:USERPROFILE\Desktop\PSWinDocumentation-Report.xlsx"
        Sections      = [ordered] @{
            SectionForest = [ordered] @{

            }
            SectionDomain = [ordered] @{
                SectionExcelDomainUsers               = [ordered] @{
                    Use                     = $true
                    ExcelExport             = $false
                    ExcelWorkSheet          = '<Domain> - Users'
                    ExcelData               = [PSWinDocumentation.ActiveDirectory]::DomainUsers

                    SqlExport               = $true
                    SqlServer               = 'localhost'
                    SqlDatabase             = 'SSAE18'
                    SqlData                 = [PSWinDocumentation.ActiveDirectory]::DomainUsers
                    SqlTable                = 'dbo.[Users]'
                    SqlTableCreate          = $true
                    DisabledSqlTableMapping = [ordered] @{  # SqlTableMapping is proper name
                        # Left Side is data in PSWinReporting
                        # Right Side is column name in SQL
                        # Changing makes sense only for left side...
                        # Use this if you need to have different mapping
                        'Name'                              = 'Name'
                        'UserPrincipalName'                 = 'UserPrincipalName'
                        'SamAccountName'                    = 'SamAccountName'
                        'Display Name'                      = 'Display Name'
                        'Given Name'                        = 'Given Name'
                        'Surname'                           = 'Surname'
                        'EmailAddress'                      = 'EmailAddress'
                        'PasswordExpired'                   = 'PasswordExpired'
                        'PasswordLastSet'                   = 'PasswordLastSet'
                        'PasswordNotRequired'               = 'PasswordNotRequired'
                        'PasswordNeverExpires'              = 'PasswordNeverExpires'
                        'Enabled'                           = 'Enabled'
                        'Manager'                           = 'Manager'
                        'Manager Email'                     = 'Manager Email'
                        'DateExpiry'                        = 'DateExpiry'
                        'DaysToExpire'                      = 'DaysToExpire'
                        'AccountExpirationDate'             = 'AccountExpirationDate'
                        'AccountLockoutTime'                = 'AccountLockoutTime'
                        'AllowReversiblePasswordEncryption' = 'AllowReversiblePasswordEncryption'
                        'BadLogonCount'                     = 'BadLogonCount'
                        'CannotChangePassword'              = 'CannotChangePassword'
                        'CanonicalName'                     = 'CanonicalName'
                        'Description'                       = 'Description'
                        'DistinguishedName'                 = 'DistinguishedName'
                        'EmployeeID'                        = 'EmployeeID'
                        'EmployeeNumber'                    = 'EmployeeNumber'
                        'LastBadPasswordAttempt'            = 'LastBadPasswordAttempt'
                        'LastLogonDate'                     = 'LastLogonDate'
                        'Created'                           = 'Created'
                        'Modified'                          = 'Modified'
                        'Protected'                         = 'Protected'
                        'Primary Group'                     = 'Primary Group'
                        'Member Of'                         = 'Member Of'
                        'AddedWhen'                         = 'AddedWhen'# ColumnsToTrack when it was added to database and by who / not part of event
                        'AddedWho'                          = 'AddedWho'   # ColumnsToTrack when it was added to database and by who / not part of event
                    }
                }

                SectionExcelDomainUsersAll            = [ordered] @{
                    Use            = $false
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Users All'
                    ExcelData      = [PSWinDocumentation.ActiveDirectory]::DomainUsersAll
                }
                SectionExcelDomainUsersSystemAccounts = [ordered] @{
                    Use            = $false
                    ExcelExport    = $false
                    ExcelWorkSheet = '<Domain> - Users System'
                    ExcelData      = [PSWinDocumentation.ActiveDirectory]::DomainUsersSystemAccounts
                }
                SectionExcelDomainUsersNeverExpiring  = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $false
                    ExcelWorkSheet        = '<Domain> - Never Expiring'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainUsersNeverExpiring

                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlTableCreate        = $true
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainUsersNeverExpiring
                    SqlTable              = 'dbo.[UsersNeverExpiring]'
                    SqlTableAlterIfNeeded = $true
                }
                SectionExcelDomainUsersFullList       = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $false
                    ExcelWorkSheet        = '<Domain> - Users List Full'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainUsersFullList
                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainUsersFullList
                    SqlTable              = 'dbo.[DomainUsersFullList]'
                    SqlTableCreate        = $true
                    SqlTableAlterIfNeeded = $true
                }

                SectionExcelDomainComputersFullList   = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $false
                    ExcelWorkSheet        = '<Domain> - Computers List'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainComputersFullList

                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlTableCreate        = $true
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainComputersFullList
                    SqlTable              = 'dbo.[DomainComputersFullList]'
                    SqlTableAlterIfNeeded = $true
                }
                SectionExcelDomainGroupsFullList      = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $false
                    ExcelWorkSheet        = '<Domain> - Groups List'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainGroupsFullList

                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlTableCreate        = $true
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainGroupsFullList
                    SqlTable              = 'dbo.[DomainGroupsFullList]'
                    SqlTableAlterIfNeeded = $true
                }
                SectionExcelDomainGroupsRest          = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $true
                    ExcelWorkSheet        = '<Domain> - Groups'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainGroups

                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlTableCreate        = $true
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainGroups
                    SqlTable              = 'dbo.[DomainGroups]'
                    SqlTableAlterIfNeeded = $true
                }

                SectionExcelDomainGroupsSpecial       = [ordered] @{
                    Use                   = $true
                    ExcelExport           = $true
                    ExcelWorkSheet        = '<Domain> - Groups Special'
                    ExcelData             = [PSWinDocumentation.ActiveDirectory]::DomainGroupsSpecial

                    SqlExport             = $true
                    SqlServer             = 'localhost'
                    SqlDatabase           = 'SSAE18'
                    SqlData               = [PSWinDocumentation.ActiveDirectory]::DomainGroupsSpecial
                    SqlTableCreate        = $true
                    SqlTable              = 'dbo.[DomainGroupsSpecial]'
                    SqlTableAlterIfNeeded = $true
                }
            }
        }

    }
}

Start-Documentation -Document $Document -Verbose