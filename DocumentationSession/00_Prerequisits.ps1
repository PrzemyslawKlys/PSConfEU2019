# Modules that will be needed for 1st part of the session
# All those modules should install automatically after you install PSWinDocumentation
# But just in case leaving it here
$Modules = @(
    'PSWinDocumentation'
    'PSWinDocumentation.AD'
    'PSWindocumentation.O365'
    'PSWinDocumentation.Exchange'
    'PSWriteWord'
    'PSWriteExcel'
    'dbatools'
    'PSSharedGoods'
)

foreach ($M in $Modules) {
    if (Get-Module -ListAvailable -Name $M) {
        Update-Module -Name $M
    } else {
        Install-Module -Name -Force
    }
}

# This is part of PSSharedGoods
Install-WinConnectity -All