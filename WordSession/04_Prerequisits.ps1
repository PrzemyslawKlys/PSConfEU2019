# Modules that will be needed for 1st part of the session
$Modules = @(
    'PSWriteWord'
    'PSSharedGoods'
)

foreach ($M in $Modules) {
    if (Get-Module -ListAvailable -Name $M) {
        Update-Module -Name $M
    } else {
        Install-Module -Name -Force
    }
}