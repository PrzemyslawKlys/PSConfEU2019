# Modules that will be needed for this session
$Modules = @(
    'PSWriteWord'
    'Documentimo'
)

foreach ($M in $Modules) {
    if (Get-Module -ListAvailable -Name $M) {
        Update-Module -Name $M
    } else {
        Install-Module -Name -Force
    }
}