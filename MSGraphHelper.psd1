@{
    ModuleVersion = '1.1.1'
    GUID = '9bde95f6-5a22-4e54-9c5e-93312fe34737'
    Author = 'Yanik Maurer'
    PowerShellVersion = '5.1'
    RootModule = 'MSGraphHelper.psm1'
    FunctionsToExport = @('Get-RequiredModules', 'Get-AccessTokenMSAL-ApplicationPermission', 'Get-AccessTokenMSAL-DelegatedPermissions', 'New-LocalSecret', 'Get-LocalSecret', 'Read-Calendar', 'Send-Email')
    RequiredModules = @('MSAL.PS')
    Description = 'Module that makes it easy to call Graph Ressources'
}