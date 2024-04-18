# MSGraphHelper
Lets you easily connect to Graph API via PowerShell using EntraID App Registrations and stores secrets if needed.
Essentially acts as a wrapper for "MSAL.PS" that further simplifies its use.

# Application vs. delegated permissions
Entra ID App registrations offers two ways of how API permissions can be applied and authenticated.
- **Delegated Permissions:**
  When using delegated permissions, the application accesses the permitted API as the signed-in user. This launches a sign-in and authorization prompt that the user needs to complete.
  
- **Application Permissions:**
  Application permissions on the other hand does not sign-in as the user. More effectively, when using GraphAPI permissions, application permissions provide global permissions to an API Action (eg. "Read all Users calendar"). This requires the use of a ClientSecret from the app registration, which potentially needs to be stored locally in secure manner (if you intend to run your automation unattended).

  When using application permissions, it is advisable to limit the scope of the permissions by other means such as for example Exchange Policies (limit permission to only mailboxes, not all).

## Requirements
MSGraphHelper requires the following:
- PowerShell 5.1 or higher
- An Entra ID Application registration (advisable to create one per use-case)
  - clientId
  - TennatId
  - clientSecret (Only when using Application Permissions)
 
## Dependencies 
- PowerShell Module "MSAL.PS"
- PowerShell Module "CredentialManager"


## Install the Module
To use the functions made available via the "MSGraphHelper" module, you need to install the module from PowerShell Gallery first, by using 'Install-Module'.

1. Open PowerShell as admin and run `Install-Module`:
  ```powershell
  Install-Module MSGraphHelper
  ```
  If the dependency "MSAL.PS" is not present, it will be installed and imported.

## How to obtain an MS Graph Access Token when using application permissions
Lets assume we want to build some tooling that uses GraphAPI to retrieve calendar entries of the mailboxes from our tenant to simplify absence planning.
As this app needs to be able to access all mailboxes, we need to use Application Permissions.

In this scenario, we assume you have an EntraID App Registration to be used.

1. Import the Module by running 
  ```powershell
  Import-Module MSGraphHelper
  ```
2. Define `tenantId` and `clientId`
  ```powershell
$tenantId = "xz651813-aac4-56az-8475-a24r14783104"
$clientId = "2sda6r16-r1g9-784a-4r07-rt91d41d4r19" 
  ```
3. As this script will run every night, we need a way to store the `clientSecret` in a secure manner on the system the script will run. We can do this by using the modules function `New-LocalSecret`
  ```powershell
  New-LocalSecret -clientId $clientId
   ```
4. This brings up a Credential prompt and asks for the `clientSecret`. Once provided, it will be stored in the Windows Credential Manager (default scope = system-wide so that any user on that system can access it).
5. When your script runs trough, you can retrieve the store `clientSecret` from Credential Manager using `Get-LocalSecrect`:
  ```powershell
  $clientSecretSecure = Get-LocalSecret -clientId $clientId
   ```
6. Then you can use that to retrieve a Graph API Access token:
  ```powershell
  $accessToken = Get-AccessTokenMSAL-ApplicationPermission -clientId $clientId -clientSecret $clientSecretSecure -tenantId $tenantId
   ```
