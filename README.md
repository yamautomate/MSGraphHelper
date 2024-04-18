# MSGraphHelper
Lets you easily connect to Graph API using an App Regisrtation and store secrets if needed.

## Application vs. Delegated Permissions
Azure App Registration have two ways of procigind API permissions.
- Application
- Delegated

When using application permissions the application runs a background service without signin-in as a user. This requires the use of a clientSecret.
Using delegated permissions on the other hand, the application accesses the API as the signed-in user. This does not require a clientSecret but will prompt the user to SignIn to the App registration and approve it. 

