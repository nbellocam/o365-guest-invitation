# Inviting guest to an Office 365 using Azure Functions and Graph API

The solution has two parts, the Azure Function itself and a console app that call the function once for each line of the csv file passed by parameter.

The Azure Function invites the user to the Office tenant and then add the new external user to a Security Group. The Security Group is created dynamically taking into consideration the maximum of 5000 users that each group can have.

## Configuring the Azure Function code

### Register your application in your tenant

1. Sign in to the Azure portal though the Office 365 Admin Portal by selecting the Azure AD admin site.
1. In the left-hand navigation pane, choose More Services, click **App Registrations**, and click **Add**.
1. Follow the prompts and create a new application.
1. Select **Web App / API** as the Application Type.
1. Provide any redirect URI (e.g. https://GraphAPI) as it's not relevant for this app.
1. The application will now show up in the list of applications, click on it to obtain the **Application ID** (also known as Client ID). Copy it as you'll need it in the azure function code.
1. In the **Settings** menu, click on **Keys** and add a new key (also known as client secret). Also copy it for use in the azure function code.

### Configure create, read and update permissions for your application

Now you need to configure your application to get all the required permissions to create, read and update users and groups.

App required permissions:
- Group.ReadWrite.All
- Directory.ReadWrite.All
- User.Invite.All

#### Steps

1. In the Azure portal's App Registrations menu, select your application.
1. In the Settings menu, click on **Required permissions**.
1. In the Required permissions menu, click on **Microsoft Graph** (or add it if it is not listed yet).
1. In the Enable Access menu, select the **Read and write directory data**, **Read and write all groups** and **Invite guest users to the organization** permissions from **Application Permissions** and click **Save**.
1. Finally, back in the Required permissions menu, click on the **Grant Permissions** button.


### Configuring the apps

Open the Azure Function code (the _index.js_ file inside the _invite-guest-azFunc_ folder) and replace the following placeholders:

- `{tenant-name-here}`: the Office 365 tenant name
- `{Application-id-here}`: the AD app id created before
- `{Application-key-here}`: the AD app key created before
- `{collection-here}`: the remaining part of the url of the SharePoint site

After updating those placeholder, create an Azure Function triggered by http and deploy it (remember to deploy the dependencies too).

Now, for the inviter code, open the _Program.cs_ file and update the following placeholders:

- `{app-name-here}`: the name of the Function App.
- `{fuction-name-here}`: the name of the Function itself (i.e. _HttpTriggerJS1_)
- `{azure-function-code-here}`: the code token provided by Azure to call the function.

### Running the console app

To run the console app just execute the following:

```
dotnet run .\guests.csv
```

> **Note**: The csv file contains the full name and the email splitted with a `;`.