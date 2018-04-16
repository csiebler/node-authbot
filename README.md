# AuthBot for Node.js with Microsoft Bot Framework

_A bot that authenticates users and get profile information and the latest email for the logged in user_.

This bot enables users to authenticate with their Microsoft ID and/or their AD domain user. After authentication, the bot uses access tokens to retrieve the latest email for the user.

This repo is based on [node-authbot](https://github.com/CatalystCode/node-authbot), but has been updated to use HTTPS and features overall simplification.

## Features

* Supports Azure Active Directory endpoint V2 (now supports both AD accounts and Microsoft accounts)
* HTTPS-only, as the [Azure Bot Service](https://azure.microsoft.com/en-us/services/bot-service/) requires HTTPS for authentication
Allow easy and secure sign in, even in chat sessions including multiple users

## Installation

Via command line:

```
git clone https://github.com/csiebler/node-authbot.git
cd node-authbot
npm install
```

If you haven't done it yet, register your bot through the [Apps Dev Portal](https://apps.dev.microsoft.com) (make sure you're signing in with the correct Azure Active Directory user):

1. Open [Apps Dev Portal](https://apps.dev.microsoft.com) and click `Add an app` (the portal is used for V2 Azure Active Directory endpoints)
1. Give your app a name
1. Click on the new app and note the Application Id (this is your `AZUREAD_APP_ID`)
1. Click `Generate New Password` under `Application Secrets` and save it somewhere (this is your `AZUREAD_APP_PASSWORD`)
1. Under `Platforms`, click `Add Platform` and choose `Web`
1. Set `Redirect URLs` to `https://localhost:3979/api/OAuthCallback`

Next, `cp .env.template .env` populate environment variables in `.env`:

* `MICROSOFT_APP_ID` - The AppId of your Bot Service
* `MICROSOFT_APP_PASSWORD` - The associated password for your Bot Service
* `AZUREAD_APP_ID` - Your AppId from [Apps Dev Portal](https://apps.dev.microsoft.com)
* `AZUREAD_APP_PASSWORD` - The associated password from the [Apps Dev Portal](https://apps.dev.microsoft.com)
* `AZUREAD_APP_REALM` - Leave set to `common` (in most cases)
* `AUTHBOT_CALLBACKHOST` - For testing, set to `https://localhost:3979`, in production, set it to the hostname of the bot service
* `USE_EMULATOR` - For local testing, use `development`, else `production`
* `STORAGE_CONNECTION_STRING` - Connection string to for your bot's storage account
* `STATE_TABLE` - Name of the state table for your bot

Start the bot locally via:

```
node app.js
```

and then connect to it via the [Microsoft Bot Framework Emulator](https://github.com/Microsoft/BotFramework-Emulator) using the following address:

```
https://localhost:3979/api/messages
```

For setting up channels for the bot (Facebook Messenger, Kik, Skype, Microsoft Teams, etc.), follow the [instructions on the official Bot Service website](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-manage-channels).

## Deployment on Azure

Firstly, update the `Redirect URLs` in the [Apps Dev Portal](https://apps.dev.microsoft.com) to return you to the correct URL.

For production use, you also want to replace the server's private key and TLS certificate with your own:

```
etc/ssl/server.crt
etc/ssl/server.key
```

If the bot is hosted on Azure App Service, make sure to increase the `maxQueryString` limit by updating the `web.config` file in `$HOME/site/wwwroot` like this:

```
<security>
  <requestFiltering>
    <requestLimits maxQueryString="10000"/>
    ...
  </requestFiltering>
</security>
```

## Acknowledgement

Many thanks go to the authors and co-authors of [node-authbot](https://github.com/CatalystCode/node-authbot), namely [@ritazh](https://github.com/ritazh), [@sozercan](https://github.com/sozercan), and [@GeekTrainer](https://github.com/GeekTrainer). 

## License

Licensed using the MIT License (MIT); For more information, please see [LICENSE](LICENSE).
