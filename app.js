'use strict';

const restify = require('restify');
const builder = require('botbuilder');
const azure = require("botbuilder-azure");

const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const expressSession = require('express-session');
const crypto = require('crypto');
const querystring = require('querystring');
const request = require('request');

require('dotenv').config();

const HttpsServer = require('./HttpsServer');
const AuthHelper = require('./AuthHelper');
const OfficeHelper = require('./OfficeHelper');

const STORAGE_CONNECTION_STRING = process.env.STORAGE_CONNECTION_STRING;
const STATE_TABLE = process.env.STATE_TABLE;

//bot application identity
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

//oauth details
const AZUREAD_APP_ID = process.env.AZUREAD_APP_ID;
const AZUREAD_APP_PASSWORD = process.env.AZUREAD_APP_PASSWORD;
const AZUREAD_APP_REALM = process.env.AZUREAD_APP_REALM;
const AUTHBOT_CALLBACKHOST = process.env.AUTHBOT_CALLBACKHOST;

const USE_EMULATOR = (process.env.USE_EMULATOR == 'development');

//=========================================================
// Bot Setup
//=========================================================

// Setup Https Server
var server = new HttpsServer();

// Create chat bot
console.log(`Starting with AppId ${MICROSOFT_APP_ID}`)

let azureTableClient = new azure.AzureTableClient(STATE_TABLE, STORAGE_CONNECTION_STRING);
let tableStorage = new azure.AzureBotStorage({ gzipData: false }, azureTableClient);

var connector = USE_EMULATOR ? new builder.ChatConnector() : new azure.BotServiceConnector({
  appId: MICROSOFT_APP_ID,
  appPassword: MICROSOFT_APP_PASSWORD
});

var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);


server.post('/api/messages', connector.listen());
server.get('/', restify.serveStatic({
  'directory': __dirname,
  'default': 'index.html'
}));

//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: 'keyboard cat', resave: true, saveUninitialized: false }));
server.use(passport.initialize());

server.get('/login', function (req, res, next) {
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login', customState: req.query.address, resourceURL: process.env.MICROSOFT_RESOURCE }, function (err, user, info) {
    console.log('login');
    if (err) {
      console.log(err);
      return next(err);
    }
    if (!user) {
      return res.redirect('/login');
    }
    req.logIn(user, function (err) {
      if (err) {
        return next(err);
      } else {
        return res.send('Welcome ' + req.user.displayName);
      }
    });
  })(req, res, next);
});

server.get('/api/OAuthCallback/',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/login' }),
  (req, res) => {
    console.log('OAuthCallback');
    console.log(req);
    const address = JSON.parse(req.query.state);
    const magicCode = crypto.randomBytes(4).toString('hex');
    const messageData = { magicCode: magicCode, accessToken: req.user.accessToken, refreshToken: req.user.refreshToken, userId: address.user.id, name: req.user.displayName, email: req.user.preferred_username };

    var continueMsg = new builder.Message().address(address).text(JSON.stringify(messageData));
    console.log(continueMsg.toMessage());

    bot.receive(continueMsg.toMessage());
    res.send('Welcome ' + req.user.displayName + '! Please copy this number and paste it back to your chat so your authentication can complete: ' + magicCode);
  });

passport.serializeUser(function (user, done) {
  done(null, user);
});

passport.deserializeUser(function (id, done) {
  done(null, id);
});

let oidStrategyv2 = AuthHelper.getStrategy();

passport.use(new OIDCStrategy(oidStrategyv2,
  (req, iss, sub, profile, accessToken, refreshToken, done) => {
    if (!profile.displayName) {
      return done(new Error("No oid found"), null);
    }
    // asynchronous verification, for effect...
    process.nextTick(() => {
      profile.accessToken = accessToken;
      profile.refreshToken = refreshToken;
      return done(null, profile);
    });
  }
));

//=========================================================
// Bots Dialogs
//=========================================================

function login(session) {
  // Generate signin link
  const address = session.message.address;

  // TODO: Encrypt the address string
  const link = AUTHBOT_CALLBACKHOST + '/login?address=' + querystring.escape(JSON.stringify(address));

  var msg = new builder.Message(session)
    .attachments([
      new builder.SigninCard(session)
        .text("Please click this link to sign in first.")
        .button("signin", link)
    ]);
  session.send(msg);
  builder.Prompts.text(session, "You must first sign into your account.");
}

bot.dialog('signin', [
  (session, results) => {
    console.log('signin callback: ' + results);
    session.endDialog();
  }
]);

bot.dialog('/', [
  (session, args, next) => {
    if (!(session.userData.userName && session.userData.accessToken && session.userData.refreshToken)) {
      session.send("Welcome! This bot retrieves the latest email for you after you login.");
      session.beginDialog('signinPrompt');
    } else {
      next();
    }
  },
  (session, results, next) => {
    if (session.userData.userName && session.userData.accessToken && session.userData.refreshToken) {
      // They're logged in
      builder.Prompts.text(session, "Welcome " + session.userData.userName + "! You are currently logged in. To get the latest email, type 'email'. To quit, type 'quit'. To log out, type 'logout'. ");
    } else {
      session.endConversation("Goodbye.");
    }
  },
  (session, results, next) => {
    var resp = results.response;
    if (resp === 'email') {
      session.beginDialog('workPrompt');
    } else if (resp === 'quit') {
      session.endConversation("Goodbye.");
    } else if (resp === 'logout') {
      session.userData.loginData = null;
      session.userData.userName = null;
      session.userData.accessToken = null;
      session.userData.refreshToken = null;
      session.endConversation("You have logged out. Goodbye.");
    } else {
      next();
    }
  },
  (session, results) => {
    session.replaceDialog('/');
  }
]);

bot.dialog('workPrompt', [
  (session) => {
    OfficeHelper.getUserLatestEmail(session.userData.accessToken,
      function (requestError, result) {
        if (result && result.value && result.value.length > 0) {
          const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
          session.send(responseMessage);
          builder.Prompts.confirm(session, "Retrieve the latest email again?");
        } else {
          console.log('no user returned');
          if (requestError) {
            console.log('requestError');
            console.error(requestError);
            // Get a new valid access token with refresh token
            AuthHelper.getAccessTokenWithRefreshToken(session.userData.refreshToken, (err, body, res) => {

              if (err || body.error) {
                console.log(err);
                session.send("Error while getting a new access token. Please try logout and login again. Error: " + err);
                session.endDialog();
              } else {
                session.userData.accessToken = body.accessToken;
                OfficeHelper.getUserLatestEmail(session.userData.accessToken,
                  function (requestError, result) {
                    if (result && result.value && result.value.length > 0) {
                      const responseMessage = 'Your latest email is: "' + result.value[0].Subject + '"';
                      session.send(responseMessage);
                      builder.Prompts.confirm(session, "Retrieve the latest email again?");
                    }
                  }
                );
              }

            });
          }
        }
      }
    );
  },
  (session, results) => {
    var prompt = results.response;
    if (prompt) {
      session.replaceDialog('workPrompt');
    } else {
      session.endDialog();
    }
  }
]);

bot.dialog('signinPrompt', [
  (session, args) => {
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "please click the signin link.");
    } else {
      login(session);
    }
  },
  (session, results) => {
    //resuming
    session.userData.loginData = JSON.parse(results.response);
    if (session.userData.loginData && session.userData.loginData.magicCode && session.userData.loginData.accessToken) {
      session.beginDialog('validateCode');
    } else {
      session.replaceDialog('signinPrompt', { invalid: true });
    }
  },
  (session, results) => {
    if (results.response) {
      //code validated
      session.userData.userName = session.userData.loginData.name;
      session.endDialogWithResult({ response: true });
    } else {
      session.endDialogWithResult({ response: false });
    }
  }
]);

bot.dialog('validateCode', [
  (session) => {
    builder.Prompts.text(session, "Please enter the code you received or type 'quit' to end. ");
  },
  (session, results) => {
    const code = results.response;
    if (code === 'quit') {
      session.endDialogWithResult({ response: false });
    } else {
      if (code === session.userData.loginData.magicCode) {
        // Authenticated, save
        session.userData.accessToken = session.userData.loginData.accessToken;
        session.userData.refreshToken = session.userData.loginData.refreshToken;

        session.endDialogWithResult({ response: true });
      } else {
        session.send("hmm... Looks like that was an invalid code. Please try again.");
        session.replaceDialog('validateCode');
      }
    }
  }
]);


