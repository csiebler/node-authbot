'use strict';
const restify = require('restify');
const builder = require('botbuilder');
const azure = require("botbuilder-azure");
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const crypto = require('crypto');

const HttpsServer = require('./helpers/HttpsServer');
const AuthHelper = require('./helpers/AuthHelper');
const OfficeHelper = require('./helpers/OfficeHelper');
const DialogHelper = require('./helpers/DialogHelper');

require('dotenv').config();

const STORAGE_CONNECTION_STRING = process.env.STORAGE_CONNECTION_STRING;
const STATE_TABLE = process.env.STATE_TABLE;

//bot application identity
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

const USE_EMULATOR = (process.env.USE_EMULATOR == 'development');

//=========================================================
// Bot Setup
//=========================================================

// Setup Https Server
var server = new HttpsServer();

// Create chat bot
console.log(`Starting with AppId=${MICROSOFT_APP_ID}`)

let azureTableClient = new azure.AzureTableClient(STATE_TABLE, STORAGE_CONNECTION_STRING);
let tableStorage = new azure.AzureBotStorage({ gzipData: false }, azureTableClient);

var connector = USE_EMULATOR ? new builder.ChatConnector() : new azure.BotServiceConnector({
  appId: MICROSOFT_APP_ID,
  appPassword: MICROSOFT_APP_PASSWORD
});

var bot = new builder.UniversalBot(connector);
bot.set('storage', tableStorage);

server.post('/api/messages', connector.listen());

//=========================================================
// Auth Setup
//=========================================================

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
    //res.send(`Welcome ${req.user.displayName}! Please copy this number and paste it back to your chat so your authentication can complete ${magicCode}`);

    var body = `<html><body>Welcome ${req.user.displayName}! Please copy this number and paste it back to your chat so your authentication can complete: ${magicCode}</body></html>`;
    res.writeHead(200, {
      'Content-Length': Buffer.byteLength(body),
      'Content-Type': 'text/html'
    });
    res.write(body);
    res.end();

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
      builder.Prompts.text(session, `Welcome ${session.userData.userName}! You are currently logged in. To get the latest email, type 'email'. To quit, type 'quit'. To log out, type 'logout'.`);
    } else {
      session.endConversation("Goodbye.");
    }
  },
  (session, results, next) => {
    var resp = results.response;
    if (resp === 'email') {
      session.beginDialog('email');
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

bot.dialog('email', DialogHelper.emailDialog);
bot.dialog('signinPrompt', DialogHelper.signInDialog);
bot.dialog('validateCode', DialogHelper.validateCodeDialog);
