'use strict';
const querystring = require('querystring');
const builder = require('botbuilder');
const OfficeHelper = require('./OfficeHelper');

//bot application identity
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

//oauth details
const AZUREAD_APP_ID = process.env.AZUREAD_APP_ID;
const AZUREAD_APP_PASSWORD = process.env.AZUREAD_APP_PASSWORD;
const AZUREAD_APP_REALM = process.env.AZUREAD_APP_REALM;
const AUTHBOT_CALLBACKHOST = process.env.AUTHBOT_CALLBACKHOST;

exports.validateCodeDialog = [
  (session) => {
    builder.Prompts.text(session, "Please enter the code you received or type 'quit' to end.");
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
        session.send("hmm...Looks like that was an invalid code. Please try again or type 'quit' to end.");
        session.replaceDialog('validateCode');
      }
    }
  }
];

exports.signInDialog = [
  (session, args) => {
    if (args && args.invalid) {
      // Re-prompt the user to click the link
      builder.Prompts.text(session, "Please click the sign-in link.");
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
];

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

exports.signout = [];


exports.emailDialog = [
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
];