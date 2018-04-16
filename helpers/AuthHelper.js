'use strict';
const request = require('request');
require('dotenv').config();

//bot application identity
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

//oauth details
const AZUREAD_APP_ID = process.env.AZUREAD_APP_ID;
const AZUREAD_APP_PASSWORD = process.env.AZUREAD_APP_PASSWORD;
const AZUREAD_APP_REALM = process.env.AZUREAD_APP_REALM;
const AUTHBOT_CALLBACKHOST = process.env.AUTHBOT_CALLBACKHOST;

exports.getAccessTokenWithRefreshToken = function (refreshToken, callback) {
  console.log("getAccessTokenWithRefreshToken");
  var data = 'grant_type=refresh_token'
    + '&refresh_token=' + refreshToken
    + '&client_id=' + AZUREAD_APP_ID
    + '&client_secret=' + encodeURIComponent(AZUREAD_APP_PASSWORD)

  var options = {
    method: 'POST',
    url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
    body: data,
    json: true,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  };

  request(options, function (err, res, body) {
    if (err) return callback(err, body, res);
    if (parseInt(res.statusCode / 100, 10) !== 2) {
      if (body.error) {
        return callback(new Error(res.statusCode + ': ' + (body.error.message || body.error)), body, res);
      }
      if (!body.access_token) {
        return callback(new Error(res.statusCode + ': refreshToken error'), body, res);
      }
      return callback(null, body, res);
    }
    callback(null, {
      accessToken: body.access_token,
      refreshToken: body.refresh_token
    }, res);
  });
};



exports.getStrategy = function () {
  // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
  // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
  var realm = AZUREAD_APP_REALM;
  let oidStrategyv2 = {
    redirectUrl: AUTHBOT_CALLBACKHOST + '/api/OAuthCallback',
    realm: AZUREAD_APP_REALM,
    clientID: AZUREAD_APP_ID,
    clientSecret: AZUREAD_APP_PASSWORD,
    identityMetadata: 'https://login.microsoftonline.com/' + AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
    skipUserProfile: false,
    validateIssuer: false,
    //allowHttpForRedirectUrl: true,
    responseType: 'code',
    responseMode: 'query',
    scope: ['email', 'profile', 'offline_access', 'https://outlook.office.com/mail.read'],
    passReqToCallback: true
  };
  return oidStrategyv2;
};

