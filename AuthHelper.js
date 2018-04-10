'use strict';

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
}