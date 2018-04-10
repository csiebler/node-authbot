'use strict';

const https = require('https');


exports.getUserLatestEmail = function (accessToken, callback) {

  var options = {
    host: 'outlook.office.com',
    path: '/api/v2.0/me/MailFolders/Inbox/messages?$top=1',
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json',
      Authorization: 'Bearer ' + accessToken
    }
  };
  https.get(options, function (response) {
    var body = '';
    response.on('data', function (d) {
      body += d;
    });
    response.on('end', function () {
      var error;
      if (response.statusCode === 200) {
        callback(null, JSON.parse(body));
      } else {
        error = new Error();
        error.code = response.statusCode;
        console.log(response.statusMessage);
        error.message = response.statusMessage;
        // The error body sometimes includes an empty space
        // before the first character, remove it or it causes an error.
        // body = body.trim();
        // console.log(body);
        // error = body;
        callback(response.statusMessage, null);
      }
    });
  }).on('error', function (e) {
    callback(e, null);
  });
};

