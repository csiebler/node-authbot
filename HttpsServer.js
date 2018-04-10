'use strict';

const fs = require('fs');
const path = require('path');
const restify = require('restify');


module.exports = function () {

  var https_options = {
    key: fs.readFileSync(path.join(__dirname, 'etc/ssl/server.key')),
    certificate: fs.readFileSync(path.join(__dirname, 'etc/ssl/server.crt'))
  };

  var server = restify.createServer(https_options);
  server.listen(process.env.port || process.env.PORT || 3979, function () {
    console.log('%s listening to %s', server.name, server.url); 
  });
  
  return server;
};