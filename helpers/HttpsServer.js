'use strict';
const fs = require('fs');
const path = require('path');
const restify = require('restify');
const expressSession = require('express-session');

module.exports = function () {

  var https_options = {
    key: fs.readFileSync(path.join(__dirname, '../etc/ssl/server.key')),
    certificate: fs.readFileSync(path.join(__dirname, '../etc/ssl/server.crt'))
  };

  var server = restify.createServer(https_options);
  server.listen(process.env.port || process.env.PORT || 3979, function () {
    console.log('%s listening to %s', server.name, server.url); 
  });

  server.get('/', restify.serveStatic({
    'directory': __dirname,
    'default': 'index.html'
  }));

  server.use(restify.queryParser());
  server.use(restify.bodyParser());
  server.use(expressSession({ secret: 'keyboard cat', resave: true, saveUninitialized: false }));

  
  return server;
};