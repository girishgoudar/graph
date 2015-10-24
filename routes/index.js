var express = require('express');
var router = express.Router();
var authContext = require('adal-node').AuthenticationContext;
var authHelper = require('../authHelper.js');
var https = require('https')

/* GET home page. */
router.get('/', function(req, res, next) {
  if(req.cookies.TOKEN_CACHE_KEY === undefined){
    res.redirect('/login')
  }
  else{
    authHelper.getTokenFromRefreshToken('https://graph.microsoft.com/',req.cookies.TOKEN_CACHE_KEY1,function(token){
      getJson('graph.microsoft.com','/beta/me/drive/root/children',token.accessToken, function(json){
        res.render('index', { title: 'MyFiles',files:JSON.parse(json) });
      });
    });
  }
 });

router.get('/login', function(req, res, next) {
  if(req.query.code === undefined){
    res.render('login',{title :'login to office 365', authRedirect:authHelper.getAuthUrl('https://graph.microsoft.com/')})
  }
  else{
    authHelper.getTokenFromCode('https://graph.microsoft.com/',req.query.code,function(token){
     res.cookie(authHelper.TOKEN_CACHE_KEY , token.refreshToken);
     res.cookie(authHelper.TENANT_CAHCE_KEY, token.tenantId) 
     res.redirect('/')
    });
  }
});

function getJson(host, path, token, callback) {
  var options = {
    host: host, 
    path: path, 
    method: 'GET', 
    headers: { 
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'Authorization': 'Bearer ' + token 
      }
    };

  https.get(options, function(response) {
    var body = "";
    response.on('data', function(d) {
      body += d;
    });
    response.on('end', function() {
      callback(body);
    });
    response.on('error', function(e) {
      callback(null);
    });
  });
};

module.exports = router;
