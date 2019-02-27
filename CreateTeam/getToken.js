var adal = require('adal-node');
var settings = require('./settings');

// getToken() - Return promise of auth token for app
// permission (client credentials flow)
module.exports = function getToken(context) {

  return new Promise((resolve, reject) => {

    // Get the ADAL client
    const authContext = new adal.AuthenticationContext(
      settings().AUTH_URL + settings().TENANT);
    const clientId = settings().CLIENT_ID;
    const clientSecret = settings().CLIENT_SECRET;
      
    authContext.acquireTokenWithClientCredentials(
      settings().GRAPH_URL, settings().CLIENT_ID, secret, 
      (err, tokenRes) => {
        if (err) {
          reject(err); 
        } else {
          resolve(tokenRes.accessToken);
        }
    });
  });
}