var request = require('request');

// postCreate() - Returns a promise to clone a Team
// with the provided token
module.exports = function postCreate(context, token, teamId, newTeam) {

  return new Promise((resolve, reject) => {

    const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/clone`;

    context.log (`Cloning team with this URL ${url}`)
    request.post(url, {
        'auth': { 'bearer': token },
        'headers': {'Content-Type': 'application/json' },
        'body': JSON.stringify({
            "displayName": newTeam,
            "description": "Test cloning via graph API 2",
            "mailNickname": newTeam,
            "partsToClone": "apps,tabs,settings,channels,members",
            "visibility": "public"
        })
    }, (error, response, body) => {

      context.log (`Received a response with status code ${response.statusCode} error=${error}`);
      context.log (`Response ${response}`);
      context.log (`boolean ${(response && response.statusCode == 202)}`);

      if (response && response.statusCode == 202) {

        // If here we were successful
        context.log(`Parsing JSON response=${response}`);
        const result = JSON.parse(response.headers);
        context.log (`result is ${result}`);
        const opUrl = result.opUrl;
        context.log(`Operation ${opUrl}`);

        resolve(opUrl);

      } else {
        
        context.log (`Exception path response ${response.statusCode}`);
        // If here something went wrong, reject with an error
        // message
        if (error) {
            reject(error);
        } else {
          let b = JSON.parse(response.body);
          reject(`${b.error.code} - ${b.error.message}`);
        }
      }

    });
  });
}