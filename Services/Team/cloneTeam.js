var request = require('request');

// Retry settings for asynchronous call
const NUMBER_OF_RETRIES = 20;
const RETRY_TIME_MSEC = 5 * 1000; // 5 sec

// cloneTeam() - Returns a promise to clone a Team and return its ID
module.exports = function cloneTeam(context, token, teamId, newTeam) {

  return new Promise((resolve, reject) => {

    const url = `https://graph.microsoft.com/beta/teams/${teamId}/clone`;

    context.log(`Cloning team with this URL ${url}`)
    request.post(url, {
      'auth': { 'bearer': token },
      'headers': { 'Content-Type': 'application/json' },
      'body': JSON.stringify({
        "displayName": newTeam,
        "description": "",
        "mailNickname": newTeam,
        "partsToClone": "apps,tabs,settings,channels,members",
        "visibility": "public"
      })
    }, (error, response, body) => {

      context.log(`Received a response with status code ${response.statusCode} error=${error}`);
      context.log(`Response ${response}`);
      context.log(`boolean ${(response && response.statusCode == 202)}`);

      if (response && response.statusCode == 202) {

        // If here we successfully issued the request
        const opUrl = `https://graph.microsoft.com/beta${response.headers.location}`;
        context.log(`operation url is ${opUrl}`);

        pollUntilDone(resolve, reject, opUrl, token, NUMBER_OF_RETRIES);

      } else {

        context.log(`Exception path response ${response.statusCode}`);
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

  function pollUntilDone(resolve, reject, opUrl, token, retryCount) {

    if (retryCount > 0) {

      // Poll the operation URL
      request.get(opUrl, {
        'auth': {
          'bearer': token
        }
      }, (error, response, body) => {

        context.log('Received response ' + response.statusCode);

        if (!error && response && response.statusCode == 200) {

          // If here we have a valid result
          const result = JSON.parse(response.body);
          if (result.status.toLowerCase() === 'succeeded') {
            // Success - resolve the promise
            resolve(result.targetResourceId);
          } else {
            // Not success - try again after waiting a few seconds
            console.log(`Received status ${result.status}`);
            setTimeout(() => {
              pollUntilDone(resolve, reject, opUrl, token, retryCount-1);
            }, RETRY_TIME_MSEC);
          }
        } else if (error) {
          context.log('Received HTTP error ' + error);
          reject(error);
        } else {
          context.log(`Invalid response ${response.statusCode}`);
          reject(`Invalid response ${response.statusCode}`);
        }

      });
    }
  }
}