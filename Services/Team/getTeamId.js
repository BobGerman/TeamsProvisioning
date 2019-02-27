var request = require('request');

// getListId() - Return promise of Team ID with the specified name
module.exports = function getTeamId(context, token, teamName) {

    context.log('Getting team ID for ' + teamName);

    return new Promise((resolve, reject) => {

        const url = `https://graph.microsoft.com/beta/groups` +
                        `?$filter=(resourceProvisioningOptions/Any(x:x eq 'Team')) ` +
                        ` and mailNickname eq '${teamName}'`
        try {

            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                context.log('Received response ' + response.statusCode);

                if (!error && response && response.statusCode == 200) {

                    // If here we have all the lists
                    const result = JSON.parse(response.body);
                    if (result.value && result.value[0]) {
                        resolve(result.value[0].id);
                    } else {
                        reject("Team not found");
                    }
                } else {
                    if (error) {
                        context.log('Received error ' + error);
                        reject(error);
                    } else {
                        let b = JSON.parse(response.body);
                        context.log('Received error ' + error);

                        reject(`${b.error.code} - ${b.error.message} - ${token}`);
                    }
                }

            });

            context.log('Requested ' + url);

        }
        catch (ex) {
            context.log('Error ' + ex);
        }

    });
}