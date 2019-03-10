var request = require('request');

// getChannelId() - Return promise of Channel ID with the specified name
module.exports = function getChannelId(context, token, teamId, channelName) {

    return new Promise((resolve, reject) => {

        const url = `https://graph.microsoft.com/beta/teams/${teamId}/channels?$filter=displayName+eq+'${channelName}'`;
        try {

            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    const result = JSON.parse(response.body);
                    if (result.value && result.value[0]) {
                        resolve(result.value[0].id);
                    } else {
                        reject("Channel not found in getChannelId");
                    }
                } else {
                    if (error) {
                        reject(`Error in getChannelId: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in getChannelId: ${b.error.message}`);
                    }
                }

            });

        }
        catch (ex) {
            context.log(`Error in getChannelId: ${ex}`);
        }

    });
}