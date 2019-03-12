var request = require('request');

module.exports = function getSiteId(context, token, tenantName,
    serverRelativeUrl) {

    return new Promise((resolve, reject) => {

        const url = `https://graph.microsoft.com/v1.0/sites/` +
            `${tenantName}.sharepoint.com:${serverRelativeUrl}`;
        try {

            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    const result = JSON.parse(response.body);
                    if (result.id) {
                        resolve(result.id);
                    } else {
                        reject(`Site not found: ${serverRelativeUrl}`);
                    }

                } else {

                    if (error) {
                        reject(`Error in getSiteId: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in getSiteId: ${b.error.message}`);
                    }

                }
            });
        } catch (ex) {
            reject(`Error in getSiteId: ${ex}`);
        }
    });
}