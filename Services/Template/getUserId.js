var request = require('request');

module.exports = function getUserId(context, token, userPrincipalName) {

    return new Promise((resolve, reject) => {

        if (userPrincipalName.indexOf('@') < 0) {

            // If there's no @ sign, assume this is already a GUID
            resolve(userPrincipalName);
            
        } else {

            const url = `https://graph.microsoft.com/v1.0/users/${userPrincipalName}`;
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
                            reject(`User not found: ${userPrincipalName}`);
                        }

                    } else {

                        if (error) {
                            reject(`Error in getUserId: ${error}`);
                        } else {
                            let b = JSON.parse(response.body);
                            reject(`Error ${b.error.code} in getUserId: ${b.error.message}`);
                        }

                    }
                });
            } catch (ex) {
                reject(`Error in getUserId: ${ex}`);
            }
        }
    });
}