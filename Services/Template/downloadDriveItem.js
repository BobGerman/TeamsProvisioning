var request = require('request');

module.exports = function downloadDriveItem(context, token, url) {

    return new Promise((resolve, reject) => {

        try {

            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    if (response.body) {
                        resolve(response.body);
                    } else {
                        reject(`Download returned no body in response: ${url}`);
                    }

                } else {

                    if (error) {
                        reject(`Error in downloadDriveItem: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in downloadDriveItem: ${b.error.message}`);
                    }
                    
                }
            });
        } catch (ex) {
            reject(`Error in downloadDriveItem: ${ex}`);
        }
    });
}