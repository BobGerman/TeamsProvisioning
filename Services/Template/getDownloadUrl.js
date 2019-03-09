var request = require('request');

module.exports = function getDownloadUrl(context, token, driveId, fileName) {

    return new Promise((resolve, reject) => {

        const url = `https://graph.microsoft.com/v1.0/drives/` + 
                    `${driveId}/root/children?$filter=name+eq+'${fileName}'`;

        try {

            request.get(url, {
                'auth': {
                    'bearer': token
                }
            }, (error, response, body) => {

                if (!error && response && response.statusCode == 200) {

                    const result = JSON.parse(response.body);
                    if (result.value && result.value[0]) {
                        resolve(result.value[0]["@microsoft.graph.downloadUrl"]);
                    } else {
                        reject(`File not found in getDownloadUrl: ${fileName}`);
                    }

                } else {

                    if (error) {
                        reject(`Error in getDownloadUrl: ${error}`);
                    } else {
                        let b = JSON.parse(response.body);
                        reject(`Error ${b.error.code} in getDownloadUrl: ${b.error.message}`);
                    }
                    
                }
            });
        } catch (ex) {
            reject(`Error in getDownloadUrl: ${ex}`);
        }


    });
}