const getToken = require('../Services/Token/getToken');
const getTemplate = require('../Services/Template/getTemplate');
const createTeam = require('../Services/Team/createTeam');

module.exports = async function (context, myQueueItem) {

    var token = "";

    return new Promise((resolve, reject) => {

        context.log('JavaScript queue trigger function processed work item', myQueueItem);

        if (myQueueItem && 
            myQueueItem.emailName && 
            myQueueItem.owners && 
            myQueueItem.jsonTemplate) {

            try {

                var emailName = myQueueItem.emailName;
                var displayName = myQueueItem.displayName || "New team";
                var description = myQueueItem.description || "";
                var owners = myQueueItem.owners;
                var jsonTemplate = myQueueItem.jsonTemplate;

                context.log(`Creating Team ${emailName} using ${jsonTemplate} json template`);

                getToken(context)
                .then ((accessToken) => {
                    context.log(`Got access token of ${accessToken.length} characters`);
                    token = accessToken;
                    return getTemplate(context, token, emailName, displayName,
                        description, owners, jsonTemplate);
                })                  
                .then((templateString) => {
                    return createTeam(context, token, templateString);
                })
                .then((newTeamId) => {
                    context.log(`createTeam created team ${newTeamId}`);
                    context.bindings.myOutputQueueItem = [newTeamId];
                    resolve();
                })
                .catch((error) => {
                    context.log(`ERROR: ${error}`);
                    //reject(error);
                    context.bindings.myOutputQueueItem = [error];
                    resolve();
                });

            } catch (ex) {
                context.log(`Error: ${ex}`);
                    //reject(ex);
                    context.bindings.myOutputQueueItem = [ex];
                    resolve();
            }
        } else {
            context.log('Skipping empty or invalid queue entry');
        }

    });

};