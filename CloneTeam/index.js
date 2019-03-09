const getToken = require('../Services/Token/getToken');
const getTeamId = require('../Services/Team/getTeamId');
const cloneTeam = require('../Services/Team/cloneTeam');

module.exports = async function (context, myQueueItem) {

    var token = "";

    return new Promise((resolve, reject) => {

        context.log('JavaScript queue trigger function processed work item', myQueueItem);

        if (myQueueItem && myQueueItem.oldTeam && myQueueItem.newTeam) {

            try {
                var oldTeam = myQueueItem.oldTeam;
                var newTeam = myQueueItem.newTeam;
                context.log(`Cloning ${oldTeam} to ${newTeam}`);

                getToken(context)
                    .then((accessToken) => {
                        context.log(`Got access token of ${accessToken.length} characters`);
                        token = accessToken;
                        return getTeamId(context, token, oldTeam);
                    })
                    .then((teamId) => {
                        context.log(`Got team ID ${teamId}`);
                        return cloneTeam(context, token, teamId, newTeam);
                    })
                    .then((newTeamId) => {
                        context.log(`cloneTeam created team ${newTeamId}`);
                        context.bindings.myOutputQueueItem = {
                            success: true,
                            teamId: newTeamId,
                            error: ''
                        };
                        resolve();
                    })
                    .catch((error) => {
                        context.log(`ERROR: ${error}`);
                        context.bindings.myOutputQueueItem = {
                            success: false,
                            teamId: '',
                            error: error
                        };
                        resolve();
                    })

            } catch (ex) {
                context.log(`ERROR: ${error}`);
                context.bindings.myOutputQueueItem = {
                    success: false,
                    teamId: '',
                    error: error
                };
                resolve();
            }
        } else {
            context.log('Skipping empty queue entry');
        }

    });

};