var getTeamId = require('./getTeamId');
var postCreate = require('./postCreate');

module.exports = async function (context, myQueueItem) {

    return new Promise((resolve, reject) => {

        context.log('JavaScript queue trigger function processed work item', myQueueItem);

        if (myQueueItem && myQueueItem.oldTeam && myQueueItem.newTeam) {

            try {

                var oldTeam = myQueueItem.oldTeam;
                var newTeam = myQueueItem.newTeam;
                context.log(`Cloning ${oldTeam} to ${newTeam}`);

                let token = context.bindings.graphToken;

                getTeamId(context, token, oldTeam)
                    .then((teamId) => {
                        context.log(`Got team ID ${teamId}`);
                        return postCreate(context, token, teamId, newTeam);
                    })
                    .then((opUrl) => {
                        context.log('postCreate resolved its promise');
                        context.bindings.myOutputQueueItem = [opUrl, "message 3"];
                        resolve();
                        //            context.done();
                    })
                    .catch((error) => {
                        context.log(`ERROR: ${error}`);
                        reject(error);
                        //            context.done();
                    })

            } catch (ex) {
                context.log(`Error: ${ex}`);
                reject(ex);
                //        context.done();
            }
        } else {
            context.log('Skipping empty queue entry');
        }

    });

};