var adal = require('adal-node');
var request = require('request');

const GRAPH_URL = "https://graph.microsoft.com";
const TENANT = "{tenant-name-here}.onmicrosoft.com";
const CLIENT_ID = "{Application-id-here}";
const CLIENT_SECRET = "{Application-key-here}";
const invitationRedirectUrl ="https://{tenant-name-here}.sharepoint.com/sites/{collection-here}";

function getToken() {
    return new Promise((resolve, reject) => {
        const authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/${TENANT}`);
        authContext.acquireTokenWithClientCredentials(GRAPH_URL, CLIENT_ID, CLIENT_SECRET, (err, tokenRes) => {
            if (err) { reject(err); }
            resolve(tokenRes.accessToken);
        });
    });
}

function createResponse(body, status) {
    return {
        status: status || 200,
        body: body,
        headers: {
            'Content-Type': 'application/json'
        }
    };
}

function createGraphAPIRequestOptions(token, method, url, body) {
    return {
        method: method,
        url: `https://graph.microsoft.com/v1.0/${url}`,
        headers: {
            'Authorization': 'Bearer ' + token,
            'content-type': 'application/json'
        },
        body: JSON.stringify(body)
    };
}

function getGroups(token) {
    return new Promise((resolve, reject) => {
        const options = createGraphAPIRequestOptions(token, 'GET', "groups?$filter=securityEnabled+eq+true+and+mailEnabled+eq+false+and+startswith(displayName,'guests-')");

        request(options, (error, response, body) => {
            const result = JSON.parse(body);
            if (!error && response.statusCode == 200) {
                resolve(result.value);
            } else {
                reject(result);
            }
        });
    });
}

function countGroupMembers(token, groupId) {
    return new Promise((resolve, reject) => {
        const options = createGraphAPIRequestOptions(token, 'GET', `groups/${groupId}/members`);

        request(options, (error, response, body) => {
            const result = JSON.parse(body);
            if (!error && response.statusCode == 200) {
                resolve(result.value.length);
            } else {
                reject(result);
            }
        });
    });
}

function createGroup(token, name) {
    return new Promise((resolve, reject) => {
        const options = createGraphAPIRequestOptions(token, 'POST', `groups/`, {
            "displayName": name,
            "mailNickname": name,
            "mailEnabled": false,
            "securityEnabled": true
        });

        request(options, (error, response, body) => {
            const result = JSON.parse(body);
            if (!error && response.statusCode == 204) {
                resolve(result.value);
            } else {
                reject(result);
            }
        });
    });
}

function getGroupID(token) {
    return new Promise((resolve, reject) => {
        getGroups(token)
            .then(groups => {
                if (groups && groups.length > 0) {
                    groups.sort((g1, g2) => {
                        const nameA = g1.displayName.toLowerCase();
                        const nameB = g2.displayName.toLowerCase();
                        if (nameA > nameB) return -1;
                        if (nameA < nameB) return 1;
                        return 0;
                    });

                    const lastGroup = groups[0];
                    countGroupMembers(token, lastGroup.id)
                        .then(groupMembersAmount => {
                            if (groupMembersAmount >= 4800) {
                                const groupNumber = ("00" + groups.length).slice(-2);
                                createGroup(token, `guests-${groups.length}` )
                                    .then(groupInfo => {
                                        resolve(groupInfo.id);
                                    })
                                    .catch(() => reject());
                            } else {
                                resolve(lastGroup.id);
                            }
                        })
                        .catch(() => reject());;
                } else {
                    createGroup(token, `guests-00`)
                        .then(groupInfo => {
                            resolve(groupInfo.id);
                        })
                        .catch(() => reject());
                }
            })
            .catch(() => reject());
    });
}

function addUserToGroupId(token, groupId, invitedUserId, context) {
    return new Promise((resolve, reject) => {
        const options = createGraphAPIRequestOptions(token, 'POST', `groups/${groupId}/members/$ref`, {
            "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${invitedUserId}`
        });

        request(options, (error, response, body) => {
            if (!error && response.statusCode == 204) {
                resolve(result);
            } else {
                reject(result);
            }
        });
    });
}

module.exports = function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    if (req.query.email && req.query.name) {
        const userEmail = req.query.email;
        const userDisplayName = req.query.name;
        
        const sendInvitationMessage = req.query.invitation && req.query.invitation == "true";
        
        getToken().then(token => {
            /* INVITE A USER TO YOUR TENANT */
            const options = createGraphAPIRequestOptions(token, 'POST', `invitations`, {
                "invitedUserDisplayName": userDisplayName,
                "invitedUserEmailAddress": userEmail,
                "inviteRedirectUrl": invitationRedirectUrl,
                "sendInvitationMessage": sendInvitationMessage,
                "invitedUserMessageInfo": {
                    "customizedMessageBody": "This is a custom body message"
                }
            });

            request(options, (error, response, body) => {
                if (!error && response.statusCode == 201) {
                    const result = JSON.parse(body);

                    getGroupID(token)
                        .then(groupId => {
                            addUserToGroupId(token, groupId, result.invitedUser.id)
                                .then(() => {
                                    context.res = createResponse({
                                        id: result.invitedUser.id,
                                        inviteRedeemUrl: result.inviteRedeemUrl,
                                        status: result.status
                                    });
                                    context.done();
                                }).catch(() => {
                                    context.done();
                                });
                        }).catch(() => {
                            context.done();
                        });
                } else {
                    context.res = createResponse(result, response.statusCode);
                    context.done();
                }
            });
        });
    } else {
        context.res = {
            status: 400,
            body: "Please pass a name and email on the query string or in the request body"
        };
        
        context.done();
    }
};