// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, CardFactory } = require('botbuilder');

var sp = require("@pnp/sp").sp;
var SPFetchClient = require("@pnp/nodejs").SPFetchClient;

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient("https://m365x226741.sharepoint.com/", "1e572fff-b6b7-4ecc-aec2-43442779327e", "Fw2AZOWV0vhzk18vArdRWokrKO7zWILurEiqIImemo8=");
        },
    },
});

const EscalationManagementCard = require('./resources/EscalationManagement.json')


class EchoBot extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {             
            if (context.activity.value !== undefined) {
                switch (context.activity.value.type) {
                    case 'CreateTeams':
                        {
                            var escalationdetails = CardFactory.adaptiveCard(EscalationManagementCard).content.body[1].items[1].facts;
                            try {
                                sp.web.lists.getByTitle("Escalation Team Creation Request").items.add({
                                    "Title": escalationdetails[2].value + "Case" + escalationdetails[0].value,
                                    "TeamDescription": escalationdetails[1].value,
                                    "EscalationType":  escalationdetails[2].value
                                }).then(iar => {                                    
                                }).catch(err=>{
                                    console.log(iar);
                                    
                                });
                            }
                            catch (ex) {}       
                            await context.sendActivity("Your request has been created");                                                        
                            break;
                        }
                    case 'CreateGroupChat': await context.sendActivity("Hurray You clicked Group Chat"); break;
                }
            }
            else {
                await context.sendActivity({
                    attachments: [CardFactory.adaptiveCard(EscalationManagementCard)]
                });
            }
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity('Welcome to Escalation Management Chat Bot .');
                }
            }        
            await next();
            // By calling next() you ensure that the next BotHandler is run.
    
        });
    }
}

module.exports.EchoBot = EchoBot;
