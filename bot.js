
const { ActivityHandler } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');
const msRest = require('@azure/ms-rest-js');
const qnamaker = require('@azure/cognitiveservices-qnamaker');
const updateKB = require('./kbHelper');
// const qnamaker_runtime = require('@azure/cognitiveservices-qnamaker-runtime');

class QnABot extends ActivityHandler {
    constructor() {
        super();

        const creds = new msRest.ApiKeyCredentials({ inHeader: { 'Ocp-Apim-Subscription-Key': process.env.AuthoringKey } });
        const qnaMakerClient = new qnamaker.QnAMakerClient(creds, `https://${process.env.ResourceName}.cognitiveservices.azure.com`);
        const knowledgeBaseClient = new qnamaker.Knowledgebase(qnaMakerClient);

        try {
            this.qnaMaker = new QnAMaker({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
                endpointKey: process.env.QnAEndpointKey,
                host: process.env.QnAEndpointHostName
            });
        } catch (err) {
            console.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
        }

        // If a new user is added to the conversation, send them a greeting message
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let cnt = 0; cnt < membersAdded.length; cnt++) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(`Welcome to Rogers Digital Media! I'll be your Virtual Assistant, let me know how I can help you through your onboarding experience 🙂`);
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        // When a user sends a message, perform a call to the QnA Maker service to retrieve matching Question and Answer pairs.
        this.onMessage(async (context, next) => {
            if (!process.env.QnAKnowledgebaseId || !process.env.QnAEndpointKey || !process.env.QnAEndpointHostName) {
                const unconfiguredQnaMessage = 'NOTE: \r\n' +
                    'QnA Maker is not configured. To enable all capabilities, add `QnAKnowledgebaseId`, `QnAEndpointKey` and `QnAEndpointHostName` to the .env file. \r\n' +
                    'You may visit www.qnamaker.ai to create a QnA Maker knowledge base.';

                await context.sendActivity(unconfiguredQnaMessage);
            } else if (context._activity.text.includes('/')) {
                console.log('Adding Q to QnA Knowlwdge Base');
                const textArr = context._activity.text.split('/');
                const { kbID } = await updateKB.updateKnowledgeBase(qnaMakerClient, knowledgeBaseClient, process.env.QnAKnowledgebaseId, textArr[0], textArr[1], textArr[2].trim().toLowerCase());
                await updateKB.publishKnowledgeBase(knowledgeBaseClient, process.env.QnAKnowledgebaseId);
                await context.sendActivity(kbID + ' has been updated and published');
            } else {
                console.log('Calling QnA Maker');
                console.log(context);
                const qnaResults = !context._activity.text.includes(':') ? await this.qnaMaker.getAnswers(context) : await this.qnaMaker.getAnswers(context, {
                    strictFilters: [
                        {
                            name: 'Category', value: context._activity.text.split(':')[0].trim().toLowerCase()
                        }
                    ]
                });

                // If an answer was received from QnA Maker, send the answer back to the user.
                if (qnaResults[0]) {
                    await context.sendActivity(qnaResults[0].answer);

                // If no answers were returned from QnA Maker, reply with help.
                } else {
                    await context.sendActivity('No QnA Maker answers were found.');
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.QnABot = QnABot;
