/* eslint-disable camelcase */
const wait_for_operation = async (qnaClient, operation_id) => {
    let state = 'NotStarted';
    // eslint-disable-next-line no-undef-init
    let operationResult = undefined;

    while (state === 'Running' || state === 'NotStarted') {
        operationResult = await qnaClient.operations.getDetails(operation_id);
        state = operationResult.operationState;

        console.log(`Operation state - ${ state }`);

        await delayTimer(1000);
    }

    return operationResult;
};
const delayTimer = async (timeInMs) => {
    return await new Promise((resolve) => {
        setTimeout(resolve, timeInMs);
    });
};
const publishKnowledgeBase = async (kbclient, kb_id) => {
    const results = await kbclient.publish(kb_id);

    if (!results._response.status.toString().indexOf('2', 0) === -1) {
        console.log(`Publish request failed - HTTP status ${ results._response.status }`);
        return false;
    }

    console.log(`Publish request succeeded - HTTP status ${ results._response.status }`);

    return { kbID: kb_id, isPublished: true };
};
const updateKnowledgeBase = async (qnaClient, kbclient, kb_id, q, a, m) => {
    const qna = {
        answer: a,
        questions: [
            q
        ],
        metadata: [
            { name: 'Category', value: m.toLowerCase() }
        ]
    };

    // Add new Q&A lists, URLs, and files to the KB.
    const kb_add_payload = {
        qnaList: [
            qna
        ],
        files: []
    };

    // Bundle the add, update, and delete requests.
    const update_kb_payload = {
        add: kb_add_payload,
        update: null,
        delete: null,
        defaultAnswerUsedForExtraction: 'No answer found. Please rephrase your question.'
    };

    console.log(JSON.stringify(update_kb_payload));

    const results = await kbclient.update(kb_id, update_kb_payload);

    if (!results._response.status.toString().indexOf('2', 0) === -1) {
        console.log(`Update request failed - HTTP status ${ results._response.status }`);
        return false;
    }

    const operationResult = await wait_for_operation(qnaClient, results.operationId);

    if (operationResult.operationState !== 'Succeeded') {
        console.log(`Update operation state failed - HTTP status ${ operationResult._response.status }`);
        return { kbID: kb_id, result: operationResult.operationState };
    }

    console.log(`Update operation state ${ operationResult._response.status } - HTTP status ${ operationResult._response.status }`);
    return { kbID: kb_id, result: operationResult.operationState };
};

module.exports.updateKnowledgeBase = updateKnowledgeBase;
module.exports.publishKnowledgeBase = publishKnowledgeBase;
