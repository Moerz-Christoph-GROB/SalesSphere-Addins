"use strict";

const MAX_RETRY_COUNT = 1;
const RETRY_DELAY_MS = 300;

/**
 * Registers event-based handlers after Office is ready.
 */
Office.onReady(function () {

    console.log("launcheventhandler.js: Office is ready.");

    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
});

/**
 * Handles the new message compose launch event.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onNewMessageComposeHandler(event) {

    console.log("onNewMessageComposeHandler: started.");
    handleComposeEvent(event);
}

/**
 * Handles the new appointment compose launch event.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onNewAppointmentOrganizer(event) {

    console.log("onNewAppointmentOrganizer: started.");
    handleComposeEvent(event);
}

/**
 * Handles the compose launch event and completes it when finished.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
async function handleComposeEvent(event) {

    try {
        await replaceItemContentWithRetry(0);
    }
    catch (objError) {
        console.error("handleComposeEvent: failed.", objError);
    }
    finally {
        event.completed();
    }
}

/**
 * Tries to replace the current item content and retries once if Office is still starting.
 * @param {number} intAttempt - The current retry attempt.
 * @returns {Promise<void>}
 */
async function replaceItemContentWithRetry(intAttempt) {

    const objReplaceItemContentNamespace = globalThis.SalesSphere
        && globalThis.SalesSphere.Outlook
        && globalThis.SalesSphere.Outlook.ReplaceItemContent;

    if (!objReplaceItemContentNamespace || typeof objReplaceItemContentNamespace.replaceItemContentFromCompose !== "function") {
        console.warn("replaceItemContentWithRetry: replaceItemContentFromCompose is not available.");
        return;
    }

    const objItem = Office.context && Office.context.mailbox
        ? Office.context.mailbox.item
        : null;

    if (!objItem || !objItem.subject || !objItem.body) {
        console.warn("replaceItemContentWithRetry: compose APIs are not ready. Attempt:", intAttempt);

        if (intAttempt < MAX_RETRY_COUNT) {
            await delayAsync(RETRY_DELAY_MS);
            await replaceItemContentWithRetry(intAttempt + 1);
        }

        return;
    }

    try {
        const blnWasReplaced = await objReplaceItemContentNamespace.replaceItemContentFromCompose();

        if (blnWasReplaced) {
            console.log("replaceItemContentWithRetry: item content replaced successfully.");
        }
        else {
            console.log("replaceItemContentWithRetry: no replacement was required.");
        }
    }
    catch (objError) {
        console.error("replaceItemContentWithRetry: replaceItemContentFromCompose failed.", objError);

        if (intAttempt < MAX_RETRY_COUNT) {
            await delayAsync(RETRY_DELAY_MS);
            await replaceItemContentWithRetry(intAttempt + 1);
            return;
        }

        throw objError;
    }
}

/**
 * Waits for the specified delay.
 * @param {number} delayMs - The delay in milliseconds.
 * @returns {Promise<void>}
 */
function delayAsync(delayMs) {

    return new Promise(function (resolve) {
        setTimeout(resolve, delayMs);
    });
}
