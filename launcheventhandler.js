"use strict";

/**
 * Registers event-based handlers after Office is ready.
 */
Office.onReady(function () {

    console.log("launcheventhandler.js: Office is ready.");

    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
});

/**
 * Handles the new message compose launch event.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onNewMessageComposeHandler(event) {

    console.log("onNewMessageComposeHandler: started.");
    setSubjectWithRetry(event);
}

/**
 * Handles the new appointment compose launch event.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onNewAppointmentComposeHandler(event) {

    console.log("onNewAppointmentComposeHandler: started.");
    setSubjectWithRetry(event);
}

/**
 * Sets the item subject. Retries once because OWA can be early during compose startup.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function setSubjectWithRetry(event) {

    trySetSubject(event, 0);
}

/**
 * Tries to set the subject and retries once if needed.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 * @param {number} intAttempt - The current retry attempt.
 */
function trySetSubject(event, intAttempt) {

    var objItem = Office.context.mailbox.item;

    if (!objItem || !objItem.subject || typeof objItem.subject.setAsync !== "function") {
        console.warn("trySetSubject: subject API not ready. Attempt:", intAttempt);

        if (intAttempt < 1) {
            setTimeout(function () {
                trySetSubject(event, intAttempt + 1);
            }, 300);

            return;
        }

        event.completed();
        return;
    }

    objItem.subject.setAsync(
        "Set by an event-based add-in!",
        function (asyncResult) {

            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("trySetSubject: setAsync failed:", asyncResult.error);

                if (intAttempt < 1) {
                    setTimeout(function () {
                        trySetSubject(event, intAttempt + 1);
                    }, 300);

                    return;
                }
            }
            else {
                console.log("trySetSubject: subject set successfully.");
            }

            event.completed();
        }
    );
}
