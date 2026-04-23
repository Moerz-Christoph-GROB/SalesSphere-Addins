// launcheventhandler.js
"use strict";

function onNewMessageComposeHandler(event) {
    setSubject(event);
}

function onNewAppointmentComposeHandler(event) {
    setSubject(event);
}

function setSubject(event) {
    Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // If it fails, there is no way to see this in the OWA console.
            }
            // You MUST call completed() so OWA doesn't hang the compose window.
            event.completed(); 
        }
    );
}

// Global Registration
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
