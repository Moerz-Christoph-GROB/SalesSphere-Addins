"use strict";

console.log("processing launcheventhandler.js");

/**
 * Handles the OnNewMessageCompose event before the task pane is opened.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onNewMessageComposeHandler(event) {

    console.log("onNewMessageComposeHandler: New message compose started.");

    Office.context.mailbox.item.subject.setAsync(
        "[VT-PR] ",
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("onNewMessageComposeHandler: Failed to set subject:", asyncResult.error.message);
            }

            // Always call event.completed - even on failure
            event.completed();
        }
    );
}

/**
 * Handles the OnMessageRecipientsChanged event.
 * @param {Office.AddinCommands.Event} event - The Office event object.
 */
function onItemChangedHandler(event) {

    console.log("onItemChangedHandler: Item changed.");
    event.completed();
}


/**
 * Entry point for event-based activation.
 * Registers all supported launch event handlers.
 */
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
// Office.actions.associate("onItemChangedHandler", onItemChangedHandler);
