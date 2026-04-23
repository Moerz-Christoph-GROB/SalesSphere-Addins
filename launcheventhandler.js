"use strict";

console.log("processing launcheventhandler.js");

/**
 * Executes when the Office Add-in has finished initializing.
 */
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office Add-in successfully loaded and ready.");
        // Add your Outlook-specific initialization code here
    }
});


function onNewMessageComposeHandler(event) {
  setSubject(event);
}
function onNewAppointmentComposeHandler(event) {
  setSubject(event);
}
function setSubject(event) {
  Office.context.mailbox.item.subject.setAsync(
    "Set by an event-based add-in!",
    {
      "asyncContext": event
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
      asyncResult.asyncContext.completed();
    });
}


/**
 * Entry point for event-based activation.
 * Registers all supported launch event handlers.
 */
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
// Office.actions.associate("onItemChangedHandler", onItemChangedHandler);

console.log("finished launcheventhandler.js");
