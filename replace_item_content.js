"use strict";

(function (objRoot) {

    console.log("running replace_item_content.js");

    const INTEGRATION_SUBJECT = "salescockpitoutlookintegration";
    const ITEM_TYPE_EMAIL = "Email";
    const ITEM_TYPE_APPOINTMENT = "Appointment";


    objRoot.SalesSphere = objRoot.SalesSphere || {};
    objRoot.SalesSphere.Outlook = objRoot.SalesSphere.Outlook || {};

    objRoot.SalesSphere.Outlook.ReplaceItemContent = {
        replaceItemContentFromCompose: replaceItemContentFromCompose
    };

    /**
     * Replace the current Outlook compose item content when the integration trigger is detected.
     * Call this function from an Outlook compose event handler.
     * @returns {Promise<boolean>} True when the item was replaced; otherwise false.
     */
    async function replaceItemContentFromCompose() {


        if (typeof Office === "undefined") {
            console.error("replaceItemContentFromCompose: Office.js is not available.");
            return false;
        }

        if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
            console.warn("replaceItemContentFromCompose: No Outlook compose item is available.");
            return false;
        }

        const objItem = Office.context.mailbox.item;
        const strCurrentSubject = await getSubjectAsync(objItem);

        if (strCurrentSubject !== INTEGRATION_SUBJECT) {
            console.log("replaceItemContentFromCompose: Current subject does not match integration subject.");
            return false;
        }

        const strCurrentBodyText = await getBodyTextAsync(objItem);
        const objTriggerData = parseTriggerBody(strCurrentBodyText);

        if (objTriggerData === null) {
            console.warn("replaceItemContentFromCompose: Trigger body is invalid.");
            return false;
        }

        const objEmailData = await loadEmailDataAsync(objTriggerData.guid);

        if (objEmailData === null) {
            console.error("replaceItemContentFromCompose: No email data was returned from the server.");
            return false;
        }

        if (objTriggerData.itemType === ITEM_TYPE_EMAIL) {
            const blnIsMessage = objItem.itemType === Office.MailboxEnums.ItemType.Message;

            if (!blnIsMessage) {
                console.warn("replaceItemContentFromCompose: Email payload requires a message compose item.");
                return false;
            }

            await replaceMessageComposeContentAsync(objItem, objEmailData);

            console.log("replaceItemContentFromCompose: Message compose item replaced successfully.");
            return true;
        }

        if (objTriggerData.itemType === ITEM_TYPE_APPOINTMENT) {
            const blnIsAppointment = objItem.itemType === Office.MailboxEnums.ItemType.Appointment;

            if (!blnIsAppointment) {
                console.warn("replaceItemContentFromCompose: Appointment payload requires an appointment compose item.");
                return false;
            }

            await replaceAppointmentComposeContentAsync(objItem, objEmailData);

            console.log("replaceItemContentFromCompose: Appointment compose item replaced successfully.");
            return true;
        }

        console.warn("replaceItemContentFromCompose: Unsupported item type:", objTriggerData.itemType);
        return false;
    }

    /**
     * Replace a message compose item with cached server data.
     * @param {Office.MessageCompose} item
     * @param {Object} emailData
     * @returns {Promise<void>}
     */
    async function replaceMessageComposeContentAsync(item, emailData) {

        const arrToRecipients = splitRecipients(emailData.to);
        const arrCcRecipients = splitRecipients(emailData.cc);
        const strSubject = getStringValue(emailData.subject);
        const strBody = getStringValue(emailData.body);
        const arrAttachments = getAttachments(emailData.attachments);

        await setRecipientsAsync(item.to, arrToRecipients, "to.setAsync");
        await setRecipientsAsync(item.cc, arrCcRecipients, "cc.setAsync");
        await setSubjectAsync(item, strSubject);
        await setBodyHtmlAsync(item, strBody);
        await addAttachmentsAsync(item, arrAttachments);
    }

    /**
     * Replace an appointment compose item with cached server data.
     * @param {Office.AppointmentCompose} item
     * @param {Object} emailData
     * @returns {Promise<void>}
     */
    async function replaceAppointmentComposeContentAsync(item, emailData) {

        const arrRequiredAttendees = splitRecipients(emailData.to);
        const arrOptionalAttendees = splitRecipients(emailData.cc);
        const strSubject = getStringValue(emailData.subject);
        const strBody = getStringValue(emailData.body);
        const arrAttachments = getAttachments(emailData.attachments);

        await setRecipientsAsync(item.requiredAttendees, arrRequiredAttendees, "requiredAttendees.setAsync");
        await setRecipientsAsync(item.optionalAttendees, arrOptionalAttendees, "optionalAttendees.setAsync");
        await setSubjectAsync(item, strSubject);
        await setBodyHtmlAsync(item, strBody);
        await addAttachmentsAsync(item, arrAttachments);
    }

    async function loadEmailDataAsync(guid) {

        const strGuid = getStringValue(guid);

        if (strGuid.length === 0) {
            console.error("loadEmailDataAsync: GUID is missing.");
            return null;
        }

        const strRequestUrl = "https://salescockpit.grob.local/Email/api/LoadEmailFromCache/" + encodeURIComponent(strGuid);

        console.log("fetichng email content with url: " + strRequestUrl)

        const objResponse = await fetch(strRequestUrl, {
            method: "GET",
            credentials: "include"
        });

        if (!objResponse.ok) {
            const strErrorMessage = "loadEmailDataAsync: Request failed with status " + objResponse.status + ".";
            console.error(strErrorMessage);
            throw new Error(strErrorMessage);
        }

        const objResponseData = await objResponse.json();

        return objResponseData.payload;
    }

    /**
     * Parse the trigger body created by the mailto link.
     * @param {string} bodyText
     * @returns {{ itemType: string, guid: string }|null}
     */
    function parseTriggerBody(bodyText) {

        const strBodyText = getStringValue(bodyText);

        const objItemTypeMatch = strBodyText.match(/ItemType:\s*(Email|Appointment)\s*(?=GUID:|$)/i);
        const objGuidMatch = strBodyText.match(/GUID:\s*([0-9a-fA-F-]{36})/i);

        if (objItemTypeMatch === null || objGuidMatch === null) {
            return null;
        }

        return {
            itemType: objItemTypeMatch[1],
            guid: objGuidMatch[1]
        };
    }

    /**
     * Split a semicolon-separated recipient string into an array of recipient objects.
     * @param {string} recipientText
     * @returns {{ emailAddress: string, displayName: string }[]}
     */
    function splitRecipients(recipientText) {

        const strRecipientText = getStringValue(recipientText);

        if (strRecipientText.length === 0) {
            return [];
        }

        const arrRawRecipients = strRecipientText.split(";");
        const arrRecipients = [];

        for (let intIndex = 0; intIndex < arrRawRecipients.length; intIndex++) {
            const strRecipient = arrRawRecipients[intIndex].trim();

            if (strRecipient.length > 0) {
                arrRecipients.push({
                    emailAddress: strRecipient//,
                    //displayName: strRecipient
                });
            }
        }

        return arrRecipients;
    }

    /**
     * Normalize the attachments array.
     * @param {any} attachments
     * @returns {Array}
     */
    function getAttachments(attachments) {

        if (!Array.isArray(attachments)) {
            return [];
        }

        return attachments;
    }

    /**
     * Safely get a string value.
     * @param {any} value
     * @returns {string}
     */
    function getStringValue(value) {

        if (typeof value === "string") {
            return value;
        }

        if (value === null || value === undefined) {
            return "";
        }

        return String(value);
    }

    /**
     * Read the compose subject.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @returns {Promise<string>}
     */
    function getSubjectAsync(item) {

        return new Promise(function (resolve, reject) {

            try {
                item.subject.getAsync(function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(getStringValue(asyncResult.value));
                        return;
                    }

                    reject(createOfficeError("subject.getAsync", asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Read the compose body as plain text.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @returns {Promise<string>}
     */
    function getBodyTextAsync(item) {

        return new Promise(function (resolve, reject) {

            try {
                item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(getStringValue(asyncResult.value));
                        return;
                    }

                    reject(createOfficeError("body.getAsync", asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Write the compose subject.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @param {string} subject
     * @returns {Promise<void>}
     */
    function setSubjectAsync(item, subject) {

        return new Promise(function (resolve, reject) {

            const strSubject = getStringValue(subject);

            try {
                item.subject.setAsync(strSubject, function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                        return;
                    }

                    reject(createOfficeError("subject.setAsync", asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Replace the compose body with HTML.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @param {string} htmlBody
     * @returns {Promise<void>}
     */
    function setBodyHtmlAsync(item, htmlBody) {

        return new Promise(function (resolve, reject) {

            const strHtmlBody = getStringValue(htmlBody);
            const objOptions = {
                coercionType: Office.CoercionType.Html
            };

            try {
                item.body.setAsync(strHtmlBody, objOptions, function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                        return;
                    }

                    reject(createOfficeError("body.setAsync", asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Replace a recipient field.
     * @param {Object} recipientField
     * @param {string[]} recipients
     * @param {string} functionName
     * @returns {Promise<void>}
     */
    function setRecipientsAsync(recipientField, recipients, functionName) {

        return new Promise(function (resolve, reject) {

            const arrRecipients = Array.isArray(recipients) ? recipients : [];

            try {
                recipientField.setAsync(arrRecipients, function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                        return;
                    }

                    reject(createOfficeError(functionName, asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Add all attachments to the current compose item.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @param {Array} attachments
     * @returns {Promise<void>}
     */
    async function addAttachmentsAsync(item, attachments) {

        for (let intIndex = 0; intIndex < attachments.length; intIndex++) {
            const objAttachment = attachments[intIndex];

            await addAttachmentAsync(item, objAttachment);
        }
    }

    /**
     * Add a single base64 attachment to the current compose item.
     * @param {Office.MessageCompose|Office.AppointmentCompose} item
     * @param {Object} attachment
     * @returns {Promise<void>}
     */
    function addAttachmentAsync(item, attachment) {

        return new Promise(function (resolve, reject) {

            const strFileName = getStringValue(attachment.fileName);
            const strContent = getStringValue(attachment.content);

            if (strFileName.length === 0 || strContent.length === 0) {
                console.warn("addAttachmentAsync: Attachment was skipped because fileName or content is missing.", attachment);
                resolve();
                return;
            }

            try {
                item.addFileAttachmentFromBase64Async(strContent, strFileName, function (asyncResult) {

                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                        return;
                    }

                    reject(createOfficeError("addFileAttachmentFromBase64Async", asyncResult.error));
                });
            } catch (objError) {
                reject(objError);
            }
        });
    }

    /**
     * Create a normalized Office error.
     * @param {string} functionName
     * @param {any} officeError
     * @returns {Error}
     */
    function createOfficeError(functionName, officeError) {

        const strFunctionName = getStringValue(functionName);
        let strMessage = strFunctionName + " failed.";

        if (officeError && officeError.message) {
            strMessage = strFunctionName + " failed: " + officeError.message;
        }

        console.error(strFunctionName + ": Office.js error.", officeError);

        return new Error(strMessage);
    }

}(globalThis));