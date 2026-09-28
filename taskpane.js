"use strict";

(function (objRoot) {

    console.log("running taskpane.js");

    objRoot.SalesSphere = objRoot.SalesSphere || {};
    objRoot.SalesSphere.Outlook = objRoot.SalesSphere.Outlook || {};

    objRoot.SalesSphere.Outlook.CreateTaskPane = {
        createTaskPane: createTaskPane
    };

    // -------------------------------------------------------------------------
    // Initialization
    // -------------------------------------------------------------------------

    /**
     * Initialize the task pane: build the DOM and bind all button events.
     * @returns {void}
     */
    async function createTaskPane() {

        const divBodyContainer = document.getElementById("body-container");

        if (!divBodyContainer) {
            console.error("createTaskPane: Element 'body-container' was not found.");
            return;
        }

        divBodyContainer.appendChild(await buildTemplatesDiv());
        //divBodyContainer.appendChild(await buildDevelopmentDiv());
        divBodyContainer.appendChild(await buildStatusDiv());

        console.log("createTaskPane: Task pane initialized.");
    }

    async function buildTemplatesDiv() {

        const divTemplates = document.createElement("div");
        divTemplates.id = "div-templates";
        divTemplates.className = "section";

        const headerTemplates = document.createElement("h1");
        divTemplates.appendChild(headerTemplates);
        headerTemplates.textContent = "Templates";

        const divTemplatesList = document.createElement("div");
        divTemplates.appendChild(divTemplatesList);

        divTemplatesList.style.display = "flex";
        divTemplatesList.style.flexDirection = "column";


        const lstTemplates = await subFetchTemplateNames();

        for (let i = 0; i < lstTemplates.length; i++) {

            const currentTemplate = lstTemplates[i];

            const strName = currentTemplate.Name;
            const dblSize = currentTemplate.Size;
            const datModificationDate = currentTemplate.ModificationDate;

            const buttonTemplatePlaceholder = document.createElement("button");
            divTemplatesList.appendChild(buttonTemplatePlaceholder);
            buttonTemplatePlaceholder.id = "cmbTemplatePlaceholder" + strName;
            buttonTemplatePlaceholder.textContent = strName;
            buttonTemplatePlaceholder.onclick = function () { loadTemplateIntoActiveEmail(strName) };

        }

        return divTemplates;
    }

    async function buildDevelopmentDiv() {

        const divDevelopment = document.createElement("div");
        divDevelopment.id = "div-development";
        divDevelopment.className = "section";

        const headerDevelopment = document.createElement("h1");
        divDevelopment.appendChild(headerDevelopment);
        headerDevelopment.textContent = "Development";

        const buttonFetchAPIData = document.createElement("button");
        divDevelopment.appendChild(buttonFetchAPIData);
        buttonFetchAPIData.id = "cmbFetchAPIData";
        buttonFetchAPIData.textContent = "Fetch API data";

        buttonFetchAPIData.onclick = subFetchAPIData;

        return divDevelopment;
    }

    async function buildStatusDiv() {

        const divStatus = document.createElement("div");
        divStatus.id = "section-status";
        divStatus.className = "section";

        const headerStatus = document.createElement("h1");
        divStatus.appendChild(headerStatus);
        headerStatus.textContent = "Status";

        const paragraphStatusMessage = document.createElement("p");
        divStatus.appendChild(paragraphStatusMessage);
        paragraphStatusMessage.id = "status-message";

        return divStatus;
    }


    // -------------------------------------------------------------------------
    // Templates section
    // -------------------------------------------------------------------------

    /**
     * Placeholder handler for the Templates section.
     * @returns {void}
     */
    function subTemplatePlaceholder() {

        showStatus("Template placeholder clicked.");
    }

    async function loadTemplateIntoActiveEmail(TemplateName) {

        console.log("loading template " + TemplateName);

        const objTemplate = await subFetchTemplate(TemplateName);

        const strTo = objTemplate.To
        const strCc = objTemplate.Cc
        const strSubject = objTemplate.Subject
        const strHtmlBody = objTemplate.Body

        setMailToAsync(strTo);
        setMailCcAsync(strCc);
        setMailSubjectAsync(strSubject);
        //Office.context.mailbox.item.Subject.setAsync(strSubject);
        //Office.context.mailbox.item.Subject.setAsync(strSubject);
        //Office.context.mailbox.item.subject.setAsync(strSubject, function (asyncResult) {
        //    if (asyncResult.status === "failed") {
        //        console.log("Action failed with error: " + asyncResult.error.message);
        //    }
        //});

        setMailBodyAsync(strHtmlBody);

        showStatus("Template loaded.");

    }

    async function subFetchTemplateNames() {

        try {
            const strApiUrl = "https://salescockpit.grob.local/Email/api/GetEmailTemplates?sub_path=client";

            showStatus("Fetching data from API...");

            const objResponse = await fetch(strApiUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json"
                },
                credentials: "include",
                cache: "no-store"
            });

            if (!objResponse.ok) {
                throw new Error("HTTP status " + objResponse.status);
            }

            showStatus("Processing response data...");

            const objData = await objResponse.json();
            const arrRecords = objData.payload;
            const intRecordCount = arrRecords ? arrRecords.length : 0;

            if (intRecordCount < 1) {
                showStatus("No records found.");
                return;
            }

            showStatus("Done. Total template records: " + intRecordCount);

            return arrRecords;

        } catch (objError) {
            console.error("subFetchAPIData: Failed to fetch template records.", objError);
            showStatus("Error: " + objError.message);
        }
    }

    async function subFetchTemplate(TemplateName) {

        try {
            const strApiUrl = "https://salescockpit.grob.local/Email/api/LoadEmailTemplate?Name=" + TemplateName + "&sub_path=client";

            showStatus("Fetching data from API...");

            const objResponse = await fetch(strApiUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json"
                },
                credentials: "include",
                cache: "no-store"
            });

            if (!objResponse.ok) {
                throw new Error("HTTP status " + objResponse.status);
            }

            showStatus("Processing response data...");

            const objData = await objResponse.json();
            const objTemplate = objData.payload;


            return objTemplate;

        } catch (objError) {
            console.error("subFetchAPIData: Failed to fetch template.", objError);
            showStatus("Error: " + objError.message);
        }
    }


    // -------------------------------------------------------------------------
    // Development section – Urlaubsantrag
    // -------------------------------------------------------------------------

    /**
     * Fetch project records from the API and write the first record to the mail body.
     * @returns {Promise<void>}
     */
    async function subFetchAPIData() {

        try {
            const strApiUrl = "https://salescockpit.grob.local/ProjectReferences/api/ProjectRecords";

            showStatus("Fetching data from API...");

            const objResponse = await fetch(strApiUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json"
                },
                credentials: "include",
                cache: "no-store"
            });

            if (!objResponse.ok) {
                throw new Error("HTTP status " + objResponse.status);
            }

            showStatus("Processing response data...");

            const objData = await objResponse.json();
            const arrRecords = objData.payload;
            const intRecordCount = arrRecords ? arrRecords.length : 0;

            if (intRecordCount < 1) {
                showStatus("No records found.");
                return;
            }

            const objFirstRecord = arrRecords[0];
            const strFirstRecordJson = JSON.stringify(objFirstRecord, null, 2);
            const strMailBody = "<pre>" + strFirstRecordJson + "</pre>";

            await setMailBodyAsync(strMailBody);

            showStatus("Done. First record written to mail body. Total records: " + intRecordCount);

        } catch (objError) {
            console.error("subFetchAPIData: Failed to fetch project reference records.", objError);
            showStatus("Error: " + objError.message);
        }
    }


    // -------------------------------------------------------------------------
    // Office helpers
    // -------------------------------------------------------------------------

    /**
     * Read the current mail body as HTML.
     * @returns {Promise<string>}
     */
    function getMailBodyAsync() {

        return new Promise(function (resolve, reject) {

            Office.context.mailbox.item.body.getAsync("html", function (objResult) {

                if (objResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(objResult.value);
                    return;
                }

                reject(new Error(objResult.error.message));
            });
        });
    }

    /**
     * Replace the current mail body with HTML content.
     * @param {string} strHtmlContent - The HTML content to write.
     * @returns {Promise<void>}
     */
    function setMailBodyAsync(strHtmlContent) {

        return new Promise(function (resolve, reject) {

            Office.context.mailbox.item.body.setAsync(strHtmlContent, { coercionType: "html" }, function (objAsyncResult) {

                if (objAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                    return;
                }

                reject(new Error(objAsyncResult.error.message));
            });
        });
    }


    /**
 * Set the mail subject.
 * @param {string} strSubject - The subject text to set.
 * @returns {Promise<void>}
 */
    function setMailSubjectAsync(strSubject) {

        return new Promise(function (resolve, reject) {

            Office.context.mailbox.item.subject.setAsync(strSubject, function (objAsyncResult) {

                if (objAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                    return;
                }

                reject(new Error(objAsyncResult.error.message));
            });
        });
    }

    /**
     * Set the mail To recipients from a semicolon-separated string of email addresses.
     * @param {string} strEmailAddresses - Semicolon-separated email addresses (e.g. "a@x.com;b@x.com").
     * @returns {Promise<void>}
     */
    function setMailToAsync(strEmailAddresses) {

        const arrAddresses = strEmailAddresses.split(";");

        return new Promise(function (resolve, reject) {

            Office.context.mailbox.item.to.setAsync(arrAddresses, function (objAsyncResult) {

                if (objAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                    return;
                }

                reject(new Error(objAsyncResult.error.message));
            });
        });
    }

    /**
     * Set the mail CC recipients from a semicolon-separated string of email addresses.
     * @param {string} strEmailAddresses - Semicolon-separated email addresses (e.g. "a@x.com;b@x.com").
     * @returns {Promise<void>}
     */
    function setMailCcAsync(strEmailAddresses) {

        const arrAddresses = strEmailAddresses.split(";");

        return new Promise(function (resolve, reject) {

            Office.context.mailbox.item.cc.setAsync(arrAddresses, function (objAsyncResult) {

                if (objAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                    return;
                }

                reject(new Error(objAsyncResult.error.message));
            });
        });
    }


    // -------------------------------------------------------------------------
    // Status helpers
    // -------------------------------------------------------------------------

    /**
     * Display a status message in the status section.
     * @param {string} strMessage - The message to display.
     * @returns {void}
     */
    function showStatus(strMessage) {

        const pStatusMessage = document.getElementById("status-message");

        if (!pStatusMessage) {
            console.warn("showStatus: Element 'status-message' was not found.");
            return;
        }

        pStatusMessage.innerHTML = strMessage;
    }

}(globalThis));