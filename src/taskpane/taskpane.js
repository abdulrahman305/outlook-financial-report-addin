/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

  export async function run() {
    // Scan the user's inbox for all emails
    const mailbox = Office.context.mailbox;
    let insertAt = document.getElementById("item-subject");
    insertAt.innerHTML = "<b>Scanning inbox for financial emails...</b><br>";

    // Use the Outlook REST API to get all messages (local only, no external calls)
    // Office.js provides limited access; we use getCallbackTokenAsync for REST API
    mailbox.getCallbackTokenAsync({ isRest: true }, async function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const accessToken = result.value;
        const restUrl = mailbox.restUrl + "/v2.0/me/mailfolders/inbox/messages?$top=50";
        // Fetch messages using local REST API
        try {
          const response = await fetch(restUrl, {
            headers: {
              "Authorization": "Bearer " + accessToken,
              "Accept": "application/json"
            }
          });
          const data = await response.json();
          if (data.value && data.value.length > 0) {
            insertAt.innerHTML += `<b>Found ${data.value.length} emails.</b><br>`;
            // List subjects and prepare for financial data extraction
            data.value.forEach(email => {
              insertAt.innerHTML += `<b>Subject:</b> ${email.Subject}<br>`;
              // TODO: Extract financial data from email.Body
              // TODO: Check for PDF attachments and parse them
            });
          } else {
            insertAt.innerHTML += "No emails found in inbox.<br>";
          }
        } catch (err) {
          insertAt.innerHTML += "Error scanning inbox: " + err.message + "<br>";
        }
      } else {
        insertAt.innerHTML += "Failed to get access token for inbox scan.<br>";
      }
    });
  }
