/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, atob */
import * as pdfjsLib from "pdfjs-dist/build/pdf";

// Set workerSrc to load the PDF worker from a CDN
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

let financialDataStore = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("generate-quarterly").onclick = () => generateReport('quarterly');
    document.getElementById("generate-semi-annual").onclick = () => generateReport('semi-annual');
    document.getElementById("generate-annual").onclick = () => generateReport('annual');
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
        const restUrl = mailbox.restUrl + "/v2.0/me/mailfolders/inbox/messages?$top=10&$select=Subject,Body,HasAttachments"; // Request body and attachments
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
              
              // Extract financial data from email body
              const financialData = extractFinancialData(email.Body.Content);
              if (financialData.length > 0) {
                insertAt.innerHTML += `Found financial data: ${financialData.join(", ")}<br>`;
                normalizeAndStoreData(financialData, "Email");
              }

              // Check for PDF attachments
              if (email.HasAttachments) {
                getAttachments(mailbox, email.Id, accessToken);
              }
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
  
  function extractFinancialData(emailBody) {
    const currencyRegex = /\$[0-9,]+(\.[0-9]{2})?/g;
    const matches = emailBody.match(currencyRegex);
    return matches || [];
  }

  async function getAttachments(mailbox, itemId, accessToken) {
    const restUrl = mailbox.restUrl + `/v2.0/me/messages/${itemId}/attachments`;
    try {
      const response = await fetch(restUrl, {
        headers: {
          "Authorization": "Bearer " + accessToken,
          "Accept": "application/json"
        }
      });
      const data = await response.json();
      if (data.value && data.value.length > 0) {
        data.value.forEach(attachment => {
          if (attachment.ContentType === "application/pdf") {
            document.getElementById("item-subject").innerHTML += `Found PDF: ${attachment.Name}<br>`;
            // We need to get the attachment content
            getAttachmentContent(mailbox, attachment.Id, accessToken);
          }
        });
      }
    } catch (err) {
      document.getElementById("item-subject").innerHTML += "Error getting attachments: " + err.message + "<br>";
    }
  }

  async function getAttachmentContent(mailbox, attachmentId, accessToken) {
      const restUrl = mailbox.restUrl + `/v2.0/me/attachments/${attachmentId}`;
      try {
          const response = await fetch(restUrl, {
              headers: {
                  "Authorization": "Bearer " + accessToken,
                  "Accept": "application/json"
              }
          });
          const attachment = await response.json();
          parsePdf(attachment.ContentBytes);
      } catch (err) {
          document.getElementById("item-subject").innerHTML += "Error getting attachment content: " + err.message + "<br>";
      }
  }

  async function parsePdf(pdfData) {
    const pdf = await pdfjsLib.getDocument({ data: atob(pdfData) }).promise;
    let pdfText = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      pdfText += textContent.items.map(item => item.str).join(" ");
    }
    
    const financialData = extractFinancialData(pdfText);
    if (financialData.length > 0) {
      document.getElementById("item-subject").innerHTML += `Found financial data in PDF: ${financialData.join(", ")}<br>`;
      normalizeAndStoreData(financialData, "PDF");
    }
  }

  function normalizeAndStoreData(data, sourceType) {
    data.forEach(item => {
        const amount = parseFloat(item.replace(/[^0-9.-]+/g,""));
        financialDataStore.push({
            amount: amount,
            source: sourceType,
            date: new Date().toLocaleDateString()
        });
    });
    renderFinancialDataTable();
  }

  function renderFinancialDataTable() {
    const container = document.getElementById("financial-data-container");
    let table = '<h3>Organized Financial Data</h3><table class="ms-Table"><thead><tr><th>Date</th><th>Amount</th><th>Source</th></tr></thead><tbody>';
    financialDataStore.forEach(item => {
        table += `<tr><td>${item.date}</td><td>$${item.amount.toFixed(2)}</td><td>${item.source}</td></tr>`;
    });
    table += '</tbody></table>';
    container.innerHTML = table;
  }

  function generateReport(period) {
    const container = document.getElementById("report-container");
    const now = new Date();
    let filteredData = [];

    if (period === 'quarterly') {
        const quarter = Math.floor(now.getMonth() / 3);
        const start = new Date(now.getFullYear(), quarter * 3, 1);
        const end = new Date(now.getFullYear(), start.getMonth() + 3, 0);
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= start && itemDate <= end;
        });
    } else if (period === 'semi-annual') {
        const semi = now.getMonth() < 6 ? 0 : 6;
        const start = new Date(now.getFullYear(), semi, 1);
        const end = new Date(now.getFullYear(), start.getMonth() + 6, 0);
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= start && itemDate <= end;
        });
    } else if (period === 'annual') {
        const start = new Date(now.getFullYear(), 0, 1);
        const end = new Date(now.getFullYear(), 11, 31);
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= start && itemDate <= end;
        });
    }

    const total = filteredData.reduce((sum, item) => sum + item.amount, 0);

    let report = `<h3>${period.charAt(0).toUpperCase() + period.slice(1)} Financial Report</h3>`;
    report += `<p>Total: $${total.toFixed(2)}</p>`;
    report += '<table class="ms-Table"><thead><tr><th>Date</th><th>Subject</th><th>Amount</th><th>Source</th></tr></thead><tbody>';
    filteredData.forEach(item => {
        report += `<tr><td>${item.date}</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td><td>${item.source}</td></tr>`;
    });
    report += '</tbody></table>';
    container.innerHTML = report;
}