/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, atob */
import * as pdfjsLib from "pdfjs-dist";

// Set workerSrc to load the PDF worker
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
                normalizeAndStoreData(financialData, "Email", email);
              }

              // Check for PDF attachments
              if (email.HasAttachments) {
                getAttachments(mailbox, email.Id, accessToken, email);
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

  async function getAttachments(mailbox, itemId, accessToken, email) {
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
            getAttachmentContent(mailbox, attachment.Id, accessToken, email);
          }
        });
      }
    } catch (err) {
      document.getElementById("item-subject").innerHTML += "Error getting attachments: " + err.message + "<br>";
    }
  }

  async function getAttachmentContent(mailbox, attachmentId, accessToken, email) {
      const restUrl = mailbox.restUrl + `/v2.0/me/attachments/${attachmentId}`;
      try {
          const response = await fetch(restUrl, {
              headers: {
                  "Authorization": "Bearer " + accessToken,
                  "Accept": "application/json"
              }
          });
          const attachment = await response.json();
          parsePdf(attachment.ContentBytes, email);
      } catch (err) {
          document.getElementById("item-subject").innerHTML += "Error getting attachment content: " + err.message + "<br>";
      }
  }

  async function parsePdf(pdfData, email) {
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
      normalizeAndStoreData(financialData, "PDF", email);
    }
  }

  function normalizeAndStoreData(data, sourceType, email) {
    data.forEach(item => {
        const amount = parseFloat(item.replace(/[^0-9.-]+/g,""));
        const category = categorizeTransaction(item, email?.Subject || '');
        financialDataStore.push({
            amount: amount,
            source: sourceType,
            date: email?.ReceivedDateTime ? new Date(email.ReceivedDateTime).toLocaleDateString() : new Date().toLocaleDateString(),
            subject: email?.Subject || 'Unknown',
            category: category
        });
    });
    renderFinancialDataTable();
  }

  function categorizeTransaction(amountStr, subject) {
    const subjectLower = subject.toLowerCase();
    if (subjectLower.includes('revenue') || subjectLower.includes('sales') || subjectLower.includes('income')) {
        return 'Revenue';
    } else if (subjectLower.includes('expense') || subjectLower.includes('cost') || subjectLower.includes('payment')) {
        return 'Expenses';
    } else if (subjectLower.includes('asset') || subjectLower.includes('inventory') || subjectLower.includes('equipment')) {
        return 'Assets';
    } else if (subjectLower.includes('liability') || subjectLower.includes('debt') || subjectLower.includes('loan')) {
        return 'Liabilities';
    } else if (subjectLower.includes('equity') || subjectLower.includes('capital') || subjectLower.includes('investment')) {
        return 'Equity';
    }
    return 'Uncategorized';
  }

  function renderFinancialDataTable() {
    const container = document.getElementById("financial-data-container");
    let table = '<h3>Extracted Financial Data</h3><table class="ms-Table"><thead><tr><th>Date</th><th>Subject</th><th>Category</th><th>Amount</th><th>Source</th></tr></thead><tbody>';
    financialDataStore.sort((a, b) => new Date(b.date) - new Date(a.date));
    financialDataStore.forEach(item => {
        table += `<tr><td>${item.date}</td><td>${item.subject}</td><td>${item.category}</td><td>$${item.amount.toFixed(2)}</td><td>${item.source}</td></tr>`;
    });
    table += '</tbody></table>';
    container.innerHTML = table;
  }

  function generateReport(period) {
    const container = document.getElementById("report-container");
    const now = new Date();
    let filteredData = [];
    let periodLabel = '';
    let startDate, endDate;

    if (period === 'quarterly') {
        const quarter = Math.floor(now.getMonth() / 3);
        startDate = new Date(now.getFullYear(), quarter * 3, 1);
        endDate = new Date(now.getFullYear(), startDate.getMonth() + 3, 0);
        periodLabel = `Q${quarter + 1} ${now.getFullYear()}`;
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= startDate && itemDate <= endDate;
        });
    } else if (period === 'semi-annual') {
        const semi = now.getMonth() < 6 ? 0 : 6;
        startDate = new Date(now.getFullYear(), semi, 1);
        endDate = new Date(now.getFullYear(), startDate.getMonth() + 6, 0);
        periodLabel = `${semi === 0 ? 'H1' : 'H2'} ${now.getFullYear()}`;
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= startDate && itemDate <= endDate;
        });
    } else if (period === 'annual') {
        startDate = new Date(now.getFullYear(), 0, 1);
        endDate = new Date(now.getFullYear(), 11, 31);
        periodLabel = `FY ${now.getFullYear()}`;
        filteredData = financialDataStore.filter(item => {
            const itemDate = new Date(item.date);
            return itemDate >= startDate && itemDate <= endDate;
        });
    }

    // Categorize data
    const categories = {
        'Assets': [],
        'Liabilities': [],
        'Equity': [],
        'Revenue': [],
        'Expenses': [],
        'Uncategorized': []
    };

    filteredData.forEach(item => {
        if (categories[item.category]) {
            categories[item.category].push(item);
        } else {
            categories['Uncategorized'].push(item);
        }
    });

    // Calculate totals
    const totalAssets = categories['Assets'].reduce((sum, item) => sum + item.amount, 0);
    const totalLiabilities = categories['Liabilities'].reduce((sum, item) => sum + item.amount, 0);
    const totalEquity = categories['Equity'].reduce((sum, item) => sum + item.amount, 0);
    const totalRevenue = categories['Revenue'].reduce((sum, item) => sum + item.amount, 0);
    const totalExpenses = categories['Expenses'].reduce((sum, item) => sum + item.amount, 0);
    const netIncome = totalRevenue - totalExpenses;

    // Generate report
    let report = `<div class="report-header">
        <h2>Financial Statement</h2>
        <p>Period: ${periodLabel}</p>
        <p>From ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}</p>
    </div>`;

    // Financial Summary
    report += `<div class="financial-summary">
        <p>Total Assets: $${totalAssets.toFixed(2)}</p>
        <p>Total Liabilities: $${totalLiabilities.toFixed(2)}</p>
        <p>Total Equity: $${totalEquity.toFixed(2)}</p>
        <p>Total Revenue: $${totalRevenue.toFixed(2)}</p>
        <p>Total Expenses: $${totalExpenses.toFixed(2)}</p>
        <p style="color: ${netIncome >= 0 ? 'green' : 'red'};">Net Income: $${netIncome.toFixed(2)}</p>
    </div>`;

    // Balance Sheet
    report += '<div class="report-section"><h4>Balance Sheet</h4>';
    report += '<table class="ms-Table"><thead><tr><th>Category</th><th>Description</th><th>Amount</th></tr></thead><tbody>';
    
    // Assets
    if (categories['Assets'].length > 0) {
        report += '<tr class="subtotal-row"><td colspan="3">ASSETS</td></tr>';
        categories['Assets'].forEach(item => {
            report += `<tr><td>Asset</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td></tr>`;
        });
        report += `<tr class="total-row"><td colspan="2">Total Assets</td><td>$${totalAssets.toFixed(2)}</td></tr>`;
    }

    // Liabilities
    if (categories['Liabilities'].length > 0) {
        report += '<tr class="subtotal-row"><td colspan="3">LIABILITIES</td></tr>';
        categories['Liabilities'].forEach(item => {
            report += `<tr><td>Liability</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td></tr>`;
        });
        report += `<tr class="total-row"><td colspan="2">Total Liabilities</td><td>$${totalLiabilities.toFixed(2)}</td></tr>`;
    }

    // Equity
    if (categories['Equity'].length > 0) {
        report += '<tr class="subtotal-row"><td colspan="3">EQUITY</td></tr>';
        categories['Equity'].forEach(item => {
            report += `<tr><td>Equity</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td></tr>`;
        });
        report += `<tr class="total-row"><td colspan="2">Total Equity</td><td>$${totalEquity.toFixed(2)}</td></tr>`;
    }

    report += '</tbody></table></div>';

    // Income Statement
    report += '<div class="report-section"><h4>Income Statement</h4>';
    report += '<table class="ms-Table"><thead><tr><th>Category</th><th>Description</th><th>Amount</th></tr></thead><tbody>';
    
    // Revenue
    if (categories['Revenue'].length > 0) {
        report += '<tr class="subtotal-row"><td colspan="3">REVENUE</td></tr>';
        categories['Revenue'].forEach(item => {
            report += `<tr><td>Revenue</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td></tr>`;
        });
        report += `<tr class="total-row"><td colspan="2">Total Revenue</td><td>$${totalRevenue.toFixed(2)}</td></tr>`;
    }

    // Expenses
    if (categories['Expenses'].length > 0) {
        report += '<tr class="subtotal-row"><td colspan="3">EXPENSES</td></tr>';
        categories['Expenses'].forEach(item => {
            report += `<tr><td>Expense</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td></tr>`;
        });
        report += `<tr class="total-row"><td colspan="2">Total Expenses</td><td>$${totalExpenses.toFixed(2)}</td></tr>`;
    }

    report += `<tr class="total-row" style="background-color: ${netIncome >= 0 ? '#d4f4dd' : '#f4d4d4'} !important;"><td colspan="2">Net Income</td><td>$${netIncome.toFixed(2)}</td></tr>`;
    report += '</tbody></table></div>';

    // Uncategorized Items
    if (categories['Uncategorized'].length > 0) {
        report += '<div class="report-section"><h4>Uncategorized Items</h4>';
        report += '<table class="ms-Table"><thead><tr><th>Date</th><th>Subject</th><th>Amount</th><th>Source</th></tr></thead><tbody>';
        categories['Uncategorized'].forEach(item => {
            report += `<tr><td>${item.date}</td><td>${item.subject}</td><td>$${item.amount.toFixed(2)}</td><td>${item.source}</td></tr>`;
        });
        report += '</tbody></table></div>';
    }

    container.innerHTML = report;
}