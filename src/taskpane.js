const SHEETS = {
  transactions: "Transactions",
  template: "Invoice Template",
  configuration: "Configuration",
  currencies: "Currencies",
};

const CELLS = {
  draftStatus: "C5",
  issueDate: "C6",
  dueDate: "C7",
  clientName: "D11",
  clientEmail: "D13",
  clientDisplay: "D10",
  lastInvoiceNumber: "C6",
  paymentTerms: "C4",
  dateFormatLabel: "C10",
  currencySymbol: "C12",
};

const TRANSACTION_COLUMNS = {
  invoiceNumber: "B",
  issueDate: "C",
  client: "D",
  dueDate: "E",
  total: "G",
  tax: "H",
  netAmount: "I",
};

const TRANSACTION_START_ROW = 14;
const TRANSACTION_FORMAT_END_ROW = 2000;
const REQUIRED_SHEETS = Object.values(SHEETS);

Office.onReady(() => {
  bindEvents();
  setStatus("Office runtime ready.\nRun workbook validation before the first invoice action.");
});

function bindEvents() {
  document.getElementById("validateWorkbook").addEventListener("click", () => runAction(validateWorkbook));
  document.getElementById("updateFormats").addEventListener("click", () => runAction(updateFormats));
  document.getElementById("createInvoice").addEventListener("click", () => runAction(createNewInvoice));
  document.getElementById("submitInvoice").addEventListener("click", () => runAction(submitInvoice));
}

async function runAction(action) {
  toggleButtons(true);
  try {
    await action();
  } catch (error) {
    console.error(error);
    setStatus(formatError(error));
  } finally {
    toggleButtons(false);
  }
}

function toggleButtons(disabled) {
  for (const button of document.querySelectorAll("button")) {
    button.disabled = disabled;
  }
}

function setStatus(message) {
  document.getElementById("statusMessage").textContent = message;
}

function formatError(error) {
  const parts = [`Error: ${error.message || error}`];

  if (error.code) {
    parts.push(`Code: ${error.code}`);
  }

  if (error.debugInfo) {
    if (error.debugInfo.errorLocation) {
      parts.push(`Location: ${error.debugInfo.errorLocation}`);
    }

    if (error.debugInfo.statement) {
      parts.push(`Statement: ${error.debugInfo.statement}`);
    }

    if (error.debugInfo.surroundingStatements) {
      parts.push(`Context: ${error.debugInfo.surroundingStatements.join(" | ")}`);
    }
  }

  if (error.stack) {
    parts.push(`Stack: ${error.stack}`);
  }

  return parts.join("\n");
}

function appendStage(error, stage) {
  if (!error || typeof error !== "object") {
    return new Error(`[${stage}] ${String(error)}`);
  }

  error.message = `[${stage}] ${error.message || error}`;
  return error;
}

async function validateWorkbook() {
  await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const names = worksheets.items.map((sheet) => sheet.name);
    const missing = REQUIRED_SHEETS.filter((name) => !names.includes(name));

    if (missing.length > 0) {
      throw new Error(`Missing required sheets: ${missing.join(", ")}`);
    }

    setStatus(`Workbook validated.\nSheets found: ${names.join(", ")}`);
  });
}

async function updateFormats() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const config = workbook.worksheets.getItem(SHEETS.configuration);
    const template = workbook.worksheets.getItem(SHEETS.template);
    const transactions = workbook.worksheets.getItem(SHEETS.transactions);

    const currencySymbolRange = config.getRange(CELLS.currencySymbol);
    const dateFormatLabelRange = config.getRange(CELLS.dateFormatLabel);

    currencySymbolRange.load("values");
    dateFormatLabelRange.load("values");
    await context.sync();

    const currencySymbol = currencySymbolRange.values[0][0] || "$";
    const dateFormatLabel = dateFormatLabelRange.values[0][0];
    const dateFormat = dateFormatLabel === "Day / Month / Year" ? "dd/mm/yyyy" : "mm/dd/yyyy";
    const currencyFormat = `${currencySymbol}#,##0.00`;

    template.getRange("G16:H21").numberFormat = repeatFormatMatrix(6, 2, currencyFormat);
    template.getRange("H23:H27").numberFormat = [[currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat]];
    template.getRange("C6:C7").numberFormat = [[dateFormat], [dateFormat]];

    transactions.getRange(`G13:I${TRANSACTION_FORMAT_END_ROW}`).numberFormat = repeatFormatMatrix(
      TRANSACTION_FORMAT_END_ROW - 12,
      3,
      currencyFormat
    );
    transactions.getRange(`C13:E${TRANSACTION_FORMAT_END_ROW}`).numberFormat = repeatFormatMatrix(
      TRANSACTION_FORMAT_END_ROW - 12,
      3,
      dateFormat
    );
    transactions.getRange("C3:C10").numberFormat = [[currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat]];

    await context.sync();
    setStatus(`Formats applied.\nCurrency: ${currencySymbol}\nDate format: ${dateFormatLabel || "Month / Day / Year"}`);
  });
}

async function createNewInvoice() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const template = worksheets.getItem(SHEETS.template);
    const config = worksheets.getItem(SHEETS.configuration);
    const paymentTermsRange = config.getRange(CELLS.paymentTerms);
    const dateFormatLabelRange = config.getRange(CELLS.dateFormatLabel);

    worksheets.load("items/name");
    paymentTermsRange.load("values");
    dateFormatLabelRange.load("values");
    await context.sync();

    const names = worksheets.items.map((sheet) => sheet.name);
    const paymentTerms = Number(paymentTermsRange.values[0][0] || 0);
    const dateFormatLabel = dateFormatLabelRange.values[0][0];
    const dateFormat = dateFormatLabel === "Day / Month / Year" ? "dayFirst" : "monthFirst";

    let draftNumber = 1;
    while (names.includes(`New Invoice ${draftNumber}`)) {
      draftNumber += 1;
    }

    const copiedSheet = template.copy(Excel.WorksheetPositionType.after, worksheets.getItem(SHEETS.configuration));
    copiedSheet.name = `New Invoice ${draftNumber}`;

    const today = new Date();
    const dueDate = new Date(today);
    dueDate.setDate(dueDate.getDate() + paymentTerms);
    const issueDateSerial = toExcelDateSerial(today);
    const dueDateSerial = toExcelDateSerial(dueDate);
    const dateNumberFormat = dateFormat === "dayFirst" ? "dd/mm/yyyy" : "mm/dd/yyyy";

    copiedSheet.getRange(CELLS.draftStatus).values = [["DRAFT"]];
    copiedSheet.getRange(CELLS.issueDate).values = [[issueDateSerial]];
    copiedSheet.getRange(CELLS.dueDate).values = [[dueDateSerial]];
    copiedSheet.getRange(CELLS.draftStatus).format.fill.color = "#ffcccc";
    copiedSheet.getRange(CELLS.draftStatus).format.font.color = "#cc0000";
    copiedSheet.getRange(CELLS.draftStatus).format.font.bold = true;
    copiedSheet.getRange("C6:C7").numberFormat = [
      [dateNumberFormat],
      [dateNumberFormat],
    ];

    copiedSheet.activate();
    copiedSheet.getRange(CELLS.clientDisplay).select();
    await context.sync();

    setStatus(`Draft created: New Invoice ${draftNumber}`);
  });
}

async function submitInvoice() {
  await Excel.run(async (context) => {
    try {
      const workbook = context.workbook;
      let activeSheet = workbook.worksheets.getActiveWorksheet();
      const config = workbook.worksheets.getItem(SHEETS.configuration);
      const transactions = workbook.worksheets.getItem(SHEETS.transactions);
      const statusRange = activeSheet.getRange(CELLS.draftStatus);
      const lastInvoiceRange = config.getRange(CELLS.lastInvoiceNumber);

      activeSheet.load("name");
      workbook.worksheets.load("items/name");
      statusRange.load("values");
      lastInvoiceRange.load("values");
      await context.sync();

      const sheetName = activeSheet.name;
      if (!sheetName.startsWith("New Invoice") && !sheetName.startsWith("Invoice ")) {
        throw new Error("Open an invoice sheet first.");
      }

      let invoiceNumber = statusRange.values[0][0];
      if (invoiceNumber === "DRAFT" || sheetName.startsWith("New Invoice")) {
        const lastInvoiceNumber = Number(lastInvoiceRange.values[0][0] || 0);
        invoiceNumber = lastInvoiceNumber + 1;
        const nextSheetName = `Invoice ${invoiceNumber}`;
        const existingNames = workbook.worksheets.items.map((sheet) => sheet.name);
        if (existingNames.includes(nextSheetName)) {
          throw new Error(`Sheet ${nextSheetName} already exists.`);
        }

        statusRange.values = [[invoiceNumber]];
        activeSheet.name = nextSheetName;
        lastInvoiceRange.values = [[invoiceNumber]];
        statusRange.format.fill.color = "#ffffff";
        statusRange.format.font.color = "#000000";
        await context.sync();

        activeSheet = workbook.worksheets.getItem(nextSheetName);
      } else {
        const existing = await findInvoiceRow(context, transactions, invoiceNumber, "submit:find-existing");
        if (existing !== null) {
          throw new Error(`Invoice #${invoiceNumber} is already in Transactions.`);
        }
      }

      await addToTransactionsInternal(context, activeSheet, transactions, invoiceNumber);
      await context.sync();

      setStatus(
        `Invoice #${invoiceNumber} submitted.\nWorkbook actions are complete.\nEmail and PDF steps should be added in phase 2.`
      );
    } catch (error) {
      throw appendStage(error, "submitInvoice");
    }
  });
}

async function addToTransactionsInternal(context, invoiceSheet, transactionsSheet, invoiceNumber) {
  try {
    const existing = await findInvoiceRow(context, transactionsSheet, invoiceNumber, "addToTransactions:find-existing");
    if (existing !== null) {
      return;
    }

    const invoiceData = invoiceSheet.getUsedRange();
    invoiceData.load("values");

    const clientRange = invoiceSheet.getRange(CELLS.clientDisplay);
    clientRange.load("values");
    await context.sync();

    const rows = invoiceData.values;
    const client = clientRange.values[0][0];
    const issueDate = getCellValueFromMatrix(rows, "Date of Issue");
    const dueDate = getCellValueFromMatrix(rows, "Date Due");
    const total = getCellValueFromMatrix(rows, "Total");
    const tax = getCellValueFromMatrix(rows, "Tax");
    const netAmount = getCellValueFromMatrix(rows, "Adjusted Subtotal");

    if (issueDate === null || dueDate === null) {
      throw new Error("Could not read invoice dates from the sheet.");
    }

    const nextRow = await getNextTransactionRow(context, transactionsSheet);

    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.invoiceNumber}${nextRow}`).values = [[invoiceNumber]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.issueDate}${nextRow}`).values = [[issueDate]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.client}${nextRow}`).values = [[client]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.dueDate}${nextRow}`).values = [[dueDate]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.total}${nextRow}`).values = [[total]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.tax}${nextRow}`).values = [[tax]];
    transactionsSheet.getRange(`${TRANSACTION_COLUMNS.netAmount}${nextRow}`).values = [[netAmount]];
  } catch (error) {
    throw appendStage(error, "addToTransactionsInternal");
  }
}

async function findInvoiceRow(context, transactionsSheet, invoiceNumber, stageLabel = "findInvoiceRow") {
  try {
    const lastOccupiedRow = await getLastOccupiedTransactionRow(context, transactionsSheet);
    if (lastOccupiedRow < TRANSACTION_START_ROW) {
      return null;
    }

    const range = transactionsSheet.getRange(`B${TRANSACTION_START_ROW}:B${lastOccupiedRow}`);
    range.load("values");
    await context.sync();

    for (let index = 0; index < range.values.length; index += 1) {
      if (String(range.values[index][0]) === String(invoiceNumber)) {
        return TRANSACTION_START_ROW + index;
      }
    }

    return null;
  } catch (error) {
    throw appendStage(error, stageLabel);
  }
}

async function getNextTransactionRow(context, transactionsSheet) {
  try {
    const lastOccupiedRow = await getLastOccupiedTransactionRow(context, transactionsSheet);
    return Math.max(TRANSACTION_START_ROW, lastOccupiedRow + 1);
  } catch (error) {
    throw appendStage(error, "getNextTransactionRow");
  }
}

async function getLastOccupiedTransactionRow(context, transactionsSheet) {
  try {
    const usedRange = transactionsSheet.getUsedRange();
    usedRange.load("rowCount");
    await context.sync();

    if (usedRange.rowCount < TRANSACTION_START_ROW) {
      return TRANSACTION_START_ROW - 1;
    }

    const probeRange = transactionsSheet.getRange(`B${TRANSACTION_START_ROW}:B${usedRange.rowCount}`);
    probeRange.load("values");
    await context.sync();

    let lastUsedOffset = -1;
    for (let index = 0; index < probeRange.values.length; index += 1) {
      const value = probeRange.values[index][0];
      if (value !== "" && value !== null && value !== undefined) {
        lastUsedOffset = index;
      }
    }

    if (lastUsedOffset < 0) {
      return TRANSACTION_START_ROW - 1;
    }

    return TRANSACTION_START_ROW + lastUsedOffset;
  } catch (error) {
    throw appendStage(error, "getLastOccupiedTransactionRow");
  }
}

function repeatFormatMatrix(rows, columns, format) {
  const matrix = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = [];
    for (let columnIndex = 0; columnIndex < columns; columnIndex += 1) {
      row.push(format);
    }
    matrix.push(row);
  }
  return matrix;
}

function toExcelDateSerial(date) {
  const utcMidnight = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return utcMidnight / 86400000 + 25569;
}

function getCellValueFromMatrix(matrix, label) {
  for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
    for (let colIndex = 0; colIndex < matrix[rowIndex].length; colIndex += 1) {
      if (matrix[rowIndex][colIndex] === label) {
        for (let offset = 1; offset <= 3; offset += 1) {
          const value = matrix[rowIndex][colIndex + offset];
          if (value !== "" && value !== null && value !== undefined) {
            return value;
          }
        }
      }
    }
  }

  return null;
}
