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
    setStatus(`Error: ${error.message || error}`);
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
    transactions.getRange("D3:D10").numberFormat = [[currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat], [currencyFormat]];

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

    copiedSheet.getRange(CELLS.draftStatus).values = [["DRAFT"]];
    copiedSheet.getRange(CELLS.issueDate).values = [[today]];
    copiedSheet.getRange(CELLS.dueDate).values = [[dueDate]];
    copiedSheet.getRange(CELLS.draftStatus).format.fill.color = "#ffcccc";
    copiedSheet.getRange(CELLS.draftStatus).format.font.color = "#cc0000";
    copiedSheet.getRange(CELLS.draftStatus).format.font.bold = true;
    copiedSheet.getRange("C6:C7").numberFormat = [
      [dateFormat === "dayFirst" ? "dd/mm/yyyy" : "mm/dd/yyyy"],
      [dateFormat === "dayFirst" ? "dd/mm/yyyy" : "mm/dd/yyyy"],
    ];

    copiedSheet.activate();
    copiedSheet.getRange(CELLS.clientDisplay).select();
    await context.sync();

    setStatus(`Draft created: New Invoice ${draftNumber}`);
  });
}

async function submitInvoice() {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const activeSheet = workbook.worksheets.getActiveWorksheet();
    const config = workbook.worksheets.getItem(SHEETS.configuration);
    const transactions = workbook.worksheets.getItem(SHEETS.transactions);

    activeSheet.load("name");
    workbook.worksheets.load("items/name");
    config.getRange(CELLS.lastInvoiceNumber).load("values");
    activeSheet.getRange(CELLS.draftStatus).load("values");
    await context.sync();

    const sheetName = activeSheet.name;
    if (!sheetName.startsWith("New Invoice") && !sheetName.startsWith("Invoice ")) {
      throw new Error("Open an invoice sheet first.");
    }

    let invoiceNumber = activeSheet.getRange(CELLS.draftStatus).values[0][0];
    if (invoiceNumber === "DRAFT" || sheetName.startsWith("New Invoice")) {
      const lastInvoiceNumber = Number(config.getRange(CELLS.lastInvoiceNumber).values[0][0] || 0);
      invoiceNumber = lastInvoiceNumber + 1;
      const nextSheetName = `Invoice ${invoiceNumber}`;
      const existingNames = workbook.worksheets.items.map((sheet) => sheet.name);
      if (existingNames.includes(nextSheetName)) {
        throw new Error(`Sheet ${nextSheetName} already exists.`);
      }

      activeSheet.getRange(CELLS.draftStatus).values = [[invoiceNumber]];
      activeSheet.name = nextSheetName;
      config.getRange(CELLS.lastInvoiceNumber).values = [[invoiceNumber]];
      activeSheet.getRange(CELLS.draftStatus).format.fill.color = "#ffffff";
      activeSheet.getRange(CELLS.draftStatus).format.font.color = "#000000";
    } else {
      const existing = await findInvoiceRow(context, transactions, invoiceNumber);
      if (existing !== null) {
        throw new Error(`Invoice #${invoiceNumber} is already in Transactions.`);
      }
    }

    await addToTransactionsInternal(context, activeSheet, transactions, invoiceNumber);
    await context.sync();

    setStatus(
      `Invoice #${invoiceNumber} submitted.\nWorkbook actions are complete.\nEmail and PDF steps should be added in phase 2.`
    );
  });
}

async function addToTransactionsInternal(context, invoiceSheet, transactionsSheet, invoiceNumber) {
  const existing = await findInvoiceRow(context, transactionsSheet, invoiceNumber);
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

  const nextRow = await getNextTransactionRow(context, transactionsSheet);

  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.invoiceNumber}${nextRow}`).values = [[invoiceNumber]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.issueDate}${nextRow}`).values = [[issueDate]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.client}${nextRow}`).values = [[client]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.dueDate}${nextRow}`).values = [[dueDate]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.total}${nextRow}`).values = [[total]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.tax}${nextRow}`).values = [[tax]];
  transactionsSheet.getRange(`${TRANSACTION_COLUMNS.netAmount}${nextRow}`).values = [[netAmount]];
}

async function findInvoiceRow(context, transactionsSheet, invoiceNumber) {
  const nextRow = await getNextTransactionRow(context, transactionsSheet);
  const range = transactionsSheet.getRange(`B${TRANSACTION_START_ROW}:B${Math.max(nextRow - 1, TRANSACTION_START_ROW)}`);
  range.load("values");
  await context.sync();

  for (let index = 0; index < range.values.length; index += 1) {
    if (String(range.values[index][0]) === String(invoiceNumber)) {
      return TRANSACTION_START_ROW + index;
    }
  }

  return null;
}

async function getNextTransactionRow(context, transactionsSheet) {
  const usedRange = transactionsSheet.getUsedRange();
  usedRange.load("rowCount");
  await context.sync();

  if (usedRange.rowCount < TRANSACTION_START_ROW) {
    return TRANSACTION_START_ROW;
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

  return TRANSACTION_START_ROW + lastUsedOffset + 1;
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
