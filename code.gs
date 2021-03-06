function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Facturen')
    .addItem('Maak factuur van huidige regel', 'createFromCurrentRow')
    .addSeparator()
    .addItem('Maak facturen met factuurnummers', 'createAllInvoices')
    .addItem('Maak voldane facturen', 'createPaidInvoices')
    .addItem('Maak niet voldane facturen', 'createUnpaidInvoices')
    .addItem('Maak herinnering facturen', 'createReminderInvoices')
    .addSeparator()
    .addItem('Maak etiketten', 'createAddressLabels')
    .addItem('Maak bedankstikkers', 'createThankYouNotes')
    .addItem('Maak infoStikkers', 'createInfoNotes')
    .addSeparator()
    .addItem('Run tests', 'runTests')
    .addToUi();
}

function getPageInfo() {
  return {
    columns: 3,
    rows: 9,
    margin: 8,
    factor: 6,
    bcc: 'mastersofkyokushin@gmail.com',
    subjectPostfix: ' sponsoring Masters of Kyokushin Gala 2018',
    bodyStartText: 'Geachte sponsor,',
    thankYouText: 'Dank voor uw sponsorbijdrage. Wij stellen uw bijdrage zeer op prijs en hopen u volgend jaar weer te mogen verwelkomen als sponsor!',
    infoText: 'Geachte sponsor, bijgevoegd uw toegangskaarten. Als u meer nodig heeft, kunt u altijd contact met ons opnemen.',
    paymentTerm: 14,
    maxDataColumn: 10
  };
}

function createAllInvoices() { createInvoices({paid:true,unpaid:true}); }
function createPaidInvoices() { createInvoices({paid:true}); }
function createUnpaidInvoices() { createInvoices({unpaid:true}); }
function createReminderInvoices() { createInvoices({overdue:true}); }
function createThankYouNotes() { createTextNotes('Bedankjes', getPageInfo().thankYouText); }
function createInfoNotes() { createTextNotes('Info', getPageInfo().infoText); }
function createFromCurrentRow() { 
  var sheet = SpreadsheetApp.getActiveSheet();
  createInvoice(sheet.getRange(sheet.getActiveRange().getRow(), 1, 1, getPageInfo().maxDataColumn).getValues(), {paid:true,unpaid:true}, 0); 
}

function createAddressLabels() {
  var getText = function (page, row, column, pageInfo, data) { 
    var index = getDataIndex(pageInfo, page, row, column);
    return index < data.length ? getLabel(getRowInfo(data[index])) : null;
  }
  createGrid('Etiketten', getText);
}

function getLabel(info) {
  return info.naam + '\r\r' + info.adres + '\r' + info.pcStad;
}

function createTextNotes(name, text) {
  var getText = function (page, row, column, pageInfo, data) { 
    var index = getDataIndex(pageInfo, page, row, column);
    return index < data.length ? text : null;
  }
  createGrid(name, getText);
}

function getDataIndex(pageInfo, page, row, column) {
  return 1 + page * pageInfo.rows * pageInfo.columns + row * pageInfo.columns + column;
}

function createGrid(name, getText) {
  var doc = DocumentApp.create(name);
  var body = doc.getBody();
  var pageInfo = getPageInfo();
  body.setMarginLeft(pageInfo.margin);
  body.setMarginRight(0);
  body.setMarginTop(pageInfo.margin);
  body.setMarginBottom(0);
  var data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  for (var page = 0; page < Math.floor(data.length / (pageInfo.rows * pageInfo.columns)) + 1; page++)
    createLabelPage(body, page, data, getText);
}

function createLabelPage(body, page, data, getText) {
  var pageInfo = getPageInfo();
  var table = body.appendTable();
  table.setBorderWidth(0);
  for (var row = 0; row < pageInfo.rows; row++)
    createLabelRow(body, page, data, row, table, pageInfo, getText);
  body.appendPageBreak();
}

function createLabelRow(body, page, data, row, table, pageInfo, getText) {
  var tableRow = table.appendTableRow();
  if (row < pageInfo.rows - 1)
    tableRow.setMinimumHeight((body.getPageHeight() - pageInfo.factor * pageInfo.margin) / pageInfo.rows);
  for (var column = 0; column < pageInfo.columns; column++) {
    var label = getText(page, row, column, pageInfo, data);
    if (label) tableRow.appendTableCell(label);
  }
}

function createInvoices(flags) {
  var data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++)
    createInvoice(data, flags, i);
  activateRow(2);
}

function activateRow(row) {
  var range = SpreadsheetApp.getActiveSheet().getRange(row, 1, 1);
  range.activate();
  SpreadsheetApp.flush();
}

function createInvoice(data, flags, i) {
  activateRow(i + 1);
  var info = getRowInfo(data[i]);
  if ((info.betaald > 0 && flags.paid) || (info.betaald == 0 && flags.unpaid) || (info.dagenOpenstaand > getPageInfo().paymentTerm && flags.overdue))
    createInvoiceDraft(info);
}

function createInvoiceDraft(info) {
  if (info.factuurnummer == '') return;
  var doc = getTemplateDoc(info.sjabloon);
  var copyDoc = doc.makeCopy();
  var copyUrl = copyDoc.getUrl();
  var copyDocument = DocumentApp.openByUrl(copyUrl);
  var copyBody = copyDocument.getBody();

  copyBody.replaceText('{factuurnummer}', info.factuurnummer);
  copyBody.replaceText('{naam}', info.naam);
  copyBody.replaceText('{adres}', info.adres);
  copyBody.replaceText('{postcode, plaats}', info.pcStad);
  copyBody.replaceText('{bedrag}', info.bedrag);
  copyBody.replaceText('{datum}', formatDate(info.datum));
  copyDocument.saveAndClose();
  
  var copyBlob = copyDoc.getAs('application/pdf').setName(info.factuurnummer + ' ' + info.naam + '.pdf');
  var body = copyBody.getText().trim();
  var text = body.substring(body.indexOf(info.bodyStartText));
  var pageInfo = getPageInfo();
  var flags = { attachments: [copyBlob], bcc: pageInfo.bcc };
  var subjectPrefix = info.dagenOpenstaand > getPageInfo().paymentTerm ? 'Herinnering factuur ' : 'Factuur ';
  GmailApp.createDraft(info.email, subjectPrefix + info.factuurnummer + pageInfo.subjectPostfix, text, flags);
  
  copyDoc.setTrashed(true);
}

function formatDate(date) {
  return date.getDate() + '-' + (1 + date.getMonth()) + '-' + date.getFullYear();
}

function getRowInfo(row) {
  return {
    factuurnummer: row[0],
    betaald: row[7],
    naam: row[1],
    adres: row[2],
    pcStad: row[3],
    tav: row[4],
    email: row[5],
    bedrag: row[6],
	datum: row[8],
    dagenOpenstaand: row[9],
    sjabloon: row[7] > 0 ? 'Factuur Voldaan Sjabloon' : row[9] > getPageInfo().paymentTerm ? 'Factuur Herinnering Sjabloon' : 'Factuur Sjabloon',
  };
}

function getTemplateDoc(name) {
  var files = DriveApp.getFilesByName(name);
  while (files.hasNext())
    return files.next();
  return null;
}

function runTests() {
  var info = {rows:9, columns:4};
  var html = 
      '<div style=\"background-color: blue;color:white;\">info={rows:9, columns:4}</div>' +
      tester('getDataIndex(info, 0, 0, 0)', getDataIndex(info, 0, 0, 0), 1) +
      tester('getDataIndex(info, 0, 0, 1)', getDataIndex(info, 0, 0, 1), 2) +
      tester('getDataIndex(info, 0, 1, 0)', getDataIndex(info, 0, 1, 0), 5) +
      tester('getDataIndex(info, 1, 0, 0)', getDataIndex(info, 1, 0, 0), 37);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle('Tests'));
}

function tester(text, result, expected) {
  if (result == expected)
    return '<div style=\"color: white;background-color:green;\">' + text + ' == ' + expected + '</div>';
  return '<div style=\"color: white;background-color:red;\">' + text + ' == ' + expected + ' but got ' + result + '</div>';
}
