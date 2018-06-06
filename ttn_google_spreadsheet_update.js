// 2017 by Daniel Eichhorn, https://blog.squix.org
// Inspired by https://gist.github.com/bmcbride/7069aebd643944c9ee8b
// Create or open an existing Sheet and click Tools > Script editor and enter the code below
// 1. Enter sheet name where data is to be written below
var SHEET_NAME = "Sheet1";
// 2. Run > setup
// 3. Publish > Deploy as web app
//    - enter Project Version name and click 'Save New Version'
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously)
// 4. Copy the 'Current web app URL' and post this in your form/script action

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return handleResponse(e);
}

function doPost(e){
  return handleResponse(e);
}

function handleResponse(e) {
  Logger.log("Waiting for lock")
  var lock = LockService.getPublicLock();
  lock.waitLock(30000); // wait 30 seconds before conceding defeat.
  Logger.log("Got  lock")

  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(SHEET_NAME); 
    // we'll assume header is in row 1 but you can override with header_row in GET/POST data

    //var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var nextRow = sheet.getLastRow()+1; // get next row
    var row = [];
    var headerRow = [];
    // loop through the header columns
    var jsonData = JSON.parse(e.postData.contents);

    headerRow.push("jsonData.app_id");
    headerRow.push("jsonData.dev_id");
    headerRow.push("jsonData.hardware_serial");
    headerRow.push("jsonData.port");
    headerRow.push("jsonData.counter");
    headerRow.push("jsonData.payload_raw");
    headerRow.push("jsonData.payload_decoded");
    headerRow.push("jsonData.metadata.time");
    headerRow.push("jsonData.metadata.frequency");
    headerRow.push("jsonData.metadata.modulation");
    headerRow.push("jsonData.metadata.data_rate");
    headerRow.push("jsonData.metadata.coding_rate");
    headerRow.push("jsonData.metadata.downlink_url");
    for (var i = 0; i < jsonData.metadata.gateways.length; i++) {
      var gateway = jsonData.metadata.gateways[i];
      headerRow.push("gateway.gtw_id");
      headerRow.push("gateway.timestamp");
      headerRow.push("gateway.channel");
      headerRow.push("gateway.rssi");
      headerRow.push("gateway.snr");
      headerRow.push("gateway.latitude");
      headerRow.push("gateway.longitude");
      headerRow.push("gateway.altitude");
    }
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

    row.push(jsonData.app_id);
    row.push(jsonData.dev_id);
    row.push(jsonData.hardware_serial);
    row.push(jsonData.port);
    row.push(jsonData.counter);
    row.push(jsonData.payload_raw);
    var raw = Utilities.base64Decode(jsonData.payload_raw);

//    var decoded = Utilities.newBlob(raw).getDataAsString();
    var decoded = toHexString(raw).toUpperCase();

    row.push(decoded);

    var formattedDate = Utilities.formatDate(new Date(getDateFromIso(jsonData.metadata.time)), "GMT", "yyyy-MM-dd HH:mm:ss");
    row.push(formattedDate);

    row.push(jsonData.metadata.frequency);
    row.push(jsonData.metadata.modulation);
    row.push(jsonData.metadata.data_rate);
    row.push(jsonData.metadata.coding_rate);
    row.push(jsonData.metadata.downlink_url);
    for (var i = 0; i < jsonData.metadata.gateways.length; i++) {
      var gateway = jsonData.metadata.gateways[i];
      row.push(gateway.gtw_id);
      row.push(gateway.timestamp);
      row.push(gateway.channel);
      row.push(gateway.rssi);
      row.push(gateway.snr);
      row.push(gateway.latitude);
      row.push(gateway.longitude);
      row.push(gateway.altitude);

    }

    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {

    Logger.log(JSON.stringify({"result":"error", "error": e}));

    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function setup() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", doc.getId());
}

// http://delete.me.uk/2005/03/iso8601.html
function getDateFromIso(string) {
  try{
    var aDate = new Date();
    var regexp = "([0-9]{4})(-([0-9]{2})(-([0-9]{2})" +
        "(T([0-9]{2}):([0-9]{2})(:([0-9]{2})(\\.([0-9]+))?)?" +
        "(Z|(([-+])([0-9]{2}):([0-9]{2})))?)?)?)?";
    var d = string.match(new RegExp(regexp));

    var offset = 0;
    var date = new Date(d[1], 0, 1);

    if (d[3]) { date.setMonth(d[3] - 1); }
    if (d[5]) { date.setDate(d[5]); }
    if (d[7]) { date.setHours(d[7]); }
    if (d[8]) { date.setMinutes(d[8]); }
    if (d[10]) { date.setSeconds(d[10]); }
    if (d[12]) { date.setMilliseconds(Number("0." + d[12]) * 1000); }
    if (d[14]) {
      offset = (Number(d[16]) * 60) + Number(d[17]);
      offset *= ((d[15] == '-') ? 1 : -1);
    }

    offset -= date.getTimezoneOffset();
    time = (Number(date) + (offset * 60 * 1000));
    return aDate.setTime(Number(time));
  } catch(e){
    return;
  }
}

function toHexString(byteArray) {
  var s = '';
  byteArray.forEach(function(byte) {
    s += ('0' + (byte & 0xFF).toString(16)).slice(-2) + ' ';
  });
  return s.trim();
}

