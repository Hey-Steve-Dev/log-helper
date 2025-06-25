function processEmailLoggerRows() {

  Logger.log('Triggered at: ' + new Date());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Email Logger');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var allowedDomainsMap = loadAllowedDomains(); // {domain: emailLogSheetUrl}
  var orgDomain = 'ubihere.com';

  for (var i = lastRow; i >= 2; i--) {
    var row = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0];
    var from = row[3]; // From (col D)
    var to = row[4];   // To (col E)

    var fromDomain = extractDomainFromEmail(from);
    var toDomain = extractDomainFromEmail(to);

    var externalDomain = null;
    if (fromDomain === orgDomain && toDomain !== orgDomain) {
      externalDomain = toDomain;
    } else if (toDomain === orgDomain && fromDomain !== orgDomain) {
      externalDomain = fromDomain;
    }

    if (!externalDomain) {
      Logger.log('No external domain found for row ' + i + '. Deleting row.');
      sheet.deleteRow(i);
      continue;
    }

    Logger.log('Row ' + i + ': extracted domain = ' + externalDomain);

    if (allowedDomainsMap.hasOwnProperty(externalDomain)) {
      var clientSheetUrl = allowedDomainsMap[externalDomain];
      Logger.log('✅ Domain allowed. Logging to client sheet: ' + clientSheetUrl);

      try {
        appendRowToClientSheet(clientSheetUrl, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], row);
      } catch (e) {
        Logger.log('❌ Error appending row to client sheet: ' + e.message);
      }
    } else {
      Logger.log('❌ Domain not allowed: ' + externalDomain);
    }

    sheet.deleteRow(i);
    Logger.log('🧹 Deleted row ' + i + ' from Email Logger');
  }
}

function loadAllowedDomains() {
  var masterSheetId = '1LtZgk5aehWblrMRa42xMzy0baOACS5ofi_tlKpn7m3I';  // Master Client List ID
  var ss = SpreadsheetApp.openById(masterSheetId);
  var infoSheet = ss.getSheetByName('Info');
  if (!infoSheet) {
    Logger.log("❌ Master Info sheet not found");
    return {};
  }

  var data = infoSheet.getDataRange().getValues(); // includes header
  var allowedDomainsMap = {};

  for (var i = 1; i < data.length; i++) { // skip header
    var emailLogUrlRaw = data[i][17]; // R: email log url
    var companyUrlRaw = data[i][18];  // S: company url

    if (companyUrlRaw && emailLogUrlRaw) {
      var domain = extractDomainFromWebsiteUrl(companyUrlRaw.toString().trim());
      if (domain) {
        allowedDomainsMap[domain] = emailLogUrlRaw.toString().trim();
      }
    }
  }
  Logger.log("Allowed domains loaded: " + JSON.stringify(Object.keys(allowedDomainsMap)));
  return allowedDomainsMap;
}

function extractDomainFromWebsiteUrl(url) {
  if (!url) return null;
  url = url.replace(/(^\w+:|^)\/\//, '').replace(/^www\./, ''); // strip protocol & www
  return url.split('/')[0].toLowerCase() || null;
}

function extractDomainFromEmail(emailStr) {
  if (!emailStr) return null;
  var emailRegex = /<([^>]+)>/;
  var match = emailStr.match(emailRegex);
  var email = match ? match[1] : emailStr;
  var domainMatch = email.match(/@([^>\s]+)/);
  return domainMatch && domainMatch[1] ? domainMatch[1].toLowerCase() : null;
}

function appendRowToClientSheet(sheetUrl, headers, rowData) {
  try {
    var sheetId = extractSheetIdFromUrl(sheetUrl);
    var clientSheet = SpreadsheetApp.openById(sheetId).getSheets()[0]; // assuming first sheet

    // Ensure headers exist in client sheet
    ensureHeaders(clientSheet, headers);

    var clientHeaders = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];
    var mappedRow = [];

    for (var i = 0; i < clientHeaders.length; i++) {
      var colName = clientHeaders[i];
      var srcIndex = headers.indexOf(colName);
      if (srcIndex !== -1) {
        mappedRow.push(rowData[srcIndex]);
      } else {
        mappedRow.push('');
      }
    }

    clientSheet.appendRow(mappedRow);
    Logger.log('✔ Appended row to client sheet: ' + sheetUrl);
  } catch (e) {
    Logger.log('❌ Error appending row to client sheet: ' + e.message);
  }
}


function extractSheetIdFromUrl(url) {
  var regex = /\/d\/([a-zA-Z0-9-_]+)(\/|$)/;
  var match = url.match(regex);
  if (match && match[1]) return match[1];
  throw new Error('Invalid sheet URL: ' + url);
}

function ensureHeaders(clientSheet, expectedHeaders) {
  var existingHeaders = clientSheet.getRange(1, 1, 1, clientSheet.getLastColumn()).getValues()[0];
  var isEmpty = existingHeaders.every(function(h) { return h === "" || h === null; });
  if (isEmpty) {
    clientSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    Logger.log('Headers added to client sheet: ' + clientSheet.getName());
  }
}

function addAllOrgUsersAsViewers(file) {
  var users = getAllDomainUserEmails();
  for (var i = 0; i < users.length; i++) {
    try {
      file.addViewer(users[i]);
    } catch (e) {
      Logger.log('❌ Error adding viewer: ' + users[i] + ' → ' + e.message);
    }
  }
}

function getAllDomainUserEmails() {
  var users = [];
  var pageToken;
  do {
    var response = AdminDirectory.Users.list({
      domain: 'ubihere.com',
      maxResults: 100,
      pageToken: pageToken
    });
    var pageUsers = response.users;
    if (pageUsers && pageUsers.length > 0) {
      for (var i = 0; i < pageUsers.length; i++) {
        if (pageUsers[i].primaryEmail && pageUsers[i].suspended !== true) {
          users.push(pageUsers[i].primaryEmail);
        }
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  return users;
}



