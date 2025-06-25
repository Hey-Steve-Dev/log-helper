function buildDraftsHomeCard() {
  var userEmail = Session.getActiveUser().getEmail();
  Logger.log("👤 Active user email: " + userEmail);

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('💬 Suggested Drafts'));

  var section = CardService.newCardSection();
  var rows = getAllDraftRowsForUser(userEmail);
  Logger.log("📋 Draft rows returned: " + rows.length);

  if (rows.length === 0) {
    section.addWidget(CardService.newTextParagraph().setText('No drafts available yet.'));
    Logger.log("🕳 No draft rows available to show.");
  } else {
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];

      Logger.log("📨 Row " + i + " — Subject: " + row.subject + ", From: " + row.from + ", To: " + row.to);
      Logger.log("🔗 threadId: " + row.threadId + ", sheetUrl: " + row.sheetUrl + ", needsDraft: " + row.needsDraft);

      section.addWidget(
        CardService.newTextParagraph().setText(
          '<b>📬 ' + row.subject + '</b><br><i>From:</i> ' + row.from +
          '<br><i>To:</i> ' + row.to + '<br><br>' + (row.needsDraft || '(Draft missing)')
        )
      );

      section.addWidget(
        CardService.newTextButton()
          .setText('✍ Generate Draft')
          .setOnClickAction(
            CardService.newAction()
              .setFunctionName('handleGenerateDraft')
              .setParameters({
                threadId: row.threadId,
                sheetUrl: row.sheetUrl
              })
          )
      );

      if (row.needsDraft) {
        Logger.log("✅ Row " + i + " has a draft. Showing Send Draft button.");
        section.addWidget(
          CardService.newTextButton()
            .setText('📤 Send Draft')
            .setOnClickAction(
              CardService.newAction()
                .setFunctionName('handleSendDraft')
                .setParameters({
                  to: row.to,
                  subject: row.subject,
                  body: row.needsDraft
                })
            )
        );
      } else {
        Logger.log("⚠️ Row " + i + " has no draft content.");
      }
    }
  }

  section.addWidget(
    CardService.newTextButton()
      .setText('🔄 Refresh')
      .setOnClickAction(CardService.newAction().setFunctionName('handleRefresh'))
  );

  card.addSection(section);
  Logger.log("✅ Card built and ready to return.");
  return card.build();
}




function getAllDraftRowsForUser(userEmail) {
  var masterSheetId = '1LtZgk5aehWblrMRa42xMzy0baOACS5ofi_tlKpn7m3I';
  var infoSheet = SpreadsheetApp.openById(masterSheetId).getSheetByName('Info');
  var data = infoSheet.getDataRange().getValues();
  var rows = [];

  Logger.log("🔍 Logged-in user: " + userEmail);

  for (var i = 1; i < data.length; i++) {
    var sheetUrl = data[i][17]; // Column R = email log url
    if (!sheetUrl || sheetUrl.indexOf('docs.google.com') === -1) {
      Logger.log("⛔ Skipping invalid or empty URL at row " + (i+1));
      continue;
    }

    try {
      Logger.log("🌐 Raw sheet URL: " + sheetUrl);
      var match = sheetUrl.match(/[-\w]{25,}/);
      if (!match) {
        Logger.log("⛔ Could not extract sheet ID from URL: " + sheetUrl);
        continue;
      }

      var sheetId = match[0];
      Logger.log("📄 Opening sheet ID: " + sheetId);
      var logSheet = SpreadsheetApp.openById(sheetId).getSheetByName('Log');
      if (!logSheet) {
        Logger.log("⛔ 'Email Logger' sheet not found in: " + sheetId);
        continue;
      }

      var lastRow = logSheet.getLastRow();
      var lastCol = logSheet.getLastColumn();
      Logger.log("📏 Detected range — Last row: " + lastRow + ", Last col: " + lastCol);

      if (lastRow <= 1) {
        Logger.log("⚠️ Skipping sheet due to low row count: " + lastRow);
        continue;
      }

      var logData = logSheet.getRange(1, 1, lastRow, lastCol).getValues();
      Logger.log("📋 logData total rows: " + logData.length);

      var headers = logData[0].map(function(h) { return (h || '').toString().trim(); });
      Logger.log("🔎 Headers found: " + headers.join(" | "));

      var toIndex = headers.indexOf('To');
      var fromIndex = headers.indexOf('From');
      var subjectIndex = headers.indexOf('Subject');
      var draftIndex = headers.indexOf('Needs Draft');
      var threadIndex = headers.indexOf('Thread ID');

      Logger.log('📌 Header Indexes — To: ' + toIndex + ', From: ' + fromIndex + ', Subject: ' + subjectIndex + ', Draft: ' + draftIndex + ', Thread ID: ' + threadIndex);

      if (toIndex === -1 || fromIndex === -1 || subjectIndex === -1 || draftIndex === -1 || threadIndex === -1) {
        Logger.log("❌ Missing required columns. Skipping this sheet.");
        continue;
      }

      var threadMap = {};

      for (var j = 1; j < logData.length; j++) {
        var row = logData[j];
        var threadId = row[threadIndex];
        Logger.log("📌 Row " + j + ": Thread ID = " + threadId);

        if (!threadId) continue;

        if (!threadMap[threadId] || j > threadMap[threadId].rowIndex) {
          threadMap[threadId] = { data: row, rowIndex: j };
        }
      }

      Logger.log("🧵 Total unique threads found: " + Object.keys(threadMap).length);

      Object.keys(threadMap).forEach(function(threadId) {
        var rowObj = threadMap[threadId];
        var row = rowObj.data;

        var from = extractEmail(row[fromIndex]);
        var to = (row[toIndex] || '').trim();
        var subject = row[subjectIndex];
        var draft = row[draftIndex];

        Logger.log("🔍 Thread: " + threadId + " | From: " + from + " | To: " + to + " | Subject: " + subject);

        if (to === userEmail || from === userEmail) {
          Logger.log("✅ Matched thread for user: " + userEmail);
          rows.push({
            subject: subject,
            to: to,
            from: from,
            needsDraft: draft,
            threadId: threadId,
            sheetUrl: sheetUrl
          });
        } else {
          Logger.log("⛔ Skipped thread — userEmail not found in To/From: " + userEmail);
        }
      });

    } catch (e) {
      Logger.log('❌ Exception processing sheet (' + sheetUrl + '): ' + e.message);
    }
  }

  Logger.log("🧮 Rows matched: " + rows.length);
  return rows;
}



function testGetDraftRows() {
  var userEmail = Session.getActiveUser().getEmail();  // or hardcode for testing: "you@yourdomain.com"
  Logger.log("🔍 Testing for user: " + userEmail);
  

  var rows = getAllDraftRowsForUser(userEmail);

  Logger.log("🧮 Total rows found: " + rows.length);

  rows.forEach(function(row, i) {
    Logger.log("[" + i + "] Subject: " + row.subject +
               " | From: " + row.from +
               " | To: " + row.to +
               " | NeedsDraft: " + row.needsDraft +
               " | ThreadID: " + row.threadId);
  });
}


function extractEmail(str) {
  var match = str && str.match(/<(.+?)>/);
  return match ? match[1].trim() : (str || '').trim();
}





