function handleRefresh(e) {
  return CardService.newNavigation()
    .updateCard(buildDraftsHomeCard());
}

function handleGenerateDraft(e) {
  Logger.log("📩 Incoming e.parameters:");
  Logger.log(JSON.stringify(e.parameters));

  var threadId = e.parameters.threadId;
  var sheetUrl = e.parameters.sheetUrl;

  var sheetIdMatch = sheetUrl.match(/[-\w]{25,}/);
  if (!sheetIdMatch) {
    Logger.log("❌ Invalid sheet URL");
    return;
  }

  var sheetId = sheetIdMatch[0];
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Log');
  if (!sheet) {
    Logger.log("❌ Sheet 'Log' not found");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var threadIndex = headers.indexOf('Thread ID');
  var subjectIndex = headers.indexOf('Subject');
  var toIndex = headers.indexOf('To');
  var fromIndex = headers.indexOf('From');
  var bodyIndex = headers.indexOf('Body');
  var draftIndex = headers.indexOf('Needs Draft');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[threadIndex] === threadId) {
      var subject = row[subjectIndex];
      var to = row[toIndex];
      var body = row[bodyIndex];
      var from = row[fromIndex];

      // 🛠 Fallback if 'from' is missing or is just an email address
      if (!from || from.indexOf('@') !== -1) {
        from = extractNameFromBody(body);
      }

      // 👤 Get user's first name
      var userName = getUserFirstName();

      // ✅ Generate Gemini reply
      var senderName = extractNameFromBody(body);
      var userFirstName = getUserFirstName();
      var generatedText = generateGeminiReply(subject, body, senderName, userFirstName);

      Logger.log("✅ Gemini reply:\n" + generatedText);

      // ✅ Create Gmail draft
      GmailApp.createDraft(to, subject, generatedText);
      Logger.log("📬 Draft created in Gmail");

      // ✅ Write to sheet
      sheet.getRange(i + 1, draftIndex + 1).setValue(generatedText);

      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText('✅ Draft generated.'))
        .build();
    }
  }

  Logger.log("❌ No matching thread found in sheet.");
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText('❌ Thread not found.'))
    .build();
}



function getUserFirstName() {
  var email = Session.getActiveUser().getEmail();
  var namePart = email.split('@')[0]; // e.g., "steve.glick" or "steve"
  var firstSegment = namePart.split(/[._]/)[0]; // split on dot or underscore
  return firstSegment.charAt(0).toUpperCase() + firstSegment.slice(1);
}




function generateGeminiReply(subject, body, senderName, userFirstName) {
  var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!GEMINI_API_KEY) {
    throw new Error('❌ GEMINI_API_KEY not found');
  }

  var url = 'https://generativelanguage.googleapis.com/v1/models/gemini-1.5-pro:generateContent?key=' + GEMINI_API_KEY;

  var prompt = "Write a professional reply to the following email.\n\n" +
               "Subject: " + subject + "\n\n" +
               "From: " + senderName + "\n\n" +
               body + "\n\n" +
               'Reply should begin with: "Hi ' + senderName + '," and end with: "Best regards, ' + userFirstName + '"';

  Logger.log("📨 Prompt being sent to Gemini:\n" + prompt);

  var payload = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (json && json.candidates && json.candidates.length > 0) {
    Logger.log("✅ Gemini candidate found: " + JSON.stringify(json.candidates[0]));
    return json.candidates[0].content.parts[0].text;
  } else {
    Logger.log("❌ Gemini API full response:\n" + response.getContentText());
    return "(Error generating draft)";
  }
}


function extractNameFromBody(body) {
  if (!body) return "there";

  var lines = String(body).split('\n');

  for (var i = lines.length - 1; i >= 0; i--) {
    var rawLine = lines[i];
    var line = rawLine && typeof rawLine.toString === 'function'
      ? rawLine.toString().trim()
      : '';

    // Skip empty or common valedictions
    if (line === '') continue;

    var lowerLine = line.toLowerCase();
    if (
      lowerLine === 'regards' || lowerLine === 'best' || lowerLine === 'thanks' ||
      lowerLine === 'sincerely' || lowerLine === 'cheers'
    ) {
      continue;
    }

    // Skip lines that look like email addresses
    if (line.indexOf('@') !== -1) continue;

    // Try to extract a capitalized first name (first word only)
    var words = line.split(/\s+/);
    if (words.length > 0) {
      var firstWord = words[0];
      if (/^[A-Z][a-z]+$/.test(firstWord)) {
        return firstWord;
      }
    }
  }

  // Fallback
  return "there";
}


function testGenerateGeminiReply() {
  var subject = "Follow-up Meeting";
  var body = "Hi Steve, just checking in on next week's follow-up. Let me know what times work for you.";
  var senderName = "Jacob Fain";
  var userName = "Steve";

  var draft = generateGeminiReply(subject, body, senderName, userName);
  Logger.log("📝 Final Draft:\n" + draft);
}


function handleSendDraft(e) {
  var to = e.parameters.to;
  var subject = e.parameters.subject;
  var body = e.parameters.body;

  GmailApp.sendEmail(to, subject, body);

  return CardService.newNotification()
    .setText('✅ Draft sent!')
    .build();
}

function testHandleGenerateDraft() {
  var e = {
    parameters: {
      threadId: '1979e328e794a517',
      sheetUrl: 'https://docs.google.com/spreadsheets/d/1U8Z02OzmOFc6rI8zFtfVwo1vB1ZoISJIL1EvA8R7YKE'
    }
  };
  handleGenerateDraft(e);
}









