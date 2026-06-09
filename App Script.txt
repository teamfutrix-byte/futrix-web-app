/**
 * Futrix Pilot Exam Portal - Google Apps Script Backend
 * 
 * ── GOOGLE SHEET COLUMN MAP (First Sheet - Registration / Responses) ──
 * The Google Sheet columns map as:
 * - Column A: Timestamp
 * - Column B: Full Name
 * - Column C: Email Address
 * - Column D: Phone Number
 * - Column E: Date of Birth
 * - Column F: Father's / Guardian's Name
 * - Column G: Father's / Guardian's Contact Number
 * - Column H: City
 * - Column I: Institute Name
 * - Column J: Pin Code
 * - Column K: Preparation For
 * - Column L: If Any Referral
 * - Column M: XP Point
 * - Column N: Referral XP
 * 
 * Note: The existing verifyLogin and checkRegistration checks find column indexes dynamically,
 * which ensures no Apps Script logic shifts are required.
 */

// ── EMAIL DELIVERY CONFIGURATION (Select your preferred mailing service)
// Options:
// - "GMAIL"  : Uses GmailApp.sendEmail. Free, but personal @gmail.com accounts may go to Spam.
// - "BREVO"  : Uses Brevo (Sendinblue) API. Free 300 emails/day. Highly recommended to land in Inbox.
// - "RESEND" : Uses Resend.com API. Free 3,000 emails/month. Highly recommended to land in Inbox.
var EMAIL_SERVICE  = "GMAIL"; 
var BREVO_API_KEY  = "YOUR_BREVO_API_KEY_HERE";
var RESEND_API_KEY = "YOUR_RESEND_API_KEY_HERE";
var SENDER_EMAIL   = "your_verified_sender_email@gmail.com"; // Required for Brevo / Resend

function doGet(e) {
  var callback = e.parameter.callback || '';
  var action   = e.parameter.action   || '';
  var result;
  try {
    if      (action === 'questions') result = getQuestions();
    else if (action === 'submit')    result = saveResponse(e);
    else if (action === 'getConfig') result = getExamConfig();
    else if (action === 'checkAttempt') result = checkAttempt(e);
    else if (action === 'checkRegistration') result = checkRegistration(e);
    else if (action === 'register') result = saveRegistration(e);
    else if (action === 'sendOTP')   result = sendOTP(e);
    else if (action === 'verifyOTP') result = verifyOTP(e);
    else                             result = verifyLogin(e);
  } catch (err) {
    result = { success: false, message: 'Error: ' + err.message };
  }
  var json = JSON.stringify(result);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function getXPColumn(sheet) {
  var lastCol = sheet.getLastColumn(); if (lastCol === 0) return 6;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) { 
    var h = String(headers[i]).trim();
    if (h === 'XP Points' || h === 'XP Point') return i + 1; 
  }
  var newCol = lastCol + 1; sheet.getRange(1, newCol).setValue('XP Point'); sheet.getRange(1, newCol).setFontWeight('bold'); return newCol;
}

function getReferralXPColumn(sheet) {
  var lastCol = sheet.getLastColumn(); if (lastCol === 0) return 12;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) { 
    var h = String(headers[i]).trim();
    if (h === 'Referral XP' || h === 'Referral XP Point') return i + 1; 
  }
  var newCol = lastCol + 1; sheet.getRange(1, newCol).setValue('Referral XP'); sheet.getRange(1, newCol).setFontWeight('bold'); return newCol;
}

function getColumnIndexByName(sheet, name, defaultCol) {
  var lastCol = sheet.getLastColumn(); if (lastCol === 0) return defaultCol;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h.indexOf(name.toLowerCase()) !== -1) return i + 1;
  }
  return defaultCol;
}

function verifyLogin(e) {
  var email = (e.parameter.email || '').toLowerCase().trim();
  var phone = (e.parameter.phone || '').trim();
  if (!email || !phone) return { success: false, message: 'Please enter both email and mobile number.' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data  = sheet.getDataRange().getValues();
  
  var nameCol  = getColumnIndexByName(sheet, 'name', 2);
  var emailCol = getColumnIndexByName(sheet, 'email', 3);
  var phoneCol = getColumnIndexByName(sheet, 'phone', 4);
  var xpCol    = getXPColumn(sheet);
  var refXpCol = getReferralXPColumn(sheet);
  var prepCol  = getColumnIndexByName(sheet, 'preparation', 9);

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol - 1]).toLowerCase().trim();
    var rowPhone = String(data[i][phoneCol - 1]).trim();
    if (rowEmail === email && rowPhone === phone) {
      var xp = parseFloat(data[i][xpCol - 1]) || 0;
      if (xp === 0) { xp = 100; sheet.getRange(i + 1, xpCol).setValue(100); }
      var referralXp = parseFloat(data[i][refXpCol - 1]) || 0;
      var prep = String(data[i][prepCol - 1] || 'Other').trim();
      return { success: true, name: String(data[i][nameCol - 1]).trim(), email: email, phone: phone, xp: xp, referralXp: referralXp, preparation: prep };
    }
  }
  return { success: false, message: 'Invalid email or mobile number. Please try again.' };
}

function getQuestions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Questions');
  if (!sheet) return { success: false, message: 'Sheet "Questions" not found.' };
  var data = sheet.getDataRange().getValues(); var questions = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    questions.push({ no: data[i][0]||i, question: String(data[i][1]).trim(),
      optionA: String(data[i][2]||'').trim(), optionB: String(data[i][3]||'').trim(),
      optionC: String(data[i][4]||'').trim(), optionD: String(data[i][5]||'').trim(),
      correct: String(data[i][6]||'').trim().toUpperCase(), marks: Number(data[i][7])||1, negative: Number(data[i][8])||-0.25 });
  }
  if (questions.length === 0) return { success: false, message: 'No questions found in the Questions sheet.' };
  return { success: true, questions: questions };
}

// ── EXAM CONFIG: Fetch from "Exam Config" tab ──
function getExamConfig() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Exam Config');
  if (!sheet) return { success: false, message: 'Sheet "Exam Config" not found.' };
  var data = sheet.getDataRange().getValues();
  var cfg  = {};
  for (var i = 0; i < data.length; i++) {
    cfg[String(data[i][0]).trim()] = data[i][1];
  }
  return {
    success:   true,
    seriesId:  String(cfg['Series ID']       || '#FX-0001'),
    questions: Number(cfg['Questions']        || 20),
    duration:  Number(cfg['Duration (mins)'] || 30),
    maxMarks:  Number(cfg['Max Marks']        || 20)
  };
}

function saveResponse(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName('Exam Responses');
  var headers = ['Timestamp','Candidate Name','Email','Phone','Series ID','Score','Correct','Wrong','Skipped','Total Questions','Time Taken','Answers'];
  if (!sheet) {
    sheet = ss.insertSheet('Exam Responses'); sheet.appendRow(headers);
    var hr = sheet.getRange(1,1,1,headers.length); hr.setFontWeight('bold'); hr.setBackground('#1a73e8'); hr.setFontColor('#ffffff'); sheet.setFrozenRows(1);
  } else {
    if (String(sheet.getRange(1,1).getValue()) !== 'Timestamp') {
      sheet.insertRowBefore(1); sheet.getRange(1,1,1,headers.length).setValues([headers]);
      var hr2 = sheet.getRange(1,1,1,headers.length); hr2.setFontWeight('bold'); hr2.setBackground('#1a73e8'); hr2.setFontColor('#ffffff'); sheet.setFrozenRows(1);
    }
  }
  var candidateEmail = (e.parameter.candidateEmail || '').toLowerCase().trim();
  sheet.appendRow([new Date(), e.parameter.candidateName||'', candidateEmail, e.parameter.candidatePhone||'',
    e.parameter.seriesId||'', Number(e.parameter.score||0), Number(e.parameter.correct||0), Number(e.parameter.wrong||0),
    Number(e.parameter.skipped||0), Number(e.parameter.totalQuestions||0), e.parameter.timeTaken||'', e.parameter.answers||'{}']);
  
  // Log entry in "Exam Attempts" sheet
  try {
    var attemptSheet = ss.getSheetByName('Exam Attempts');
    var attemptHeaders = ['Timestamp', 'Series ID', 'Candidate Name', 'Email'];
    if (!attemptSheet) {
      attemptSheet = ss.insertSheet('Exam Attempts'); attemptSheet.appendRow(attemptHeaders);
      var ahr = attemptSheet.getRange(1,1,1,attemptHeaders.length); ahr.setFontWeight('bold'); ahr.setBackground('#34a853'); ahr.setFontColor('#ffffff'); attemptSheet.setFrozenRows(1);
    } else {
      if (String(attemptSheet.getRange(1,1).getValue()) !== 'Timestamp') {
        attemptSheet.insertRowBefore(1); attemptSheet.getRange(1,1,1,attemptHeaders.length).setValues([attemptHeaders]);
        var ahr2 = attemptSheet.getRange(1,1,1,attemptHeaders.length); ahr2.setFontWeight('bold'); ahr2.setBackground('#34a853'); ahr2.setFontColor('#ffffff'); attemptSheet.setFrozenRows(1);
      }
    }
    attemptSheet.appendRow([new Date(), e.parameter.seriesId||'', e.parameter.candidateName||'', candidateEmail]);
  } catch (err) {
    Logger.log('Error saving to Exam Attempts sheet: ' + err.message);
  }

  // Log detailed question responses
  try {
    var detailedSheet = ss.getSheetByName('Detailed Question Responses');
    var detailedHeaders = ['Timestamp', 'Candidate Name', 'Email', 'Phone', 'Series ID', 'Question No', 'Question Text', 'Topic', 'Candidate Response', 'Correct Answer', 'Status'];
    if (!detailedSheet) {
      detailedSheet = ss.insertSheet('Detailed Question Responses'); detailedSheet.appendRow(detailedHeaders);
      var dhr = detailedSheet.getRange(1,1,1,detailedHeaders.length); dhr.setFontWeight('bold'); dhr.setBackground('#673ab7'); dhr.setFontColor('#ffffff'); detailedSheet.setFrozenRows(1);
    } else {
      if (String(detailedSheet.getRange(1,1).getValue()) !== 'Timestamp') {
        detailedSheet.insertRowBefore(1); detailedSheet.getRange(1,1,1,detailedHeaders.length).setValues([detailedHeaders]);
        var dhr2 = detailedSheet.getRange(1,1,1,detailedHeaders.length); dhr2.setFontWeight('bold'); dhr2.setBackground('#673ab7'); dhr2.setFontColor('#ffffff'); detailedSheet.setFrozenRows(1);
      }
    }
    
    var qSheet = ss.getSheetByName('Questions');
    if (qSheet) {
      var qData = qSheet.getDataRange().getValues();
      var parsedAnswers = JSON.parse(e.parameter.answers || '{}');
      var qIndex = 0;
      for (var i = 1; i < qData.length; i++) {
        if (!qData[i][1]) continue;
        var qNo = qData[i][0] || (qIndex + 1);
        var qText = String(qData[i][1]).trim();
        var correctAns = String(qData[i][6] || '').trim().toUpperCase();
        var topic = String(qData[i][9] || 'General').trim();
        
        var candidateAns = parsedAnswers[qIndex];
        var status = '';
        if (candidateAns === undefined || candidateAns === null || candidateAns === '') {
          candidateAns = '';
          status = 'Skipped';
        } else {
          candidateAns = String(candidateAns).trim().toUpperCase();
          status = (candidateAns === correctAns) ? 'Correct' : 'Wrong';
        }
        
        detailedSheet.appendRow([
          new Date(),
          e.parameter.candidateName || '',
          candidateEmail,
          e.parameter.candidatePhone || '',
          e.parameter.seriesId || '',
          qNo,
          qText,
          topic,
          candidateAns,
          correctAns,
          status
        ]);
        qIndex++;
      }
    }
  } catch (err) {
    Logger.log('Error saving detailed question responses: ' + err.message);
  }

  var xpEarned = parseFloat(e.parameter.xpEarned || 0);
  if (xpEarned !== 0) {
    var regSheet = ss.getSheets()[0]; var regData = regSheet.getDataRange().getValues(); var xpCol = getXPColumn(regSheet);
    var emailCol = getColumnIndexByName(regSheet, 'email', 3);
    for (var i = 1; i < regData.length; i++) {
      if (String(regData[i][emailCol-1]).toLowerCase().trim() === candidateEmail) {
        var newXP = Math.max(0, parseFloat(((parseFloat(regData[i][xpCol-1])||100) + xpEarned).toFixed(2)));
        regSheet.getRange(i+1, xpCol).setValue(newXP); break;
      }
    }
  }
  return { success: true };
}

// ── CHECK ATTEMPT: Check if test is already taken once ──
function checkAttempt(e) {
  // Support multiple parameter aliases for email and series
  var email = (e.parameter.email || e.parameter.candidateEmail || e.parameter.emailAddress || '').toLowerCase().trim();
  var seriesId = (e.parameter.seriesId || e.parameter.series || '').trim();

  if (!email || !seriesId)
    return { attempted: false };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Exam Attempts');

  if (!sheet) return { attempted: false };

  var data = sheet.getDataRange().getValues();

  var emailCol  = getColumnIndexByName(sheet, 'email', 4);
  var seriesCol = getColumnIndexByName(sheet, 'series', 2);

  for (var i = 1; i < data.length; i++) {
    var rowSeries = String(data[i][seriesCol - 1] || '').trim();
    var rowEmail  = String(data[i][emailCol - 1] || '').toLowerCase().trim();
    if (rowEmail === email && rowSeries === seriesId) {
      return { attempted: true, series: seriesId };
    }
  }

  return { attempted: false };
}

// ── PREVENT DUPLICATE REGISTRATION: Check if email or phone is already registered ──
function checkRegistration(e) {
  var email = (e.parameter.email || '').toLowerCase().trim();
  var phone = (e.parameter.phone || '').trim();

  if (!email && !phone) {
    return { success: false, message: 'Please provide email or mobile number to check.' };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data  = sheet.getDataRange().getValues();

  var emailCol = getColumnIndexByName(sheet, 'email', 3);
  var phoneCol = getColumnIndexByName(sheet, 'phone', 4);

  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol - 1] || '').toLowerCase().trim();
    var rowPhone = String(data[i][phoneCol - 1] || '').trim();

    if (email && rowEmail === email) {
      return { success: false, message: 'You are already registered with the same email ID.' };
    }
    if (phone && rowPhone === phone) {
      return { success: false, message: 'You are already registered with the same mobile number.' };
    }
  }

  return { success: true };
}

// ── SAVE REGISTRATION: Register a new candidate directly to the sheet ──
function saveRegistration(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  var fullName = (e.parameter.fullName || '').trim();
  var email = (e.parameter.email || '').toLowerCase().trim();
  var phone = (e.parameter.phone || '').trim();
  var dob = (e.parameter.dob || '').trim();
  var guardianName = (e.parameter.guardianName || '').trim();
  var guardianContact = (e.parameter.guardianContact || '').trim();
  var city = (e.parameter.city || '').trim();
  var instituteName = (e.parameter.instituteName || '').trim();
  var pinCode = (e.parameter.pinCode || '').trim();
  var preparation = (e.parameter.preparation || '').trim();
  var referral = (e.parameter.referral || '').trim();
  var xpPoint = 100; // Registration defaults to 100 XP point
  
  if (!fullName || !email || !phone) {
    return { success: false, message: 'Required fields (Full Name, Email Address, and Phone Number) are missing.' };
  }

  // Validate mobile numbers to prevent fake registrations
  if (isFakePhone(phone)) {
    return { success: false, message: 'Invalid phone number. Please enter a valid 10-digit mobile number starting with 6, 7, 8, or 9.' };
  }
  if (guardianContact && isFakePhone(guardianContact)) {
    return { success: false, message: 'Invalid Father\'s/Guardian\'s contact number. Please enter a valid 10-digit mobile number starting with 6, 7, 8, or 9.' };
  }

  // Verify that the email was successfully verified in the OTPs sheet within the last 15 minutes (900,000 ms)
  var otpSheet = ss.getSheetByName('OTPs');
  var emailVerified = false;
  if (otpSheet) {
    var otpData = otpSheet.getDataRange().getValues();
    var nowTime = new Date().getTime();
    for (var k = otpData.length - 1; k >= 1; k--) {
      var oEmail = String(otpData[k][1]).toLowerCase().trim();
      var oStatus = String(otpData[k][3]).trim();
      var oTime = new Date(otpData[k][0]).getTime();
      if (oEmail === email && oStatus === 'verified') {
        if (nowTime - oTime < 900000) { // 15 mins
          emailVerified = true;
          break;
        }
      }
    }
  }
  if (!emailVerified) {
    return { success: false, message: 'Your email address is not verified. Please verify it using OTP first.' };
  }
  
  // Prevent duplicate email/phone check (similar to checkRegistration)
  var data = sheet.getDataRange().getValues();
  var emailCol = getColumnIndexByName(sheet, 'email', 3);
  var phoneCol = getColumnIndexByName(sheet, 'phone', 4);
  var xpCol = getXPColumn(sheet);
  
  // Verify referral exists in the registered phone numbers
  var referrerRowIdx = -1;
  if (referral) {
    var cleanReferral = referral.replace(/[^0-9]/g, '');
    for (var i = 1; i < data.length; i++) {
      var rowPhone = String(data[i][phoneCol - 1] || '').replace(/[^0-9]/g, '');
      if (rowPhone && cleanReferral && rowPhone.indexOf(cleanReferral) !== -1 || cleanReferral.indexOf(rowPhone) !== -1) {
        if (rowPhone.length >= 10 && cleanReferral.length >= 10) {
          referrerRowIdx = i + 1; // 1-indexed row number
          break;
        }
      }
    }
    if (referrerRowIdx === -1) {
      return { success: false, message: 'Referrer mobile number is not registered. Please enter a valid registered mobile number or leave it blank.' };
    }
  }
  
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol - 1] || '').toLowerCase().trim();
    var rowPhone = String(data[i][phoneCol - 1] || '').trim();
    if (email && rowEmail === email) {
      return { success: false, message: 'You are already registered with the same email ID.' };
    }
    if (phone && rowPhone === phone) {
      return { success: false, message: 'You are already registered with the same mobile number.' };
    }
  }
  
  // Header order: Timestamp, Full Name, Email Address, Phone Number, Date of Birth, Father's / Guardian's Name, Father's / Guardian's Contact Number, City, Institute Name, Pin Code, Preparation For, If Any Referral, XP Point, Referral XP
  var headers = [
    'Timestamp',
    'Full Name',
    'Email Address',
    'Phone Number',
    'Date of Birth',
    "Father's / Guardian's Name",
    "Father's / Guardian's Contact Number",
    'City',
    'Institute Name',
    'Pin Code',
    'Preparation For',
    'If Any Referral',
    'XP Point',
    'Referral XP'
  ];
  
  // If the sheet is completely empty, setup the headers
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setFontWeight('bold');
    hr.setBackground('#1a73e8');
    hr.setFontColor('#ffffff');
  } else {
    // If headers exist, let's verify if they match our desired order, if not, write/align them
    var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var match = true;
    for (var j = 0; j < headers.length; j++) {
      if (j >= existingHeaders.length || String(existingHeaders[j]).trim().toLowerCase() !== headers[j].toLowerCase()) {
        match = false;
        break;
      }
    }
    // Only rewrite headers if there are no rows or they don't match, to preserve existing data structure
    if (!match && sheet.getLastRow() === 1) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var hr2 = sheet.getRange(1, 1, 1, headers.length);
      hr2.setFontWeight('bold');
      hr2.setBackground('#1a73e8');
      hr2.setFontColor('#ffffff');
    }
  }
  
  sheet.appendRow([
    new Date(),
    fullName,
    email,
    phone,
    dob,
    guardianName,
    guardianContact,
    city,
    instituteName,
    pinCode,
    preparation,
    referral,
    xpPoint,
    0 // Referral XP default is 0
  ]);
  
  // Award Referral bonus of +100 XP to referrer if referral is verified
  if (referrerRowIdx !== -1) {
    try {
      var refXpCol = getReferralXPColumn(sheet);
      var referrerXP = parseFloat(sheet.getRange(referrerRowIdx, xpCol).getValue()) || 100;
      sheet.getRange(referrerRowIdx, xpCol).setValue(referrerXP + 100);
      var referrerRefXP = parseFloat(sheet.getRange(referrerRowIdx, refXpCol).getValue()) || 0;
      sheet.getRange(referrerRowIdx, refXpCol).setValue(referrerRefXP + 100);
    } catch (err) {
      Logger.log('Error adding referral bonus: ' + err.message);
    }
  }
  
  return { success: true, message: 'Registration successful!' };
}

// ── SEND OTP: Generates and emails a 6-digit verification code
function sendOTP(e) {
  var email = (e.parameter.email || '').toLowerCase().trim();
  if (!email) return { success: false, message: 'Email address is required.' };
  
  // Generate 6-digit OTP
  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OTPs');
  if (!sheet) {
    sheet = ss.insertSheet('OTPs');
    sheet.appendRow(['Timestamp', 'Email', 'OTP', 'Status']);
    var hr = sheet.getRange(1, 1, 1, 4);
    hr.setFontWeight('bold');
    hr.setBackground('#ea4335');
    hr.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  
  // Save OTP in the sheet
  sheet.appendRow([new Date(), email, otp, 'pending']);
  
  // Send email
  try {
    var subject = "Verify your email address - Futrix Exam Portal";
    
    var htmlBody = "<div style=\"background-color:#f8fafc;padding:30px 15px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;\">" +
                   "  <div style=\"max-width:480px;margin:0 auto;background-color:#ffffff;border:1px solid #e2e8f0;border-radius:16px;overflow:hidden;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05),0 2px 4px -1px rgba(0,0,0,0.03);\">" +
                   "    <div style=\"background-color:#1e3a8a;padding:24px;text-align:center;\">" +
                   "      <h1 style=\"color:#ffffff;margin:0;font-size:24px;font-weight:700;letter-spacing:1px;\">FUTRIX</h1>" +
                   "      <p style=\"color:#93c5fd;margin:4px 0 0 0;font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:0.5px;\">Pilot Exam Portal</p>" +
                   "    </div>" +
                   "    <div style=\"padding:30px 24px;\">" +
                   "      <p style=\"font-size:16px;color:#1e293b;margin:0 0 16px 0;line-height:1.5;\">Hello Pilot,</p>" +
                   "      <p style=\"font-size:14px;color:#475569;margin:0 0 24px 0;line-height:1.6;\">Thank you for initiating your registration at the Futrix Pilot Portal. To verify your email address, please enter the one-time verification code below:</p>" +
                   "      <div style=\"text-align:center;margin:24px 0;padding:16px;background-color:#f1f5f9;border-radius:12px;\">" +
                   "        <span style=\"font-family:'Courier New',Courier,monospace;font-size:32px;font-weight:700;color:#1e3a8a;letter-spacing:6px;display:inline-block;padding-left:6px;\">" + otp + "</span>" +
                   "      </div>" +
                   "      <p style=\"font-size:12px;color:#64748b;margin:0 0 24px 0;line-height:1.5;text-align:center;\">This verification code is valid for 10 minutes. For security reasons, please do not share this code with anyone.</p>" +
                   "      <div style=\"border-top:1px solid #e2e8f0;padding-top:20px;text-align:center;\">" +
                   "        <p style=\"font-size:11px;color:#94a3b8;margin:0;line-height:1.6;\">This is an automated security notification from Futrix.<br>If you did not request this code, you can safely ignore this email.</p>" +
                   "      </div>" +
                   "    </div>" +
                   "    <div style=\"background-color:#f8fafc;padding:16px;text-align:center;border-top:1px solid #e2e8f0;\">" +
                   "      <p style=\"font-size:10px;color:#94a3b8;margin:0;\">&copy; 2026 Futrix Exam Portal. All rights reserved.</p>" +
                   "    </div>" +
                   "  </div>" +
                   "</div>";
                   
    var sent = sendEmailViaProvider(email, subject, otp, htmlBody);
    if (sent) {
      return { success: true, message: 'OTP sent successfully to your email.' };
    } else {
      return { success: false, message: 'Failed to dispatch email verification. Please check settings.' };
    }
  } catch (err) {
    return { success: false, message: 'Error sending email: ' + err.message };
  }
}

// ── EMAIL PROVIDER ROUTER: Dispatches the email based on the configured service
function sendEmailViaProvider(email, subject, otp, htmlBody) {
  var service = EMAIL_SERVICE.toUpperCase().trim();
  
  if (service === "BREVO" && BREVO_API_KEY && BREVO_API_KEY !== "YOUR_BREVO_API_KEY_HERE") {
    var url = "https://api.brevo.com/v3/smtp/email";
    var payload = {
      sender: { name: "FUTRIX Pilot Portal", email: SENDER_EMAIL },
      to: [{ email: email }],
      subject: subject,
      htmlContent: htmlBody
    };
    var options = {
      method: "POST",
      headers: {
        "api-key": BREVO_API_KEY,
        "content-type": "application/json",
        "accept": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    var resCode = response.getResponseCode();
    if (resCode === 200 || resCode === 201) return true;
    Logger.log("Brevo API Error: " + response.getContentText());
  } 
  
  if (service === "RESEND" && RESEND_API_KEY && RESEND_API_KEY !== "YOUR_RESEND_API_KEY_HERE") {
    var url = "https://api.resend.com/emails";
    var payload = {
      from: "FUTRIX <onboarding@resend.dev>",
      to: [email],
      subject: subject,
      html: htmlBody
    };
    var options = {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + RESEND_API_KEY,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    var resCode = response.getResponseCode();
    if (resCode === 200 || resCode === 201) return true;
    Logger.log("Resend API Error: " + response.getContentText());
  }

  // Fallback to standard GmailApp
  var plainText = "Hello Pilot,\n\nYour email verification code for FUTRIX is: " + otp + "\n\nThis code is valid for 10 minutes.";
  GmailApp.sendEmail(email, subject, plainText, {
    name: "FUTRIX Pilot Portal",
    htmlBody: htmlBody
  });
  return true;
}

// ── PHONE VALIDATION: Validates Indian mobile number format & prevents dummy numbers
function isFakePhone(phone) {
  if (!phone) return true;
  var clean = phone.replace(/[^0-9]/g, '');
  if (clean.length !== 10) return true;
  if (!/^[6-9]/.test(clean)) return true;
  if (/(\d)\1{4,}/.test(clean)) return true;
  if (/^(\d{2})\1{4}$/.test(clean)) return true;
  if (/^(\d{3})\1{2}\d$/.test(clean)) return true;
  
  var sequentialUp = "0123456789";
  var sequentialDown = "9876543210";
  if (sequentialUp.indexOf(clean) !== -1 || sequentialDown.indexOf(clean) !== -1) return true;
  
  var testPatterns = [
    "1234512345", "9876598765", "6789067890", "1234567890", "0123456789", 
    "9876543210", "8765432109", "7654321098", "6543210987", "5432109876"
  ];
  if (testPatterns.indexOf(clean) !== -1) return true;
  
  return false;
}

// ── VERIFY OTP: Validates the entered OTP code
function verifyOTP(e) {
  var email = (e.parameter.email || '').toLowerCase().trim();
  var otp = (e.parameter.otp || '').trim();
  if (!email || !otp) return { success: false, message: 'Email and OTP are required.' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('OTPs');
  if (!sheet) return { success: false, message: 'No OTP records found.' };
  
  var data = sheet.getDataRange().getValues();
  var now = new Date().getTime();
  
  // Search from the end to get the latest OTP
  for (var i = data.length - 1; i >= 1; i--) {
    var rowEmail = String(data[i][1]).toLowerCase().trim();
    var rowOtp = String(data[i][2]).trim();
    var rowStatus = String(data[i][3]).trim();
    var rowTime = new Date(data[i][0]).getTime();
    
    if (rowEmail === email && rowOtp === otp) {
      if (rowStatus === 'verified') {
        return { success: true, message: 'Email already verified.' };
      }
      // Check if OTP is less than 10 minutes old (600,000 ms)
      if (now - rowTime < 600000) {
        sheet.getRange(i + 1, 4).setValue('verified');
        return { success: true, message: 'Email verified successfully!' };
      } else {
        return { success: false, message: 'OTP has expired. Please request a new one.' };
      }
    }
  }
  return { success: false, message: 'Invalid OTP code. Please try again.' };
}

// ── AUTHORIZE SCRIPT: Run this function once in the Apps Script Editor to trigger the OAuth Permission dialog
function authorizeScript() {
  Logger.log("Requesting authorization for email sending...");
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      GmailApp.sendEmail(email, "FUTRIX Pilot Portal - Authorization Success", "Your Google Apps Script has been successfully authorized to send email verification OTPs using GmailApp!");
      Logger.log("Test email sent to " + email + " successfully. Authorization complete!");
    } else {
      Logger.log("Active user email not found. Please ensure you are logged into your Google Account.");
    }
  } catch (err) {
    Logger.log("Authorization error or GmailApp failure: " + err.message);
  }
}
