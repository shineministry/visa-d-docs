// ============================================
// Visa D Application - Google Apps Script Backend
// Generates formal itinerary documents
// ============================================

const SHEET_NAME = 'Applications';

// -----------------------------------------------
// doPost - Main entry point
// -----------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'demo') {
      return sendDemoEmail(data.email);
    }

    return handleFormSubmission(data);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------
// Handle Form Submission
// -----------------------------------------------
function handleFormSubmission(data) {
  saveToSheet(data);

  const itineraryHtml = buildItineraryDocument(data);
  const subject = `REPUBLIC OF AUSTRIA - Travel Itinerary - ${(data.firstName || '').toUpperCase()} ${(data.surname || '').toUpperCase()}`;

  GmailApp.sendEmail(data.email, subject, '', {
    htmlBody: itineraryHtml,
    name: 'Austrian Visa Application Portal'
  });

  const adminEmail = Session.getActiveUser().getEmail();
  if (adminEmail && adminEmail !== data.email) {
    GmailApp.sendEmail(adminEmail, `New Application: ${data.firstName} ${data.surname}`, '', {
      htmlBody: buildAdminNotificationHtml(data)
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Application submitted successfully' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------
// Send Demo Email
// -----------------------------------------------
function sendDemoEmail(email) {
  const demoData = {
    surname: '{{surname}}',
    firstName: '{{firstName}}',
    nationality: '{{nationality}}',
    currentNationality: '{{nationality}}',
    passportNumber: '{{passportNumber}}',
    travelDocType: '{{travelDocType}}',
    purposeOfJourney: '{{purpose}}',
    arrivalDate: '{{arrivalDate}}',
    departureDate: '{{departureDate}}',
    stayDuration: '{{stayDuration}}',
    borderCrossing: '{{borderCrossing}}',
    invitingPerson: '{{hostName}}',
    invitingAddress: '{{hostAddress}}',
    invitingPhone: '{{hostPhone}}',
    homeAddress: '{{homeAddress}}',
    email: email,
    dateOfBirth: '{{dateOfBirth}}',
    placeOfBirth: '{{placeOfBirth}}',
    countryOfBirth: '{{countryOfBirth}}',
    sex: '{{sex}}',
    maritalStatus: '{{maritalStatus}}',
    docNumber: '{{passportNumber}}',
    docIssueDate: '{{passportIssueDate}}',
    docValidUntil: '{{passportExpiry}}',
    docIssuedBy: '{{passportIssuer}}',
    telephone: '{{telephone}}',
    currentOccupation: '{{occupation}}',
    employerInfo: '{{employerDetails}}',
    entriesRequested: '{{entries}}',
    costCoveredBy: '{{fundingSource}}',
    meansOfSupport: '{{meansOfSupport}}',
    place: '{{signingPlace}}',
    date: '{{signingDate}}'
  };

  const html = buildItineraryDocument(demoData);
  const subject = `[DEMO] REPUBLIC OF AUSTRIA - Travel Itinerary`;

  GmailApp.sendEmail(email, subject, '', {
    htmlBody: html,
    name: 'Austrian Visa Application Portal [DEMO]'
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Demo email sent' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------
// Build Formal Itinerary Document
// -----------------------------------------------
function buildItineraryDocument(data) {
  const surname = (data.surname || '{{surname}}').toUpperCase();
  const firstName = (data.firstName || '{{firstName}}').toUpperCase();
  const fullName = `${surname} ${firstName}`;
  const arrival = formatDate(data.arrivalDate);
  const departure = formatDate(data.departureDate);
  const nationality = data.currentNationality || data.nationality || '{{nationality}}';
  const passportNo = data.docNumber || data.passportNumber || '{{passportNumber}}';
  const place = data.place || '{{signingPlace}}';
  const purpose = data.purposeOfJourney || '{{purpose}}';
  const invitingPerson = data.invitingPerson || '{{hostName}}';
  const invitingAddress = data.invitingAddress || '{{hostAddress}}';
  const employerInfo = data.employerInfo || '{{employerDetails}}';
  const borderCrossing = data.borderCrossing || '{{borderCrossing}}';

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: 'Times New Roman', Times, serif;
      background: #f5f5f5;
      margin: 0;
      padding: 20px;
      color: #000;
    }
    .document {
      max-width: 700px;
      margin: 0 auto;
      background: white;
      padding: 40px 50px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      line-height: 1.6;
    }
    .doc-header {
      text-align: center;
      border-bottom: 3px double #000;
      padding-bottom: 16px;
      margin-bottom: 24px;
    }
    .doc-header h1 {
      font-size: 16px;
      font-weight: bold;
      letter-spacing: 2px;
      margin: 0 0 4px 0;
      text-transform: uppercase;
    }
    .doc-header h2 {
      font-size: 13px;
      font-weight: normal;
      margin: 0;
      color: #333;
    }
    .section-title {
      font-size: 13px;
      font-weight: bold;
      text-transform: uppercase;
      letter-spacing: 1px;
      margin: 24px 0 12px 0;
      padding-bottom: 4px;
      border-bottom: 1px solid #ccc;
    }
    .info-table {
      width: 100%;
      border-collapse: collapse;
      margin: 8px 0 16px 0;
    }
    .info-table td {
      padding: 6px 12px;
      font-size: 13px;
      vertical-align: top;
    }
    .info-table td:first-child {
      font-weight: bold;
      width: 180px;
      color: #333;
    }
    .itinerary-table {
      width: 100%;
      border-collapse: collapse;
      margin: 12px 0;
      font-size: 12px;
    }
    .itinerary-table th {
      background: #f0f0f0;
      padding: 8px 10px;
      text-align: left;
      font-weight: bold;
      border: 1px solid #ccc;
      font-size: 11px;
      text-transform: uppercase;
    }
    .itinerary-table td {
      padding: 8px 10px;
      border: 1px solid #ccc;
      vertical-align: top;
    }
    .paragraph {
      font-size: 13px;
      margin: 10px 0;
      text-align: justify;
    }
    .declaration {
      margin-top: 30px;
      padding-top: 20px;
      border-top: 1px solid #ccc;
    }
    .signature-block {
      margin-top: 40px;
      display: flex;
      justify-content: space-between;
    }
    .sig-line {
      width: 200px;
      border-top: 1px solid #000;
      padding-top: 4px;
      font-size: 12px;
      text-align: center;
    }
    .footer-note {
      margin-top: 30px;
      padding-top: 12px;
      border-top: 2px solid #000;
      font-size: 10px;
      color: #666;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="document">

    <div class="doc-header">
      <h1>REPUBLIC OF AUSTRIA</h1>
      <h2>PROPOSED TRAVEL ITINERARY</h2>
    </div>

    <div class="section-title">APPLICANT INFORMATION</div>
    <table class="info-table">
      <tr><td>Surname</td><td>${surname}</td></tr>
      <tr><td>Given Name</td><td>${firstName}</td></tr>
      <tr><td>Nationality</td><td>${nationality}</td></tr>
      <tr><td>Passport Nationality</td><td>${nationality}</td></tr>
      <tr><td>Visa Category</td><td>Residence Permit (Pupil) – National Visa (D)</td></tr>
      <tr><td>Country of Destination</td><td>Austria</td></tr>
      <tr><td>Passport No</td><td>${passportNo}</td></tr>
    </table>

    <div class="section-title">INTERNATIONAL FLIGHT ITINERARY</div>
    <table class="itinerary-table">
      <tr>
        <th>Travel Sector</th>
        <th>Date</th>
        <th>Departure</th>
        <th>Arrival</th>
        <th>Purpose</th>
      </tr>
      <tr>
        <td>Outbound Journey</td>
        <td>${arrival}</td>
        <td>${borderCrossing}</td>
        <td>Vienna International Airport (VIE), Austria</td>
        <td>Entry into Austria for ${purpose}</td>
      </tr>
    </table>

    <div class="section-title">DETAILED TRAVEL PLAN</div>

    <p class="paragraph">
      <strong>Departure:</strong>
      The applicant intends to depart from ${borderCrossing} on ${arrival} for Vienna, Austria.
    </p>

    <p class="paragraph">
      <strong>Arrival in Austria:</strong>
      The applicant will arrive at Vienna International Airport (VIE), Austria for the purpose of ${purpose} under a Residence Permit (Pupil) and National Visa (D).
    </p>

    <p class="paragraph">
      <strong>Internal Travel within Austria:</strong>
      Upon arrival in Vienna, the applicant will be received by: ${invitingPerson}.
      The applicant will thereafter travel from Vienna to destination by private vehicle arranged by ${invitingPerson}.
    </p>

    <div class="section-title">INSTITUTION DETAILS</div>
    <table class="info-table">
      <tr><td>School</td><td>${employerInfo}</td></tr>
      <tr><td>Purpose of Stay</td><td>${purpose}</td></tr>
      <tr><td>Country of Studies</td><td>Austria</td></tr>
    </table>

    <div class="section-title">ACCOMMODATION / DESTINATION ADDRESS</div>
    <p class="paragraph" style="white-space: pre-line;">${invitingAddress}</p>

    <div class="section-title">PURPOSE OF TRAVEL</div>
    <p class="paragraph">
      The purpose of travel is to enter Austria legally under a National Visa (D) and Residence Permit (Pupil) for ${purpose} purposes and to reside in Austria in accordance with Austrian immigration and residence regulations.
    </p>

    <div class="section-title">SUPPORTING TRAVEL INFORMATION</div>
    <p class="paragraph">
      • Valid ${nationality} Passport held by applicant<br>
      • Confirmed destination and accommodation details in Austria<br>
      • Internal transportation arranged from Vienna to destination<br>
      • Compliance with Austrian visa and immigration requirements
    </p>

    <div class="declaration">
      <div class="section-title">DECLARATION</div>
      <p class="paragraph">
        I hereby declare that the above-mentioned travel information and itinerary are true and correct to the best of my knowledge and are submitted in support of the Austria National Visa (D) and Residence Permit (Pupil) application.
      </p>

      <div class="signature-block">
        <div class="sig-line">Applicant Signature</div>
        <div class="sig-line">Date</div>
        <div class="sig-line">Place: ${place}</div>
      </div>

      <p style="margin-top: 30px; font-size: 12px;">
        Name of Applicant: <strong>${fullName}</strong>
      </p>
    </div>

    <div class="footer-note">
      This document was generated by the Austrian Visa Application Portal<br>
      Application Reference: VD-${Date.now().toString(36).toUpperCase()}
    </div>

  </div>
</body>
</html>`;
}

// -----------------------------------------------
// Build Admin Notification
// -----------------------------------------------
function buildAdminNotificationHtml(data) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif; padding:20px;">
  <h2 style="color:#673ab7;">New Visa D Application Received</h2>
  <table style="border-collapse:collapse; width:100%; max-width:500px;">
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Name</td><td style="padding:8px; border-bottom:1px solid #eee;">${data.firstName} ${data.surname}</td></tr>
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Email</td><td style="padding:8px; border-bottom:1px solid #eee;">${data.email}</td></tr>
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Nationality</td><td style="padding:8px; border-bottom:1px solid #eee;">${data.currentNationality}</td></tr>
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Purpose</td><td style="padding:8px; border-bottom:1px solid #eee;">${data.purposeOfJourney}</td></tr>
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Arrival</td><td style="padding:8px; border-bottom:1px solid #eee;">${formatDate(data.arrivalDate)}</td></tr>
    <tr><td style="padding:8px; font-weight:bold; border-bottom:1px solid #eee;">Departure</td><td style="padding:8px; border-bottom:1px solid #eee;">${formatDate(data.departureDate)}</td></tr>
    <tr><td style="padding:8px; font-weight:bold;">Stay Duration</td><td style="padding:8px;">${data.stayDuration} days</td></tr>
  </table>
</body>
</html>`;
}

// -----------------------------------------------
// Save to Google Sheet
// -----------------------------------------------
function saveToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp', 'Email', 'Surname', 'First Name', 'Date of Birth',
      'Place of Birth', 'Country of Birth', 'Nationality', 'Sex',
      'Marital Status', 'Travel Doc Type', 'Doc Number', 'Doc Issue Date',
      'Doc Valid Until', 'Issued By', 'Home Address', 'Telephone',
      'Occupation', 'Employer', 'Purpose', 'Border Crossing',
      'Entries', 'Stay Duration', 'Arrival Date', 'Departure Date',
      'Inviting Person', 'Inviting Address', 'Cost Covered By',
      'Means of Support', 'Place', 'Date'
    ]);
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#673ab7');
    headerRange.setFontColor('white');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date().toISOString(),
    data.email || '',
    data.surname || '',
    data.firstName || '',
    data.dateOfBirth || '',
    data.placeOfBirth || '',
    data.countryOfBirth || '',
    data.currentNationality || '',
    data.sex || '',
    data.maritalStatus || '',
    data.travelDocType || '',
    data.docNumber || '',
    data.docIssueDate || '',
    data.docValidUntil || '',
    data.docIssuedBy || '',
    data.homeAddress || '',
    data.telephone || '',
    data.currentOccupation || '',
    data.employerInfo || '',
    data.purposeOfJourney || '',
    data.borderCrossing || '',
    data.entriesRequested || '',
    data.stayDuration || '',
    data.arrivalDate || '',
    data.departureDate || '',
    data.invitingPerson || '',
    data.invitingAddress || '',
    data.costCoveredBy || '',
    data.meansOfSupport || '',
    data.place || '',
    data.date || ''
  ]);
}

// -----------------------------------------------
// Utility: Format Date
// -----------------------------------------------
function formatDate(dateStr) {
  if (!dateStr) return '{{date}}';
  try {
    const d = new Date(dateStr);
    return d.toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'long',
      year: 'numeric'
    });
  } catch (e) {
    return dateStr;
  }
}

// -----------------------------------------------
// doGet - Health check
// -----------------------------------------------
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Visa D Application API is running',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------
// PREVIEW - Run this to see the document in browser
// -----------------------------------------------
function previewDocument() {
  const testData = {
    surname: '{{surname}}',
    firstName: '{{firstName}}',
    nationality: '{{nationality}}',
    currentNationality: '{{nationality}}',
    passportNumber: '{{passportNumber}}',
    purposeOfJourney: '{{purpose}}',
    arrivalDate: '{{arrivalDate}}',
    departureDate: '{{departureDate}}',
    stayDuration: '{{stayDuration}}',
    borderCrossing: '{{borderCrossing}}',
    invitingPerson: '{{hostName}}',
    invitingAddress: '{{hostAddress}}',
    employerInfo: '{{employerDetails}}',
    place: '{{signingPlace}}',
    date: '{{signingDate}}'
  };

  const html = buildItineraryDocument(testData);

  const output = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(output, 'Itinerary Preview');
}

// -----------------------------------------------
// PREVIEW IN NEW TAB - Opens preview in new browser tab
// -----------------------------------------------
function previewInNewTab() {
  const testData = {
    surname: '{{surname}}',
    firstName: '{{firstName}}',
    nationality: '{{nationality}}',
    currentNationality: '{{nationality}}',
    passportNumber: '{{passportNumber}}',
    purposeOfJourney: '{{purpose}}',
    arrivalDate: '{{arrivalDate}}',
    departureDate: '{{departureDate}}',
    stayDuration: '{{stayDuration}}',
    borderCrossing: '{{borderCrossing}}',
    invitingPerson: '{{hostName}}',
    invitingAddress: '{{hostAddress}}',
    employerInfo: '{{employerDetails}}',
    place: '{{signingPlace}}',
    date: '{{signingDate}}'
  };

  const html = buildItineraryDocument(testData);

  const blob = HtmlService.createHtmlOutput(html).getBlob();
  const file = DriveApp.createFile(blob).setName('Itinerary Preview.html');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log('Preview URL: ' + file.getUrl());
  return file.getUrl();
}
