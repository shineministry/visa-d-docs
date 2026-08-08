// ============================================
// Visa D Application - Google Apps Script Backend
// Uses existing Google Doc template, fills placeholders, sends copy
// ============================================

const SHEET_NAME = 'Applications';
const FOLDER_NAME = 'Visa D Itineraries';

// Template document ID from your Google Docs URL
// https://docs.google.com/document/d/1vhlz9vgiV3Eeqg0xjzse25W99CpP59HPrz2tWcl3vpI/edit
const TEMPLATE_DOC_ID = '1pbK2IfT1eTrScej1yUWRavA49qvCkAHlNUYRX2KFfvM';

// -----------------------------------------------
// doPost - Main entry point
// -----------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'demo') {
      return sendDemoEmail(data.email);
    }

    if (data.action === 'preview') {
      return handlePreview(data);
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

  const doc = fillTemplate(data);
  const docUrl = doc.getUrl();
  const pdfBlob = doc.getAs('application/pdf').setName(doc.getName() + '.pdf');

  const subject = `REPUBLIC OF AUSTRIA - Travel Itinerary - ${(data.firstName || '').toUpperCase()} ${(data.surname || '').toUpperCase()}`;

  GmailApp.sendEmail(data.email, subject, '', {
    htmlBody: buildEmailBody(data, docUrl),
    attachments: [pdfBlob],
    name: 'Austrian Visa Application Portal'
  });

  const adminEmail = Session.getActiveUser().getEmail();
  if (adminEmail && adminEmail !== data.email) {
    GmailApp.sendEmail(adminEmail, `New Application: ${data.firstName} ${data.surname}`, '', {
      htmlBody: buildAdminNotificationHtml(data, docUrl)
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Application submitted successfully', docUrl: docUrl }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------
// Fill Template - Copy doc and replace placeholders
// -----------------------------------------------
function fillTemplate(data) {
  const surname = (data.surname || '').toUpperCase();
  const firstName = (data.firstName || '').toUpperCase();
  const fullName = surname + ' ' + firstName;
  const nationality = data.currentNationality || data.nationality || '';
  const passportNo = data.docNumber || data.passportNumber || '';
  const visaCategory = data.visaCategory || 'Residence Permit (Pupil) – National Visa (D)';
  const arrival = formatDate(data.arrivalDate);
  const borderCrossing = data.borderCrossing || '';
  const purpose = data.purposeOfJourney || '';
  const invitingPerson = data.invitingPerson || '';
  const invitingAddress = data.invitingAddress || '';
  const employerInfo = data.employerInfo || '';
  const place = data.place || '';
  const destination = data.internalTo || 'Salzburg';
  const stayDuration = data.stayDuration || '';
  const departure = formatDate(data.departureDate);

  // Get or create folder
  let folder;
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(FOLDER_NAME);
  }

  // Copy template document
  const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
  const newFile = templateFile.makeCopy(`${surname} ${firstName} - Travel Itinerary`, folder);
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  // Replace all placeholders
  const replacements = {
    '{{SURNAME}}': surname,
    '{{FIRSTNAME}}': firstName,
    '{{fullName}}': fullName,
    '{{nationality}}': nationality,
    '{{passportCountry}}': nationality === 'Indian' ? 'India' : nationality,
    '{{passportNumber}}': passportNo,
    '{{visaCategory}}': visaCategory,
    '{{arrivalDate}}': arrival,
    '{{departureDate}}': departure,
    '{{borderCrossing}}': borderCrossing,
    '{{arrivalAirport}}': data.flights && data.flights[0] ? (data.flights[0].arrivalAirport || 'Vienna International Airport (VIE)') + ', ' + (data.flights[0].arrivalCountry || 'Austria') : 'Vienna International Airport (VIE), Austria',
    '{{flightSector}}': data.flights && data.flights[0] ? data.flights[0].sector || 'Outbound Journey' : 'Outbound Journey',
    '{{flightPurpose}}': data.flights && data.flights[0] ? data.flights[0].purpose || 'Entry into Austria for ' + purpose : 'Entry into Austria for ' + purpose,
    '{{purpose}}': purpose,
    '{{hostName}}': invitingPerson,
    '{{hostAddress}}': invitingAddress,
    '{{employerDetails}}': employerInfo,
    '{{institutionLabel}}': purpose.toLowerCase().includes('study') || purpose.toLowerCase().includes('pupil') ? 'School' : purpose.toLowerCase().includes('employ') ? 'Employer' : 'Host Institution',
    '{{institutionName}}': employerInfo,
    '{{destination}}': destination,
    '{{travelMethod}}': data.travelMethod || 'private vehicle',
    '{{internalFrom}}': data.internalFrom || 'Vienna International Airport (VIE)',
    '{{internalTo}}': destination,
    '{{arrangedBy}}': data.arrangedBy || invitingPerson,
    '{{email}}': data.email || '',
    '{{dateOfBirth}}': formatDate(data.dateOfBirth),
    '{{placeOfBirth}}': data.placeOfBirth || '',
    '{{countryOfBirth}}': data.countryOfBirth || '',
    '{{sex}}': data.sex || '',
    '{{maritalStatus}}': data.maritalStatus || '',
    '{{passportIssueDate}}': formatDate(data.docIssueDate),
    '{{passportExpiry}}': formatDate(data.docValidUntil),
    '{{passportIssuer}}': data.docIssuedBy || '',
    '{{telephone}}': data.telephone || '',
    '{{homeAddress}}': data.homeAddress || '',
    '{{occupation}}': data.currentOccupation || '',
    '{{entries}}': data.entriesRequested || '',
    '{{fundingSource}}': data.costCoveredBy || '',
    '{{meansOfSupport}}': data.meansOfSupport || '',
    '{{hostPhone}}': data.invitingPhone || '',
    '{{signingPlace}}': place,
    '{{signingDate}}': formatDate(data.date || new Date().toISOString().split('T')[0])
  };

  // Replace each placeholder in the document body
  for (const [placeholder, value] of Object.entries(replacements)) {
    body.replaceText(placeholder, value);
  }

  doc.saveAndClose();

  // Set sharing
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log('Document created: ' + doc.getUrl());
  return doc;
}

// -----------------------------------------------
// Handle Preview - Create temporary doc for preview
// -----------------------------------------------
function handlePreview(data) {
  const doc = fillTemplate(data);
  const docUrl = doc.getUrl();

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, docUrl: docUrl }))
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
    visaCategory: 'Residence Permit (Pupil) – National Visa (D)',
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
    date: '{{signingDate}}',
    internalTo: '{{destination}}'
  };

  const doc = fillTemplate(demoData);
  const docUrl = doc.getUrl();
  const pdfBlob = doc.getAs('application/pdf').setName(doc.getName() + '.pdf');

  const subject = '[DEMO] REPUBLIC OF AUSTRIA - Travel Itinerary';

  GmailApp.sendEmail(email, subject, '', {
    htmlBody: buildEmailBody(demoData, docUrl),
    attachments: [pdfBlob],
    name: 'Austrian Visa Application Portal [DEMO]'
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'Demo email sent', docUrl: docUrl }))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------
// Build Email Body
// -----------------------------------------------
function buildEmailBody(data, docUrl) {
  const fullName = ((data.surname || '').toUpperCase() + ' ' + (data.firstName || '').toUpperCase()).trim();

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;padding:20px;background:#f5f5f5;"><div style="max-width:600px;margin:0 auto;background:white;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.1);"><div style="background:linear-gradient(135deg,#673ab7,#7c4dff);padding:24px;text-align:center;"><h1 style="color:white;font-size:20px;margin:0;">REPUBLIC OF AUSTRIA</h1><p style="color:rgba(255,255,255,0.85);font-size:12px;margin:4px 0 0;">Travel Itinerary Document</p></div><div style="padding:24px;"><p style="font-size:14px;color:#333;">Dear <strong>' + fullName + '</strong>,</p><p style="font-size:13px;color:#555;line-height:1.6;">Your travel itinerary document has been generated. Please find the PDF attached to this email.</p><p style="font-size:13px;color:#555;line-height:1.6;">You can also view the document online:</p><div style="text-align:center;margin:20px 0;"><a href="' + docUrl + '" style="display:inline-block;background:#673ab7;color:white;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:13px;font-weight:500;">View Document</a></div><p style="font-size:11px;color:#999;margin-top:20px;border-top:1px solid #eee;padding-top:12px;">This is an automated email from the Austrian Visa Application Portal.<br>Application Reference: VD-' + Date.now().toString(36).toUpperCase() + '</p></div></div></body></html>';
}

// -----------------------------------------------
// Build Admin Notification
// -----------------------------------------------
function buildAdminNotificationHtml(data, docUrl) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="font-family:Arial,sans-serif;padding:20px;"><h2 style="color:#673ab7;">New Visa D Application Received</h2><table style="border-collapse:collapse;width:100%;max-width:500px;"><tr><td style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Name</td><td style="padding:8px;border-bottom:1px solid #eee;">' + data.firstName + ' ' + data.surname + '</td></tr><tr><td style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Email</td><td style="padding:8px;border-bottom:1px solid #eee;">' + data.email + '</td></tr><tr><td style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Nationality</td><td style="padding:8px;border-bottom:1px solid #eee;">' + data.currentNationality + '</td></tr><tr><td style="padding:8px;font-weight:bold;border-bottom:1px solid #eee;">Purpose</td><td style="padding:8px;border-bottom:1px solid #eee;">' + data.purposeOfJourney + '</td></tr><tr><td style="padding:8px;font-weight:bold;">Document</td><td style="padding:8px;"><a href="' + docUrl + '">View Itinerary</a></td></tr></table></body></html>';
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
// PREVIEW - Creates doc and returns URL
// -----------------------------------------------
function previewDocument() {
  const testData = {
    surname: '{{surname}}',
    firstName: '{{firstName}}',
    nationality: '{{nationality}}',
    currentNationality: '{{nationality}}',
    passportNumber: '{{passportNumber}}',
    visaCategory: 'Residence Permit (Pupil) – National Visa (D)',
    purposeOfJourney: '{{purpose}}',
    arrivalDate: '{{arrivalDate}}',
    departureDate: '{{departureDate}}',
    stayDuration: '{{stayDuration}}',
    borderCrossing: '{{borderCrossing}}',
    invitingPerson: '{{hostName}}',
    invitingAddress: '{{hostAddress}}',
    employerInfo: '{{employerDetails}}',
    place: '{{signingPlace}}',
    date: '{{signingDate}}',
    internalTo: '{{destination}}',
    borderCountry: '{{borderCountry}}'
  };

  const doc = fillTemplate(testData);
  Logger.log('Preview URL: ' + doc.getUrl());
  return doc.getUrl();
}

// -----------------------------------------------
// SETUP TEMPLATE - Run this ONCE to update your Google Doc
// -----------------------------------------------
function setupTemplate() {
  const doc = DocumentApp.openById(TEMPLATE_DOC_ID);
  const body = doc.getBody();

  // Clear entire document
  body.clear();

  // Set margins
  body.setMarginTop(72);
  body.setMarginBottom(72);
  body.setMarginLeft(72);
  body.setMarginRight(72);

  // Build document content
  const content = [
    { text: '{{SURNAME}} {{FIRSTNAME}} - Travel Itinerary', style: 'title' },
    { text: '', style: 'normal' },
    { text: 'REPUBLIC OF AUSTRIA', style: 'centered' },
    { text: 'RESIDENCE PERMIT (PUPIL) – STUDENT VISA', style: 'centered' },
    { text: 'PROPOSED TRAVEL ITINERARY', style: 'centered' },
    { text: '', style: 'normal' },
    { text: 'APPLICANT INFORMATION', style: 'heading' },
    { text: '', style: 'normal' },
    { text: 'table:Particulars|Details|Surname|{{SURNAME}}|Given Name|{{FIRSTNAME}}|Nationality|{{nationality}}|Passport Nationality|Republic of {{passportCountry}}|Visa Category|{{visaCategory}}|Country of Destination|Austria|Passport No|{{passportNumber}}', style: 'table' },
    { text: '', style: 'normal' },
    { text: 'INTERNATIONAL FLIGHT ITINERARY', style: 'heading' },
    { text: '', style: 'normal' },
    { text: 'table:Travel Sector|Date|Departure|Arrival|Purpose|{{flightSector}}|{{arrivalDate}}|{{borderCrossing}}|{{arrivalAirport}}|{{flightPurpose}}', style: 'table' },
    { text: '', style: 'normal' },
    { text: 'DETAILED TRAVEL PLAN', style: 'heading' },
    { text: '', style: 'normal' },
    { text: 'Departure from {{passportCountry}}', style: 'subheading' },
    { text: 'The applicant intends to depart from {{borderCrossing}} on {{arrivalDate}}, along with his parents, for Vienna, Austria.', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'Arrival in Austria', style: 'subheading' },
    { text: 'The applicant will arrive at {{arrivalAirport}} for the purpose of {{purpose}} under a {{visaCategory}}.', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'Internal Travel within Austria', style: 'subheading' },
    { text: 'Upon arrival in Vienna, the applicant will be received by:', style: 'normal' },
    { text: '{{hostName}}', style: 'normal' },
    { text: 'The applicant will thereafter travel from Vienna to {{destination}} by {{travelMethod}} arranged by {{arrangedBy}}.', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'INSTITUTION DETAILS', style: 'heading' },
    { text: '', style: 'normal' },
    { text: 'table:Particulars|Details|{{institutionLabel}}|{{institutionName}}|Purpose of Stay|{{purpose}}|Country of Studies|Austria', style: 'table' },
    { text: '', style: 'normal' },
    { text: 'ACCOMMODATION / DESTINATION ADDRESS', style: 'heading' },
    { text: '{{hostAddress}}', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'PURPOSE OF TRAVEL', style: 'heading' },
    { text: 'The purpose of travel is to enter Austria legally under a {{visaCategory}} for {{purpose}} and to reside in Austria in accordance with Austrian immigration and residence regulations.', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'SUPPORTING TRAVEL INFORMATION', style: 'heading' },
    { text: 'Valid {{nationality}} Passport held by applicant', style: 'normal' },
    { text: 'Confirmed destination and accommodation details in Austria', style: 'normal' },
    { text: 'Internal transportation arranged from Vienna to {{destination}}', style: 'normal' },
    { text: 'Compliance with Austrian visa and immigration requirements', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'DECLARATION', style: 'heading' },
    { text: 'I hereby declare that the above-mentioned travel information and itinerary are true and correct to the best of my knowledge and are submitted in support of the Austria {{visaCategory}} application.', style: 'normal' },
    { text: '', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'Applicant Signature: ___________________________', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'Name of Applicant: {{SURNAME}} {{FIRSTNAME}}', style: 'normal' },
    { text: '', style: 'normal' },
    { text: 'Date: {{signingDate}}', style: 'normal' },
    { text: 'Place: {{signingPlace}}', style: 'normal' }
  ];

  // Build document
  content.forEach(item => {
    if (item.style === 'table' && item.text.startsWith('table:')) {
      const tableData = item.text.replace('table:', '').split('|');
      const rows = [];
      for (let i = 0; i < tableData.length; i += 2) {
        rows.push([tableData[i], tableData[i + 1]]);
      }
      const table = body.appendTable(rows);
      // Style header row
      const headerRow = table.getRow(0);
      for (let c = 0; c < headerRow.getNumCells(); c++) {
        const cell = headerRow.getCell(c);
        cell.setBold(true);
        cell.setFontFamily('Play');
        cell.setFontSize(11);
      }
      // Style all cells
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          cell.setFontFamily('Play');
          cell.setFontSize(11);
        }
      }
    } else if (item.style === 'title') {
      const p = body.appendParagraph(item.text);
      p.setFontFamily('Play');
      p.setFontSize(20);
      p.setBold(true);
      p.setForegroundColor('#1a73e8');
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else if (item.style === 'centered') {
      const p = body.appendParagraph(item.text);
      p.setFontFamily('Play');
      p.setFontSize(13);
      p.setBold(true);
      p.setForegroundColor('#1a73e8');
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    } else if (item.style === 'heading') {
      const p = body.appendParagraph(item.text);
      p.setFontFamily('Play');
      p.setFontSize(14);
      p.setBold(true);
      p.setForegroundColor('#1a73e8');
      p.setSpacingBefore(16);
      p.setSpacingAfter(6);
      body.appendHorizontalRule();
    } else if (item.style === 'subheading') {
      const p = body.appendParagraph(item.text);
      p.setFontFamily('Play');
      p.setFontSize(11);
      p.setBold(true);
      p.setSpacingAfter(4);
    } else {
      const p = body.appendParagraph(item.text);
      p.setFontFamily('Play');
      p.setFontSize(11);
      p.setSpacingAfter(4);
    }
  });

  doc.saveAndClose();
  Logger.log('Template updated: ' + doc.getUrl());
  return doc.getUrl();
}
