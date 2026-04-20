const SETTINGS = {
  SPREADSHEET_ID: '1bZJGa5Gq0EBCrJKKUM57oJzKT5jIXDFCTHsL9kvWSiE',
  SHEET_NAME: 'Orders',
  COMPANY_CODE: 'ZY',
  DRIVE_FOLDER_ID: '1eNOAvweC8HEvZcA5x5B8SMlNSWxPnBF_',
  MAX_ATTACHMENT_BYTES: 8 * 1024 * 1024,
};

const ORDER_HEADERS = [
  'Submitted At',
  'Order ID',
  'Name',
  'Phone Number',
  'Email ID',
  'Product',
  'Quantity',
  'Address',
  'Customization Notes',
  'Attachment Name',
  'Attachment Link',
  'Attachment Preview',
];



function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = submitOrder(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, ...result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function submitOrder(order) {
  const cleanOrder = validateOrder_(order);
  const sheet = getOrdersSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const submittedAt = new Date();
    const orderId = generateOrderId_(sheet, submittedAt);
    const attachment = saveCustomizationAttachment_(cleanOrder.customizationFile, orderId);
    const nextRow = sheet.getLastRow() + 1;

    sheet.getRange(nextRow, 1, 1, ORDER_HEADERS.length).setValues([[
      submittedAt,
      orderId,
      cleanOrder.name,
      cleanOrder.phone,
      cleanOrder.email,
      cleanOrder.product,
      cleanOrder.quantity,
      cleanOrder.address,
      cleanOrder.customizationNotes,
      attachment.name,
      attachment.linkFormula,
      attachment.previewFormula,
    ]]);
    if (attachment.previewFormula) {
      sheet.setRowHeight(nextRow, 130);
    }
    SpreadsheetApp.flush();

    return {
      orderId: orderId,
      submittedAt: submittedAt.toISOString(),
      rowNumber: nextRow,
      attachmentUrl: attachment.url,
    };
  } finally {
    lock.releaseLock();
  }
}

function getOrdersSheet_() {
  const configuredSpreadsheetId = getSetting_('SPREADSHEET_ID', SETTINGS.SPREADSHEET_ID);
  const spreadsheetId = parseSpreadsheetId_(configuredSpreadsheetId);
  if (!spreadsheetId || spreadsheetId === 'PASTE_YOUR_SPREADSHEET_ID_HERE') {
    throw new Error('Set a valid sheet ID in SETTINGS.SPREADSHEET_ID or Script Properties (SPREADSHEET_ID).');
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(SETTINGS.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS.SHEET_NAME);
  }

  ensureOrderHeaders_(sheet);
  return sheet;
}

function ensureOrderHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, ORDER_HEADERS.length).getValues()[0];
  const orderSheetDetected = String(currentHeaders[1] || '').trim().toLowerCase() === 'order id';
  if (!orderSheetDetected) {
    return;
  }

  const requiresUpdate = ORDER_HEADERS.some(function (header, index) {
    return String(currentHeaders[index] || '').trim() !== header;
  });
  if (requiresUpdate) {
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
  }
}

function parseSpreadsheetId_(value) {
  return extractDriveFileId_(value);
}

function getSetting_(key, fallbackValue) {
  const scriptValue = PropertiesService.getScriptProperties().getProperty(key);
  return String(scriptValue || fallbackValue || '').trim();
}

function extractDriveFileId_(value) {
  const text = String(value || '').trim();
  if (!text) {
    return '';
  }

  const patterns = [
    /\/d\/([-\w]{25,})/,
    /[?&]id=([-\w]{25,})/,
    /^([-\w]{25,})$/,
  ];

  for (let i = 0; i < patterns.length; i += 1) {
    const match = text.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }

  return '';
}

function generateOrderId_(sheet, submittedAt) {
  const timeZone = Session.getScriptTimeZone();
  const datePart = Utilities.formatDate(submittedAt, timeZone, 'yyyyMMdd');
  const orderPrefix = SETTINGS.COMPANY_CODE + datePart;
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return orderPrefix + '01';
  }

  const existingIds = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let maxSequence = 0;

  for (let i = 0; i < existingIds.length; i += 1) {
    const orderId = String(existingIds[i][0] || '').trim();
    if (!orderId || !orderId.startsWith(orderPrefix)) {
      continue;
    }

    const sequence = Number(orderId.slice(orderPrefix.length));
    if (Number.isInteger(sequence) && sequence > maxSequence) {
      maxSequence = sequence;
    }
  }

  const nextSequence = String(maxSequence + 1).padStart(2, '0');
  return orderPrefix + nextSequence;
}

function saveCustomizationAttachment_(file, orderId) {
  if (!file) {
    return { name: '', url: '', linkFormula: '', previewFormula: '' };
  }

  const bytes = Utilities.base64Decode(file.base64Data);
  if (!bytes || bytes.length === 0) {
    throw new Error('Attachment could not be decoded.');
  }
  if (bytes.length > SETTINGS.MAX_ATTACHMENT_BYTES) {
    throw new Error('Attachment is too large. Keep it under 8 MB.');
  }

  const safeName = sanitizeFileName_(file.name || 'attachment');
  const finalName = orderId + '_' + safeName;
  const blob = Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', finalName);
  const configuredFolderId = extractDriveFileId_(getSetting_('DRIVE_FOLDER_ID', SETTINGS.DRIVE_FOLDER_ID));

  let savedFile;
  if (configuredFolderId) {
    savedFile = DriveApp.getFolderById(configuredFolderId).createFile(blob);
  } else {
    savedFile = DriveApp.createFile(blob);
  }
  savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = savedFile.getId();
  const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing';
  const imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
  const isImage = String(file.mimeType || '').toLowerCase().startsWith('image/');

  return {
    name: savedFile.getName(),
    url: fileUrl,
    linkFormula: '=HYPERLINK("' + fileUrl + '","View File")',
    previewFormula: isImage ? '=IMAGE("' + imageUrl + '")' : '',
  };
}

function sanitizeFileName_(name) {
  const cleanedName = String(name || '')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
  return cleanedName || 'attachment';
}

function validateOrder_(order) {
  if (!order || typeof order !== 'object') {
    throw new Error('Order data is required.');
  }

  const name = String(order.name || '').trim();
  const phone = String(order.phone || '').trim();
  const email = String(order.email || '').trim();
  const product = String(order.product || '').trim();
  const quantity = String(order.quantity || '').trim();
  const address = String(order.address || '').trim();
  const customizationNotes = String(order.customizationNotes || '').trim();
  const customizationFile = validateCustomizationFile_(order.customizationFile);

  if (!name) {
    throw new Error('Name is required.');
  }
  if (!phone) {
    throw new Error('Phone number is required.');
  }
  if (!email) {
    throw new Error('Email ID is required.');
  }
  if (!product) {
    throw new Error('Please select a product.');
  }
  if (!quantity || isNaN(Number(quantity)) || Number(quantity) < 1) {
    throw new Error('Please enter a valid quantity.');
  }
  if (!address) {
    throw new Error('Address is required.');
  }

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailPattern.test(email)) {
    throw new Error('Enter a valid email ID.');
  }

  return {
    name: name,
    phone: phone,
    email: email,
    product: product,
    quantity: quantity,
    address: address,
    customizationNotes: customizationNotes,
    customizationFile: customizationFile,
  };
}

function validateCustomizationFile_(file) {
  if (!file) {
    return null;
  }
  if (typeof file !== 'object') {
    throw new Error('Invalid attachment data.');
  }

  const name = String(file.name || '').trim();
  const mimeType = String(file.mimeType || '').trim() || 'application/octet-stream';
  const base64Data = String(file.base64Data || '').trim();
  const size = Number(file.size || 0);

  if (!name || !base64Data) {
    throw new Error('Invalid attachment data.');
  }

  const estimatedSize = estimateBase64Size_(base64Data);
  if (estimatedSize > SETTINGS.MAX_ATTACHMENT_BYTES || size > SETTINGS.MAX_ATTACHMENT_BYTES) {
    throw new Error('Attachment is too large. Keep it under 8 MB.');
  }

  return {
    name: name,
    mimeType: mimeType,
    base64Data: base64Data,
    size: estimatedSize,
  };
}

function estimateBase64Size_(base64Data) {
  const length = base64Data.length;
  let padding = 0;
  if (base64Data.endsWith('==')) {
    padding = 2;
  } else if (base64Data.endsWith('=')) {
    padding = 1;
  }
  return Math.floor((length * 3) / 4) - padding;
}

function backfillAttachmentPreviews() {
  const sheet = getOrdersSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, ORDER_HEADERS.length).getValues();
  let updatedRows = 0;

  for (let i = 0; i < rows.length; i += 1) {
    const rowNumber = i + 2;
    const attachmentName = String(rows[i][9] || '').trim();
    const linkCellValue = String(rows[i][10] || '').trim();
    const previewCellValue = String(rows[i][11] || '').trim();

    if (!attachmentName || previewCellValue) {
      continue;
    }

    const fileId = extractDriveFileId_(linkCellValue);
    if (!fileId) {
      continue;
    }

    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view?usp=sharing';
    const imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    const isImage = String(file.getMimeType() || '').toLowerCase().startsWith('image/');

    sheet.getRange(rowNumber, 11).setFormula('=HYPERLINK("' + fileUrl + '","View File")');
    if (isImage) {
      sheet.getRange(rowNumber, 12).setFormula('=IMAGE("' + imageUrl + '")');
      sheet.setRowHeight(rowNumber, 130);
    }

    updatedRows += 1;
  }

  SpreadsheetApp.flush();
  return updatedRows;
}
