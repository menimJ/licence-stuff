const PAYMENT_CONFIG = Object.freeze({
  AMOUNT: null,

  PROVIDER: 'Revop',

  MAX_FILE_SIZE_BYTES: 5 * 1024 * 1024,

  ALLOWED_MIME_TYPES: [
    'application/pdf',
    'image/jpeg',
    'image/png',
  ],

  RETENTION_DAYS_AFTER_CONFIRMATION: 90,

  PAYMENT_PROOF_FOLDER_NAME:
    'CRFFN Payment Proofs',

  PAYMENT_URL:
    'https://revop.gov.ng/payments/generate-bill?org=0229006001000',
});

/**
 * Returns payment details for the applicant portal.
 */

/**
 * Cached payment summary for the applicant portal.
 */
function getApplicantPaymentData(
  applicationId,
  secureToken
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  const cleanToken =
    String(
      secureToken || ''
    ).trim();

  if (
    !cleanId ||
    !cleanToken
  ) {
    return getApplicantPaymentDataUncached_(
      cleanId,
      cleanToken
    );
  }

  const key =
    crffnVersionedApplicantKey_(
      crffnPaymentCacheKey_(
        cleanId,
        cleanToken
      ),
      cleanId
    );

  const cached =
    crffnGetCachedJson_(
      key
    );

  if (cached) {
    cached.performanceCache =
      'HIT';

    return cached;
  }

  const result =
    getApplicantPaymentDataUncached_(
      cleanId,
      cleanToken
    );

  if (
    result &&
    result.ok === true
  ) {
    crffnPutCachedJson_(
      key,
      result,
      CRFFN_PERFORMANCE_CONFIG
        .PAYMENT_TTL_SECONDS
    );
  }

  if (result) {
    result.performanceCache =
      'MISS';
  }

  return result;
}


/**
 * Write wrapper: clear visible applicant data after a
 * successful payment-proof submission/replacement.
 */
function submitApplicantPaymentProof(
  applicationId,
  secureToken,
  paymentReference,
  filePayload
) {
  const result =
    submitApplicantPaymentProofUncached_(
      applicationId,
      secureToken,
      paymentReference,
      filePayload
    );

  if (
    result &&
    result.ok === true
  ) {
    invalidateApplicantPerformanceCache_(
      applicationId,
      secureToken
    );

    invalidateApplicantCacheByApplicationId_(
      applicationId
    );
  }

  return result;
}


function getApplicantPaymentDataUncached_(
  applicationId,
  secureToken
) {
  const applicant = findApplicantRow_(
    applicationId,
    secureToken
  );

  if (!applicant.ok) {
    return applicant;
  }

  const rowObject = applicant.rowObject;

  const informationStatus =
    String(
      rowObject[
        'Application Information Status'
      ] || ''
    )
      .trim()
      .toLowerCase();

  const documentStatus =
    String(
      rowObject[
        'Document Status'
      ] || ''
    )
      .trim()
      .toLowerCase();

  const paymentStatus =
    String(
      rowObject[
        'Payment Status'
      ] ||
      'Not Available'
    ).trim();

  const paymentAllowed =
    informationStatus ===
      'confirmed' &&
    documentStatus ===
      'complete' &&
    paymentStatus
      .toLowerCase() !==
      'not available';

  return {
    ok: true,

    amount: PAYMENT_CONFIG.AMOUNT,
    provider: PAYMENT_CONFIG.PROVIDER,
    paymentUrl:
      PAYMENT_CONFIG.PAYMENT_URL,
    paymentAllowed:
      paymentAllowed,

    informationStatus:
      rowObject[
        'Application Information Status'
      ] || 'Pending',

    documentStatus:
      rowObject[
        'Document Status'
      ] || 'Not Submitted',

    paymentStatus:
      paymentStatus,

    paymentReference:
      rowObject['Payment Reference'] || '',

    receiptPdfUrl:
      safePortalDocumentUrl_(
        rowObject['Receipt PDF URL']
      ),

    uploadedAt:
      rowObject['Payment Proof Uploaded At'] || '',

    deletionStatus:
      rowObject[
        'Payment Proof Deletion Status'
      ] || '',
  };
}

/**
 * Accepts an applicant payment reference and proof.
 *
 * filePayload structure:
 * {
 *   fileName: "receipt.pdf",
 *   mimeType: "application/pdf",
 *   base64Data: "BASE64_CONTENT"
 * }
 */
function submitApplicantPaymentProofUncached_(
  applicationId,
  secureToken,
  paymentReference,
  filePayload
) {
  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const cleanReference = String(
      paymentReference || ''
    ).trim();

    if (!cleanReference) {
      return {
        ok: false,
        errorCode: 'MISSING_PAYMENT_REFERENCE',
        message:
          'Please enter your Revop payment reference.',
      };
    }

    if (cleanReference.length < 6) {
      return {
        ok: false,
        errorCode: 'INVALID_PAYMENT_REFERENCE',
        message:
          'The payment reference entered is too short.',
      };
    }

    const applicant = findApplicantRow_(
      applicationId,
      secureToken
    );

    if (!applicant.ok) {
      return applicant;
    }

    const sheet = applicant.sheet;
    const rowNumber = applicant.rowNumber;
    const headerMap = applicant.headerMap;
    const rowObject = applicant.rowObject;

    const informationStatus =
      String(
        rowObject[
          'Application Information Status'
        ] || ''
      )
        .trim()
        .toLowerCase();

    const documentStatus =
      String(
        rowObject[
          'Document Status'
        ] || ''
      )
        .trim()
        .toLowerCase();

    const currentPaymentStatus =
      String(
        rowObject[
          'Payment Status'
        ] || ''
      )
        .trim()
        .toLowerCase();

    if (
      informationStatus !==
        'confirmed' ||
      documentStatus !==
        'complete' ||
      currentPaymentStatus ===
        'not available' ||
      !currentPaymentStatus
    ) {
      return {
        ok: false,
        errorCode:
          'PAYMENT_NOT_AVAILABLE',
        message:
          'Payment is not yet available. CRFFN will notify you after your Application Information and Supporting Documents are approved.'
      };
    }

    const existingStatus = String(
      rowObject['Payment Status'] || ''
    ).trim();

    if (existingStatus === 'Confirmed') {
      return {
        ok: false,
        errorCode: 'PAYMENT_ALREADY_CONFIRMED',
        message:
          'This payment has already been confirmed.',
      };
    }

    const duplicatePayment =
      findDuplicatePaymentReference_(
        sheet,
        headerMap,
        cleanReference,
        applicationId
      );

    if (duplicatePayment) {
      writeRowValues_(
        sheet,
        rowNumber,
        headerMap,
        {
          'Payment Reference': cleanReference,
          'Payment Status': 'Duplicate Reference',
          'Review Note':
            'Payment reference matches another application.',
        }
      );

      SpreadsheetApp.flush();

      return {
        ok: false,
        errorCode: 'DUPLICATE_REFERENCE',
        message:
          'This payment reference is already attached to another application.',
      };
    }

    const validatedFile =
      validatePaymentProofFile_(filePayload);

    if (!validatedFile.ok) {
      return validatedFile;
    }

    const oldFileId = String(
      rowObject['Payment Proof File ID'] || ''
    ).trim();

    const uploadedFile =
      savePaymentProofFile_(
        applicationId,
        cleanReference,
        validatedFile
      );

    const uploadedAt = new Date();

    writeRowValues_(
      sheet,
      rowNumber,
      headerMap,
      {
        'Payment Reference': cleanReference,
        'Payment Status': 'Pending Verification',

        'Receipt PDF URL':
          uploadedFile.fileUrl,

        'Payment Proof File ID':
          uploadedFile.fileId,

        'Payment Proof Uploaded At':
          uploadedAt,

        'Payment Verified At': '',
        'Payment Verified By': '',
        'Payment Proof Delete After': '',

        'Payment Proof Deletion Status':
          'Retained Pending Verification',

        'Payment Proof Deleted At': '',

        'Review Note':
          'Payment proof submitted and awaiting verification.',
      }
    );

    SpreadsheetApp.flush();

    /*
     * Remove the previous proof only after the new proof
     * has been saved successfully.
     */
    if (
      oldFileId &&
      oldFileId !== uploadedFile.fileId
    ) {
      safelyTrashDriveFile_(oldFileId);
    }

    return {
      ok: true,
      message:
        'Your payment proof has been submitted successfully and is awaiting verification.',

      paymentStatus: 'Pending Verification',
      paymentReference: cleanReference,
      uploadedAt: uploadedAt.toISOString(),
    };
  } catch (error) {
    console.error(error);

    return {
      ok: false,
      errorCode: 'PAYMENT_SUBMISSION_FAILED',
      message:
        'The payment proof could not be submitted. Please try again.',
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Finds an applicant by Application ID and secure token.
 */
function findApplicantRow_(
  applicationId,
  secureToken
) {
  const cleanApplicationId = String(
    applicationId || ''
  ).trim();

  const cleanSecureToken = String(
    secureToken || ''
  ).trim();

  if (!cleanApplicationId || !cleanSecureToken) {
    return {
      ok: false,
      errorCode: 'MISSING_CREDENTIALS',
      message:
        'The applicant portal link is incomplete.',
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getResponseSheet_(ss);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    return {
      ok: false,
      errorCode: 'APPLICATION_NOT_FOUND',
      message:
        'The application could not be found.',
    };
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value || '').trim());

  const headerMap = getHeaderMap_(headers);

  const requiredColumns = [
    'Application ID',
    'Secure Token',
    'Payment Status',
    'Payment Reference',
    'Receipt PDF URL',
    'Payment Proof File ID',
    'Payment Proof Uploaded At',
    'Payment Verified At',
    'Payment Verified By',
    'Payment Proof Delete After',
    'Payment Proof Deletion Status',
    'Payment Proof Deleted At',
  ];

  const missingColumns = requiredColumns.filter(
    columnName => !headerMap[columnName]
  );

  if (missingColumns.length > 0) {
    throw new Error(
      'Missing payment columns: ' +
      missingColumns.join(', ')
    );
  }

  const values = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    )
    .getDisplayValues();

  const applicationIdIndex =
    headerMap['Application ID'] - 1;

  const secureTokenIndex =
    headerMap['Secure Token'] - 1;

  let matchingIndex =
    -1;

  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        values[index][
          applicationIdIndex
        ] || ''
      ).trim() ===
        cleanApplicationId &&
      String(
        values[index][
          secureTokenIndex
        ] || ''
      ).trim() ===
        cleanSecureToken
    ) {
      matchingIndex =
        index;

      break;
    }
  }

  if (matchingIndex === -1) {
    return {
      ok: false,
      errorCode: 'INVALID_ACCESS',
      message:
        'The applicant portal credentials are invalid.',
    };
  }

  const rowNumber = matchingIndex + 2;
  const matchingRow = values[matchingIndex];

  return {
    ok: true,
    sheet: sheet,
    rowNumber: rowNumber,
    headers: headers,
    headerMap: headerMap,
    rowObject: rowToObject_(
      headers,
      matchingRow
    ),
  };
}

/**
 * Validates an uploaded payment-proof file.
 */
function validatePaymentProofFile_(filePayload) {
  if (
    !filePayload ||
    !filePayload.fileName ||
    !filePayload.mimeType ||
    !filePayload.base64Data
  ) {
    return {
      ok: false,
      errorCode: 'MISSING_PAYMENT_PROOF',
      message:
        'Please select a payment-proof file.',
    };
  }

  const fileName = String(
    filePayload.fileName
  ).trim();

  const mimeType = String(
    filePayload.mimeType
  ).trim();

  if (
    !PAYMENT_CONFIG.ALLOWED_MIME_TYPES.includes(
      mimeType
    )
  ) {
    return {
      ok: false,
      errorCode: 'INVALID_FILE_TYPE',
      message:
        'Only PDF, JPG and PNG files are allowed.',
    };
  }

  let decodedBytes;

  try {
    decodedBytes = Utilities.base64Decode(
      String(filePayload.base64Data)
    );
  } catch (error) {
    return {
      ok: false,
      errorCode: 'INVALID_FILE_DATA',
      message:
        'The selected file could not be processed.',
    };
  }

  if (
    decodedBytes.length >
    PAYMENT_CONFIG.MAX_FILE_SIZE_BYTES
  ) {
    return {
      ok: false,
      errorCode: 'FILE_TOO_LARGE',
      message:
        'The payment proof must not exceed 5 MB.',
    };
  }

  if (decodedBytes.length === 0) {
    return {
      ok: false,
      errorCode: 'EMPTY_FILE',
      message:
        'The selected file is empty.',
    };
  }

  return {
    ok: true,
    fileName: fileName,
    mimeType: mimeType,
    bytes: decodedBytes,
  };
}

/**
 * Saves one payment-proof file in Google Drive.
 */
function savePaymentProofFile_(
  applicationId,
  paymentReference,
  validatedFile
) {
  const folder = getPaymentProofFolder_();

  const extension =
    getFileExtensionForMimeType_(
      validatedFile.mimeType
    );

  const safeApplicationId =
    sanitizeFileNamePart_(applicationId);

  const safeReference =
    sanitizeFileNamePart_(paymentReference);

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );

  const fileName = [
    safeApplicationId,
    safeReference,
    timestamp,
  ].join('_') + extension;

  const blob = Utilities.newBlob(
    validatedFile.bytes,
    validatedFile.mimeType,
    fileName
  );

  const file = folder.createFile(blob);

  file.setDescription(
    [
      'CRFFN licensing payment proof',
      'Application ID: ' + applicationId,
      'Payment reference: ' + paymentReference,
    ].join('\n')
  );

  /*
   * The file remains private.
   * Do not call setSharing with public access.
   */
  return {
    fileId: file.getId(),
    fileUrl: file.getUrl(),
  };
}

/**
 * Returns or creates the restricted payment-proof folder.
 */
function getPaymentProofFolder_() {
  const docProps =
    PropertiesService.getDocumentProperties();

  const propertyKey =
    'PAYMENT_PROOF_FOLDER_ID';

  const storedFolderId =
    docProps.getProperty(propertyKey);

  if (storedFolderId) {
    try {
      return DriveApp.getFolderById(
        storedFolderId
      );
    } catch (error) {
      docProps.deleteProperty(propertyKey);
    }
  }

  const existingFolders =
    DriveApp.getFoldersByName(
      PAYMENT_CONFIG.PAYMENT_PROOF_FOLDER_NAME
    );

  let folder;

  if (existingFolders.hasNext()) {
    folder = existingFolders.next();
  } else {
    folder = DriveApp.createFolder(
      PAYMENT_CONFIG.PAYMENT_PROOF_FOLDER_NAME
    );
  }

  docProps.setProperty(
    propertyKey,
    folder.getId()
  );

  return folder;
}

/**
 * Checks whether another application already uses
 * the submitted payment reference.
 */
function findDuplicatePaymentReference_(
  sheet,
  headerMap,
  paymentReference,
  currentApplicationId
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return false;
  }

  const applicationIdColumn =
    headerMap['Application ID'];

  const paymentReferenceColumn =
    headerMap['Payment Reference'];

  const numberOfRows = lastRow - 1;

  const applicationIds = sheet
    .getRange(
      2,
      applicationIdColumn,
      numberOfRows,
      1
    )
    .getDisplayValues();

  const paymentReferences = sheet
    .getRange(
      2,
      paymentReferenceColumn,
      numberOfRows,
      1
    )
    .getDisplayValues();

  const normalizedReference = String(
    paymentReference
  )
    .trim()
    .toLowerCase();

  for (
    let index = 0;
    index < numberOfRows;
    index++
  ) {
    const storedApplicationId = String(
      applicationIds[index][0] || ''
    ).trim();

    const storedReference = String(
      paymentReferences[index][0] || ''
    )
      .trim()
      .toLowerCase();

    if (
      storedReference === normalizedReference &&
      storedApplicationId !== currentApplicationId
    ) {
      return true;
    }
  }

  return false;
}

/**
 * Confirms an applicant payment.
 *
 * Run this from an admin function or future admin dashboard.
 */
function confirmApplicantPayment(
  applicationId,
  verificationNote
) {
  return updateApplicantPaymentVerification_(
    applicationId,
    'Confirmed',
    verificationNote
  );
}

/**
 * Rejects an applicant payment.
 */
function rejectApplicantPayment(
  applicationId,
  verificationNote
) {
  return updateApplicantPaymentVerification_(
    applicationId,
    'Rejected',
    verificationNote
  );
}

/**
 * Updates payment verification and retention information.
 */
function updateApplicantPaymentVerification_(
  applicationId,
  newStatus,
  verificationNote
) {
  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const cleanApplicationId = String(
      applicationId || ''
    ).trim();

    if (!cleanApplicationId) {
      throw new Error(
        'Application ID is required.'
      );
    }

    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const sheet = getResponseSheet_(ss);

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(value =>
        String(value || '').trim()
      );

    const headerMap = getHeaderMap_(headers);

    const applicationIds = sheet
      .getRange(
        2,
        headerMap['Application ID'],
        lastRow - 1,
        1
      )
      .getDisplayValues();

    const matchingIndex =
      applicationIds.findIndex(row =>
        String(row[0] || '').trim() ===
        cleanApplicationId
      );

    if (matchingIndex === -1) {
      throw new Error(
        'Application not found: ' +
        cleanApplicationId
      );
    }

    const rowNumber = matchingIndex + 2;
    const verifiedAt = new Date();

    let deleteAfter = '';
    let deletionStatus =
      'Retained Pending Resolution';

    if (newStatus === 'Confirmed') {
      deleteAfter = new Date(
        verifiedAt.getTime() +
        PAYMENT_CONFIG
          .RETENTION_DAYS_AFTER_CONFIRMATION *
          24 *
          60 *
          60 *
          1000
      );

      deletionStatus =
        'Scheduled for Deletion';
    }

    const verifiedBy =
      Session.getEffectiveUser().getEmail() ||
      'Administrator';

    writeRowValues_(
      sheet,
      rowNumber,
      headerMap,
      {
        'Payment Status': newStatus,
        'Payment Verified At': verifiedAt,
        'Payment Verified By': verifiedBy,
        'Payment Proof Delete After':
          deleteAfter,
        'Payment Proof Deletion Status':
          deletionStatus,
        'Review Note':
          String(verificationNote || '').trim(),
      }
    );

    SpreadsheetApp.flush();

    return {
      ok: true,
      applicationId: cleanApplicationId,
      paymentStatus: newStatus,
      verifiedAt: verifiedAt.toISOString(),
      deleteAfter: deleteAfter
        ? deleteAfter.toISOString()
        : '',
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes confirmed payment proofs whose
 * retention period has expired.
 *
 * Intended to run from a daily time trigger.
 */
function cleanupExpiredPaymentProofs() {
  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const ss =
      SpreadsheetApp.getActiveSpreadsheet();

    const sheet = getResponseSheet_(ss);

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2) {
      return {
        ok: true,
        deletedCount: 0,
      };
    }

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(value =>
        String(value || '').trim()
      );

    const headerMap = getHeaderMap_(headers);

    const values = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();

    const now = new Date();

    const paymentStatusIndex =
      headerMap['Payment Status'] - 1;

    const fileIdIndex =
      headerMap['Payment Proof File ID'] - 1;

    const deleteAfterIndex =
      headerMap[
        'Payment Proof Delete After'
      ] - 1;

    const deletionStatusIndex =
      headerMap[
        'Payment Proof Deletion Status'
      ] - 1;

    const receiptUrlIndex =
      headerMap['Receipt PDF URL'] - 1;

    const deletedAtIndex =
      headerMap[
        'Payment Proof Deleted At'
      ] - 1;

    let deletedCount = 0;

    values.forEach(row => {
      const paymentStatus = String(
        row[paymentStatusIndex] || ''
      ).trim();

      const fileId = String(
        row[fileIdIndex] || ''
      ).trim();

      const deletionStatus = String(
        row[deletionStatusIndex] || ''
      ).trim();

      const deleteAfterValue =
        row[deleteAfterIndex];

      if (
        paymentStatus !== 'Confirmed' ||
        !fileId ||
        deletionStatus === 'Deleted'
      ) {
        return;
      }

      const deleteAfter =
        deleteAfterValue instanceof Date
          ? deleteAfterValue
          : new Date(deleteAfterValue);

      if (
        isNaN(deleteAfter.getTime()) ||
        deleteAfter > now
      ) {
        return;
      }

      const deleted =
        safelyTrashDriveFile_(fileId);

      if (!deleted) {
        return;
      }

      row[receiptUrlIndex] =
        'Deleted after retention period';

      row[fileIdIndex] = '';

      row[deletionStatusIndex] =
        'Deleted';

      row[deletedAtIndex] = now;

      deletedCount += 1;
    });

    sheet
      .getRange(
        2,
        1,
        values.length,
        lastColumn
      )
      .setValues(values);

    SpreadsheetApp.flush();

    console.log(
      'Deleted ' +
      deletedCount +
      ' expired payment proof(s).'
    );

    return {
      ok: true,
      deletedCount: deletedCount,
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Installs the daily payment-proof cleanup trigger.
 */
function installPaymentProofCleanupTrigger() {
  ScriptApp
    .getProjectTriggers()
    .filter(trigger =>
      trigger.getHandlerFunction() ===
      'cleanupExpiredPaymentProofs'
    )
    .forEach(trigger =>
      ScriptApp.deleteTrigger(trigger)
    );

  ScriptApp
    .newTrigger(
      'cleanupExpiredPaymentProofs'
    )
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  SpreadsheetApp.getUi().alert(
    'Daily payment-proof cleanup trigger installed.'
  );
}

/**
 * Safely moves a Drive file to Trash.
 */
function safelyTrashDriveFile_(fileId) {
  try {
    DriveApp
      .getFileById(fileId)
      .setTrashed(true);

    return true;
  } catch (error) {
    console.error(
      'Could not delete Drive file ' +
      fileId +
      ': ' +
      error.message
    );

    return false;
  }
}

/**
 * Returns a safe extension for the uploaded file.
 */
function getFileExtensionForMimeType_(
  mimeType
) {
  const extensions = {
    'application/pdf': '.pdf',
    'image/jpeg': '.jpg',
    'image/png': '.png',
  };

  return extensions[mimeType] || '';
}

/**
 * Removes unsafe characters from file-name parts.
 */
function sanitizeFileNamePart_(value) {
  return String(value || '')
    .trim()
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .slice(0, 80);
}

// temp stuff
function setupPaymentProofFolder() {
  const folder = getPaymentProofFolder_();

  console.log(
    'Payment proof folder created or found: ' +
    folder.getName()
  );

  console.log(
    'Folder URL: ' +
    folder.getUrl()
  );

  return {
    folderName: folder.getName(),
    folderUrl: folder.getUrl(),
  };
}
