/**
 * ============================================================
 * SUPPORTING DOCUMENTS WORKFLOW
 * ============================================================
 *
 * Complete backend for applicant supporting documents.
 *
 * Replacement logic:
 * 1. authenticate applicant;
 * 2. find latest active row;
 * 3. read old File ID;
 * 4. upload replacement first;
 * 5. update Sheet to replacement;
 * 6. flush Sheet;
 * 7. only then move old Drive file to Trash.
 */

const SUPPORTING_DOCUMENT_UPLOAD_MAX_BYTES =
  5 * 1024 * 1024;

const SUPPORTING_DOCUMENT_CONFIG =
  Object.freeze({
    cac: Object.freeze({
      type: 'cac',
      label: 'CAC Document',
      aliases: [
        'cac',
        'cac_document',
        'cac document',
        'cac-document'
      ],
      urlHeader: 'CAC Document URL',
      fileIdHeader: 'CAC Document File ID',
      uploadedAtHeader: 'CAC Document Uploaded At',
      reviewStatusHeader: 'CAC Document Review Status',
      reviewNotesHeader: 'CAC Document Review Notes',
      reviewedAtHeader: 'CAC Document Reviewed At',
      reviewedByHeader: 'CAC Document Reviewed By'
    }),

    passport: Object.freeze({
      type: 'passport',
      label: 'Passport Photograph',
      aliases: [
        'passport',
        'passport_photograph',
        'passport photograph',
        'passport_photo',
        'passport-photo'
      ],
      urlHeader: 'Passport Photograph URL',
      fileIdHeader: 'Passport Photograph File ID',
      uploadedAtHeader: 'Passport Photograph Uploaded At',
      reviewStatusHeader: 'Passport Photograph Review Status',
      reviewNotesHeader: 'Passport Photograph Review Notes',
      reviewedAtHeader: 'Passport Photograph Reviewed At',
      reviewedByHeader: 'Passport Photograph Reviewed By'
    }),

    education: Object.freeze({
      type: 'education',
      label: 'Educational Certificates',
      aliases: [
        'education',
        'educational_certificates',
        'educational certificates',
        'certificate',
        'certificates'
      ],
      urlHeader: 'Educational Certificates URL',
      fileIdHeader: 'Educational Certificates File ID',
      uploadedAtHeader: 'Educational Certificates Uploaded At',
      reviewStatusHeader: 'Educational Certificates Review Status',
      reviewNotesHeader: 'Educational Certificates Review Notes',
      reviewedAtHeader: 'Educational Certificates Reviewed At',
      reviewedByHeader: 'Educational Certificates Reviewed By'
    }),

    experience: Object.freeze({
      type: 'experience',
      label: 'Proof of Experience',
      aliases: [
        'experience',
        'proof_of_experience',
        'proof of experience',
        'proof-experience'
      ],
      urlHeader: 'Proof of Experience URL',
      fileIdHeader: 'Proof of Experience File ID',
      uploadedAtHeader: 'Proof of Experience Uploaded At',
      reviewStatusHeader: 'Proof of Experience Review Status',
      reviewNotesHeader: 'Proof of Experience Review Notes',
      reviewedAtHeader: 'Proof of Experience Reviewed At',
      reviewedByHeader: 'Proof of Experience Reviewed By'
    }),

    identification: Object.freeze({
      type: 'identification',
      label: 'Valid Means of Identification',
      aliases: [
        'identification',
        'means_of_identification',
        'means of identification',
        'valid means of identification',
        'id'
      ],
      urlHeader: 'Means of Identification URL',
      fileIdHeader: 'Means of Identification File ID',
      uploadedAtHeader: 'Means of Identification Uploaded At',
      reviewStatusHeader: 'Means of Identification Review Status',
      reviewNotesHeader: 'Means of Identification Review Notes',
      reviewedAtHeader: 'Means of Identification Reviewed At',
      reviewedByHeader: 'Means of Identification Reviewed By'
    })
  });


function getApplicantSupportingDocuments(
  applicationId,
  secureToken
) {
  const cleanId = String(applicationId || '').trim();
  const cleanToken = String(secureToken || '').trim();

  if (!cleanId || !cleanToken) {
    return getApplicantSupportingDocumentsUncached_(
      cleanId,
      cleanToken
    );
  }

  const key =
    crffnVersionedApplicantKey_(
      crffnSupportingDocumentsCacheKey_(
        cleanId,
        cleanToken
      ),
      cleanId
    );

  const cached = crffnGetCachedJson_(key);

  if (cached) {
    cached.performanceCache = 'HIT';
    return cached;
  }

  const result =
    getApplicantSupportingDocumentsUncached_(
      cleanId,
      cleanToken
    );

  if (result && result.ok === true) {
    crffnPutCachedJson_(
      key,
      result,
      CRFFN_PERFORMANCE_CONFIG
        .SUPPORTING_DOCUMENTS_TTL_SECONDS
    );
  }

  if (result) {
    result.performanceCache = 'MISS';
  }

  return result;
}


function getApplicantSupportingDocumentsUncached_(
  applicationId,
  secureToken
) {
  const access =
    findApplicantSupportingDocumentRecord_(
      applicationId,
      secureToken
    );

  const record = access.record;

  const documents =
    Object.keys(
      SUPPORTING_DOCUMENT_CONFIG
    ).map(function(key) {
      const config = SUPPORTING_DOCUMENT_CONFIG[key];

      const fileId = String(
        getSupportingDocumentRecordValue_(
          record,
          config.fileIdHeader
        ) || ''
      ).trim();

      const url = String(
        getSupportingDocumentRecordValue_(
          record,
          config.urlHeader
        ) || ''
      ).trim();

      const uploaded = Boolean(fileId || url);

      const uploadedAt =
        getSupportingDocumentRecordValue_(
          record,
          config.uploadedAtHeader
        );

      const reviewStatus = String(
        getSupportingDocumentRecordValue_(
          record,
          config.reviewStatusHeader
        ) || (
          uploaded
            ? 'Pending Review'
            : 'Not Submitted'
        )
      ).trim();

      const reviewNotes = String(
        getSupportingDocumentRecordValue_(
          record,
          config.reviewNotesHeader
        ) || ''
      ).trim();

      return {
        type: config.type,
        label: config.label,
        uploaded: uploaded,
        uploadedAt:
          formatSupportingDocumentDate_(
            uploadedAt
          ),
        reviewStatus: reviewStatus,
        reviewNotes: reviewNotes,
        canUpload:
          !uploaded ||
          reviewStatus
            .toLowerCase() ===
            'correction required'
      };
    });

  const uploadedCount =
    documents.filter(function(document) {
      return document.uploaded === true;
    }).length;

  const approvedCount =
    documents.filter(function(document) {
      return String(
        document.reviewStatus || ''
      )
        .trim()
        .toLowerCase() === 'approved';
    }).length;

  return {
    ok: true,
    applicationId: String(
      getSupportingDocumentRecordValue_(
        record,
        'Application ID'
      ) || ''
    ).trim(),
    documentStatus: String(
      getSupportingDocumentRecordValue_(
        record,
        'Document Status'
      ) || 'Not Submitted'
    ).trim(),
    reviewNotes: String(
      getSupportingDocumentRecordValue_(
        record,
        'Document Review Notes'
      ) || ''
    ).trim(),
    uploadedCount: uploadedCount,
    approvedCount: approvedCount,
    totalCount: documents.length,
    documents: documents
  };
}


function uploadApplicantSupportingDocument(
  payload
) {
  const input = payload || {};

  const result =
    uploadApplicantSupportingDocumentUncached_(
      input
    );

  if (result && result.ok === true) {
    invalidateApplicantPerformanceCache_(
      input.applicationId,
      input.secureToken
    );

    invalidateApplicantCacheByApplicationId_(
      input.applicationId
    );
  }

  return result;
}


function uploadApplicantSupportingDocumentUncached_(
  payload
) {
  const input = payload || {};

  const applicationId = String(
    input.applicationId || ''
  ).trim();

  const secureToken = String(
    input.secureToken || ''
  ).trim();

  const documentType = String(
    input.documentType || ''
  ).trim();

  const fileName = String(
    input.fileName || ''
  ).trim();

  const mimeType =
    normalizeSupportingDocumentMimeType_(
      input.mimeType,
      fileName
    );

  const base64Data = String(
    input.base64Data || ''
  ).trim();

  if (!applicationId || !secureToken) {
    throw new Error(
      'The Practitioner Portal link is incomplete.'
    );
  }

  if (!documentType) {
    throw new Error(
      'Document type is required.'
    );
  }

  if (!fileName || !base64Data) {
    throw new Error(
      'Select a file to upload.'
    );
  }

  const config =
    resolveSupportingDocumentConfig_(
      documentType
    );

  if (!config) {
    throw new Error(
      'Unsupported supporting document type: ' +
      documentType
    );
  }

  const allowedMimeTypes = [
    'application/pdf',
    'image/jpeg',
    'image/png'
  ];

  if (
    allowedMimeTypes.indexOf(mimeType) === -1
  ) {
    throw new Error(
      'Only PDF, JPG and PNG files are allowed.'
    );
  }

  let bytes;

  try {
    bytes = Utilities.base64Decode(
      base64Data
    );
  } catch (error) {
    throw new Error(
      'The selected file could not be read.'
    );
  }

  if (
    bytes.length >
    SUPPORTING_DOCUMENT_UPLOAD_MAX_BYTES
  ) {
    throw new Error(
      'The file must not exceed 5 MB.'
    );
  }

  if (!bytes.length) {
    throw new Error(
      'The selected file is empty.'
    );
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let newFile = null;
  let sheetUpdateCompleted = false;

  try {
    const access =
      findApplicantSupportingDocumentRecord_(
        applicationId,
        secureToken
      );

    const sheet = access.sheet;
    const record = access.record;

    const licenceStatus = String(
      getSupportingDocumentRecordValue_(
        record,
        'Licence Status'
      ) || ''
    )
      .trim()
      .toLowerCase();

    const blockedLicenceStatuses = [
      'generated',
      'printed',
      'stamped',
      'uploaded',
      'released'
    ];

    if (
      blockedLicenceStatuses.indexOf(
        licenceStatus
      ) !== -1
    ) {
      throw new Error(
        'Supporting documents can no longer be changed because licence processing has already started.'
      );
    }

    const oldFileId = String(
      getSupportingDocumentRecordValue_(
        record,
        config.fileIdHeader
      ) || ''
    ).trim();

    const currentDocumentReviewStatus =
      String(
        getSupportingDocumentRecordValue_(
          record,
          config.reviewStatusHeader
        ) || ''
      )
        .trim()
        .toLowerCase();

    if (
      oldFileId &&
      currentDocumentReviewStatus !==
        'correction required'
    ) {
      throw new Error(
        currentDocumentReviewStatus ===
          'approved'
          ? config.label +
            ' has already been approved and is locked.'
          : config.label +
            ' is already submitted and awaiting review. It can only be replaced after an administrator requests a correction.'
      );
    }

    const applicationFolder =
      getOrCreateApplicantSupportingDocumentFolder_(
        applicationId
      );

    const storedFileName =
      buildSupportingDocumentStorageFileName_(
        applicationId,
        config.label,
        fileName
      );

    const blob = Utilities.newBlob(
      bytes,
      mimeType,
      storedFileName
    );

    newFile = applicationFolder.createFile(
      blob
    );

    const uploadedAt = new Date();

    setSupportingDocumentRecordMemory_(
      record,
      config.urlHeader,
      newFile.getUrl()
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.fileIdHeader,
      newFile.getId()
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.uploadedAtHeader,
      uploadedAt
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.reviewStatusHeader,
      'Pending Review'
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.reviewNotesHeader,
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.reviewedAtHeader,
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      config.reviewedByHeader,
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Document Correction Email Sent At',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Document Correction Email Sent By',
      ''
    );

    const allDocumentsUploaded =
      areAllSupportingDocumentsUploadedAfterChange_(
        record,
        config,
        newFile
      );

    setSupportingDocumentRecordMemory_(
      record,
      'Document Status',
      allDocumentsUploaded
        ? 'Pending Review'
        : 'In Progress'
    );

    const currentPaymentStatus =
      String(
        getSupportingDocumentRecordValue_(
          record,
          'Payment Status'
        ) || ''
      )
        .trim()
        .toLowerCase();

    if (
      currentPaymentStatus ===
        'awaiting payment' ||
      currentPaymentStatus ===
        'not available' ||
      currentPaymentStatus ===
        'not submitted' ||
      !currentPaymentStatus
    ) {
      setSupportingDocumentRecordMemory_(
        record,
        'Payment Status',
        'Not Available'
      );

      setSupportingDocumentRecordMemory_(
        record,
        'Payment Invitation Sent At',
        ''
      );
    }

    setSupportingDocumentRecordMemory_(
      record,
      'Document Review Notes',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Documents Verified At',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Documents Verified By',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Verification Status',
      'Pending'
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Verification Notes',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Verification Completed At',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Verification Completed By',
      ''
    );

    setSupportingDocumentRecordMemory_(
      record,
      'Licence Status',
      'Not Generated'
    );

    commitSupportingDocumentRecord_(
      sheet,
      record
    );

    SpreadsheetApp.flush();
    sheetUpdateCompleted = true;

    let previousFileTrashed = false;

    if (
      oldFileId &&
      oldFileId !== newFile.getId()
    ) {
      try {
        const oldFile =
          DriveApp.getFileById(
            oldFileId
          );

        if (!oldFile.isTrashed()) {
          oldFile.setTrashed(true);
        }

        previousFileTrashed = true;
      } catch (cleanupError) {
        console.warn(
          'Replacement saved, but previous supporting document could not be moved to Trash: ' +
          cleanupError.message
        );
      }
    }

    const uploadedCount =
      Object.keys(
        SUPPORTING_DOCUMENT_CONFIG
      ).filter(function(key) {
        const itemConfig =
          SUPPORTING_DOCUMENT_CONFIG[key];

        return Boolean(
          String(
            getSupportingDocumentRecordValue_(
              record,
              itemConfig.fileIdHeader
            ) ||
            getSupportingDocumentRecordValue_(
              record,
              itemConfig.urlHeader
            ) || ''
          ).trim()
        );
      }).length;

    return {
      ok: true,
      applicationId: applicationId,
      documentType: config.type,
      documentLabel: config.label,
      uploadedAt:
        formatSupportingDocumentDate_(
          uploadedAt
        ),
      replaced: Boolean(oldFileId),
      previousFileTrashed:
        previousFileTrashed,
      documentStatus:
        allDocumentsUploaded
          ? 'Pending Review'
          : 'In Progress',
      reviewStatus: 'Pending Review',
      uploadedCount: uploadedCount,
      totalCount:
        Object.keys(
          SUPPORTING_DOCUMENT_CONFIG
        ).length,
      message:
        oldFileId
          ? config.label +
            ' was replaced successfully.'
          : config.label +
            ' was uploaded successfully.'
    };
  } catch (error) {
    if (
      newFile &&
      !sheetUpdateCompleted
    ) {
      try {
        newFile.setTrashed(true);
      } catch (cleanupError) {
        console.warn(
          'Could not clean up failed new supporting-document upload: ' +
          cleanupError.message
        );
      }
    }

    throw error;
  } finally {
    lock.releaseLock();
  }
}


function findApplicantSupportingDocumentRecord_(
  applicationId,
  secureToken
) {
  const normalizedApplicationId =
    String(applicationId || '').trim();

  const normalizedToken =
    String(secureToken || '').trim();

  if (
    !normalizedApplicationId ||
    !normalizedToken
  ) {
    throw new Error(
      'The Practitioner Portal link is incomplete.'
    );
  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(spreadsheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      'The application could not be found.'
    );
  }

  ensureSupportingDocumentColumns_(
    sheet
  );

  const lastColumn =
    sheet.getLastColumn();

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function(header) {
        return String(
          header || ''
        ).trim();
      });

  const headerMap =
    getHeaderMap_(headers);

  const applicationIdColumn =
    headerMap['Application ID'];

  const secureTokenColumn =
    headerMap['Secure Token'];

  if (
    !applicationIdColumn ||
    !secureTokenColumn
  ) {
    throw new Error(
      'Application ID or Secure Token column is missing.'
    );
  }

  const numberOfRows =
    lastRow - 1;

  /*
   * Read only the two credential columns while locating the
   * newest row. This avoids loading every column for every
   * application merely to authenticate one portal request.
   */
  const applicationIds =
    sheet
      .getRange(
        2,
        applicationIdColumn,
        numberOfRows,
        1
      )
      .getDisplayValues();

  const secureTokens =
    sheet
      .getRange(
        2,
        secureTokenColumn,
        numberOfRows,
        1
      )
      .getDisplayValues();

  let rowNumber = 0;

  for (
    let index = numberOfRows - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        applicationIds[index][0] || ''
      ).trim() ===
        normalizedApplicationId &&
      String(
        secureTokens[index][0] || ''
      ).trim() ===
        normalizedToken
    ) {
      rowNumber = index + 2;
      break;
    }
  }

  if (!rowNumber) {
    throw new Error(
      'This Practitioner Portal link is invalid or no longer matches an application.'
    );
  }

  const rowRange =
    sheet.getRange(
      rowNumber,
      1,
      1,
      lastColumn
    );

  const rowValues =
    rowRange.getValues()[0];

  const rowFormulas =
    rowRange.getFormulas()[0];

  /* Preserve formula cells during the grouped row commit. */
  rowFormulas.forEach(
    function(formula, index) {
      if (formula) {
        rowValues[index] = formula;
      }
    }
  );

  return {
    spreadsheet: spreadsheet,
    sheet: sheet,
    record: {
      rowNumber: rowNumber,
      headers: headers,
      headerMap: headerMap,
      rowValues: rowValues
    }
  };
}


function ensureSupportingDocumentColumns_(
  sheet
) {
  const requiredHeaders = [
    'Document Status',
    'Document Review Notes',
    'Payment Status',
    'Payment Invitation Sent At',
    'Documents Verified At',
    'Documents Verified By',
    'Document Correction Email Sent At',
    'Document Correction Email Sent By',

    'CAC Document URL',
    'CAC Document File ID',
    'CAC Document Uploaded At',
    'CAC Document Review Status',
    'CAC Document Review Notes',
    'CAC Document Reviewed At',
    'CAC Document Reviewed By',

    'Passport Photograph URL',
    'Passport Photograph File ID',
    'Passport Photograph Uploaded At',
    'Passport Photograph Review Status',
    'Passport Photograph Review Notes',
    'Passport Photograph Reviewed At',
    'Passport Photograph Reviewed By',

    'Educational Certificates URL',
    'Educational Certificates File ID',
    'Educational Certificates Uploaded At',
    'Educational Certificates Review Status',
    'Educational Certificates Review Notes',
    'Educational Certificates Reviewed At',
    'Educational Certificates Reviewed By',

    'Proof of Experience URL',
    'Proof of Experience File ID',
    'Proof of Experience Uploaded At',
    'Proof of Experience Review Status',
    'Proof of Experience Review Notes',
    'Proof of Experience Reviewed At',
    'Proof of Experience Reviewed By',

    'Means of Identification URL',
    'Means of Identification File ID',
    'Means of Identification Uploaded At',
    'Means of Identification Review Status',
    'Means of Identification Review Notes',
    'Means of Identification Reviewed At',
    'Means of Identification Reviewed By',

    'Verification Status',
    'Verification Notes',
    'Verification Completed At',
    'Verification Completed By',

    'Licence Status'
  ];

  const lastColumn =
    sheet.getLastColumn();

  const existingHeaders =
    lastColumn > 0
      ? sheet
          .getRange(
            1,
            1,
            1,
            lastColumn
          )
          .getDisplayValues()[0]
          .map(function(header) {
            return String(
              header || ''
            ).trim();
          })
      : [];

  const missingHeaders =
    requiredHeaders.filter(
      function(header) {
        return (
          existingHeaders.indexOf(
            header
          ) === -1
        );
      }
    );

  if (!missingHeaders.length) {
    return;
  }

  sheet
    .getRange(
      1,
      lastColumn + 1,
      1,
      missingHeaders.length
    )
    .setValues([
      missingHeaders
    ])
    .setFontWeight('bold');

  SpreadsheetApp.flush();
}


function getSupportingDocumentRecordValue_(
  record,
  header
) {
  const column =
    record.headerMap[header];

  if (!column) {
    return '';
  }

  return record.rowValues[
    column - 1
  ];
}


function setSupportingDocumentRecordMemory_(
  record,
  header,
  value
) {
  const column =
    record.headerMap[header];

  if (!column) {
    throw new Error(
      'Required supporting-document column is missing: ' +
      header
    );
  }

  record.rowValues[
    column - 1
  ] = value;
}


function commitSupportingDocumentRecord_(
  sheet,
  record
) {
  const width = record.headers.length;

  if (
    record.rowValues.length < width
  ) {
    while (
      record.rowValues.length < width
    ) {
      record.rowValues.push('');
    }
  }

  sheet
    .getRange(
      record.rowNumber,
      1,
      1,
      width
    )
    .setValues([
      record.rowValues.slice(0, width)
    ]);
}


function setSupportingDocumentRecordValue_(
  sheet,
  record,
  header,
  value
) {
  let column =
    record.headerMap[header];

  if (!column) {
    column =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        column
      )
      .setValue(header)
      .setFontWeight('bold');

    record.headerMap[header] =
      column;

    record.headers[
      column - 1
    ] = header;
  }

  sheet
    .getRange(
      record.rowNumber,
      column
    )
    .setValue(value);

  record.rowValues[
    column - 1
  ] = value;
}


function areAllSupportingDocumentsUploadedAfterChange_(
  record,
  changedConfig,
  newFile
) {
  return Object.keys(
    SUPPORTING_DOCUMENT_CONFIG
  ).every(function(key) {
    const config =
      SUPPORTING_DOCUMENT_CONFIG[key];

    if (
      config.type ===
      changedConfig.type
    ) {
      return Boolean(
        newFile &&
        newFile.getId()
      );
    }

    const fileId = String(
      getSupportingDocumentRecordValue_(
        record,
        config.fileIdHeader
      ) || ''
    ).trim();

    const url = String(
      getSupportingDocumentRecordValue_(
        record,
        config.urlHeader
      ) || ''
    ).trim();

    return Boolean(fileId || url);
  });
}


function resolveSupportingDocumentConfig_(
  documentType
) {
  const normalized =
    String(documentType || '')
      .trim()
      .toLowerCase();

  const keys =
    Object.keys(
      SUPPORTING_DOCUMENT_CONFIG
    );

  for (
    let index = 0;
    index < keys.length;
    index++
  ) {
    const config =
      SUPPORTING_DOCUMENT_CONFIG[
        keys[index]
      ];

    if (
      config.type === normalized
    ) {
      return config;
    }

    const aliases =
      config.aliases.map(
        function(alias) {
          return String(alias)
            .trim()
            .toLowerCase();
        }
      );

    if (
      aliases.indexOf(normalized) !==
      -1
    ) {
      return config;
    }
  }

  return null;
}


function getOrCreateApplicantSupportingDocumentFolder_(
  applicationId
) {
  const rootFolder =
    getSupportingDocumentsRootFolder_();

  const folderName =
    String(applicationId || '').trim();

  if (!folderName) {
    throw new Error(
      'Application ID is required before a supporting-document folder can be created.'
    );
  }

  const folders =
    rootFolder.getFoldersByName(
      folderName
    );

  if (folders.hasNext()) {
    return folders.next();
  }

  return rootFolder.createFolder(
    folderName
  );
}


/**
 * Returns an accessible Supporting Documents root folder.
 *
 * Important:
 * - validates saved Document Properties instead of blindly trusting them;
 * - falls back to APP_CONFIG / LICENCE_CONFIG when a saved ID is stale;
 * - repairs SUPPORTING_DOCUMENTS_FOLDER_ID when a valid fallback is found;
 * - gives a clear configuration error instead of exposing Drive's
 *   "No item with the given ID could be found" exception to applicants.
 */
function getSupportingDocumentsRootFolder_() {
  const candidates =
    getSupportingDocumentsFolderCandidates_();

  const attempted = [];

  for (
    let index = 0;
    index < candidates.length;
    index++
  ) {
    const candidate =
      candidates[index];

    if (!candidate.id) {
      continue;
    }

    try {
      const folder =
        DriveApp.getFolderById(
          candidate.id
        );

      // Force an actual access check.
      const name =
        folder.getName();

      if (!name) {
        throw new Error(
          'Folder name could not be read.'
        );
      }

      // Keep Document Properties synchronized with the working folder.
      PropertiesService
        .getDocumentProperties()
        .setProperty(
          'SUPPORTING_DOCUMENTS_FOLDER_ID',
          candidate.id
        );

      return folder;
    } catch (error) {
      attempted.push(
        candidate.source +
        ': ' +
        candidate.id
      );

      console.warn(
        'Supporting Documents folder candidate could not be accessed (' +
        candidate.source +
        '): ' +
        (
          error &&
          error.message
            ? error.message
            : String(error)
        )
      );
    }
  }

  throw new Error(
    [
      'The Supporting Documents Drive folder is not configured correctly or the web-app account cannot access it.',
      '',
      attempted.length
        ? 'Folder IDs checked: ' +
          attempted.join(' | ')
        : 'No Supporting Documents folder ID is currently configured.',
      '',
      'Open the spreadsheet, run setSupportingDocumentsFolderId(), and choose/paste a Google Drive folder that the account running this web app can edit.'
    ].join('\n')
  );
}


/**
 * Backward-compatible helper for code that needs only the folder ID.
 * Returns the first configured candidate without attempting Drive access.
 */
function getSupportingDocumentsRootFolderId_() {
  const candidates =
    getSupportingDocumentsFolderCandidates_();

  return candidates.length
    ? candidates[0].id
    : '';
}


/**
 * Collects possible Supporting Documents folder IDs in priority order.
 * Duplicate IDs are removed.
 */
function getSupportingDocumentsFolderCandidates_() {
  const candidates = [];

  const addCandidate =
    function(source, value) {
      const id =
        extractGoogleDriveFolderId_(
          value
        );

      if (!id) {
        return;
      }

      const exists =
        candidates.some(
          function(item) {
            return item.id === id;
          }
        );

      if (!exists) {
        candidates.push({
          source: source,
          id: id
        });
      }
    };

  addCandidate(
    'Document Properties',
    PropertiesService
      .getDocumentProperties()
      .getProperty(
        'SUPPORTING_DOCUMENTS_FOLDER_ID'
      )
  );

  try {
    if (
      typeof APP_CONFIG !==
        'undefined' &&
      APP_CONFIG
    ) {
      addCandidate(
        'APP_CONFIG',
        APP_CONFIG
          .SUPPORTING_DOCUMENTS_FOLDER_ID
      );
    }
  } catch (error) {}

  try {
    if (
      typeof LICENCE_CONFIG !==
        'undefined' &&
      LICENCE_CONFIG
    ) {
      addCandidate(
        'LICENCE_CONFIG',
        LICENCE_CONFIG
          .SUPPORTING_DOCUMENTS_FOLDER_ID
      );
    }
  } catch (error) {}

  return candidates;
}


/**
 * Accepts either a plain Google Drive folder ID or a folder URL.
 */
function extractGoogleDriveFolderId_(
  value
) {
  const normalized =
    String(value || '').trim();

  if (!normalized) {
    return '';
  }

  // Plain ID.
  if (
    /^[A-Za-z0-9_-]{10,}$/.test(
      normalized
    )
  ) {
    return normalized;
  }

  // Typical Drive folder URL:
  // https://drive.google.com/drive/folders/FOLDER_ID
  const folderMatch =
    normalized.match(
      /\/folders\/([A-Za-z0-9_-]+)/i
    );

  if (folderMatch) {
    return folderMatch[1];
  }

  // Generic ?id=FOLDER_ID fallback.
  const idMatch =
    normalized.match(
      /[?&]id=([A-Za-z0-9_-]+)/i
    );

  return idMatch
    ? idMatch[1]
    : '';
}


function setSupportingDocumentsFolderId() {
  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      'Supporting Documents Folder',
      [
        'Paste the Google Drive folder URL or folder ID that should contain applicant supporting documents.',
        '',
        'The system will create one subfolder per Application ID.',
        '',
        'The Google account running the web app must have edit access to this folder.'
      ].join('\n'),
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const folderId =
    extractGoogleDriveFolderId_(
      response.getResponseText()
    );

  if (!folderId) {
    ui.alert(
      'The value entered does not look like a valid Google Drive folder URL or folder ID.'
    );
    return;
  }

  let folder;

  try {
    folder =
      DriveApp.getFolderById(
        folderId
      );

    // Force access now so a stale/inaccessible ID is rejected
    // before it is saved into configuration.
    folder.getName();
  } catch (error) {
    ui.alert(
      [
        'CRFFN could not access that Google Drive folder.',
        '',
        'Check that the folder exists, is not in Trash, and that this Google account has edit access.',
        '',
        'Drive error: ' +
        (
          error &&
          error.message
            ? error.message
            : String(error)
        )
      ].join('\n')
    );

    return;
  }

  PropertiesService
    .getDocumentProperties()
    .setProperty(
      'SUPPORTING_DOCUMENTS_FOLDER_ID',
      folderId
    );

  ui.alert(
    [
      'Supporting Documents folder saved successfully.',
      '',
      'Folder: ' +
      folder.getName(),
      '',
      'Folder ID: ' +
      folderId
    ].join('\n')
  );
}


/**
 * Run from Apps Script when supporting-document uploads fail.
 *
 * It reports the folder source/ID being used and whether the account
 * executing the script can actually read and write to that folder.
 */
function diagnoseSupportingDocumentsFolder() {
  const candidates =
    getSupportingDocumentsFolderCandidates_();

  const result = {
    currentAccount:
      String(
        Session
          .getEffectiveUser()
          .getEmail() ||
        ''
      ).trim(),

    candidates:
      [],

    resolvedFolder:
      null
  };

  candidates.forEach(
    function(candidate) {
      const check = {
        source:
          candidate.source,
        id:
          candidate.id,
        accessible:
          false,
        folderName:
          '',
        error:
          ''
      };

      try {
        const folder =
          DriveApp.getFolderById(
            candidate.id
          );

        check.folderName =
          folder.getName();

        check.accessible =
          true;
      } catch (error) {
        check.error =
          error &&
          error.message
            ? error.message
            : String(error);
      }

      result.candidates.push(
        check
      );
    }
  );

  try {
    const folder =
      getSupportingDocumentsRootFolder_();

    result.resolvedFolder = {
      id:
        folder.getId(),
      name:
        folder.getName(),
      accessible:
        true
    };
  } catch (error) {
    result.resolvedFolder = {
      accessible:
        false,
      error:
        error &&
        error.message
          ? error.message
          : String(error)
    };
  }

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


function normalizeSupportingDocumentMimeType_(
  mimeType,
  fileName
) {
  let normalized =
    String(mimeType || '')
      .trim()
      .toLowerCase();

  if (normalized) {
    if (normalized === 'image/jpg') {
      return 'image/jpeg';
    }

    return normalized;
  }

  const name =
    String(fileName || '')
      .trim()
      .toLowerCase();

  if (name.endsWith('.pdf')) {
    return 'application/pdf';
  }

  if (
    name.endsWith('.jpg') ||
    name.endsWith('.jpeg')
  ) {
    return 'image/jpeg';
  }

  if (name.endsWith('.png')) {
    return 'image/png';
  }

  return '';
}


function buildSupportingDocumentStorageFileName_(
  applicationId,
  documentLabel,
  originalFileName
) {
  const original =
    String(originalFileName || '')
      .trim();

  const extensionMatch =
    original.match(
      /(\.[A-Za-z0-9]+)$/
    );

  const extension =
    extensionMatch
      ? extensionMatch[1]
          .toLowerCase()
      : '';

  const safeApplicationId =
    sanitizeSupportingDocumentFileName_(
      applicationId
    );

  const safeLabel =
    sanitizeSupportingDocumentFileName_(
      documentLabel
    );

  const timestamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd-HHmmss'
    );

  return (
    safeApplicationId +
    ' - ' +
    safeLabel +
    ' - ' +
    timestamp +
    extension
  );
}


function sanitizeSupportingDocumentFileName_(
  value
) {
  return String(value || '')
    .replace(
      /[\\/:*?"<>|#%{}[\]]/g,
      '-'
    )
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 100);
}


function safeSupportingDocumentUrl_(
  url
) {
  const normalized =
    String(url || '').trim();

  if (
    !/^https?:\/\/[^\s]+$/i.test(
      normalized
    )
  ) {
    return '';
  }

  return normalized;
}


function formatSupportingDocumentDate_(
  value
) {
  if (!value) {
    return '';
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMMM yyyy, hh:mm a'
  );
}


function diagnoseSupportingDocuments_(
  applicationId,
  secureToken
) {
  const result =
    getApplicantSupportingDocuments(
      applicationId,
      secureToken
    );

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}
