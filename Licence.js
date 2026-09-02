/**
 * ============================================================
 * CRFFN LICENCE GENERATION
 * ============================================================
 *
 * This file works with the existing AdminPortal.gs.
 *
 * Do not add another approveAndGenerateLicence() function here.
 * That function already exists in AdminPortal.gs.
 */

const LICENCE_CONFIG = Object.freeze({
  TEMPLATE_DOCUMENT_ID:
    '1Mm236B2JBl-kV8MWHQ_8l_z8E8gSZD9Mi4xgeHv0V9U',

  DRAFT_LICENCES_FOLDER_ID:
    '1Li11I9uJlClhpiSEioOW88wnXolUYxvV',

  LICENCE_PREFIX:
    'CRFFN/FLSP',

  DOCUMENT_REFERENCE_PREFIX:
    'CRFFN/FFL',

  LICENCE_SEQUENCE_LENGTH:
    6,

  DOCUMENT_REFERENCE_SEQUENCE_LENGTH:
    5,

  DEFAULT_AUTHORIZED_SERVICES:
    [
      'Import/Export Logistics',
      'Customs Clearance',
      'Transportation and Haulage',
      'Barge Services',
      'Warehousing',
      'Courier Services',
      'Cold Chain Services',
      'Cargo Consolidation',
      'Last Mile Delivery',
      'Distributorship',
      'Courier and Dispatch Delivery',
      'Container Freight Stations (CFS)',
      'Value Added Services',
      'Supply Chain Management'
    ].join(', ')
});


/**
 * ============================================================
 * MAIN LICENCE GENERATOR
 * ============================================================
 *
 * Called from approveAndGenerateLicence() in AdminPortal.gs:
 *
 * generateLicenceForApplication_({
 *   applicationId,
 *   sheet,
 *   record,
 *   generatedBy
 * });
 */
function generateLicenceForApplication_(options) {
  const config =
    options || {};

  const applicationId =
    String(
      config.applicationId || ''
    ).trim();

  const sheet =
    config.sheet;

  const record =
    config.record;

  const generatedBy =
    String(
      config.generatedBy || ''
    ).trim();

  if (!applicationId) {
    throw new Error(
      'Application ID is required for licence generation.'
    );
  }

  if (
    !sheet ||
    typeof sheet.getLastRow !==
    'function'
  ) {
    throw new Error(
      'A valid application sheet was not supplied to the licence generator.'
    );
  }

  if (
    !record ||
    !record.headerMap ||
    !Array.isArray(
      record.rowValues
    ) ||
    !record.rowNumber
  ) {
    throw new Error(
      'A valid application record was not supplied to the licence generator.'
    );
  }

  /*
   * Prevent accidental duplicate generation.
   */
  const existingPdfUrl =
    String(
      getAdminRecordValue_(
        record,
        'Licence PDF URL'
      ) || ''
    ).trim();

  const existingDocumentUrl =
    String(
      getAdminRecordValue_(
        record,
        'Licence Document URL'
      ) || ''
    ).trim();

  const existingLicenceNumber =
    String(
      getAdminRecordValue_(
        record,
        'Licence Number'
      ) || ''
    ).trim();

  if (
    existingPdfUrl &&
    existingLicenceNumber
  ) {
    return {
      ok: true,

      alreadyGenerated:
        true,

      licenceNumber:
        existingLicenceNumber,

      documentReference:
        getAdminRecordValue_(
          record,
          'Document Reference'
        ),

      licenceDocumentUrl:
        existingDocumentUrl,

      documentUrl:
        existingDocumentUrl,

      licencePdfUrl:
        existingPdfUrl,

      pdfUrl:
        existingPdfUrl,

      message:
        'A draft licence has already been generated for this application.'
    };
  }

  const now =
    new Date();

  const year =
    Number(
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        'yyyy'
      )
    );

  const issueDate =
    now;

  /*
   * Licence expires one year after issue,
   * minus one day.
   */
  const validTill =
    new Date(
      issueDate.getFullYear() + 1,
      issueDate.getMonth(),
      issueDate.getDate() - 1
    );

  const licenceNumber =
    generateLicenceNumber_(
      applicationId,
      issueDate
    );

  const documentReference =
    generateDocumentReference_(
      sheet,
      year
    );

  const folder =
    getDraftLicencesFolder_();

  const templateFile =
    getLicenceTemplateFile_();

  const companyName =
    getLicenceRecordValue_(
      record,
      [
        'Company Name',
        'Registered Company Name',
        'Business Name',
        'Organisation Name',
        'Organization Name',
        'Applicant Name'
      ]
    ) ||
    applicationId;

  const safeCompanyName =
    sanitizeLicenceFileName_(
      companyName
    );

  const safeApplicationId =
    sanitizeLicenceFileName_(
      applicationId
    );

  const documentName =
    [
      'DRAFT Licence',
      safeCompanyName,
      safeApplicationId
    ].join(' - ');

  let copiedDocumentFile =
    null;

  let pdfFile =
    null;

  try {
    /*
     * Copy the licence template into the
     * Draft Licences folder.
     */
    copiedDocumentFile =
      templateFile.makeCopy(
        documentName,
        folder
      );

    const document =
      DocumentApp.openById(
        copiedDocumentFile.getId()
      );

    const replacementData =
      buildLicenceReplacementData_(
        record,
        {
          applicationId:
            applicationId,

          licenceNumber:
            licenceNumber,

          documentReference:
            documentReference,

          issueDate:
            issueDate,

          validTill:
            validTill
        }
      );

    replaceLicencePlaceholders_(
      document,
      replacementData
    );

    document.saveAndClose();

    /*
     * saveAndClose() is synchronous, so the normal path can
     * export immediately. If Drive is briefly still finalising
     * the copy, retry once after a short delay instead of always
     * waiting a full second.
     */
    pdfFile =
      createLicencePdfWithRetry_(
        copiedDocumentFile,
        folder,
        documentName
      );

    /*
     * Save all licence fields in one authoritative row commit
     * when the optimized admin helpers are available. This avoids
     * seven separate spreadsheet writes and an explicit flush.
     * The fallback preserves compatibility with older installs.
     */
    const generatedByValue =
      generatedBy ||
      getCurrentLicenceUserEmail_();

    const licenceUpdates = {
      'Licence Status':
        'Generated',
      'Licence Number':
        licenceNumber,
      'Document Reference':
        documentReference,
      'Licence Document URL':
        copiedDocumentFile.getUrl(),
      'Licence PDF URL':
        pdfFile.getUrl(),
      'Licence Generated At':
        now,
      'Licence Generated By':
        generatedByValue
    };

    if (
      typeof setAdminRecordMemory_ ===
        'function' &&
      typeof commitAdminRecord_ ===
        'function'
    ) {
      Object.keys(
        licenceUpdates
      ).forEach(
        function(header) {
          setAdminRecordMemory_(
            record,
            header,
            licenceUpdates[header]
          );
        }
      );

      commitAdminRecord_(
        sheet,
        record
      );
    } else {
      Object.keys(
        licenceUpdates
      ).forEach(
        function(header) {
          setAdminRecordValue_(
            sheet,
            record,
            header,
            licenceUpdates[header]
          );
        }
      );
    }

    return {
      ok: true,

      alreadyGenerated:
        false,

      licenceStatus:
        'Generated',

      licenceNumber:
        licenceNumber,

      documentReference:
        documentReference,

      licenceDocumentUrl:
        copiedDocumentFile.getUrl(),

      documentUrl:
        copiedDocumentFile.getUrl(),

      licencePdfUrl:
        pdfFile.getUrl(),

      pdfUrl:
        pdfFile.getUrl(),

      message:
        'Unstamped licence generated successfully.'
    };

  } catch (error) {
    /*
     * Remove incomplete files when generation fails.
     */
    if (pdfFile) {
      try {
        pdfFile.setTrashed(
          true
        );
      } catch (cleanupError) {
        console.error(
          'Could not remove incomplete PDF:',
          cleanupError
        );
      }
    }

    if (copiedDocumentFile) {
      try {
        copiedDocumentFile.setTrashed(
          true
        );
      } catch (cleanupError) {
        console.error(
          'Could not remove incomplete Google Doc:',
          cleanupError
        );
      }
    }

    console.error(
      'Licence generation failed:',
      error
    );

    throw new Error(
      error && error.message
        ? error.message
        : 'The licence could not be generated.'
    );
  }
}


/**
 * ============================================================
 * PLACEHOLDER DATA
 * ============================================================
 */
function buildLicenceReplacementData_(
  record,
  licenceData
) {
  return {
    '{{APPLICATION_ID}}':
      licenceData.applicationId,

    '{{DOCUMENT_REFERENCE}}':
      licenceData.documentReference,

    '{{LICENCE_NUMBER}}':
      licenceData.licenceNumber,

    '{{COMPANY_NAME}}':
      getLicenceRecordValue_(
        record,
        [
          'Company Name',
          'Registered Company Name',
          'Business Name',
          'Organisation Name',
          'Organization Name',
          'Applicant Name'
        ]
      ),

    '{{COMPANY_RC_NUMBER}}':
      getLicenceRecordValue_(
        record,
        [
          'Company RC Number',
          'RC Number',
          'CAC Registration Number',
          'CAC Number'
        ]
      ),

    '{{TIN}}':
      getLicenceRecordValue_(
        record,
        [
          'TIN',
          'Tax Identification Number',
          'Company TIN'
        ]
      ),

    '{{CRFFN_MEMBERSHIP_NUMBER}}':
      getLicenceRecordValue_(
        record,
        [
          'CRFFN Membership Number',

          // Compatibility with existing records
          'CRFFN Registration Number',
          'CRFFN Number',
          'CRFFN Registration No',
          'CRFFN Reg Number'
        ]
      ),

    '{{OFFICE_ADDRESS}}':
      getLicenceRecordValue_(
        record,
        [
          'Office Address',
          'Registered Office Address',
          'Company Address',
          'Business Address',
          'Address'
        ]
      ),

    '{{COMPANY_EMAIL}}':
      getLicenceRecordValue_(
        record,
        [
          'Company Email',
          'Company Email Address',
          'Email Address',
          'Email'
        ]
      ),

    '{{COMPANY_PHONE}}':
      getLicenceRecordValue_(
        record,
        [
          'Company Phone',
          'Company Phone Number',
          'Phone Number',
          'Phone',
          'Telephone Number'
        ]
      ),

    '{{AUTHORIZED_SERVICES}}':
      getLicenceRecordValue_(
        record,
        [
          'Authorized Services',
          'Authorised Services',
          'Services',
          'Service Category',
          'Service Categories',
          'Licence Category'
        ]
      ) ||
      LICENCE_CONFIG
        .DEFAULT_AUTHORIZED_SERVICES,

    '{{ISSUE_DATE}}':
      formatLicenceDate_(
        licenceData.issueDate
      ),

    '{{VALID_TILL}}':
      formatLicenceDate_(
        licenceData.validTill
      )
  };
}


/**
 * Gets the first non-empty value from a list
 * of possible spreadsheet headers.
 */
function getLicenceRecordValue_(
  record,
  possibleHeaders
) {
  for (
    let index = 0;
    index < possibleHeaders.length;
    index++
  ) {
    const value =
      getAdminRecordValue_(
        record,
        possibleHeaders[index]
      );

    if (
      value !== null &&
      value !== undefined &&
      String(value).trim()
    ) {
      return String(value).trim();
    }
  }

  return '';
}


/**
 * ============================================================
 * PLACEHOLDER REPLACEMENT
 * ============================================================
 */
function replaceLicencePlaceholders_(
  document,
  replacements
) {
  const sections = [];

  const body =
    document.getBody();

  if (body) {
    sections.push(
      body
    );
  }

  const header =
    document.getHeader();

  if (header) {
    sections.push(
      header
    );
  }

  const footer =
    document.getFooter();

  if (footer) {
    sections.push(
      footer
    );
  }

  Object.keys(
    replacements
  ).forEach(function (placeholder) {
    const replacementValue =
      String(
        replacements[placeholder] || ''
      );

    const escapedPlaceholder =
      escapeRegexForDocument_(
        placeholder
      );

    sections.forEach(function (section) {
      section.replaceText(
        escapedPlaceholder,
        replacementValue
      );
    });
  });
}


/**
 * ============================================================
 * PDF GENERATION
 * ============================================================
 */
function createLicencePdfWithRetry_(
  documentFile,
  folder,
  documentName
) {
  try {
    return createLicencePdf_(
      documentFile,
      folder,
      documentName
    );
  } catch (error) {
    Utilities.sleep(250);

    return createLicencePdf_(
      documentFile,
      folder,
      documentName
    );
  }
}


function createLicencePdf_(
  documentFile,
  folder,
  documentName
) {
  const documentBlob =
    documentFile.getBlob();

  const pdfBlob =
    documentBlob
      .getAs(
        MimeType.PDF
      )
      .setName(
        documentName + '.pdf'
      );

  return folder.createFile(
    pdfBlob
  );
}


/**
 * ============================================================
 * LICENCE NUMBER GENERATION
 * ============================================================
 */
function generateLicenceNumber_(
  applicationId,
  issueDate
) {
  const date = issueDate || new Date();

  const year = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy'
  );

  const month = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MM'
  );

  const day = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd'
  );

  const applicationDigits = String(
    applicationId || ''
  ).replace(/\D/g, '');

  if (!applicationDigits) {
    throw new Error(
      'A valid Application ID is required to generate the licence number.'
    );
  }

  const applicationSuffix =
    applicationDigits.slice(-3).padStart(3, '0');

  return [
    LICENCE_CONFIG.LICENCE_PREFIX,
    year,
    month + day + applicationSuffix
  ].join('/');
}


/**
 * ============================================================
 * DOCUMENT REFERENCE GENERATION
 * ============================================================
 */
function generateDocumentReference_(
  sheet,
  year
) {
  const prefix =
    [
      LICENCE_CONFIG
        .DOCUMENT_REFERENCE_PREFIX,
      year
    ].join('/');

  const sequence =
    getNextLicenceSequence_(
      sheet,
      'Document Reference',
      prefix,
      LICENCE_CONFIG
        .DOCUMENT_REFERENCE_SEQUENCE_LENGTH
    );

  return (
    prefix +
    '/' +
    padLicenceSequence_(
      sequence,
      LICENCE_CONFIG
        .DOCUMENT_REFERENCE_SEQUENCE_LENGTH
    )
  );
}


/**
 * Finds the highest existing sequence in the sheet
 * and returns the next number.
 *
 * This avoids restarting from 1 if Script Properties
 * are empty or the script is copied.
 */
function getNextLicenceSequence_(
  sheet,
  headerName,
  prefix,
  sequenceLength
) {
  if (
    !sheet ||
    typeof sheet.getLastRow !==
    'function'
  ) {
    throw new Error(
      'A valid sheet is required to generate licence numbers.'
    );
  }

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastColumn < 1
  ) {
    return 1;
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0]
      .map(function (header) {
        return String(header).trim();
      });

  const columnIndex =
    headers.indexOf(
      headerName
    );

  if (
    columnIndex === -1
  ) {
    throw new Error(
      'Required licence column "' +
      headerName +
      '" was not found. Run setupLicenceColumns first.'
    );
  }

  let highestSequence =
    0;

  if (
    lastRow >= 2
  ) {
    const values =
      sheet
        .getRange(
          2,
          columnIndex + 1,
          lastRow - 1,
          1
        )
        .getDisplayValues();

    const expectedPrefix =
      prefix + '/';

    values.forEach(function (row) {
      const value =
        String(
          row[0] || ''
        ).trim();

      if (
        !value ||
        value.indexOf(
          expectedPrefix
        ) !== 0
      ) {
        return;
      }

      const sequenceText =
        value.substring(
          expectedPrefix.length
        );

      const sequence =
        Number(
          sequenceText
        );

      if (
        Number.isFinite(sequence) &&
        sequence > highestSequence
      ) {
        highestSequence =
          sequence;
      }
    });
  }

  const propertyKey =
    [
      'CRFFN',
      headerName
        .toUpperCase()
        .replace(
          /\s+/g,
          '_'
        ),
      prefix
        .replace(
          /[^A-Z0-9]/gi,
          '_'
        )
    ].join('_');

  const properties =
    PropertiesService
      .getScriptProperties();

  const storedSequence =
    Number(
      properties.getProperty(
        propertyKey
      ) || 0
    );

  const nextSequence =
    Math.max(
      highestSequence,
      storedSequence
    ) + 1;

  properties.setProperty(
    propertyKey,
    String(nextSequence)
  );

  return nextSequence;
}


function padLicenceSequence_(
  value,
  length
) {
  return String(value)
    .padStart(
      length,
      '0'
    );
}


/**
 * ============================================================
 * FILE AND FOLDER HELPERS
 * ============================================================
 */
function getDraftLicencesFolder_() {
  try {
    return DriveApp.getFolderById(
      LICENCE_CONFIG
        .DRAFT_LICENCES_FOLDER_ID
    );
  } catch (error) {
    throw new Error(
      'The Draft Licences folder could not be opened. Check the folder ID and sharing permissions.'
    );
  }
}


function getLicenceTemplateFile_() {
  try {
    const file =
      DriveApp.getFileById(
        LICENCE_CONFIG
          .TEMPLATE_DOCUMENT_ID
      );

    if (
      file.getMimeType() !==
      MimeType.GOOGLE_DOCS
    ) {
      throw new Error(
        'The licence template must be a Google Docs document.'
      );
    }

    return file;

  } catch (error) {
    throw new Error(
      'The licence template could not be opened. Confirm the document ID and sharing permissions.'
    );
  }
}


/**
 * ============================================================
 * GENERAL HELPERS
 * ============================================================
 */
function formatLicenceDate_(date) {
  if (
    !date ||
    Object.prototype.toString.call(date) !==
    '[object Date]' ||
    isNaN(date.getTime())
  ) {
    return '';
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMMM yyyy'
  );
}


function sanitizeLicenceFileName_(value) {
  return String(value || '')
    .replace(
      /[\\/:*?"<>|#%{}[\]]/g,
      '-'
    )
    .replace(
      /\s+/g,
      ' '
    )
    .trim()
    .substring(
      0,
      100
    );
}


function escapeRegexForDocument_(value) {
  return String(value)
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );
}


function getCurrentLicenceUserEmail_() {
  return (
    Session
      .getActiveUser()
      .getEmail() ||
    Session
      .getEffectiveUser()
      .getEmail() ||
    ''
  );
}


/**
 * ============================================================
 * ONE-TIME LICENCE COLUMN SETUP
 * ============================================================
 *
 * Run this function manually once.
 */
function setupLicenceColumns() {
  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'The active spreadsheet could not be found.'
    );
  }

  /*
   * Uses the same helper as your existing
   * application and admin systems.
   */
  const sheet =
    getResponseSheet_(
      spreadsheet
    );

  if (
    !sheet ||
    typeof sheet.getLastRow !==
    'function'
  ) {
    throw new Error(
      'The application response sheet could not be found.'
    );
  }

  const requiredHeaders = [
    'Licence Status',
    'Licence Number',
    'Document Reference',
    'Licence Document URL',
    'Licence PDF URL',
    'Licence Generated At',
    'Licence Generated By',
    'Licence Released At',
    'Licence Released By'
  ];

  const lastColumn =
    sheet.getLastColumn();

  let existingHeaders = [];

  if (
    lastColumn > 0
  ) {
    existingHeaders =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getDisplayValues()[0]
        .map(function (header) {
          return String(header).trim();
        });
  }

  const missingHeaders =
    requiredHeaders.filter(
      function (header) {
        return existingHeaders
          .indexOf(header) === -1;
      }
    );

  if (
    missingHeaders.length === 0
  ) {
    return {
      ok: true,

      sheetName:
        sheet.getName(),

      addedColumns:
        [],

      message:
        'All licence columns already exist.'
    };
  }

  const startColumn =
    lastColumn + 1;

  sheet
    .getRange(
      1,
      startColumn,
      1,
      missingHeaders.length
    )
    .setValues([
      missingHeaders
    ]);

  SpreadsheetApp.flush();

  return {
    ok: true,

    sheetName:
      sheet.getName(),

    addedColumns:
      missingHeaders,

    message:
      missingHeaders.length +
      ' licence columns were created successfully.'
  };
}


/**
 * ============================================================
 * OPTIONAL CONNECTION TEST
 * ============================================================
 *
 * Run this manually to confirm that the folder,
 * template and application sheet are accessible.
 */
function testLicenceConfiguration() {
  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'The active spreadsheet could not be found.'
    );
  }

  const sheet =
    getResponseSheet_(
      spreadsheet
    );

  const folder =
    getDraftLicencesFolder_();

  const template =
    getLicenceTemplateFile_();

  return {
    ok: true,

    sheetName:
      sheet.getName(),

    folderName:
      folder.getName(),

    templateName:
      template.getName(),

    message:
      'Licence configuration is valid.'
  };
}