/**
 * Applicant Portal web-app router.
 *
 * Example:
 * /exec?view=applicant&ref=APP-0001&token=SECURE_TOKEN
 */
/**
 * Web-app router.
 *
 * Applicant:
 * /exec?view=applicant&ref=APP-0001&token=SECURE_TOKEN
 *
 * Admin:
 * /exec?view=admin
 *
 * Admin application:
 * /exec?view=admin&ref=APP-0001
 */
function doGet(e) {
  const parameters =
    e && e.parameter
      ? e.parameter
      : {};

  const view = String(
    parameters.view || 'applicant'
  )
    .trim()
    .toLowerCase();

  if (view === 'applicant') {
    return renderApplicantPortal_(
      parameters
    );
  }

  if (view === 'admin') {
    return renderAdminPortal_(
      parameters
    );
  }

  if (view === 'admindocument') {
    return renderAdminDocumentViewer_(
      parameters
    );
  }

  if (
    view ===
    'downloadlicence'
  ) {
    return renderLicenceDownload_(
      parameters
    );
  }

  return HtmlService
    .createHtmlOutput(
      [
        '<h2>Page not found</h2>',
        '<p>The requested page does not exist.</p>',
      ].join('')
    )
    .setTitle('Page Not Found');
}

/**
 * Secure administrator-only document viewer.
 *
 * The raw Drive URL is never opened by the administrator.
 * Apps Script validates the administrator session and reads the
 * private Drive file as the web-app owner.
 */
function renderAdminDocumentViewer_(
  parameters
) {
  const input =
    parameters || {};

  const sessionToken =
    String(
      input.session || ''
    ).trim();

  const applicationId =
    String(
      input.applicationId || ''
    ).trim();

  const documentType =
    String(
      input.documentType || ''
    )
      .trim()
      .toLowerCase();

  const action =
    String(
      input.action || 'view'
    )
      .trim()
      .toLowerCase();

  try {
    requireAdminAccess_(
      sessionToken
    );

    const file =
      getAdminPrivateDocumentFile_(
        applicationId,
        documentType,
        sessionToken
      );

    if (file.isTrashed()) {
      throw new Error(
        'The requested file is no longer available.'
      );
    }

    const blob =
      getAdminPrivateDocumentBlob_(
        file,
        documentType
      );

    const fileName =
      getAdminPrivateDocumentFileName_(
        file,
        documentType
      );

    const mimeType =
      blob.getContentType() ||
      file.getMimeType() ||
      'application/octet-stream';

    const base64Data =
      Utilities.base64Encode(
        blob.getBytes()
      );

    if (action === 'download') {
      return renderAdminPrivateDocumentDownload_(
        fileName,
        mimeType,
        base64Data
      );
    }

    if (action !== 'view') {
      throw new Error(
        'This document action is not permitted.'
      );
    }

    const template =
      HtmlService.createTemplateFromFile(
        'AdminDocumentViewer'
      );

    template.documentName =
      fileName;

    template.mimeType =
      mimeType;

    template.base64Data =
      base64Data;

    return template
      .evaluate()
      .setTitle(
        fileName
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  } catch (error) {
    return HtmlService
      .createHtmlOutput(
        [
          '<!doctype html>',
          '<html>',
          '<head>',
          '<meta name="viewport" content="width=device-width,initial-scale=1">',
          '<style>',
          'body{font-family:Arial,sans-serif;background:#f8fafc;padding:30px;color:#1d2939}',
          '.box{max-width:680px;margin:40px auto;background:#fff;border:1px solid #eaecf0;border-radius:12px;padding:24px}',
          '</style>',
          '</head>',
          '<body>',
          '<div class="box">',
          '<h2>Document unavailable</h2>',
          '<p>',
          escapeHtml_(
            error &&
            error.message
              ? error.message
              : 'The document could not be opened.'
          ),
          '</p>',
          '</div>',
          '</body>',
          '</html>'
        ].join('')
      )
      .setTitle(
        'Document unavailable'
      );
  }
}


function getAdminPrivateDocumentFile_(
  applicationId,
  documentType,
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  if (!cleanApplicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  const sheet =
    getResponseSheet_(
      SpreadsheetApp
        .getActiveSpreadsheet()
    );

  const record =
    findAdminApplicationRecord_(
      sheet,
      cleanApplicationId
    );

  if (!record) {
    throw new Error(
      'Application record was not found.'
    );
  }

  const allowedHeaders = {
    cac: [
      'CAC Document URL'
    ],
    passport: [
      'Passport Photograph URL'
    ],
    education: [
      'Educational Certificates URL'
    ],
    experience: [
      'Proof of Experience URL'
    ],
    identification: [
      'Means of Identification URL'
    ],
    payment: [
      'Receipt PDF URL'
    ],
    draftlicence: [
      'Licence PDF URL',
      'License PDF URL'
    ],
    stampedlicence: [
      'Stamped Licence File ID',
      'Stamped License File ID',
      'Stamped Licence PDF URL',
      'Stamped License PDF URL'
    ]
  };

  const headers =
    allowedHeaders[
      documentType
    ];

  if (!headers) {
    throw new Error(
      'This document type is not permitted.'
    );
  }

  let storedValue =
    '';

  for (
    let index = 0;
    index < headers.length;
    index++
  ) {
    const candidate =
      String(
        getAdminRecordValue_(
          record,
          headers[index]
        ) || ''
      ).trim();

    if (candidate) {
      storedValue =
        candidate;
      break;
    }
  }

  const fileId =
    extractPrivateDriveFileId_(
      storedValue
    );

  if (!fileId) {
    throw new Error(
      'The stored document reference is unavailable.'
    );
  }

  return DriveApp.getFileById(
    fileId
  );
}


function extractPrivateDriveFileId_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  if (!text) {
    return '';
  }

  if (
    /^[A-Za-z0-9_-]{20,}$/.test(
      text
    )
  ) {
    return text;
  }

  const patterns = [
    /\/d\/([A-Za-z0-9_-]{20,})/,
    /[?&]id=([A-Za-z0-9_-]{20,})/
  ];

  for (
    let index = 0;
    index < patterns.length;
    index++
  ) {
    const match =
      text.match(
        patterns[index]
      );

    if (
      match &&
      match[1]
    ) {
      return match[1];
    }
  }

  return '';
}


/**
 * Returns the private file as a browser-friendly blob.
 * Licence files are always served as PDF, while existing
 * supporting-document behaviour remains unchanged.
 */
function getAdminPrivateDocumentBlob_(
  file,
  documentType
) {
  const type =
    String(
      documentType || ''
    )
      .trim()
      .toLowerCase();

  if (
    type === 'draftlicence' ||
    type === 'stampedlicence'
  ) {
    const blob =
      file.getBlob();

    if (
      String(
        blob.getContentType() ||
        file.getMimeType() ||
        ''
      ).toLowerCase() ===
      'application/pdf'
    ) {
      return blob;
    }

    return blob.getAs(
      MimeType.PDF
    );
  }

  return file.getBlob();
}


function getAdminPrivateDocumentFileName_(
  file,
  documentType
) {
  const type =
    String(
      documentType || ''
    )
      .trim()
      .toLowerCase();

  let name =
    String(
      file.getName() ||
      'document'
    ).trim();

  if (
    (
      type === 'draftlicence' ||
      type === 'stampedlicence'
    ) &&
    !/\.pdf$/i.test(name)
  ) {
    name += '.pdf';
  }

  return name;
}


/**
 * Apps Script cannot stream arbitrary Drive bytes with a
 * Content-Disposition header. For secure downloads, return a
 * tiny authenticated HTML page that reconstructs the private
 * blob in the browser and immediately starts the download.
 */
function renderAdminPrivateDocumentDownload_(
  fileName,
  mimeType,
  base64Data
) {
  const safeFileName =
    JSON.stringify(
      String(
        fileName ||
        'document'
      )
    );

  const safeMimeType =
    JSON.stringify(
      String(
        mimeType ||
        'application/octet-stream'
      )
    );

  const safeBase64 =
    JSON.stringify(
      String(
        base64Data ||
        ''
      )
    );

  return HtmlService
    .createHtmlOutput(
      [
        '<!doctype html>',
        '<html><head>',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        '<title>Preparing download</title>',
        '<style>body{font-family:Arial,sans-serif;background:#f8fafc;color:#1d2939;padding:32px}.box{max-width:620px;margin:40px auto;background:#fff;border:1px solid #eaecf0;border-radius:12px;padding:24px}</style>',
        '</head><body>',
        '<div class="box"><h2>Preparing download</h2><p id="message">Your secure licence download is starting.</p></div>',
        '<script>',
        'const fileName=' + safeFileName + ';',
        'const mimeType=' + safeMimeType + ';',
        'const base64Data=' + safeBase64 + ';',
        'try{',
        'const binary=atob(base64Data);',
        'const bytes=new Uint8Array(binary.length);',
        'for(let i=0;i<binary.length;i++){bytes[i]=binary.charCodeAt(i);}',
        'const url=URL.createObjectURL(new Blob([bytes],{type:mimeType}));',
        'const link=document.createElement("a");',
        'link.href=url;link.download=fileName;document.body.appendChild(link);link.click();link.remove();',
        'setTimeout(function(){URL.revokeObjectURL(url);document.getElementById("message").textContent="Download started. You may close this tab.";},1500);',
        '}catch(error){document.getElementById("message").textContent="The download could not be prepared. Please close this tab and try again.";}',
        '</script></body></html>'
      ].join('')
    )
    .setTitle(
      'Secure Licence Download'
    );
}


/**
 * App Config.
 */
const APP_CONFIG = Object.freeze({
  PRODUCTION_WEB_APP_URL:
    'https://script.google.com/macros/s/AKfycbxnpEmnzekbkbSdTiQKixvGSiJvvN0ScdN6PH2KiFWLMdYi6FF2NgX2aaw968nCu2YXHw/exec',
});
/**
 * Renders the applicant portal.
 */
function renderApplicantPortal_(parameters) {
  const template = HtmlService.createTemplateFromFile(
    'ApplicantPortal'
  );

  const applicationId = String(
    parameters.ref || ''
  ).trim();

  const secureToken = String(
    parameters.token || ''
  ).trim();

  template.portalData = getApplicantPortalData_(
    applicationId,
    secureToken
  );
  template.applicationId =
  applicationId;

  template.secureToken =
    secureToken;

  return template
    .evaluate()
    .setTitle('CRFFN Licensing Portal')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}

/**
 * Finds and validates one applicant.
 *
 * Only safe information required by the portal is returned.
 */

/**
 * Cached wrapper for the applicant portal summary.
 */
function getApplicantPortalData_(
  applicationId,
  secureToken
) {
  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  const cleanSecureToken =
    String(
      secureToken || ''
    ).trim();

  if (
    !cleanApplicationId ||
    !cleanSecureToken
  ) {
    return getApplicantPortalDataUncached_(
      cleanApplicationId,
      cleanSecureToken
    );
  }

  const key =
    crffnVersionedApplicantKey_(
      crffnApplicantPortalCacheKey_(
        cleanApplicationId,
        cleanSecureToken
      ),
      cleanApplicationId
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
    getApplicantPortalDataUncached_(
      cleanApplicationId,
      cleanSecureToken
    );

  if (
    result &&
    result.ok === true
  ) {
    crffnPutCachedJson_(
      key,
      result,
      CRFFN_PERFORMANCE_CONFIG
        .APPLICANT_PORTAL_TTL_SECONDS
    );
  }

  if (result) {
    result.performanceCache =
      'MISS';
  }

  return result;
}


function getApplicantPortalDataUncached_(
  applicationId,
  secureToken
) {
  if (!applicationId || !secureToken) {
    return {
      ok: false,
      errorCode: 'MISSING_CREDENTIALS',
      message:
        'This applicant portal link is incomplete. Please use the complete link sent to you.',
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
        'We could not find an application matching this link.',
    };
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(header => String(header).trim());

  const headerMap = getHeaderMap_(headers);

  const requiredColumns = [
    'Application ID',
    'Secure Token',
  ];

  requiredColumns.forEach(columnName => {
    if (!headerMap[columnName]) {
      throw new Error(
        'Required column is missing: ' + columnName
      );
    }
  });

  const values = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();

  const applicationIdIndex =
    headerMap['Application ID'] - 1;

  const secureTokenIndex =
    headerMap['Secure Token'] - 1;

  let matchingRow =
    null;

  /*
   * Corrections/resubmissions keep the same Application ID
   * and Secure Token, so always use the newest matching row.
   */
  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    const storedApplicationId =
      String(
        values[index][
          applicationIdIndex
        ] || ''
      ).trim();

    const storedToken =
      String(
        values[index][
          secureTokenIndex
        ] || ''
      ).trim();

    if (
      storedApplicationId ===
        applicationId &&
      storedToken ===
        secureToken
    ) {
      matchingRow =
        values[index];

      break;
    }
  }

  if (!matchingRow) {
    return {
      ok: false,
      errorCode: 'INVALID_ACCESS',
      message:
        'This portal link is invalid or no longer matches an application.',
    };
  }

  const rowObject = rowToObject_(
    headers,
    matchingRow
  );

  return {
    ok: true,

    applicationId:
      rowObject['Application ID'] || '',

    applicantName:
      rowObject['Full Name'] || 'Applicant',

    companyName:
      rowObject['Company Name'] || '',

    email:
      maskEmail_(rowObject[EMAIL_FIELD]),

    submittedAt:
      formatPortalDate_(
        rowObject['Timestamp']
      ),

    recordStatus:
      rowObject['Record Status'] ||
      'Submitted',

    paymentStatus:
      rowObject['Payment Status'] ||
      'Not Submitted',

    paymentReviewNotes:
      rowObject[
        'Payment Review Notes'
      ] || '',

    documentStatus:
      rowObject[
        'Document Status'
      ] ||
      'Not Submitted',

    documentReviewNotes:
      rowObject[
        'Document Review Notes'
      ] || '',

    informationStatus:
      rowObject[
        'Application Information Status'
      ] ||
      'Pending',

    informationReviewNotes:
      rowObject[
        'Application Information Review Notes'
      ] || '',

    applicationCorrectionFields:
      parseApplicationCorrectionFields_(
        rowObject[
          'Application Correction Fields JSON'
        ]
      ),

    applicationCorrectionFormUrl:
      String(
        rowObject[
          'Application Information Status'
        ] || ''
      ).trim() ===
        'Correction Required'
        ? buildPrefilledCorrectionFormUrl_(
            rowObject
          )
        : '',

    verificationStatus:
      rowObject['Verification Status'] ||
      'Pending',

    verificationNotes:
      rowObject[
        'Verification Notes'
      ] || '',

    licenceStatus:
      rowObject['Licence Status'] ||
      'Not Generated',

    applicationPdfUrl:
      safePortalDocumentUrl_(
        rowObject['Application PDF URL']
      ),

    receiptPdfUrl:
      safePortalDocumentUrl_(
        rowObject['Receipt PDF URL']
      ),

    licencePdfUrl:
      safePortalDocumentUrl_(
        rowObject['Licence PDF URL']
      ),
  };
}

/**
 * Parses structured fields selected by the administrator.
 */
function parseApplicationCorrectionFields_(
  value
) {
  const text =
    String(
      value || ''
    ).trim();

  if (!text) {
    return [];
  }

  try {
    const parsed =
      JSON.parse(
        text
      );

    return Array.isArray(
      parsed
    )
      ? parsed
      : [];
  } catch (error) {
    console.warn(
      'Application correction fields could not be parsed.'
    );

    return [];
  }
}


const APPLICATION_CORRECTION_TARGET_TO_FORM_TITLE =
  Object.freeze({
    company_name:
      'Company Name',
    company_rc_number:
      'Company RC Number',
    company_tin:
      'Company TIN',
    crffn_membership_number:
      'CRFFN Membership Number',
    company_address:
      'Company Address',
    full_name:
      'Full Name',
    phone_number:
      'Phone Number',
    residential_address:
      'Residential Address',
    means_of_identification:
      'Means of Identification',
    id_number:
      'ID Number',
    expiry_date:
      'Expiry Date',
    position_held:
      'Position Held',
    area_of_practice:
      'Area of Practice',
  });


/**
 * Returns the exact correction targets for the current application.
 *
 * Structured JSON is the primary source.
 * Review-note parsing is a compatibility fallback for corrections
 * created before the JSON field was available.
 */
function resolveApplicationCorrectionItems_(
  rowObject
) {
  const structured =
    parseApplicationCorrectionFields_(
      rowObject[
        'Application Correction Fields JSON'
      ]
    );

  if (
    Array.isArray(
      structured
    ) &&
    structured.length
  ) {
    return structured;
  }

  return inferApplicationCorrectionItemsFromNotes_(
    rowObject[
      'Application Information Review Notes'
    ]
  );
}


/**
 * Recovers correction targets from readable review notes.
 *
 * Example notes:
 * Company RC Number requires correction.
 * Company TIN does not match the submitted records.
 *
 * This returns:
 * [
 *   { targetCode: 'company_rc_number' },
 *   { targetCode: 'company_tin' }
 * ]
 */
function inferApplicationCorrectionItemsFromNotes_(
  notes
) {
  const text =
    String(
      notes || ''
    ).trim();

  if (!text) {
    return [];
  }

  const titleToCode =
    {};

  Object.keys(
    APPLICATION_CORRECTION_TARGET_TO_FORM_TITLE
  ).forEach(
    function(code) {
      const title =
        APPLICATION_CORRECTION_TARGET_TO_FORM_TITLE[
          code
        ];

      titleToCode[
        title
      ] =
        code;
    }
  );

  /*
   * Match longer titles first so a shorter title cannot
   * accidentally win a prefix match.
   */
  const titles =
    Object.keys(
      titleToCode
    )
      .sort(
        function(a, b) {
          return (
            b.length -
            a.length
          );
        }
      );

  const found =
    [];

  const seen =
    {};

  /*
   * Reasons may be newline-separated, semicolon-separated,
   * or presented as bullet-like text.
   */
  const lines =
    text
      .split(
        /\r?\n|;/
      )
      .map(
        function(line) {
          return String(
            line || ''
          )
            .replace(
              /^[-•\s]+/,
              ''
            )
            .trim();
        }
      )
      .filter(Boolean);

  lines.forEach(
    function(line) {
      for (
        let index = 0;
        index < titles.length;
        index++
      ) {
        const title =
          titles[index];

        if (
          line === title ||
          line.indexOf(
            title + ' '
          ) === 0 ||
          line.indexOf(
            title + ':'
          ) === 0 ||
          line.indexOf(
            title + ' —'
          ) === 0
        ) {
          const code =
            titleToCode[
              title
            ];

          if (!seen[code]) {
            seen[code] =
              true;

            found.push({
              targetCode:
                code,
              recoveredFromNotes:
                true,
            });
          }

          break;
        }
      }
    }
  );

  return found;
}


/**
 * Diagnostic helper for correction-prefill problems.
 *
 * Run manually:
 * diagnoseCorrectionPrefill_('APP-0001')
 *
 * It logs:
 * - stored JSON
 * - review notes
 * - resolved target codes
 * - form titles that will be left blank
 */
function diagnoseCorrectionPrefill_(
  applicationId
) {
  const normalizedId =
    String(
      applicationId || ''
    ).trim();

  if (!normalizedId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      ss
    );

  const lastRow =
    sheet.getLastRow();

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
      .map(
        function(header) {
          return String(
            header || ''
          ).trim();
        }
      );

  const headerMap =
    getHeaderMap_(
      headers
    );

  const applicationIdColumn =
    headerMap[
      'Application ID'
    ];

  if (!applicationIdColumn) {
    throw new Error(
      'Application ID column is missing.'
    );
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        Math.max(
          lastRow - 1,
          1
        ),
        lastColumn
      )
      .getDisplayValues();

  let row =
    null;

  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        values[index][
          applicationIdColumn - 1
        ] || ''
      ).trim() ===
      normalizedId
    ) {
      row =
        values[index];

      break;
    }
  }

  if (!row) {
    throw new Error(
      'Application not found: ' +
      normalizedId
    );
  }

  const rowObject =
    rowToObject_(
      headers,
      row
    );

  const items =
    resolveApplicationCorrectionItems_(
      rowObject
    );

  const blankTitles =
    items
      .map(
        function(item) {
          return (
            APPLICATION_CORRECTION_TARGET_TO_FORM_TITLE[
              String(
                item.targetCode || ''
              )
            ] || ''
          );
        }
      )
      .filter(Boolean);

  const result = {
    applicationId:
      normalizedId,

    informationStatus:
      rowObject[
        'Application Information Status'
      ] || '',

    storedJson:
      rowObject[
        'Application Correction Fields JSON'
      ] || '',

    reviewNotes:
      rowObject[
        'Application Information Review Notes'
      ] || '',

    resolvedItems:
      items,

    fieldsThatWillBeBlank:
      blankTitles,
  };

  console.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Builds a Google Forms prefilled correction URL.
 *
 * All prior application answers are prefilled except:
 * 1. fields explicitly selected by the admin for correction;
 * 2. declaration/confirmation fields which must be reconfirmed.
 */
function buildPrefilledCorrectionFormUrl_(
  rowObject
) {
  const properties =
    PropertiesService
      .getDocumentProperties();

  const formId =
    String(
      properties.getProperty(
        PROP_KEYS.FORM_ID
      ) || ''
    ).trim();

  if (!formId) {
    return '';
  }

  /*
   * Prefer the structured JSON saved by the admin workflow.
   *
   * FALLBACK:
   * Older correction records may have been rejected before the
   * JSON column was introduced. In that case, recover the affected
   * fields from the human-readable Information Review Notes.
   */
  const correctionItems =
    resolveApplicationCorrectionItems_(
      rowObject
    );

  /*
   * A field is left blank ONLY because the admin selected that
   * field as a correction target.
   *
   * The issue/reason does not participate in the prefill rule.
   * Whether the issue is incomplete, incorrect, mismatch,
   * invalid, missing or other, the selected field stays blank.
   */
  if (
    !correctionItems.length
  ) {
    console.error(
      'Application Information is marked Correction Required, but no correction targets could be resolved.'
    );

    /*
     * Returning an empty URL is safer than opening a form with
     * every field prefilled, because that hides the actual
     * correction requirement from the applicant.
     */
    return '';
  }

  const fieldsToLeaveBlank =
    new Set();

  correctionItems.forEach(
    function(item) {
      const targetCode =
        String(
          item &&
          item.targetCode ||
          ''
        ).trim();

      if (!targetCode) {
        return;
      }

      const formTitle =
        APPLICATION_CORRECTION_TARGET_TO_FORM_TITLE[
          targetCode
        ];

      if (formTitle) {
        fieldsToLeaveBlank.add(
          formTitle
        );
      }
    }
  );

  const doNotPrefill =
    new Set([
      'Supporting Documents Confirmation',
      '4. DECLARATION',
    ]);

  const form =
    FormApp.openById(
      formId
    );

  const response =
    form.createResponse();

  form.getItems().forEach(
    function(item) {
      const title =
        String(
          item.getTitle() || ''
        ).trim();

      /*
       * Never create a prefilled response for a field selected
       * by the admin for correction. This guarantees that the
       * field opens blank regardless of the selected issue.
       */
      if (
        !title ||
        fieldsToLeaveBlank.has(
          title
        ) ||
        doNotPrefill.has(
          title
        )
      ) {
        return;
      }

      const previousValue =
        rowObject[
          title
        ];

      if (
        previousValue ===
          undefined ||
        previousValue ===
          null ||
        String(
          previousValue
        ).trim() ===
          ''
      ) {
        return;
      }

      try {
        const itemResponse =
          createPrefilledItemResponse_(
            item,
            previousValue
          );

        if (itemResponse) {
          response.withItemResponse(
            itemResponse
          );
        }
      } catch (error) {
        console.warn(
          'Could not prefill "' +
          title +
          '": ' +
          error.message
        );
      }
    }
  );

  return response
    .toPrefilledUrl();
}


function createPrefilledItemResponse_(
  item,
  value
) {
  const type =
    item.getType();

  const text =
    String(
      value || ''
    ).trim();

  if (!text) {
    return null;
  }

  if (
    type ===
    FormApp.ItemType.TEXT
  ) {
    return item
      .asTextItem()
      .createResponse(
        text
      );
  }

  if (
    type ===
    FormApp.ItemType.PARAGRAPH_TEXT
  ) {
    return item
      .asParagraphTextItem()
      .createResponse(
        text
      );
  }

  if (
    type ===
    FormApp.ItemType.MULTIPLE_CHOICE
  ) {
    return item
      .asMultipleChoiceItem()
      .createResponse(
        text
      );
  }

  if (
    type ===
    FormApp.ItemType.LIST
  ) {
    return item
      .asListItem()
      .createResponse(
        text
      );
  }

  if (
    type ===
    FormApp.ItemType.CHECKBOX
  ) {
    const choices =
      text
        .split(',')
        .map(function(choice) {
          return String(
            choice || ''
          ).trim();
        })
        .filter(Boolean);

    return item
      .asCheckboxItem()
      .createResponse(
        choices
      );
  }

  if (
    type ===
    FormApp.ItemType.DATE
  ) {
    const date =
      new Date(
        text
      );

    if (
      isNaN(
        date.getTime()
      )
    ) {
      return null;
    }

    return item
      .asDateItem()
      .createResponse(
        date
      );
  }

  return null;
}


/**
 * ============================================================
 * SECURE RELEASED-LICENCE DOWNLOAD
 * ============================================================
 *
 * The Drive file remains private.
 * The applicant opens a web-app URL containing their existing
 * Application ID + Secure Token.
 */
function renderLicenceDownload_(
  parameters
) {
  const input =
    parameters || {};

  const template =
    HtmlService
      .createTemplateFromFile(
        'LicenceDownload'
      );

  template.applicationId =
    String(
      input.ref ||
      input.applicationId ||
      ''
    ).trim();

  template.secureToken =
    String(
      input.token ||
      input.secureToken ||
      ''
    ).trim();

  return template
    .evaluate()
    .setTitle(
      'Download Your CRFFN Licence'
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );
}


/**
 * Called by LicenceDownload.html.
 *
 * Validates:
 * - Application ID
 * - Secure Token
 * - Licence Status = Released
 * - Stamped Licence Approval Status = Approved
 * - private stamped file still exists
 *
 * Only then is the PDF returned to the browser as base64.
 */
function getReleasedLicenceDownloadData(
  applicationId,
  secureToken
) {
  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  const cleanToken =
    String(
      secureToken || ''
    ).trim();

  if (
    !cleanApplicationId ||
    !cleanToken
  ) {
    throw new Error(
      'The licence download link is incomplete.'
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      ss
    );

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    throw new Error(
      'The application could not be found.'
    );
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
      .map(
        function(header) {
          return String(
            header || ''
          ).trim();
        }
      );

  const headerMap =
    getHeaderMap_(
      headers
    );

  const idColumn =
    headerMap[
      'Application ID'
    ];

  const tokenColumn =
    headerMap[
      'Secure Token'
    ];

  if (
    !idColumn ||
    !tokenColumn
  ) {
    throw new Error(
      'The application security columns are missing.'
    );
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getDisplayValues();

  let matchingRow =
    null;

  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    const storedId =
      String(
        values[
          index
        ][
          idColumn - 1
        ] || ''
      ).trim();

    const storedToken =
      String(
        values[
          index
        ][
          tokenColumn - 1
        ] || ''
      ).trim();

    if (
      storedId ===
        cleanApplicationId &&
      storedToken ===
        cleanToken
    ) {
      matchingRow =
        values[
          index
        ];

      break;
    }
  }

  if (!matchingRow) {
    throw new Error(
      'This licence download link is invalid.'
    );
  }

  const rowObject =
    rowToObject_(
      headers,
      matchingRow
    );

  const licenceStatus =
    String(
      rowObject[
        'Licence Status'
      ] || ''
    ).trim();

  if (
    licenceStatus !==
    'Released'
  ) {
    throw new Error(
      'This licence has not been released for download.'
    );
  }

  const approvalStatus =
    String(
      rowObject[
        'Stamped Licence Approval Status'
      ] || ''
    ).trim();

  if (
    approvalStatus !==
    'Approved'
  ) {
    throw new Error(
      'The final licence has not been approved for download.'
    );
  }

  const stampedFileId =
    String(
      rowObject[
        'Stamped Licence File ID'
      ] || ''
    ).trim();

  if (!stampedFileId) {
    throw new Error(
      'The released licence file could not be found.'
    );
  }

  const file =
    DriveApp
      .getFileById(
        stampedFileId
      );

  if (
    file.isTrashed()
  ) {
    throw new Error(
      'The released licence file is currently unavailable.'
    );
  }

  const blob =
    file.getBlob();

  const bytes =
    blob.getBytes();

  const base64 =
    Utilities
      .base64Encode(
        bytes
      );

  const licenceNumber =
    String(
      rowObject[
        'Licence Number'
      ] ||
      cleanApplicationId
    ).trim();

  const safeLicenceNumber =
    licenceNumber
      .replace(
        /[\\/:*?"<>|#%{}[\]]/g,
        '-'
      )
      .replace(
        /\s+/g,
        ' '
      )
      .trim();

  return {
    ok: true,

    applicationId:
      cleanApplicationId,

    licenceNumber:
      licenceNumber,

    fileName:
      'CRFFN Licence - ' +
      safeLicenceNumber +
      '.pdf',

    mimeType:
      'application/pdf',

    base64Data:
      base64
  };
}


/**
 * Converts one spreadsheet row into an object.
 */
function rowToObject_(headers, rowValues) {
  const result = {};

  headers.forEach((header, index) => {
    if (!header) {
      return;
    }

    result[header] = rowValues[index] || '';
  });

  return result;
}

/**
 * Hides part of the applicant email address.
 *
 * example@example.com becomes:
 * ex*****@example.com
 */
function maskEmail_(email) {
  const normalizedEmail = String(
    email || ''
  ).trim();

  if (!normalizedEmail.includes('@')) {
    return '';
  }

  const parts = normalizedEmail.split('@');
  const localPart = parts[0];
  const domain = parts.slice(1).join('@');

  const visibleCharacters = Math.min(
    2,
    localPart.length
  );

  const visiblePart = localPart.slice(
    0,
    visibleCharacters
  );

  const hiddenPart = '*'.repeat(
    Math.max(localPart.length - visibleCharacters, 3)
  );

  return (
    visiblePart +
    hiddenPart +
    '@' +
    domain
  );
}

/**
 * Formats a date value for portal display.
 */
function formatPortalDate_(value) {
  if (!value) {
    return '';
  }

  const parsedDate = new Date(value);

  if (isNaN(parsedDate.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(
    parsedDate,
    Session.getScriptTimeZone(),
    'dd MMMM yyyy, hh:mm a'
  );
}

/**
 * Allows only valid HTTP or HTTPS document links.
 */
function safePortalDocumentUrl_(url) {
  const normalizedUrl = String(
    url || ''
  ).trim();

  if (
    !/^https?:\/\/[^\s]+$/i.test(
      normalizedUrl
    )
  ) {
    return '';
  }

  return normalizedUrl;
}
/**
 * Returns the visual CSS class for a status.
 */
function getStatusClass_(status) {
  const normalizedStatus = String(
    status || ''
  )
    .trim()
    .toLowerCase();

  const successStatuses = [
    'approved',
    'confirmed',
    'generated',
    'released',
    'completed',
  ];

  const warningStatuses = [
    'new',
    'pending',
    'under review',
    'pending verification',
    'not submitted',
    'not generated',
    'resubmission',
    'needs review',
    'correction required',
  ];

  const dangerStatuses = [
    'rejected',
    'revoked',
    'suspended',
    'duplicate',
    'duplicate reference',
  ];

  if (successStatuses.includes(normalizedStatus)) {
    return 'badge badge-success';
  }

  if (dangerStatuses.includes(normalizedStatus)) {
    return 'badge badge-danger';
  }

  if (warningStatuses.includes(normalizedStatus)) {
    return 'badge badge-warning';
  }

  return 'badge badge-neutral';
}

/**
 * Safely renders one portal document row.
 */
function renderDocumentItem_(
  documentName,
  documentUrl
) {
  const safeName = escapeHtml_(
    documentName
  );

  if (!documentUrl) {
    return [
      '<div class="document-item">',
      '<div>',
      '<div class="document-name">',
      safeName,
      '</div>',
      '<div class="document-state">',
      'Not available yet',
      '</div>',
      '</div>',
      '<span class="button button-disabled">',
      'Unavailable',
      '</span>',
      '</div>',
    ].join('');
  }

  const safeUrl = escapeHtml_(
    documentUrl
  );

  return [
    '<div class="document-item">',
    '<div>',
    '<div class="document-name">',
    safeName,
    '</div>',
    '<div class="document-state">',
    'Available for download',
    '</div>',
    '</div>',
    '<a class="button" href="',
    safeUrl,
    '" target="_blank" rel="noopener noreferrer">',
    'Download',
    '</a>',
    '</div>',
  ].join('');
}

/**
 * Escapes values before inserting them into HTML strings.
 */
function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
/**
 * Saves the deployed Apps Script web-app URL.
 *
 * Run this function manually once after deploying the web app.
 */
function setWebAppUrl() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Set Web App URL',
    [
      'Paste the deployed web-app URL ending in /exec.',
      '',
      'Example:',
      'https://script.google.com/macros/s/DEPLOYMENT_ID/exec',
    ].join('\n'),
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  let webAppUrl = String(
    response.getResponseText() || ''
  ).trim();

  if (!webAppUrl) {
    ui.alert('No web-app URL was entered.');
    return;
  }

  // Remove query parameters and trailing slashes.
  webAppUrl = webAppUrl
    .split('?')[0]
    .replace(/\/+$/, '');

  const isValidWebAppUrl =
    /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/i
      .test(webAppUrl);

  if (!isValidWebAppUrl) {
    ui.alert(
      [
        'The URL entered is not a valid deployed web-app URL.',
        '',
        'It must end in /exec, not /dev.',
        '',
        'Expected format:',
        'https://script.google.com/macros/s/DEPLOYMENT_ID/exec',
      ].join('\n')
    );

    return;
  }

  PropertiesService
    .getDocumentProperties()
    .setProperty(
      PROP_KEYS.WEB_APP_URL,
      webAppUrl
    );

  ui.alert(
    [
      'Web-app URL saved successfully.',
      '',
      webAppUrl,
      '',
      'You can now run updateAllApplicantPortalUrls.',
    ].join('\n')
  );
}

/**
 * Returns the stored deployed web-app URL.
 */
function getWebAppUrl_() {
  return String(
    APP_CONFIG.PRODUCTION_WEB_APP_URL || ''
  )
    .trim()
    .split('?')[0]
    .replace(/\/+$/, '');
}

/**
 * Builds a secure applicant portal URL.
 */
function buildApplicantPortalUrl_(
  applicationId,
  secureToken
) {
  const webAppUrl = APP_CONFIG.PRODUCTION_WEB_APP_URL;

  if (!webAppUrl) {
    return 'Pending web app deployment';
  }

  if (!applicationId || !secureToken) {
    return '';
  }

  return (
    webAppUrl +
    '?view=applicant' +
    '&ref=' +
    encodeURIComponent(applicationId) +
    '&token=' +
    encodeURIComponent(secureToken)
  );
}
/**
 * Updates secure tokens and portal URLs
 * for all existing application rows.
 */
function updateAllApplicantPortalUrls() {
  const webAppUrl = getWebAppUrl_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getResponseSheet_(ss);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('No application records found.');
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value || '').trim());

  const headerMap = getHeaderMap_(headers);

  const applicationIdColumn =
    headerMap['Application ID'];

  const secureTokenColumn =
    headerMap['Secure Token'];

  const portalUrlColumn =
    headerMap['Portal URL'];

  if (
    !applicationIdColumn ||
    !secureTokenColumn ||
    !portalUrlColumn
  ) {
    throw new Error(
      'Application ID, Secure Token or Portal URL column is missing.'
    );
  }

  const numberOfRows = lastRow - 1;

  const applicationIds = sheet
    .getRange(
      2,
      applicationIdColumn,
      numberOfRows,
      1
    )
    .getDisplayValues();

  const existingTokens = sheet
    .getRange(
      2,
      secureTokenColumn,
      numberOfRows,
      1
    )
    .getDisplayValues();

  const tokenValues = [];
  const portalUrlValues = [];

  for (let index = 0; index < numberOfRows; index++) {
    const applicationId = String(
      applicationIds[index][0] || ''
    ).trim();

    let secureToken = String(
      existingTokens[index][0] || ''
    ).trim();

    if (!applicationId) {
      tokenValues.push([secureToken]);
      portalUrlValues.push(['']);
      continue;
    }

    if (!secureToken) {
      secureToken = Utilities
        .getUuid()
        .replace(/-/g, '');
    }

    const portalUrl =
      webAppUrl +
      '?view=applicant' +
      '&ref=' +
      encodeURIComponent(applicationId) +
      '&token=' +
      encodeURIComponent(secureToken);

    tokenValues.push([secureToken]);
    portalUrlValues.push([portalUrl]);
  }

  sheet
    .getRange(
      2,
      secureTokenColumn,
      numberOfRows,
      1
    )
    .setNumberFormat('@')
    .setValues(tokenValues);

  sheet
    .getRange(
      2,
      portalUrlColumn,
      numberOfRows,
      1
    )
    .setNumberFormat('@')
    .setValues(portalUrlValues);

  SpreadsheetApp.flush();

  console.log(
    'Updated ' + numberOfRows + ' applicant portal URL(s).'
  );
}
//Diagnisis
function diagnosePortalUrlColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No active spreadsheet was found. Open Apps Script from Extensions → Apps Script inside the spreadsheet.'
    );
  }

  const sheet = getResponseSheet_(ss);

  if (!sheet) {
    throw new Error(
      'getResponseSheet_ did not return a sheet.'
    );
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    throw new Error(
      'The selected response sheet has no columns.'
    );
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(header => String(header || '').trim());

  const headerMap = getHeaderMap_(headers);

  const applicationIdColumn =
    headerMap['Application ID'] || 0;

  const secureTokenColumn =
    headerMap['Secure Token'] || 0;

  const portalUrlColumn =
    headerMap['Portal URL'] || 0;

  let firstApplicationId = '';
  let firstSecureToken = '';
  let firstPortalUrl = '';

  if (lastRow >= 2) {
    if (applicationIdColumn > 0) {
      firstApplicationId = sheet
        .getRange(2, applicationIdColumn)
        .getDisplayValue();
    }

    if (secureTokenColumn > 0) {
      firstSecureToken = sheet
        .getRange(2, secureTokenColumn)
        .getDisplayValue();
    }

    if (portalUrlColumn > 0) {
      firstPortalUrl = sheet
        .getRange(2, portalUrlColumn)
        .getDisplayValue();
    }
  }

  const result = {
    spreadsheetName: ss.getName(),
    selectedResponseSheet: sheet.getName(),
    lastRow: lastRow,
    lastColumn: lastColumn,
    applicationIdColumn:
      applicationIdColumn || 'Missing',
    secureTokenColumn:
      secureTokenColumn || 'Missing',
    portalUrlColumn:
      portalUrlColumn || 'Missing',
    firstApplicationId:
      firstApplicationId || 'Empty',
    firstSecureToken:
      firstSecureToken || 'Empty',
    firstPortalUrl:
      firstPortalUrl || 'Empty',
  };

  console.log(JSON.stringify(result, null, 2));

  return result;
}
function testStep1() {
  console.log('Step 1 works');
}
// createPortalUrl 
function createPortalUrlForNewSubmission_(e) {
  if (!e || !e.range) {
    throw new Error(
      'This function must run from a spreadsheet form-submit trigger.'
    );
  }

  const sheet = e.range.getSheet();
  const rowNumber = e.range.getRow();

  if (rowNumber < 2) {
    return;
  }

  const lastColumn = sheet.getLastColumn();

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(value => String(value || '').trim());

  const headerMap = getHeaderMap_(headers);

  const applicationIdColumn =
    headerMap['Application ID'];

  const secureTokenColumn =
    headerMap['Secure Token'];

  const portalUrlColumn =
    headerMap['Portal URL'];

  if (
    !applicationIdColumn ||
    !secureTokenColumn ||
    !portalUrlColumn
  ) {
    throw new Error(
      'Application ID, Secure Token or Portal URL column is missing.'
    );
  }

  const applicationId = String(
    sheet
      .getRange(rowNumber, applicationIdColumn)
      .getDisplayValue() || ''
  ).trim();

  if (!applicationId) {
    throw new Error(
      'Application ID has not been generated for row ' +
      rowNumber
    );
  }

  let secureToken = String(
    sheet
      .getRange(rowNumber, secureTokenColumn)
      .getDisplayValue() || ''
  ).trim();

  if (!secureToken) {
    secureToken = Utilities
      .getUuid()
      .replace(/-/g, '') +
      Utilities
        .getUuid()
        .replace(/-/g, '');
  }

  const portalUrl = buildApplicantPortalUrl_(
    applicationId,
    secureToken
  );

  sheet
    .getRange(rowNumber, secureTokenColumn)
    .setNumberFormat('@')
    .setValue(secureToken);

  sheet
    .getRange(rowNumber, portalUrlColumn)
    .setNumberFormat('@')
    .setValue(portalUrl);

  SpreadsheetApp.flush();

  console.log(
    'Applicant portal created for ' + applicationId
  );

  return portalUrl;
}