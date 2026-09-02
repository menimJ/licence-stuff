/**
 * ============================================================
 * STAMPED LICENCE WORKFLOW
 * ============================================================
 *
 * Final storage rules:
 *
 * 1. Unstamped Google Doc + PDF are temporary working files.
 * 2. After a stamped PDF uploads successfully, both unstamped
 *    files are moved to Drive Trash and their Sheet URLs cleared.
 * 3. The stamped PDF remains PRIVATE in Drive.
 * 4. The applicant is NEVER added as a Drive viewer.
 * 5. The release email contains only a secure web-app download
 *    link. No Drive URL is sent to the applicant.
 *
 * Dependencies expected elsewhere in the project:
 * - requireAdminAccess_()
 * - getResponseSheet_(spreadsheet)
 * - findAdminApplicationRecord_(sheet, applicationId)
 * - getAdminRecordValue_(record, header)
 * - setAdminRecordValue_(sheet, record, header, value)
 * - APP_CONFIG.PRODUCTION_WEB_APP_URL
 * - LICENCE_CONFIG.DRAFT_LICENCES_FOLDER_ID
 */

const STAMPED_LICENCE_CONFIG =
  Object.freeze({
    MAXIMUM_FILE_SIZE_BYTES:
      10 * 1024 * 1024,

    APPROVAL_PENDING:
      'Pending Approval',

    APPROVAL_APPROVED:
      'Approved',

    SUPPORT_EMAIL:
      'licensing@crffn.gov.ng'
  });


/**
 * Run once after adding/replacing this file.
 */
function setupStampedLicenceColumns() {
  assertAdminAuthEditorSetup_();

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      spreadsheet
    );

  if (
    !sheet ||
    typeof sheet.getLastColumn !==
      'function'
  ) {
    throw new Error(
      'The application sheet could not be found.'
    );
  }

  const requiredHeaders = [
    'Stamped Licence PDF URL',
    'Stamped Licence File ID',
    'Stamped Licence Uploaded At',
    'Stamped Licence Uploaded By',
    'Stamped Licence Approval Status',
    'Stamped Licence Approved At',
    'Stamped Licence Approved By',
    'Licence Released At',
    'Licence Released By',
    'Licence Email Sent At'
  ];

  const lastColumn =
    sheet.getLastColumn();

  const headers =
    lastColumn > 0
      ? sheet
          .getRange(
            1,
            1,
            1,
            lastColumn
          )
          .getDisplayValues()[0]
          .map(
            function(value) {
              return String(
                value || ''
              ).trim();
            }
          )
      : [];

  const missingHeaders =
    requiredHeaders.filter(
      function(header) {
        return (
          headers.indexOf(
            header
          ) ===
          -1
        );
      }
    );

  if (
    !missingHeaders.length
  ) {
    return {
      ok: true,
      addedColumns: [],
      message:
        'All stamped-licence columns already exist.'
    };
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
    ]);

  SpreadsheetApp.flush();

  return {
    ok: true,
    addedColumns:
      missingHeaders,
    message:
      missingHeaders.length +
      ' stamped-licence columns were added.'
  };
}


/**
 * Receives a base64 stamped PDF from the admin portal.
 *
 * IMPORTANT:
 * The draft Google Doc/PDF are removed only AFTER the new
 * stamped PDF has been created successfully.
 */
function uploadStampedLicence(
  payload
) {
  const input =
    payload || {};

  requireAdminAccess_(
    input.sessionToken
  );

  const applicationId =
    String(
      input.applicationId || ''
    ).trim();

  const fileName =
    String(
      input.fileName || ''
    ).trim();

  const mimeType =
    String(
      input.mimeType ||
      'application/pdf'
    ).trim();

  const base64Data =
    String(
      input.base64Data || ''
    ).trim();

  if (!applicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  if (
    !fileName ||
    !base64Data
  ) {
    throw new Error(
      'Select a stamped PDF to upload.'
    );
  }

  if (
    mimeType !==
      'application/pdf' &&
    !fileName
      .toLowerCase()
      .endsWith(
        '.pdf'
      )
  ) {
    throw new Error(
      'Only PDF files are allowed.'
    );
  }

  const bytes =
    Utilities
      .base64Decode(
        base64Data
      );

  if (
    bytes.length >
    STAMPED_LICENCE_CONFIG
      .MAXIMUM_FILE_SIZE_BYTES
  ) {
    throw new Error(
      'The stamped licence must not exceed 10 MB.'
    );
  }

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      spreadsheet
    );

  const record =
    findAdminApplicationRecord_(
      sheet,
      applicationId
    );

  if (!record) {
    throw new Error(
      'The application record was not found.'
    );
  }

  const licenceStatus =
    String(
      getAdminRecordValue_(
        record,
        'Licence Status'
      ) || ''
    ).trim();

  if (
    licenceStatus !==
      'Generated' &&
    licenceStatus !==
      'Released'
  ) {
    throw new Error(
      'Generate the unstamped licence before uploading the stamped licence.'
    );
  }

  const folder =
    DriveApp
      .getFolderById(
        LICENCE_CONFIG
          .DRAFT_LICENCES_FOLDER_ID
      );

  const companyName =
    getFirstAdminRecordValue_(
      record,
      [
        'Company Name',
        'Registered Company Name',
        'Business Name',
        'Applicant Name'
      ]
    ) ||
    applicationId;

  const licenceNumber =
    getAdminRecordValue_(
      record,
      'Licence Number'
    ) ||
    applicationId;

  const safeName =
    sanitizeStampedLicenceFileName_(
      companyName
    );

  const safeLicenceNumber =
    sanitizeStampedLicenceFileName_(
      licenceNumber
    );

  const finalFileName =
    'STAMPED Licence - ' +
    safeName +
    ' - ' +
    safeLicenceNumber +
    '.pdf';

  const previousFileId =
    String(
      getAdminRecordValue_(
        record,
        'Stamped Licence File ID'
      ) || ''
    ).trim();

  /*
   * Create the replacement stamped file BEFORE removing anything.
   */
  const blob =
    Utilities.newBlob(
      bytes,
      'application/pdf',
      finalFileName
    );

  const file =
    folder.createFile(
      blob
    );

  /*
   * At this point the new stamped PDF exists safely.
   * Remove the previous stamped PDF if this is a replacement.
   */
  if (previousFileId) {
    try {
      if (
        previousFileId !==
        file.getId()
      ) {
        DriveApp
          .getFileById(
            previousFileId
          )
          .setTrashed(
            true
          );
      }
    } catch (error) {
      console.warn(
        'Previous stamped licence could not be removed:',
        error
      );
    }
  }

  const now =
    new Date();

  const adminEmail =
    getCurrentLicenceAdminEmail_();

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence PDF URL',
    file.getUrl()
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence File ID',
    file.getId()
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Uploaded At',
    now
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Uploaded By',
    adminEmail
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approval Status',
    STAMPED_LICENCE_CONFIG
      .APPROVAL_PENDING
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approved At',
    ''
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approved By',
    ''
  );

  /*
   * Uploading/replacing a stamped file returns the licence to
   * Generated until the stamped file is explicitly approved.
   */
  setAdminRecordValue_(
    sheet,
    record,
    'Licence Status',
    'Generated'
  );

  /*
   * Now that the stamped file is safely stored, clean up the
   * temporary unstamped Google Doc and PDF.
   */
  const draftCleanup =
    cleanupUnstampedLicenceFiles_(
      sheet,
      record
    );

  SpreadsheetApp.flush();

  return {
    ok: true,

    stampedLicencePdfUrl:
      file.getUrl(),

    stampedLicenceDownloadUrl:
      getDriveDownloadUrlFromFileId_(
        file.getId()
      ),

    approvalStatus:
      STAMPED_LICENCE_CONFIG
        .APPROVAL_PENDING,

    draftFilesTrashed:
      draftCleanup.trashed,

    message:
      'Stamped licence uploaded successfully. The temporary unstamped licence files were cleaned up. Review the stamped licence and approve it before release.'
  };
}


/**
 * Moves the generated unstamped Google Doc/PDF to Drive Trash
 * and clears their Sheet URLs.
 *
 * This is intentionally non-destructive to the Licence Number,
 * Document Reference, generated timestamps, etc.
 */
function cleanupUnstampedLicenceFiles_(
  sheet,
  record
) {
  const draftUrls = [
    {
      header:
        'Licence Document URL',
      url:
        String(
          getAdminRecordValue_(
            record,
            'Licence Document URL'
          ) || ''
        ).trim()
    },
    {
      header:
        'Licence PDF URL',
      url:
        String(
          getAdminRecordValue_(
            record,
            'Licence PDF URL'
          ) || ''
        ).trim()
    }
  ];

  let trashed =
    0;

  let skipped =
    0;

  const seenFileIds =
    {};

  draftUrls.forEach(
    function(item) {
      const fileId =
        extractDriveFileId_(
          item.url
        );

      if (
        fileId &&
        !seenFileIds[
          fileId
        ]
      ) {
        seenFileIds[
          fileId
        ] =
          true;

        try {
          DriveApp
            .getFileById(
              fileId
            )
            .setTrashed(
              true
            );

          trashed +=
            1;
        } catch (error) {
          skipped +=
            1;

          console.warn(
            'Temporary unstamped licence file could not be trashed:',
            fileId,
            error
          );
        }
      }

      /*
       * Clear the draft URL regardless. Once a stamped licence has
       * uploaded successfully the draft is no longer part of the
       * active workflow.
       */
      setAdminRecordValue_(
        sheet,
        record,
        item.header,
        ''
      );
    }
  );

  return {
    trashed:
      trashed,
    skipped:
      skipped
  };
}


/**
 * Approves the stamped licence, releases it and emails ONLY a
 * secure system download link.
 *
 * The stamped Drive file stays private.
 */
function approveAndReleaseStampedLicence(
  applicationId,
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

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      spreadsheet
    );

  const record =
    findAdminApplicationRecord_(
      sheet,
      cleanApplicationId
    );

  if (!record) {
    throw new Error(
      'The application record was not found.'
    );
  }

  const stampedFileId =
    String(
      getAdminRecordValue_(
        record,
        'Stamped Licence File ID'
      ) || ''
    ).trim();

  if (!stampedFileId) {
    throw new Error(
      'Upload the stamped licence before approving it.'
    );
  }

  /*
   * Ensure the private stamped file still exists.
   */
  const file =
    DriveApp.getFileById(
      stampedFileId
    );

  if (
    file.isTrashed()
  ) {
    throw new Error(
      'The stamped licence file is in Drive Trash. Upload it again before release.'
    );
  }

  const applicantEmail =
    getFirstAdminRecordValue_(
      record,
      [
        'Email',
        'Email Address',
        'Applicant Email',
        'Company Email',
        'Company Email Address'
      ]
    );

  if (!applicantEmail) {
    throw new Error(
      'The applicant email address is missing.'
    );
  }

  const secureToken =
    String(
      getAdminRecordValue_(
        record,
        'Secure Token'
      ) || ''
    ).trim();

  if (!secureToken) {
    throw new Error(
      'The applicant secure token is missing. The licence cannot be released safely.'
    );
  }

  const now =
    new Date();

  const adminEmail =
    getCurrentLicenceAdminEmail_();

  const licenceNumber =
    String(
      getAdminRecordValue_(
        record,
        'Licence Number'
      ) || ''
    ).trim();

  const applicantName =
    getFirstAdminRecordValue_(
      record,
      [
        'Applicant Name',
        'Full Name',
        'Name'
      ]
    ) ||
    'Applicant';

  const companyName =
    getFirstAdminRecordValue_(
      record,
      [
        'Company Name',
        'Registered Company Name',
        'Business Name'
      ]
    );

  const secureDownloadUrl =
    buildReleasedLicenceDownloadUrl_(
      cleanApplicationId,
      secureToken
    );

  if (!secureDownloadUrl) {
    throw new Error(
      'The secure licence download URL could not be generated.'
    );
  }

  /*
   * IMPORTANT:
   * Do NOT call file.addViewer().
   * The stamped licence remains private in Drive.
   */
  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approval Status',
    STAMPED_LICENCE_CONFIG
      .APPROVAL_APPROVED
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approved At',
    now
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Stamped Licence Approved By',
    adminEmail
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Licence Status',
    'Released'
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Licence Released At',
    now
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Licence Released By',
    adminEmail
  );

  SpreadsheetApp.flush();

  /*
   * Send the email after the release state has been persisted,
   * because the secure download endpoint requires Released status.
   */
  sendReleasedLicenceEmail_({
    recipientEmail:
      applicantEmail,

    applicantName:
      applicantName,

    companyName:
      companyName,

    applicationId:
      cleanApplicationId,

    licenceNumber:
      licenceNumber,

    secureDownloadUrl:
      secureDownloadUrl
  });

  setAdminRecordValue_(
    sheet,
    record,
    'Licence Email Sent At',
    now
  );

  SpreadsheetApp.flush();

  return {
    ok: true,

    licenceStatus:
      'Released',

    /*
     * Admin portal may continue to use the private Drive URL for
     * admins. It is NOT sent to the applicant.
     */
    stampedLicencePdfUrl:
      file.getUrl(),

    secureLicenceDownloadUrl:
      secureDownloadUrl,

    applicantEmail:
      applicantEmail,

    message:
      'The stamped licence was approved and released. A secure download link was emailed to ' +
      applicantEmail +
      '.'
  };
}


/**
 * Builds the private web-app download URL sent to the applicant.
 */
function buildReleasedLicenceDownloadUrl_(
  applicationId,
  secureToken
) {
  const baseUrl =
    String(
      APP_CONFIG
        .PRODUCTION_WEB_APP_URL ||
      ''
    )
      .trim()
      .split('?')[0]
      .replace(
        /\/+$/,
        ''
      );

  if (
    !baseUrl ||
    !applicationId ||
    !secureToken
  ) {
    return '';
  }

  return (
    baseUrl +
    '?view=downloadLicence' +
    '&ref=' +
    encodeURIComponent(
      applicationId
    ) +
    '&token=' +
    encodeURIComponent(
      secureToken
    )
  );
}


/**
 * Final licence release email.
 *
 * No Drive URL.
 * No attachment.
 * One secure system Download Licence button/link.
 */
function sendReleasedLicenceEmail_(
  data
) {
  const subject =
    'Your CRFFN Licence Is Ready' +
    (
      data.licenceNumber
        ? ' - ' +
          data.licenceNumber
        : ''
    );

  const organisation =
    'CRFFN Licensing Team';

  const companyLine =
    data.companyName
      ? (
        '<p><strong>Company:</strong> ' +
        escapeLicenceEmailHtml_(
          data.companyName
        ) +
        '</p>'
      )
      : '';

  const safeDownloadUrl =
    escapeLicenceEmailHtml_(
      data.secureDownloadUrl
    );

  const htmlBody =
    '<p>Dear ' +
    escapeLicenceEmailHtml_(
      data.applicantName
    ) +
    ',</p>' +

    '<p>Your CRFFN Licence has been approved and released.</p>' +

    '<p><strong>Application ID:</strong> ' +
    escapeLicenceEmailHtml_(
      data.applicationId
    ) +
    '</p>' +

    companyLine +

    '<p><strong>Licence Number:</strong> ' +
    escapeLicenceEmailHtml_(
      data.licenceNumber ||
      '—'
    ) +
    '</p>' +

    '<p>You can download your approved licence using the secure link below:</p>' +

    '<p>' +
      '<a href="' +
      safeDownloadUrl +
      '" style="display:inline-block;padding:12px 18px;background:#183b56;color:#ffffff;text-decoration:none;border-radius:6px;font-weight:bold;">' +
      'Download Your Licence' +
      '</a>' +
    '</p>' +

    '<p>This link is private and is associated with your application. Please do not forward or share it.</p>' +

    '<p>For enquiries, clarification or assistance, contact ' +
    '<a href="mailto:' +
    escapeLicenceEmailHtml_(
      STAMPED_LICENCE_CONFIG
        .SUPPORT_EMAIL
    ) +
    '">' +
    escapeLicenceEmailHtml_(
      STAMPED_LICENCE_CONFIG
        .SUPPORT_EMAIL
    ) +
    '</a> and include your Application ID.</p>' +

    '<p>Regards,<br>' +
    organisation +
    '</p>';

  const plainBody = [
    'Dear ' +
      data.applicantName +
      ',',
    '',
    'Your CRFFN Licence has been approved and released.',
    '',
    'Application ID: ' +
      data.applicationId,
    data.companyName
      ? (
        'Company: ' +
        data.companyName
      )
      : '',
    'Licence Number: ' +
      (
        data.licenceNumber ||
        '—'
      ),
    '',
    'Download Your Licence:',
    data.secureDownloadUrl,
    '',
    'This link is private and is associated with your application. Please do not forward or share it.',
    '',
    'For enquiries, clarification or assistance, contact ' +
      STAMPED_LICENCE_CONFIG
        .SUPPORT_EMAIL +
      ' and include your Application ID.',
    '',
    'Regards,',
    organisation
  ]
    .filter(
      function(line) {
        return line !==
          null;
      }
    )
    .join(
      '\n'
    );

  MailApp.sendEmail({
    to:
      data.recipientEmail,

    subject:
      subject,

    body:
      plainBody,

    htmlBody:
      htmlBody,

    name:
      organisation,

    replyTo:
      STAMPED_LICENCE_CONFIG
        .SUPPORT_EMAIL
  });
}


/**
 * Existing admin helper fields.
 */
function getStampedLicenceAdminFields_(
  record
) {
  return {
    licenceDocumentUrl:
      getAdminRecordValue_(
        record,
        'Licence Document URL'
      ) || '',

    licencePdfUrl:
      getAdminRecordValue_(
        record,
        'Licence PDF URL'
      ) || '',

    licenceStatus:
      getAdminRecordValue_(
        record,
        'Licence Status'
      ) ||
      'Not Generated',

    stampedLicencePdfUrl:
      getAdminRecordValue_(
        record,
        'Stamped Licence PDF URL'
      ) || '',

    stampedLicenceApprovalStatus:
      getAdminRecordValue_(
        record,
        'Stamped Licence Approval Status'
      ) || ''
  };
}


function getDriveDownloadUrlFromUrl_(
  url
) {
  const fileId =
    extractDriveFileId_(
      url
    );

  return fileId
    ? getDriveDownloadUrlFromFileId_(
        fileId
      )
    : String(
        url || ''
      );
}


function getDriveDownloadUrlFromFileId_(
  fileId
) {
  return (
    'https://drive.google.com/uc' +
    '?export=download&id=' +
    encodeURIComponent(
      String(
        fileId || ''
      )
    )
  );
}


function extractDriveFileId_(
  url
) {
  const value =
    String(
      url || ''
    ).trim();

  if (!value) {
    return '';
  }

  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];

  for (
    let index = 0;
    index <
      patterns.length;
    index++
  ) {
    const match =
      value.match(
        patterns[
          index
        ]
      );

    if (match) {
      return match[1];
    }
  }

  return '';
}


function getFirstAdminRecordValue_(
  record,
  headers
) {
  for (
    let index = 0;
    index <
      headers.length;
    index++
  ) {
    const value =
      getAdminRecordValue_(
        record,
        headers[
          index
        ]
      );

    if (
      value !== null &&
      value !== undefined &&
      String(
        value
      ).trim()
    ) {
      return String(
        value
      ).trim();
    }
  }

  return '';
}


function sanitizeStampedLicenceFileName_(
  value
) {
  return String(
    value || ''
  )
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


function getCurrentLicenceAdminEmail_() {
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


function escapeLicenceEmailHtml_(
  value
) {
  return String(
    value || ''
  )
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#039;'
    );
}
