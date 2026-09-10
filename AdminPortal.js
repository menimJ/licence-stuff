/**
 * Admin Portal configuration.
 *
 * Replace these email addresses with the real
 * authorised administrator email addresses.
 */
const ADMIN_PORTAL_CONFIG = Object.freeze({
  ADMIN_EMAILS: [
    'sammymenim@gmail.com',
    'registrarceo.crffn@gmail.com',
    'crffn08@gmail.com',
    'alonzychris@gmail.com'
  ],

  PORTAL_TITLE:
    'CRFFN Licensing Administration',

  MAX_APPLICATIONS:
    500,
});

/**
 * Renders the admin portal.
 *
 * Admin authentication is handled by AdminAuth.gs.
 */
function renderAdminPortal_(parameters) {
  const template =
    HtmlService.createTemplateFromFile(
      'AdminPortalPage'
    );

  const inputParameters =
    parameters || {};

  const sessionToken =
    String(
      inputParameters.session ||
      ''
    ).trim();

  const rawAccessResult =
    getAdminAccessResult_(
      sessionToken
    );

  const accessResult =
    rawAccessResult &&
    typeof rawAccessResult ===
      'object'
      ? rawAccessResult
      : {
          ok: false,
          mustChangePassword: false,
          email: '',
          role: '',
          message:
            'Administrator session has expired. Sign in again.',
        };

  template.adminAccess =
    accessResult;

  template.adminEmail =
    String(
      accessResult.email || ''
    );

  template.adminRole =
    String(
      accessResult.role || ''
    );

  template.adminSessionToken =
    sessionToken;

  template.adminSessionTokenEncoded =
    encodeURIComponent(
      sessionToken
    );

  template.applicationId =
    String(
      inputParameters.applicationId ||
      inputParameters.ref ||
      ''
    ).trim();

  template.adminData = {
    ok: false,
    applications: [],
    application: null,
  };

  if (
    accessResult.ok &&
    !accessResult.mustChangePassword
  ) {
    if (
      template.applicationId
    ) {
      template.adminData =
        getAdminApplicationDetail_(
          template.applicationId,
          sessionToken
        );
    } else {
      template.adminData =
        getAdminApplications_(
          sessionToken
        );
    }
  }

  return template
    .evaluate()
    .setTitle(
      ADMIN_PORTAL_CONFIG.PORTAL_TITLE
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.DEFAULT
    );
}


/**
 * Returns the current password-session access state.
 */
function getAdminAccessResult_(
  sessionToken
) {
  return getAdminSessionAccess_(
    sessionToken
  );
}


/**
 * Every sensitive Admin Portal backend function must call this.
 */
function requireAdminAccess_(
  sessionToken
) {
  return requireAdminSession_(
    sessionToken
  );
}


/**
 * Returns all applications for the
 * admin portal.
 */
function getAdminApplications_(
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(ss);

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow < 2 ||
    lastColumn < 1
  ) {
    return {
      ok: true,
      applications: [],
      application: null,
      total: 0,
      message:
        'No applications have been submitted.',
    };
  }

  const headers = sheet
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

  const numberOfRows =
    Math.min(
      lastRow - 1,
      ADMIN_PORTAL_CONFIG
        .MAX_APPLICATIONS
    );

  const values = sheet
    .getRange(
      lastRow - numberOfRows + 1,
      1,
      numberOfRows,
      lastColumn
    )
    .getDisplayValues();

  /*
   * Work from newest row to oldest row.
   * A Google Form correction/resubmission creates another row
   * with the SAME Application ID. The admin list should show
   * only the latest row for that application.
   */
  const seenApplicationIds =
    {};

  const applications = [];

  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    const rowObject =
      rowToObject_(
        headers,
        values[index]
      );

    const application =
      buildAdminApplicationSummary_(
        rowObject
      );

    const id =
      String(
        application.applicationId || ''
      ).trim();

    if (
      !id ||
      seenApplicationIds[id]
    ) {
      continue;
    }

    seenApplicationIds[id] =
      true;

    applications.push(
      application
    );
  }

  return {
    ok: true,
    applications:
      applications,
    application: null,
    total:
      applications.length,
    message: '',
  };
}


/**
 * Paginated application list for the new dashboard.
 * Legacy getAdminApplications_() remains unchanged.
 */
function getAdminApplicationsPage(sessionToken, options) {
  requireAdminAccess_(sessionToken);

  const input =
    options && typeof options === 'object'
      ? options
      : {};

  const pageSize = Math.max(
    1,
    Math.min(Number(input.pageSize || 20) || 20, 100)
  );

  const requestedPage = Math.max(
    1,
    Number(input.page || 1) || 1
  );

  const search = String(input.search || '')
    .trim()
    .toLowerCase();

  const status = String(input.status || '')
    .trim()
    .toLowerCase();

  const source = getAdminApplications_(sessionToken);

  const applications =
    source && Array.isArray(source.applications)
      ? source.applications
      : [];

  const filtered = applications.filter(function(application) {
    const searchable = [
      application.applicationId,
      application.applicantName,
      application.companyName,
      application.paymentStatus,
      application.documentStatus,
      application.informationStatus,
      application.verificationStatus,
      application.licenceStatus
    ]
      .map(function(value) {
        return String(value || '').toLowerCase();
      })
      .join(' ');

    if (search && searchable.indexOf(search) === -1) {
      return false;
    }

    if (status) {
      const statuses = [
        application.paymentStatus,
        application.documentStatus,
        application.informationStatus,
        application.verificationStatus,
        application.licenceStatus
      ]
        .map(function(value) {
          return String(value || '').toLowerCase();
        })
        .join(' ');

      if (statuses.indexOf(status) === -1) {
        return false;
      }
    }

    return true;
  });

  const totalItems = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  const page = Math.min(requestedPage, totalPages);
  const start = (page - 1) * pageSize;

  return {
    ok: true,
    items: filtered.slice(start, start + pageSize),
    page: page,
    pageSize: pageSize,
    totalItems: totalItems,
    totalPages: totalPages,
    hasPrevious: page > 1,
    hasNext: page < totalPages,
    message: totalItems
      ? ''
      : 'No applications match the current search or filter.'
  };
}



/**
 * Dashboard-only workflow queues.
 * Legacy admin portal behaviour is not changed.
 */
function getAdminDashboardLatestRows_(sessionToken) {
  requireAdminAccess_(sessionToken);

  const sheet =
    getResponseSheet_(
      SpreadsheetApp.getActiveSpreadsheet()
    );

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const headers =
    sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(function(header) {
        return String(header || '').trim();
      });

  const numberOfRows =
    Math.min(
      lastRow - 1,
      ADMIN_PORTAL_CONFIG.MAX_APPLICATIONS
    );

  const values =
    sheet
      .getRange(
        lastRow - numberOfRows + 1,
        1,
        numberOfRows,
        lastColumn
      )
      .getDisplayValues();

  const seen = {};
  const rows = [];

  for (
    let index = values.length - 1;
    index >= 0;
    index--
  ) {
    const rowObject =
      rowToObject_(
        headers,
        values[index]
      );

    const applicationId =
      getFirstAvailableValue_(
        rowObject,
        ['Application ID']
      );

    if (
      !applicationId ||
      seen[applicationId]
    ) {
      continue;
    }

    seen[applicationId] = true;
    rows.push(rowObject);
  }

  return rows;
}


function paginateAdminDashboardItems_(
  items,
  options,
  emptyMessage
) {
  const input =
    options &&
    typeof options === 'object'
      ? options
      : {};

  const pageSize =
    Math.max(
      1,
      Math.min(
        Number(input.pageSize || 20) || 20,
        100
      )
    );

  const requestedPage =
    Math.max(
      1,
      Number(input.page || 1) || 1
    );

  const totalItems = items.length;

  const totalPages =
    Math.max(
      1,
      Math.ceil(totalItems / pageSize)
    );

  const page =
    Math.min(
      requestedPage,
      totalPages
    );

  const start =
    (page - 1) * pageSize;

  return {
    ok: true,
    items:
      items.slice(
        start,
        start + pageSize
      ),
    page: page,
    pageSize: pageSize,
    totalItems: totalItems,
    totalPages: totalPages,
    hasPrevious: page > 1,
    hasNext: page < totalPages,
    message:
      totalItems
        ? ''
        : emptyMessage
  };
}


/**
 * Payment queue for the NEW dashboard.
 */
function getAdminPaymentsPage(
  sessionToken,
  options
) {
  const input =
    options &&
    typeof options === 'object'
      ? options
      : {};

  const search =
    String(input.search || '')
      .trim()
      .toLowerCase();

  const status =
    String(input.status || '')
      .trim()
      .toLowerCase();

  const items =
    getAdminDashboardLatestRows_(
      sessionToken
    )
      .map(function(rowObject) {
        return {
          applicationId:
            getFirstAvailableValue_(
              rowObject,
              ['Application ID']
            ),

          applicantName:
            getFirstAvailableValue_(
              rowObject,
              [
                'Full Name',
                'Applicant Name'
              ]
            ) || 'Applicant',

          companyName:
            getFirstAvailableValue_(
              rowObject,
              [
                'Company Name',
                'Business Name'
              ]
            ),

          paymentReference:
            getFirstAvailableValue_(
              rowObject,
              [
                'Payment Reference',
                'Remita Reference',
                'RRR'
              ]
            ),

          paymentStatus:
            getFirstAvailableValue_(
              rowObject,
              ['Payment Status']
            ) || 'Not Available',

          paymentProofUploadedAt:
            getFirstAvailableValue_(
              rowObject,
              ['Payment Proof Uploaded At']
            ),

          paymentVerifiedAt:
            getFirstAvailableValue_(
              rowObject,
              ['Payment Verified At']
            ),

          hasPaymentProof:
            Boolean(
              getFirstAvailableValue_(
                rowObject,
                [
                  'Payment Proof File ID',
                  'Receipt PDF URL',
                  'Payment Receipt URL',
                  'Payment Proof URL'
                ]
              )
            )
        };
      })
      .filter(function(item) {
        if (status) {
          const itemStatus =
            String(item.paymentStatus || '')
              .trim()
              .toLowerCase();

          if (
            itemStatus.indexOf(status) === -1
          ) {
            return false;
          }
        }

        if (search) {
          const searchable = [
            item.applicationId,
            item.applicantName,
            item.companyName,
            item.paymentReference
          ]
            .map(function(value) {
              return String(value || '')
                .toLowerCase();
            })
            .join(' ');

          if (
            searchable.indexOf(search) === -1
          ) {
            return false;
          }
        }

        return true;
      });

  return paginateAdminDashboardItems_(
    items,
    input,
    'No payments match the current search or filter.'
  );
}


/**
 * Licence queue for the NEW dashboard.
 */
function getAdminLicencesPage(
  sessionToken,
  options
) {
  const input =
    options &&
    typeof options === 'object'
      ? options
      : {};

  const search =
    String(input.search || '')
      .trim()
      .toLowerCase();

  const status =
    String(input.status || '')
      .trim()
      .toLowerCase();

  const items =
    getAdminDashboardLatestRows_(
      sessionToken
    )
      .map(function(rowObject) {
        return {
          applicationId:
            getFirstAvailableValue_(
              rowObject,
              ['Application ID']
            ),

          applicantName:
            getFirstAvailableValue_(
              rowObject,
              [
                'Full Name',
                'Applicant Name'
              ]
            ) || 'Applicant',

          companyName:
            getFirstAvailableValue_(
              rowObject,
              [
                'Company Name',
                'Business Name'
              ]
            ),

          licenceNumber:
            getFirstAvailableValue_(
              rowObject,
              [
                'Licence Number',
                'License Number'
              ]
            ),

          licenceStatus:
            getFirstAvailableValue_(
              rowObject,
              [
                'Licence Status',
                'License Status'
              ]
            ) || 'Not Generated',

          licenceGeneratedAt:
            getFirstAvailableValue_(
              rowObject,
              [
                'Licence Generated At',
                'License Generated At'
              ]
            ),

          licenceReleasedAt:
            getFirstAvailableValue_(
              rowObject,
              [
                'Licence Released At',
                'License Released At'
              ]
            )
        };
      })
      .filter(function(item) {
        if (status) {
          const itemStatus =
            String(item.licenceStatus || '')
              .trim()
              .toLowerCase();

          if (
            itemStatus.indexOf(status) === -1
          ) {
            return false;
          }
        }

        if (search) {
          const searchable = [
            item.applicationId,
            item.applicantName,
            item.companyName,
            item.licenceNumber
          ]
            .map(function(value) {
              return String(value || '')
                .toLowerCase();
            })
            .join(' ');

          if (
            searchable.indexOf(search) === -1
          ) {
            return false;
          }
        }

        return true;
      });

  return paginateAdminDashboardItems_(
    items,
    input,
    'No licences match the current search or filter.'
  );
}


/**
 * Submitted supporting documents for the NEW dashboard.
 * Missing documents are intentionally not returned.
 */
function getAdminDashboardSupportingDocuments_(
  sessionToken,
  applicationId
) {
  const cleanId =
    String(applicationId || '')
      .trim();

  const rowObject =
    getAdminDashboardLatestRows_(
      sessionToken
    ).find(function(row) {
      return (
        getFirstAvailableValue_(
          row,
          ['Application ID']
        ) === cleanId
      );
    });

  if (
    !rowObject ||
    typeof SUPPORTING_DOCUMENT_CONFIG ===
      'undefined'
  ) {
    return [];
  }

  return Object.keys(
    SUPPORTING_DOCUMENT_CONFIG
  )
    .map(function(key) {
      const config =
        SUPPORTING_DOCUMENT_CONFIG[key];

      const fileId =
        getFirstAvailableValue_(
          rowObject,
          [config.fileIdHeader]
        );

      const url =
        getFirstAvailableValue_(
          rowObject,
          [config.urlHeader]
        );

      const uploaded =
        Boolean(fileId || url);

      return {
        type: config.type,
        label: config.label,
        uploaded: uploaded,
        uploadedAt:
          getFirstAvailableValue_(
            rowObject,
            [config.uploadedAtHeader]
          ),
        reviewStatus:
          getFirstAvailableValue_(
            rowObject,
            [config.reviewStatusHeader]
          ) || (
            uploaded
              ? 'Pending Review'
              : 'Not Submitted'
          ),
        reviewNotes:
          getFirstAvailableValue_(
            rowObject,
            [config.reviewNotesHeader]
          ),
        reviewedAt:
          getFirstAvailableValue_(
            rowObject,
            [config.reviewedAtHeader]
          ),
        reviewedBy:
          getFirstAvailableValue_(
            rowObject,
            [config.reviewedByHeader]
          )
      };
    });
}


/**
 * Returns one complete application for
 * admin review.
 */
function getAdminApplicationDetail_(
  applicationId,
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const normalizedApplicationId =
    String(
      applicationId || ''
    ).trim();

  if (!normalizedApplicationId) {
    return {
      ok: false,
      applications: [],
      application: null,
      message:
        'Application ID is required.'
    };
  }

  if (
    typeof getCachedAdminApplicationDetail_ ===
      'function'
  ) {
    const cached =
      getCachedAdminApplicationDetail_(
        normalizedApplicationId
      );

    if (cached) {
      return {
        ok: true,
        applications: [],
        application:
          cached,
        message: ''
      };
    }
  }

  const sheet =
    getResponseSheet_(
      SpreadsheetApp
        .getActiveSpreadsheet()
    );

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {
    return {
      ok: false,
      applications: [],
      application: null,
      message:
        'No application records were found.'
    };
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
      .map(function(header) {
        return String(
          header || ''
        ).trim();
      });

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
      'Required column is missing: Application ID'
    );
  }

  const ids =
    sheet
      .getRange(
        2,
        applicationIdColumn,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let rowNumber =
    -1;

  for (
    let index =
      ids.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        ids[index][0] ||
        ''
      ).trim() ===
      normalizedApplicationId
    ) {
      rowNumber =
        index + 2;

      break;
    }
  }

  if (rowNumber < 2) {
    return {
      ok: false,
      applications: [],
      application: null,
      message:
        'The requested application could not be found.'
    };
  }

  const row =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const application =
    buildAdminApplicationDetail_(
      rowToObject_(
        headers,
        row
      )
    );

  if (
    typeof putCachedAdminApplicationDetail_ ===
      'function'
  ) {
    putCachedAdminApplicationDetail_(
      normalizedApplicationId,
      application
    );
  }

  return {
    ok: true,
    applications: [],
    application:
      application,
    message: ''
  };
}

/**
 * Application detail endpoint for the new dashboard.
 * Reuses the existing legacy application-detail logic.
 */
function getAdminDashboardApplicationDetail(
  sessionToken,
  applicationId
) {
  const result =
    getAdminApplicationDetail_(
      applicationId,
      sessionToken
    );

  if (
    result &&
    result.ok === true &&
    result.application
  ) {
    result.application
      .supportingDocuments =
        getAdminDashboardSupportingDocuments_(
          sessionToken,
          applicationId
        );
  }

  return result;
}


/**
 * Builds one row for the admin
 * application list.
 */
function buildAdminApplicationSummary_(
  rowObject
) {
  return {
    applicationId:
      getFirstAvailableValue_(
        rowObject,
        [
          'Application ID',
        ]
      ),

    applicantName:
      getFirstAvailableValue_(
        rowObject,
        [
          'Full Name',
          'Applicant Name',
        ]
      ) || 'Applicant',

    gender:
      getFirstAvailableValue_(
        rowObject,
        ['Gender']
      ),

    dateOfBirth:
      getFirstAvailableValue_(
        rowObject,
        ['Date of Birth']
      ),

    nationality:
      getFirstAvailableValue_(
        rowObject,
        ['Nationality']
      ),

    stateOfOrigin:
      getFirstAvailableValue_(
        rowObject,
        [
          'State of Origin — Nigerian Applicants Only',
          'State of Origin',
        ]
      ),

    residentialAddress:
      getFirstAvailableValue_(
        rowObject,
        ['Residential Address']
      ),

    meansOfIdentification:
      getFirstAvailableValue_(
        rowObject,
        ['Means of Identification']
      ),

    idNumber:
      getFirstAvailableValue_(
        rowObject,
        ['ID Number']
      ),

    areaOfPractice:
      getFirstAvailableValue_(
        rowObject,
        ['Area of Practice']
      ),

    otherAreaOfPractice:
      getFirstAvailableValue_(
        rowObject,
        ['If Other, please specify']
      ),

    companyName:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company Name',
          'Business Name',
        ]
      ),

    submittedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Timestamp',
          'Submitted At',
        ]
      ),

    paymentStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Status',
        ]
      ) || 'Not Available',

    documentStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Document Status',
          'Documents Status',
        ]
      ) || 'Pending',

    informationStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Application Information Status',
          'Information Status',
        ]
      ) || 'Pending',

    verificationStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Verification Status',
        ]
      ) || 'Pending',

    licenceStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Status',
          'License Status',
        ]
      ) || 'Not Generated',
  };
}

/**
 * Builds the complete admin application
 * detail object.
 */
function buildAdminApplicationDetail_(
  rowObject
) {
  return {
    applicationId:
      getFirstAvailableValue_(
        rowObject,
        [
          'Application ID',
        ]
      ),

    serialNumber:
      getFirstAvailableValue_(
        rowObject,
        [
          'S/N',
          'Serial Number',
          'Application Number',
        ]
      ),

    applicantName:
      getFirstAvailableValue_(
        rowObject,
        [
          'Full Name',
          'Applicant Name',
        ]
      ) || 'Applicant',

    companyName:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company Name',
          'Business Name',
        ]
      ),

    rcNumber:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company RC Number',
          'RC Number',
          'CAC Registration Number',
        ]
      ),

    tin:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company TIN',
          'Company Tax Identification Number (TIN)',
          'Tax Identification Number',
          'TIN',
        ]
      ),

    crffnMembershipNumber:
  getFirstAvailableValue_(
    rowObject,
    [
      'CRFFN Membership Number',

      // Temporary compatibility with old records
      'CRFFN Registration Number',
      'Registration Number',
    ]
  ),

    officeAddress:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company Office Address',
          'Office Address',
          'Company Address',
          'Address',
        ]
      ),

    positionHeld:
      getFirstAvailableValue_(
        rowObject,
        ['Position Held']
      ),

    email:
      getFirstAvailableValue_(
        rowObject,
        [
          'Email Address',
          'Email',
        ]
      ),

    phoneNumber:
      getFirstAvailableValue_(
        rowObject,
        [
          'Company Contact Number',
          'Phone Number',
          'Phone',
          'Contact Number',
        ]
      ),

    submittedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Timestamp',
          'Submitted At',
        ]
      ),

    recordStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Record Status',
        ]
      ) || 'Submitted',

    paymentReference:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Reference',
          'Remita Reference',
          'RRR',
        ]
      ),

    paymentStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Status',
        ]
      ) || 'Not Available',

    paymentProofUploadedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Proof Uploaded At',
        ]
      ),

    paymentVerifiedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Verified At',
        ]
      ),

    paymentVerifiedBy:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Verified By',
        ]
      ),

    paymentRejectionReason:
      getFirstAvailableValue_(
        rowObject,
        [
          'Payment Rejection Reason',
          'Payment Review Notes',
        ]
      ),

    receiptPdfUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Receipt PDF URL',
            'Payment Receipt URL',
            'Payment Proof URL',
          ]
        )
      ),

    applicationPdfUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Application PDF URL',
            'Completed Application PDF URL',
          ]
        )
      ),

    documentStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Document Status',
          'Documents Status',
        ]
      ) || 'Pending',

    documentReviewNotes:
      getFirstAvailableValue_(
        rowObject,
        [
          'Document Review Notes',
          'Document Notes',
        ]
      ),

    informationStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Application Information Status',
          'Information Status',
        ]
      ) || 'Pending',

    informationReviewNotes:
      getFirstAvailableValue_(
        rowObject,
        [
          'Application Information Review Notes',
          'Information Review Notes',
        ]
      ),

    verificationStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Verification Status',
        ]
      ) || 'Pending',

    verificationNotes:
      getFirstAvailableValue_(
        rowObject,
        [
          'Verification Notes',
        ]
      ),

    licenceStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Status',
          'License Status',
        ]
      ) || 'Not Generated',

    licenceNumber:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Number',
          'License Number',
        ]
      ),

    documentReference:
      getFirstAvailableValue_(
        rowObject,
        [
          'Document Reference',
        ]
      ),

    licenceDocumentUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Licence Document URL',
            'License Document URL',
          ]
        )
      ),

    licencePdfUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Licence PDF URL',
            'License PDF URL',
          ]
        )
      ),

    stampedLicencePdfUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Stamped Licence PDF URL',
            'Stamped License PDF URL',
          ]
        )
      ),

    stampedLicenceFileId:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence File ID',
          'Stamped License File ID',
        ]
      ),

    stampedLicenceApprovalStatus:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence Approval Status',
          'Stamped License Approval Status',
        ]
      ),

    stampedLicenceUploadedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence Uploaded At',
          'Stamped License Uploaded At',
        ]
      ),

    stampedLicenceUploadedBy:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence Uploaded By',
          'Stamped License Uploaded By',
        ]
      ),

    stampedLicenceApprovedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence Approved At',
          'Stamped License Approved At',
        ]
      ),

    stampedLicenceApprovedBy:
      getFirstAvailableValue_(
        rowObject,
        [
          'Stamped Licence Approved By',
          'Stamped License Approved By',
        ]
      ),

    licenceReleasedAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Released At',
          'License Released At',
        ]
      ),

    licenceReleasedBy:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Released By',
          'License Released By',
        ]
      ),

    licenceEmailSentAt:
      getFirstAvailableValue_(
        rowObject,
        [
          'Licence Email Sent At',
          'License Email Sent At',
        ]
      ),

    portalUrl:
      safeAdminUrl_(
        getFirstAvailableValue_(
          rowObject,
          [
            'Portal URL',
            'Applicant Portal URL',
          ]
        )
      ),

    canGenerateLicence:
      canGenerateLicenceFromRow_(
        rowObject
      ),
  };
}

/**
 * Returns the first populated value from
 * a list of possible spreadsheet headers.
 */
function getFirstAvailableValue_(
  rowObject,
  possibleHeaders
) {
  for (
    let index = 0;
    index < possibleHeaders.length;
    index++
  ) {
    const header =
      possibleHeaders[index];

    const value =
      rowObject[header];

    if (
      value !== undefined &&
      value !== null &&
      String(value).trim() !== ''
    ) {
      return String(value).trim();
    }
  }

  return '';
}

/**
 * Checks whether all final approval
 * requirements have been completed.
 */
function canGenerateLicenceFromRow_(
  rowObject
) {
  const paymentStatus =
    getFirstAvailableValue_(
      rowObject,
      [
        'Payment Status',
      ]
    )
      .trim()
      .toLowerCase();

  const documentStatus =
    getFirstAvailableValue_(
      rowObject,
      [
        'Document Status',
        'Documents Status',
      ]
    )
      .trim()
      .toLowerCase();

  const informationStatus =
    getFirstAvailableValue_(
      rowObject,
      [
        'Application Information Status',
        'Information Status',
      ]
    )
      .trim()
      .toLowerCase();

  const verificationStatus =
    getFirstAvailableValue_(
      rowObject,
      [
        'Verification Status',
      ]
    )
      .trim()
      .toLowerCase();

  const licenceStatus =
    getFirstAvailableValue_(
      rowObject,
      [
        'Licence Status',
        'License Status',
      ]
    )
      .trim()
      .toLowerCase();

  return (
    paymentStatus === 'confirmed' &&
    documentStatus === 'complete' &&
    informationStatus === 'confirmed' &&
    verificationStatus === 'approved' &&
    (
      !licenceStatus ||
      licenceStatus === 'not generated'
    )
  );
}

/**
 * Allows only valid HTTP or HTTPS URLs
 * in the admin portal.
 */
function safeAdminUrl_(url) {
  const normalizedUrl =
    String(url || '').trim();

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
 * Returns a CSS class for the admin
 * portal status badges.
 */
function getAdminStatusClass_(status) {
  const normalizedStatus =
    String(status || '')
      .trim()
      .toLowerCase();

  const successStatuses = [
    'approved',
    'confirmed',
    'complete',
    'completed',
    'correct',
    'generated',
    'released',
    'active',
  ];

  const warningStatuses = [
    'pending',
    'new',
    'submitted',
    'under review',
    'pending verification',
    'awaiting payment',
    'not available',
    'not submitted',
    'not generated',
    'correction required',
    'incomplete',
  ];

  const dangerStatuses = [
    'rejected',
    'revoked',
    'suspended',
    'duplicate',
    'duplicate reference',
  ];

  if (
    successStatuses.includes(
      normalizedStatus
    )
  ) {
    return 'badge badge-success';
  }

  if (
    dangerStatuses.includes(
      normalizedStatus
    )
  ) {
    return 'badge badge-danger';
  }

  if (
    warningStatuses.includes(
      normalizedStatus
    )
  ) {
    return 'badge badge-warning';
  }

  return 'badge badge-neutral';
}
/**
 * Renders one admin detail field.
 */
function renderAdminDetail_(
  label,
  value
) {
  const safeLabel =
    escapeHtml_(label);

  const safeValue =
    escapeHtml_(
      value || 'Not available'
    );

  return [
    '<div class="detail-item">',
    '<span class="detail-label">',
    safeLabel,
    '</span>',
    '<span class="detail-value">',
    safeValue,
    '</span>',
    '</div>',
  ].join('');
}

/**
 * Renders one admin document link.
 */


const ADMIN_DOCUMENT_REVIEW_CACHE_TTL_SECONDS =
  30;


function getAdminDocumentReviewCacheKey_(
  applicationId,
  suffix
) {
  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  const cleanSuffix =
    String(
      suffix || 'summary'
    ).trim();

  return [
    'CRFFN',
    'ADMIN_DOC_REVIEW',
    cleanApplicationId,
    cleanSuffix
  ].join(':');
}


function getAdminDocumentReviewCache_(
  applicationId,
  suffix
) {
  try {
    const cached =
      CacheService
        .getScriptCache()
        .get(
          getAdminDocumentReviewCacheKey_(
            applicationId,
            suffix
          )
        );

    return cached
      ? JSON.parse(cached)
      : null;
  } catch (error) {
    return null;
  }
}


function putAdminDocumentReviewCache_(
  applicationId,
  suffix,
  value
) {
  try {
    CacheService
      .getScriptCache()
      .put(
        getAdminDocumentReviewCacheKey_(
          applicationId,
          suffix
        ),
        JSON.stringify(value),
        ADMIN_DOCUMENT_REVIEW_CACHE_TTL_SECONDS
      );
  } catch (error) {
    // Cache failure must never block the review workflow.
  }
}


function invalidateAdminDocumentReviewCache_(
  applicationId
) {
  try {
    CacheService
      .getScriptCache()
      .removeAll([
        getAdminDocumentReviewCacheKey_(
          applicationId,
          'summary'
        ),
        getAdminDocumentReviewCacheKey_(
          applicationId,
          'render'
        )
      ]);
  } catch (error) {
    // Best-effort cache invalidation only.
  }
}


function getAdminSupportingDocumentReviewContext_(
  sheet,
  applicationId
) {
  const record =
    findAdminApplicationRecord_(
      sheet,
      applicationId
    );

  if (!record) {
    return null;
  }

  const summary =
    getAdminSupportingDocumentReviewSummary_(
      record
    );

  return {
    record:
      record,
    summary:
      summary,
    overallStatus:
      String(
        getAdminRecordValue_(
          record,
          'Document Status'
        ) || 'Not Submitted'
      ).trim()
  };
}


const ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG =
  Object.freeze({
    cac:
      Object.freeze({
        type: 'cac',
        label: 'CAC Document',
        urlHeader: 'CAC Document URL',
        fileIdHeader: 'CAC Document File ID',
        reviewStatusHeader: 'CAC Document Review Status',
        reviewNotesHeader: 'CAC Document Review Notes',
        reviewedAtHeader: 'CAC Document Reviewed At',
        reviewedByHeader: 'CAC Document Reviewed By'
      }),

    passport:
      Object.freeze({
        type: 'passport',
        label: 'Passport Photograph',
        urlHeader: 'Passport Photograph URL',
        fileIdHeader: 'Passport Photograph File ID',
        reviewStatusHeader: 'Passport Photograph Review Status',
        reviewNotesHeader: 'Passport Photograph Review Notes',
        reviewedAtHeader: 'Passport Photograph Reviewed At',
        reviewedByHeader: 'Passport Photograph Reviewed By'
      }),

    education:
      Object.freeze({
        type: 'education',
        label: 'Educational Certificates',
        urlHeader: 'Educational Certificates URL',
        fileIdHeader: 'Educational Certificates File ID',
        reviewStatusHeader: 'Educational Certificates Review Status',
        reviewNotesHeader: 'Educational Certificates Review Notes',
        reviewedAtHeader: 'Educational Certificates Reviewed At',
        reviewedByHeader: 'Educational Certificates Reviewed By'
      }),

    experience:
      Object.freeze({
        type: 'experience',
        label: 'Proof of Experience',
        urlHeader: 'Proof of Experience URL',
        fileIdHeader: 'Proof of Experience File ID',
        reviewStatusHeader: 'Proof of Experience Review Status',
        reviewNotesHeader: 'Proof of Experience Review Notes',
        reviewedAtHeader: 'Proof of Experience Reviewed At',
        reviewedByHeader: 'Proof of Experience Reviewed By'
      }),

    identification:
      Object.freeze({
        type: 'identification',
        label: 'Valid Means of Identification',
        urlHeader: 'Means of Identification URL',
        fileIdHeader: 'Means of Identification File ID',
        reviewStatusHeader: 'Means of Identification Review Status',
        reviewNotesHeader: 'Means of Identification Review Notes',
        reviewedAtHeader: 'Means of Identification Reviewed At',
        reviewedByHeader: 'Means of Identification Reviewed By'
      })
  });


function ensureAdminSupportingDocumentReviewColumns_(
  sheet
) {
  const requiredHeaders =
    Object.keys(
      ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG
    ).reduce(
      function(headers, key) {
        const config =
          ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG[key];

        return headers.concat([
          config.reviewStatusHeader,
          config.reviewNotesHeader,
          config.reviewedAtHeader,
          config.reviewedByHeader
        ]);
      },
      [
        'Document Correction Email Sent At',
        'Document Correction Email Sent By'
      ]
    );

  const lastColumn =
    Math.max(
      sheet.getLastColumn(),
      1
    );

  const existingHeaders =
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
      sheet.getLastColumn() + 1,
      1,
      missingHeaders.length
    )
    .setValues([
      missingHeaders
    ])
    .setFontWeight(
      'bold'
    );
}


function getAdminSupportingDocumentReviewSummary_(
  record
) {
  const items =
    Object.keys(
      ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG
    ).map(
      function(key) {
        const config =
          ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG[key];

        const uploaded =
          Boolean(
            String(
              getAdminRecordValue_(
                record,
                config.fileIdHeader
              ) ||
              getAdminRecordValue_(
                record,
                config.urlHeader
              ) ||
              ''
            ).trim()
          );

        const status =
          String(
            getAdminRecordValue_(
              record,
              config.reviewStatusHeader
            ) || (
              uploaded
                ? 'Pending Review'
                : 'Not Submitted'
            )
          ).trim();

        const notes =
          String(
            getAdminRecordValue_(
              record,
              config.reviewNotesHeader
            ) || ''
          ).trim();

        return {
          type:
            config.type,
          label:
            config.label,
          uploaded:
            uploaded,
          status:
            status,
          notes:
            notes
        };
      }
    );

  const approvedCount =
    items.filter(
      function(item) {
        return (
          item.status
            .toLowerCase() ===
          'approved'
        );
      }
    ).length;

  const correctionCount =
    items.filter(
      function(item) {
        return (
          item.status
            .toLowerCase() ===
          'correction required'
        );
      }
    ).length;

  const pendingCount =
    items.filter(
      function(item) {
        const status =
          item.status.toLowerCase();

        return (
          item.uploaded &&
          status !== 'approved' &&
          status !== 'correction required'
        );
      }
    ).length;

  const allUploaded =
    items.every(
      function(item) {
        return item.uploaded;
      }
    );

  const allApproved =
    allUploaded &&
    approvedCount ===
      items.length;

  return {
    items:
      items,
    approvedCount:
      approvedCount,
    correctionCount:
      correctionCount,
    pendingCount:
      pendingCount,
    totalCount:
      items.length,
    allUploaded:
      allUploaded,
    allApproved:
      allApproved
  };
}


function deriveOverallSupportingDocumentStatus_(
  summary,
  currentOverallStatus
) {
  const current =
    String(
      currentOverallStatus || ''
    )
      .trim()
      .toLowerCase();

  if (
    current ===
      'complete' &&
    summary.allApproved
  ) {
    return 'Complete';
  }

  if (
    summary.correctionCount >
      0
  ) {
    return 'Correction Required';
  }

  const uploadedCount =
    summary.items.filter(
      function(item) {
        return item.uploaded;
      }
    ).length;

  if (!uploadedCount) {
    return 'Not Submitted';
  }

  if (!summary.allUploaded) {
    return 'In Progress';
  }

  if (summary.allApproved) {
    return 'Ready for Final Approval';
  }

  return 'Pending Review';
}


function renderAdminDocument_(
  documentName,
  documentUrl,
  applicationId,
  documentType,
  sessionToken
) {
  const safeName =
    escapeHtml_(documentName);

  if (!documentUrl) {
    return [
      '<div class="document-item"><div>',
      '<div class="document-name">', safeName, '</div>',
      '<div class="document-state">Not available</div>',
      '</div><span class="button button-disabled">',
      'Unavailable</span></div>'
    ].join('');
  }

  const viewerUrl =
    ScriptApp.getService().getUrl() +
    '?view=adminDocument' +
    '&applicationId=' +
    encodeURIComponent(applicationId) +
    '&documentType=' +
    encodeURIComponent(documentType) +
    '&session=' +
    encodeURIComponent(sessionToken);

  return [
    '<div class="document-item"><div>',
    '<div class="document-name">', safeName, '</div>',
    '<div class="document-state">Available for secure review</div>',
    '</div><a class="button" href="',
    escapeHtml_(viewerUrl),
    '" target="_blank" rel="noopener noreferrer">',
    'View Document</a></div>'
  ].join('');
}


function renderAdminSupportingDocuments_(
  applicationId,
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const sheet =
    getResponseSheet_(
      SpreadsheetApp
        .getActiveSpreadsheet()
    );

  ensureAdminSupportingDocumentReviewColumns_(
    sheet
  );

  const record =
    findAdminApplicationRecord_(
      sheet,
      applicationId
    );

  if (!record) {
    return '';
  }

  const overallStatus =
    String(
      getAdminRecordValue_(
        record,
        'Document Status'
      ) || ''
    )
      .trim()
      .toLowerCase();

  return Object.keys(
    ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG
  )
    .map(
      function(key) {
        const config =
          ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG[key];

        const documentUrl =
          String(
            getAdminRecordValue_(
              record,
              config.urlHeader
            ) || ''
          ).trim();

        const fileId =
          String(
            getAdminRecordValue_(
              record,
              config.fileIdHeader
            ) || ''
          ).trim();

        const uploaded =
          Boolean(
            documentUrl ||
            fileId
          );

        const status =
          String(
            getAdminRecordValue_(
              record,
              config.reviewStatusHeader
            ) || (
              uploaded
                ? 'Pending Review'
                : 'Not Submitted'
            )
          ).trim();

        const notes =
          String(
            getAdminRecordValue_(
              record,
              config.reviewNotesHeader
            ) || ''
          ).trim();

        const locked =
          overallStatus ===
            'complete';

        const viewHtml =
          renderAdminDocument_(
            config.label,
            documentUrl,
            applicationId,
            config.type,
            sessionToken
          );

        if (!uploaded) {
          return [
            '<div class="per-document-review">',
            viewHtml,
            '<div class="document-review-status-row">',
            '<span class="detail-label">Review Status</span>',
            '<span class="badge badge-warning">Not Submitted</span>',
            '</div>',
            '</div>'
          ].join('');
        }

        const statusClass =
          status.toLowerCase() ===
            'approved'
            ? 'badge-success'
            : (
              status.toLowerCase() ===
                'correction required'
                ? 'badge-danger'
                : 'badge-warning'
            );

        const noteHtml =
          notes
            ? [
              '<div class="document-review-note">',
              '<strong>Review note:</strong> ',
              escapeHtml_(notes),
              '</div>'
            ].join('')
            : '';

        const controls =
          locked
            ? '<div class="muted">Finalised and locked.</div>'
            : [
              '<div class="document-review-controls">',
              '<div class="review-actions">',
              '<button type="button" class="button document-review-action" data-document-type="',
              escapeHtml_(config.type),
              '" data-review-decision="approve">Approve</button>',
              '<button type="button" class="button button-danger document-review-action" data-document-type="',
              escapeHtml_(config.type),
              '" data-review-decision="correction">Request Correction</button>',
              '</div>',
              '</div>'
            ].join('');

        return [
          '<div class="per-document-review">',
          viewHtml,
          '<div class="document-review-status-row">',
          '<span class="detail-label">Review Status</span>',
          '<span class="badge ',
          statusClass,
          '">',
          escapeHtml_(status),
          '</span>',
          '</div>',
          noteHtml,
          controls,
          '</div>'
        ].join('');
      }
    )
    .join('');
}


function renderAdminSupportingDocumentsSummary_(
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

  const cached =
    getAdminDocumentReviewCache_(
      cleanApplicationId,
      'summary'
    );

  if (
    cached &&
    typeof cached.html ===
      'string'
  ) {
    return cached.html;
  }

  const sheet =
    getResponseSheet_(
      SpreadsheetApp
        .getActiveSpreadsheet()
    );

  ensureAdminSupportingDocumentReviewColumns_(
    sheet
  );

  const context =
    getAdminSupportingDocumentReviewContext_(
      sheet,
      cleanApplicationId
    );

  if (!context) {
    return '';
  }

  const summary =
    context.summary;

  const overallStatus =
    context.overallStatus;

  const finalised =
    overallStatus
      .toLowerCase() ===
      'complete';

  let helper =
    '';

  if (finalised) {
    helper =
      'All five supporting documents have been approved. The supporting-document stage is Complete.';
  } else if (
    summary.allUploaded &&
    summary.pendingCount === 0 &&
    summary.correctionCount > 0
  ) {
    helper =
      'All five documents have been reviewed. The applicant has been sent one consolidated correction email listing every document that requires correction.';
  } else {
    helper =
      'Review each supporting document individually. No email is sent until all five documents have a review decision.';
  }

  const html = [
    '<div class="admin-review-card">',
    '<div>',
    '<span class="detail-label">Supporting Documents Review</span>',
    '<span id="documentsStatus" class="',
    getAdminStatusClass_(
      overallStatus
    ),
    '">',
    escapeHtml_(
      overallStatus
    ),
    '</span>',
    '<div class="document-review-summary">',
    summary.approvedCount,
    ' of ',
    summary.totalCount,
    ' documents approved',
    summary.correctionCount
      ? ' · ' +
        summary.correctionCount +
        ' correction required'
      : '',
    summary.pendingCount
      ? ' · ' +
        summary.pendingCount +
        ' pending review'
      : '',
    '</div>',
    '<div class="muted">',
    escapeHtml_(helper),
    '</div>',
    '</div>',
    '</div>',
    '<div id="documentsActionMessage" class="admin-action-message" style="display:none"></div>'
  ].join('');

  putAdminDocumentReviewCache_(
    cleanApplicationId,
    'summary',
    {
      html:
        html
    }
  );

  return html;
}

/**
 * Columns required for the complete admin-review workflow.
 */
const ADMIN_REVIEW_COLUMNS = Object.freeze([
  'Payment Status',
  'Payment Review Notes',
  'Payment Verified At',
  'Payment Verified By',
  'Payment Invitation Sent At',

  'Document Status',
  'Document Review Notes',
  'Documents Verified At',
  'Documents Verified By',
  'Document Correction Email Sent At',
  'Document Correction Email Sent By',

  'CAC Document Review Status',
  'CAC Document Review Notes',
  'CAC Document Reviewed At',
  'CAC Document Reviewed By',
  'Passport Photograph Review Status',
  'Passport Photograph Review Notes',
  'Passport Photograph Reviewed At',
  'Passport Photograph Reviewed By',
  'Educational Certificates Review Status',
  'Educational Certificates Review Notes',
  'Educational Certificates Reviewed At',
  'Educational Certificates Reviewed By',
  'Proof of Experience Review Status',
  'Proof of Experience Review Notes',
  'Proof of Experience Reviewed At',
  'Proof of Experience Reviewed By',
  'Means of Identification Review Status',
  'Means of Identification Review Notes',
  'Means of Identification Reviewed At',
  'Means of Identification Reviewed By',

  'Application Information Status',
  'Application Information Review Notes',
  'Application Information Verified At',
  'Application Information Verified By',

  'Verification Status',
  'Verification Notes',
  'Verification Completed At',
  'Verification Completed By',

  'Licence Status',
  'Licence Number',
  'Licence PDF URL',
  'Document Reference',
  'Licence Document URL',
  'Licence Generated At',
  'Licence Generated By',

]);

/**
 * Run this once from the Apps Script editor.
 *
 * It adds any missing review columns without
 * deleting existing spreadsheet information.
 */
function setupAdminReviewColumns(
  sessionToken
) {
  const cleanSessionToken =
    String(
      sessionToken || ''
    ).trim();

  if (cleanSessionToken) {
    requireAdminAccess_(
      cleanSessionToken
    );
  } else {
    assertAdminAuthEditorSetup_();
  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(spreadsheet);

  const lastColumn =
    Math.max(sheet.getLastColumn(), 1);

  const existingHeaders = sheet
    .getRange(
      1,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0]
    .map(header =>
      String(header || '').trim()
    );

  const missingHeaders =
    ADMIN_REVIEW_COLUMNS.filter(header =>
      !existingHeaders.includes(header)
    );

  if (!missingHeaders.length) {
    return {
      ok: true,
      addedColumns: [],
      message:
        'All admin-review columns already exist.',
    };
  }

  const startColumn =
    sheet.getLastColumn() + 1;

  sheet
    .getRange(
      1,
      startColumn,
      1,
      missingHeaders.length
    )
    .setValues([
      missingHeaders,
    ]);

  sheet
    .getRange(
      1,
      startColumn,
      1,
      missingHeaders.length
    )
    .setFontWeight('bold');

  return {
    ok: true,
    addedColumns: missingHeaders,
    message:
      missingHeaders.length +
      ' review columns were added.',
  };
}
/**
 * Standard rejection/correction reasons.
 *
 * The browser sends a reason code. The backend validates the
 * code here and converts it to the official human-readable
 * reason stored in the Sheet and sent to the applicant.
 */
const ADMIN_REJECTION_REASONS =
  Object.freeze({
    payment:
      Object.freeze({
        unreadable:
          'Payment proof is unclear or unreadable',
        invalid_reference:
          'Payment reference is invalid',
        incorrect_amount:
          'Payment amount is incorrect',
        duplicate_reference:
          'Payment reference has already been used',
        unverifiable:
          'Payment could not be verified',
        other:
          'Other',
      }),

    documents:
      Object.freeze({
        missing_document:
          'Required document is missing',
        unreadable:
          'Document is unclear or unreadable',
        incorrect_document:
          'Incorrect document was uploaded',
        expired_document:
          'Document has expired',
        mismatch:
          'Document information does not match the application',
        other:
          'Other',
      }),

    information:
      Object.freeze({
        incomplete:
          'Application information is incomplete',
        rc_number:
          'Company RC Number requires correction',
        tin:
          'Company TIN requires correction',
        membership_number:
          'CRFFN Membership Number requires correction',
        contact_information:
          'Contact information requires correction',
        mismatch:
          'Information does not match the submitted records',
        other:
          'Other',
      }),

    verification:
      Object.freeze({
        outstanding_issue:
          'Outstanding verification issue',
        unable_to_validate:
          'Application information could not be validated',
        regulatory_requirement:
          'A regulatory requirement has not been satisfied',
        earlier_stage:
          'An earlier review stage requires further attention',
        other:
          'Other',
      }),
  });


/**
 * Specific correction targets.
 */
const ADMIN_REJECTION_TARGETS =
  Object.freeze({
    information:
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
        other:
          'Other Application Information',
      }),

    documents:
      Object.freeze({
        cac_document:
          'CAC Document',
        passport_photograph:
          'Passport Photograph',
        educational_certificates:
          'Educational Certificates',
        proof_of_experience:
          'Proof of Experience',
        means_of_identification:
          'Valid Means of Identification',
        other:
          'Other Supporting Document',
      }),
  });


const ADMIN_TARGET_ISSUES =
  Object.freeze({
    information:
      Object.freeze({
        incomplete:
          'is incomplete',
        incorrect:
          'requires correction',
        mismatch:
          'does not match the submitted records',
        invalid:
          'appears invalid',
        missing:
          'is missing',
        other:
          'has another issue',
      }),

    documents:
      Object.freeze({
        missing:
          'is missing',
        unreadable:
          'is unclear or unreadable',
        incorrect:
          'is the incorrect document',
        expired:
          'has expired',
        mismatch:
          'does not match the application information',
        incomplete:
          'is incomplete',
        other:
          'has another issue',
      }),
  });


/**
 * Supported admin-review stages.
 */
const ADMIN_REVIEW_STAGES = Object.freeze({
  payment: {
    statusHeader:
      'Payment Status',

    notesHeader:
      'Payment Review Notes',

    reviewedAtHeader:
      'Payment Verified At',

    reviewedByHeader:
      'Payment Verified By',

    approvedStatus:
      'Confirmed',

    rejectedStatus:
      'Rejected',
  },

  documents: {
    statusHeader:
      'Document Status',

    notesHeader:
      'Document Review Notes',

    reviewedAtHeader:
      'Documents Verified At',

    reviewedByHeader:
      'Documents Verified By',

    approvedStatus:
      'Complete',

    rejectedStatus:
      'Correction Required',
  },

  information: {
    statusHeader:
      'Application Information Status',

    notesHeader:
      'Application Information Review Notes',

    reviewedAtHeader:
      'Application Information Verified At',

    reviewedByHeader:
      'Application Information Verified By',

    approvedStatus:
      'Confirmed',

    rejectedStatus:
      'Correction Required',
  },

  verification: {
    statusHeader:
      'Verification Status',

    notesHeader:
      'Verification Notes',

    reviewedAtHeader:
      'Verification Completed At',

    reviewedByHeader:
      'Verification Completed By',

    approvedStatus:
      'Approved',

    rejectedStatus:
      'Rejected',
  },
});

/**
 * Approves an admin-review stage.
 *
 * Supported stages:
 * payment
 * documents
 * information
 * verification
 */

function reviewSupportingDocument(
  applicationId,
  documentType,
  decision,
  notes,
  sessionToken
) {
  const adminAccess =
    requireAdminAccess_(
      sessionToken
    );

  const cleanApplicationId =
    String(
      applicationId || ''
    ).trim();

  const cleanType =
    String(
      documentType || ''
    )
      .trim()
      .toLowerCase();

  const cleanDecision =
    String(
      decision || ''
    )
      .trim()
      .toLowerCase();

  const cleanNotes =
    String(
      notes || ''
    ).trim();

  const config =
    ADMIN_SUPPORTING_DOCUMENT_REVIEW_CONFIG[
      cleanType
    ];

  if (!cleanApplicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  if (!config) {
    throw new Error(
      'Invalid supporting document type.'
    );
  }

  if (
    cleanDecision !== 'approve' &&
    cleanDecision !== 'correction'
  ) {
    throw new Error(
      'Invalid document-review decision.'
    );
  }

  if (
    cleanDecision === 'correction' &&
    !cleanNotes
  ) {
    throw new Error(
      'Enter the correction required for this document.'
    );
  }

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    throw new Error(
      'Another review action is still being saved. Please try again in a moment.'
    );
  }

  let result = null;
  let queuedJob = null;

  try {
    const sheet =
      getResponseSheet_(
        SpreadsheetApp
          .getActiveSpreadsheet()
      );

    ensureAdminSupportingDocumentReviewColumns_(
      sheet
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

    const existingOverallStatus =
      String(
        getAdminRecordValue_(
          record,
          'Document Status'
        ) || ''
      )
        .trim()
        .toLowerCase();

    if (
      existingOverallStatus ===
      'complete'
    ) {
      throw new Error(
        'Supporting Documents have already been completed and are locked.'
      );
    }

    const fileId =
      String(
        getAdminRecordValue_(
          record,
          config.fileIdHeader
        ) || ''
      ).trim();

    const url =
      String(
        getAdminRecordValue_(
          record,
          config.urlHeader
        ) || ''
      ).trim();

    if (!fileId && !url) {
      throw new Error(
        config.label +
        ' has not been submitted.'
      );
    }

    const existingStatus =
      String(
        getAdminRecordValue_(
          record,
          config.reviewStatusHeader
        ) || ''
      )
        .trim()
        .toLowerCase();

    const existingNotes =
      String(
        getAdminRecordValue_(
          record,
          config.reviewNotesHeader
        ) || ''
      ).trim();

    const newStatus =
      cleanDecision === 'approve'
        ? 'Approved'
        : 'Correction Required';

    const decisionChanged =
      existingStatus !==
        newStatus.toLowerCase() ||
      (
        cleanDecision === 'correction' &&
        existingNotes !== cleanNotes
      );

    if (!decisionChanged) {
      const currentSummary =
        getAdminSupportingDocumentReviewSummary_(
          record
        );

      return {
        ok: true,
        unchanged: true,
        applicationId:
          cleanApplicationId,
        documentType:
          cleanType,
        documentStatus:
          newStatus,
        overallDocumentStatus:
          String(
            getAdminRecordValue_(
              record,
              'Document Status'
            ) || 'Pending Review'
          ),
        approvedCount:
          currentSummary.approvedCount,
        totalCount:
          currentSummary.totalCount,
        paymentStatus:
          String(
            getAdminRecordValue_(
              record,
              'Payment Status'
            ) || 'Not Available'
          ),
        message:
          config.label +
          ' already has this review decision.'
      };
    }

    const now =
      new Date();

    const reviewedBy =
      String(
        adminAccess.email || ''
      ).trim();

    /*
     * All workflow changes are first made in memory.
     * The application row is committed once at the end.
     */
    setAdminRecordMemory_(
      record,
      config.reviewStatusHeader,
      newStatus
    );

    setAdminRecordMemory_(
      record,
      config.reviewNotesHeader,
      cleanDecision === 'approve'
        ? ''
        : cleanNotes
    );

    setAdminRecordMemory_(
      record,
      config.reviewedAtHeader,
      now
    );

    setAdminRecordMemory_(
      record,
      config.reviewedByHeader,
      reviewedBy
    );

    setAdminRecordMemory_(
      record,
      'Document Correction Email Sent At',
      ''
    );

    setAdminRecordMemory_(
      record,
      'Document Correction Email Sent By',
      ''
    );

    const summary =
      getAdminSupportingDocumentReviewSummary_(
        record
      );

    let nextOverallStatus =
      deriveOverallSupportingDocumentStatus_(
        summary,
        existingOverallStatus
      );

    /*
     * Final completion is decided here, in the same transaction.
     * We no longer call updateAdminReviewStage_() and reread the Sheet.
     */
    if (
      summary.allUploaded &&
      summary.pendingCount === 0
    ) {
      nextOverallStatus =
        summary.allApproved
          ? 'Complete'
          : 'Correction Required';
    }

    setAdminRecordMemory_(
      record,
      'Document Status',
      nextOverallStatus
    );

    const correctionSummary =
      summary.items
        .filter(function(item) {
          return (
            item.status.toLowerCase() ===
            'correction required'
          );
        })
        .map(function(item) {
          return (
            item.label +
            ': ' +
            item.notes
          );
        })
        .join('\n');

    setAdminRecordMemory_(
      record,
      'Document Review Notes',
      correctionSummary
    );

    if (
      nextOverallStatus ===
      'Complete'
    ) {
      setAdminRecordMemory_(
        record,
        'Documents Verified At',
        now
      );

      setAdminRecordMemory_(
        record,
        'Documents Verified By',
        reviewedBy
      );
    } else {
      setAdminRecordMemory_(
        record,
        'Documents Verified At',
        ''
      );

      setAdminRecordMemory_(
        record,
        'Documents Verified By',
        ''
      );
    }

    /*
     * Any document change invalidates Final Verification.
     */
    setAdminRecordMemory_(
      record,
      'Verification Status',
      'Pending'
    );

    setAdminRecordMemory_(
      record,
      'Verification Notes',
      ''
    );

    setAdminRecordMemory_(
      record,
      'Verification Completed At',
      ''
    );

    setAdminRecordMemory_(
      record,
      'Verification Completed By',
      ''
    );

    let paymentUnlocked =
      false;

    /*
     * If a correction is required before payment has really started,
     * keep payment unavailable.
     */
    if (
      summary.correctionCount > 0
    ) {
      const paymentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      if (
        !paymentStatus ||
        paymentStatus === 'not available' ||
        paymentStatus === 'not submitted' ||
        paymentStatus === 'awaiting payment'
      ) {
        setAdminRecordMemory_(
          record,
          'Payment Status',
          'Not Available'
        );

        setAdminRecordMemory_(
          record,
          'Payment Invitation Sent At',
          ''
        );
      }
    }

    /*
     * PAYMENT UNLOCK:
     * all five documents Approved + Information Confirmed.
     */
    if (
      nextOverallStatus ===
      'Complete'
    ) {
      const informationStatus =
        String(
          getAdminRecordValue_(
            record,
            'Application Information Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      const paymentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      if (
        informationStatus === 'confirmed' &&
        (
          !paymentStatus ||
          paymentStatus === 'not available' ||
          paymentStatus === 'not submitted'
        )
      ) {
        setAdminRecordMemory_(
          record,
          'Payment Status',
          'Awaiting Payment'
        );

        paymentUnlocked =
          true;

        const invitationSentAt =
          String(
            getAdminRecordValue_(
              record,
              'Payment Invitation Sent At'
            ) || ''
          ).trim();

        if (!invitationSentAt) {
          const paymentPayload = {
            applicationId:
              cleanApplicationId,
            applicantName:
              String(
                getAdminRecordValue_(
                  record,
                  'Full Name'
                ) || 'Applicant'
              ).trim(),
            email:
              String(
                getAdminRecordValue_(
                  record,
                  'Email Address'
                ) || ''
              ).trim(),
            portalUrl:
              String(
                getAdminRecordValue_(
                  record,
                  'Portal URL'
                ) || ''
              ).trim()
          };

          queuedJob = {
            jobType:
              'PAYMENT_INVITATION_EMAIL',
            applicationId:
              cleanApplicationId,
            jobKey:
              makeAdminBackgroundJobKey_(
                cleanApplicationId,
                'PAYMENT_INVITATION_EMAIL',
                paymentPayload,
                now
              ),
            payload:
              paymentPayload
          };
        }
      }
    }

    /*
     * If the full review cycle ends with one or more corrections,
     * queue exactly one consolidated correction email.
     */
    if (
      summary.allUploaded &&
      summary.pendingCount === 0 &&
      summary.correctionCount > 0
    ) {
      const correctionPayload = {
        applicationId:
          cleanApplicationId,
        applicantName:
          String(
            getAdminRecordValue_(
              record,
              'Full Name'
            ) || 'Applicant'
          ).trim(),
        email:
          String(
            getAdminRecordValue_(
              record,
              'Email Address'
            ) || ''
          ).trim(),
        portalUrl:
          String(
            getAdminRecordValue_(
              record,
              'Portal URL'
            ) || ''
          ).trim(),
        corrections:
          summary.items
            .filter(function(item) {
              return (
                item.status.toLowerCase() ===
                'correction required'
              );
            })
            .map(function(item) {
              return {
                label:
                  item.label,
                note:
                  item.notes
              };
            }),
        approvedDocuments:
          summary.items
            .filter(function(item) {
              return (
                item.status.toLowerCase() ===
                'approved'
              );
            })
            .map(function(item) {
              return item.label;
            })
      };

      queuedJob = {
        jobType:
          'DOCUMENT_CORRECTION_EMAIL',
        applicationId:
          cleanApplicationId,
        jobKey:
          makeAdminBackgroundJobKey_(
            cleanApplicationId,
            'DOCUMENT_CORRECTION_EMAIL',
            correctionPayload,
            now
          ),
        payload:
          correctionPayload
      };
    }

    /*
     * ONE authoritative application-row write.
     */
    commitAdminRecord_(
      sheet,
      record
    );

    invalidateAdminCachesAfterWrite_(
      cleanApplicationId
    );

    result = {
      ok: true,
      applicationId:
        cleanApplicationId,
      documentType:
        cleanType,
      documentStatus:
        newStatus,
      overallDocumentStatus:
        nextOverallStatus,
      approvedCount:
        summary.approvedCount,
      correctionCount:
        summary.correctionCount,
      pendingCount:
        summary.pendingCount,
      totalCount:
        summary.totalCount,
      reviewCycleComplete:
        (
          summary.allUploaded &&
          summary.pendingCount === 0
        ),
      paymentStatus:
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || 'Not Available'
        ),
      paymentUnlocked:
        paymentUnlocked,
      notificationQueued:
        false,
      message:
        cleanDecision === 'approve'
          ? config.label + ' approved.'
          : 'Correction requested for ' +
            config.label + '.'
    };
  } finally {
    lock.releaseLock();
  }

  /*
   * Queue slow email AFTER the critical application write and AFTER
   * releasing the application lock.
   */
  if (queuedJob) {
    try {
      const queueResult =
        queueSystemJob_(
          queuedJob
        );

      result.notificationQueued =
        Boolean(
          queueResult &&
          queueResult.ok
        );

      if (
        queuedJob.jobType ===
        'PAYMENT_INVITATION_EMAIL'
      ) {
        result.message +=
          ' Supporting Documents are Complete. Payment is now available and the payment invitation has been queued.';
      } else {
        result.message +=
          ' All five documents have been reviewed. One consolidated correction email has been queued.';
      }
    } catch (queueError) {
      result.notificationQueued =
        false;

      result.notificationQueueError =
        queueError.message;

      result.message +=
        ' The review was saved, but the notification could not be queued. An administrator can retry the failed notification from System Jobs.';
    }
  }

  return result;
}


function sendSupportingDocumentCorrections(
  applicationId,
  sessionToken
) {
  const adminAccess =
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
      SpreadsheetApp.getActiveSpreadsheet()
    );

  ensureAdminSupportingDocumentReviewColumns_(
    sheet
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

  const summary =
    getAdminSupportingDocumentReviewSummary_(
      record
    );

  if (!summary.allUploaded) {
    throw new Error(
      'All five supporting documents must be submitted before completing this review cycle.'
    );
  }

  if (summary.pendingCount > 0) {
    throw new Error(
      'Review all five supporting documents before sending the correction email.'
    );
  }

  if (summary.correctionCount <= 0) {
    throw new Error(
      'There are no supporting-document corrections to send.'
    );
  }

  const corrections =
    summary.items
      .filter(function(item) {
        return (
          item.status.toLowerCase() ===
          'correction required'
        );
      })
      .map(function(item) {
        return {
          label: item.label,
          note: item.notes
        };
      });

  const approvedDocuments =
    summary.items
      .filter(function(item) {
        return (
          item.status.toLowerCase() ===
          'approved'
        );
      })
      .map(function(item) {
        return item.label;
      });

  if (
    typeof sendSupportingDocumentsCorrectionSummaryEmail_ !==
      'function'
  ) {
    throw new Error(
      'Email.gs is missing sendSupportingDocumentsCorrectionSummaryEmail_().'
    );
  }

  sendSupportingDocumentsCorrectionSummaryEmail_({
    applicationId:
      cleanApplicationId,
    applicantName:
      String(
        getAdminRecordValue_(
          record,
          'Full Name'
        ) || 'Applicant'
      ).trim(),
    email:
      String(
        getAdminRecordValue_(
          record,
          'Email Address'
        ) || ''
      ).trim(),
    portalUrl:
      String(
        getAdminRecordValue_(
          record,
          'Portal URL'
        ) || ''
      ).trim(),
    corrections:
      corrections,
    approvedDocuments:
      approvedDocuments
  });

  setAdminRecordValue_(
    sheet,
    record,
    'Document Correction Email Sent At',
    new Date()
  );

  setAdminRecordValue_(
    sheet,
    record,
    'Document Correction Email Sent By',
    String(
      adminAccess.email || ''
    ).trim()
  );

  SpreadsheetApp.flush();

  invalidateAdminDocumentReviewCache_(
    cleanApplicationId
  );

  if (
    typeof invalidateApplicantCacheByApplicationId_ ===
      'function'
  ) {
    invalidateApplicantCacheByApplicationId_(
      cleanApplicationId
    );
  }

  return {
    ok:
      true,
    applicationId:
      cleanApplicationId,
    correctionCount:
      corrections.length,
    approvedCount:
      approvedDocuments.length,
    message:
      'One consolidated supporting-document correction email was sent to the applicant.'
  };
}


function finaliseSupportingDocuments(
  applicationId,
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const sheet =
    getResponseSheet_(
      SpreadsheetApp
        .getActiveSpreadsheet()
    );

  ensureAdminSupportingDocumentReviewColumns_(
    sheet
  );

  const record =
    findAdminApplicationRecord_(
      sheet,
      applicationId
    );

  if (!record) {
    throw new Error(
      'Application record was not found.'
    );
  }

  const summary =
    getAdminSupportingDocumentReviewSummary_(
      record
    );

  if (!summary.allApproved) {
    throw new Error(
      'All five supporting documents must be individually approved before final approval.'
    );
  }

  return updateAdminReviewStage_({
    applicationId:
      applicationId,
    stage:
      'documents',
    approved:
      true,
    notes:
      ''
  }, sessionToken);
}


function approveAdminReviewStage(
  applicationId,
  stage,
  notes,
  sessionToken
) {
  return updateAdminReviewStage_({
    applicationId:
      applicationId,

    stage:
      stage,

    approved:
      true,

    notes:
      notes,
  }, sessionToken);
}

/**
 * Rejects an admin-review stage.
 */
function rejectAdminReviewStage(
  applicationId,
  stage,
  correctionItems,
  customDetails,
  sessionToken
) {
  const normalizedStage =
    String(
      stage || ''
    )
      .trim()
      .toLowerCase();

  const details =
    String(
      customDetails || ''
    ).trim();

  if (
    normalizedStage ===
      'information' ||
    normalizedStage ===
      'documents'
  ) {
    const items =
      Array.isArray(
        correctionItems
      )
        ? correctionItems
        : [];

    if (!items.length) {
      throw new Error(
        'Select at least one item that requires correction.'
      );
    }

    const targetConfig =
      ADMIN_REJECTION_TARGETS[
        normalizedStage
      ];

    const issueConfig =
      ADMIN_TARGET_ISSUES[
        normalizedStage
      ];

    const normalizedItems =
      items.map(
        function(item) {
          const targetCode =
            String(
              item &&
              item.targetCode ||
              ''
            )
              .trim()
              .toLowerCase();

          const issueCode =
            String(
              item &&
              item.issueCode ||
              ''
            )
              .trim()
              .toLowerCase();

          const targetLabel =
            targetConfig[
              targetCode
            ];

          const issueText =
            issueConfig[
              issueCode
            ];

          if (!targetLabel) {
            throw new Error(
              'One or more selected correction items are invalid.'
            );
          }

          if (!issueText) {
            throw new Error(
              'Select an issue for every correction item.'
            );
          }

          return {
            targetCode:
              targetCode,
            targetLabel:
              targetLabel,
            issueCode:
              issueCode,
            issueText:
              issueText,
          };
        }
      );

    const containsOther =
      normalizedItems.some(
        function(item) {
          return (
            item.targetCode ===
              'other' ||
            item.issueCode ===
              'other'
          );
        }
      );

    if (
      containsOther &&
      !details
    ) {
      throw new Error(
        'Enter the specific custom reason for the item marked Other.'
      );
    }

    const formattedReasons =
      formatTargetedCorrectionReasons_(
        normalizedItems,
        details
      );

    return updateAdminReviewStage_({
      applicationId:
        applicationId,
      stage:
        normalizedStage,
      approved:
        false,
      notes:
        formattedReasons.summary,
      rejectionItems:
        normalizedItems,
      rejectionDetails:
        details,
    }, sessionToken);
  }

  const reasonCode =
    Array.isArray(
      correctionItems
    )
      ? String(
          correctionItems[0] ||
          ''
        )
      : String(
          correctionItems || ''
        );

  const normalizedReasonCode =
    reasonCode
      .trim()
      .toLowerCase();

  const stageReasons =
    ADMIN_REJECTION_REASONS[
      normalizedStage
    ];

  if (!stageReasons) {
    throw new Error(
      'Invalid admin-review stage.'
    );
  }

  const reasonLabel =
    stageReasons[
      normalizedReasonCode
    ];

  if (!reasonLabel) {
    throw new Error(
      'Select a valid rejection or correction reason.'
    );
  }

  if (
    normalizedReasonCode ===
      'other' &&
    !details
  ) {
    throw new Error(
      'Enter the custom rejection or correction reason.'
    );
  }

  const finalReason =
    details
      ? (
        reasonLabel +
        ' — ' +
        details
      )
      : reasonLabel;

  return updateAdminReviewStage_({
    applicationId:
      applicationId,
    stage:
      normalizedStage,
    approved:
      false,
    notes:
      finalReason,
    rejectionReasonCode:
      normalizedReasonCode,
    rejectionDetails:
      details,
  }, sessionToken);
}


function formatTargetedCorrectionReasons_(
  items,
  customDetails
) {
  const details =
    String(
      customDetails || ''
    ).trim();

  const sentences =
    items.map(
      function(item) {
        let sentence =
          item.targetLabel +
          ' ' +
          item.issueText;

        if (
          (
            item.targetCode ===
              'other' ||
            item.issueCode ===
              'other'
          ) &&
          details
        ) {
          sentence +=
            ': ' +
            details;
        }

        sentence =
          sentence
            .replace(
              /\s+/g,
              ' '
            )
            .trim();

        if (
          !/[.!?]$/.test(
            sentence
          )
        ) {
          sentence +=
            '.';
        }

        return sentence;
      }
    );

  return {
    sentences:
      sentences,
    summary:
      sentences.join(
        '\n'
      ),
  };
}


function previewAdminReviewCorrectionEmail(
  applicationId,
  stage,
  correctionItems,
  customDetails,
  sessionToken
) {
  requireAdminAccess_(
    sessionToken
  );

  const normalizedApplicationId =
    String(
      applicationId || ''
    ).trim();

  const normalizedStage =
    String(
      stage || ''
    )
      .trim()
      .toLowerCase();

  if (!normalizedApplicationId) {
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
      normalizedApplicationId
    );

  let reasonText =
    '';

  if (
    normalizedStage ===
      'information' ||
    normalizedStage ===
      'documents'
  ) {
    const targetConfig =
      ADMIN_REJECTION_TARGETS[
        normalizedStage
      ];

    const issueConfig =
      ADMIN_TARGET_ISSUES[
        normalizedStage
      ];

    const items =
      Array.isArray(
        correctionItems
      )
        ? correctionItems
        : [];

    if (!items.length) {
      throw new Error(
        'Select at least one correction item.'
      );
    }

    const normalizedItems =
      items.map(
        function(item) {
          const targetCode =
            String(
              item.targetCode || ''
            ).trim();

          const issueCode =
            String(
              item.issueCode || ''
            ).trim();

          if (
            !targetConfig[
              targetCode
            ] ||
            !issueConfig[
              issueCode
            ]
          ) {
            throw new Error(
              'Complete the target and issue for every correction item.'
            );
          }

          return {
            targetCode:
              targetCode,
            targetLabel:
              targetConfig[
                targetCode
              ],
            issueCode:
              issueCode,
            issueText:
              issueConfig[
                issueCode
              ],
          };
        }
      );

    const containsOther =
      normalizedItems.some(
        function(item) {
          return (
            item.targetCode ===
              'other' ||
            item.issueCode ===
              'other'
          );
        }
      );

    const details =
      String(
        customDetails || ''
      ).trim();

    if (
      containsOther &&
      !details
    ) {
      throw new Error(
        'Enter the specific custom reason for the item marked Other.'
      );
    }

    reasonText =
      formatTargetedCorrectionReasons_(
        normalizedItems,
        details
      ).summary;
  } else {
    const reasonCode =
      Array.isArray(
        correctionItems
      )
        ? correctionItems[0]
        : correctionItems;

    const stageReasons =
      ADMIN_REJECTION_REASONS[
        normalizedStage
      ];

    const reasonLabel =
      stageReasons &&
      stageReasons[
        String(
          reasonCode || ''
        ).trim()
      ];

    if (!reasonLabel) {
      throw new Error(
        'Select a valid rejection reason.'
      );
    }

    reasonText =
      reasonLabel;

    const details =
      String(
        customDetails || ''
      ).trim();

    if (details) {
      reasonText +=
        ' — ' +
        details;
    }
  }

  return buildAdminReviewCorrectionEmail_({
    applicationId:
      normalizedApplicationId,
    stage:
      normalizedStage,
    reason:
      reasonText,
    applicantName:
      String(
        getAdminRecordValue_(
          record,
          'Full Name'
        ) || 'Applicant'
      ).trim(),
    email:
      String(
        getAdminRecordValue_(
          record,
          'Email Address'
        ) || ''
      ).trim(),
    portalUrl:
      String(
        getAdminRecordValue_(
          record,
          'Portal URL'
        ) || ''
      ).trim(),
  });
}


/**
 * Internal review update function.
 *
 * Rejection emails are sent only after the Sheet update lock has
 * been released. A mail failure does NOT undo the review decision;
 * it is returned to the admin as emailSent:false so it is visible.
 */
function updateAdminReviewStage_(
  options,
  sessionToken
) {
  const adminAccess =
    requireAdminAccess_(
      sessionToken
    );

  const applicationId =
    String(
      options.applicationId || ''
    ).trim();

  const stage =
    String(
      options.stage || ''
    )
      .trim()
      .toLowerCase();

  const notes =
    String(
      options.notes || ''
    ).trim();

  if (!applicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  const stageConfig =
    ADMIN_REVIEW_STAGES[
      stage
    ];

  if (!stageConfig) {
    throw new Error(
      'Invalid admin-review stage.'
    );
  }

  const lock =
    LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    throw new Error(
      'Another administrator action is still being saved. Please try again in a moment.'
    );
  }

  let result = null;
  let queuedJob = null;

  try {
    const sheet =
      getResponseSheet_(
        SpreadsheetApp
          .getActiveSpreadsheet()
      );

    if (stage === 'documents') {
      ensureAdminSupportingDocumentReviewColumns_(
        sheet
      );
    }

    const record =
      findAdminApplicationRecord_(
        sheet,
        applicationId
      );

    if (!record) {
      throw new Error(
        'Application record was not found.'
      );
    }

    if (
      stage === 'documents' &&
      options.approved
    ) {
      const documentSummary =
        getAdminSupportingDocumentReviewSummary_(
          record
        );

      if (!documentSummary.allApproved) {
        throw new Error(
          'All five supporting documents must be individually approved before the supporting-document package can be finalised.'
        );
      }
    }

    if (
      stage === 'verification' &&
      options.approved
    ) {
      const informationStatus =
        String(
          getAdminRecordValue_(
            record,
            'Application Information Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      const documentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Document Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      const paymentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      if (
        informationStatus !== 'confirmed' ||
        documentStatus !== 'complete' ||
        paymentStatus !== 'confirmed'
      ) {
        throw new Error(
          'Final Verification cannot be approved until Application Information is Confirmed, Supporting Documents are Complete, and Payment is Confirmed.'
        );
      }
    }

    const currentLicenceStatus =
      String(
        getAdminRecordValue_(
          record,
          'Licence Status'
        ) || ''
      )
        .trim()
        .toLowerCase();

    if (!options.approved) {
      const protectedLicenceStatuses = [
        'generated',
        'printed',
        'stamped',
        'uploaded',
        'released'
      ];

      if (
        protectedLicenceStatuses.includes(
          currentLicenceStatus
        )
      ) {
        throw new Error(
          'This review cannot be changed because licence processing has already started.'
        );
      }
    }

    const now =
      new Date();

    const status =
      options.approved
        ? stageConfig.approvedStatus
        : stageConfig.rejectedStatus;

    setAdminRecordMemory_(
      record,
      stageConfig.statusHeader,
      status
    );

    setAdminRecordMemory_(
      record,
      stageConfig.notesHeader,
      notes
    );

    setAdminRecordMemory_(
      record,
      stageConfig.reviewedAtHeader,
      now
    );

    setAdminRecordMemory_(
      record,
      stageConfig.reviewedByHeader,
      adminAccess.email
    );

    if (!options.approved) {
      setAdminRecordMemory_(
        record,
        'Licence Status',
        'Not Generated'
      );

      if (stage === 'information') {
        setAdminRecordMemory_(
          record,
          'Verification Status',
          'Pending'
        );

        setAdminRecordMemory_(
          record,
          'Verification Notes',
          ''
        );

        setAdminRecordMemory_(
          record,
          'Verification Completed At',
          ''
        );

        setAdminRecordMemory_(
          record,
          'Verification Completed By',
          ''
        );
      }

      const correctionPayload = {
        applicationId:
          applicationId,
        stage:
          stage,
        reason:
          notes,
        applicantName:
          String(
            getAdminRecordValue_(
              record,
              'Full Name'
            ) || 'Applicant'
          ).trim(),
        email:
          String(
            getAdminRecordValue_(
              record,
              'Email Address'
            ) || ''
          ).trim(),
        portalUrl:
          String(
            getAdminRecordValue_(
              record,
              'Portal URL'
            ) || ''
          ).trim()
      };

      queuedJob = {
        jobType:
          'ADMIN_REVIEW_CORRECTION_EMAIL',
        applicationId:
          applicationId,
        jobKey:
          makeAdminBackgroundJobKey_(
            applicationId,
            'ADMIN_REVIEW_CORRECTION_EMAIL_' +
              stage.toUpperCase(),
            correctionPayload,
            now
          ),
        payload:
          correctionPayload
      };
    }

    let paymentUnlocked =
      false;

    if (
      options.approved &&
      (
        stage === 'information' ||
        stage === 'documents'
      )
    ) {
      const informationStatus =
        String(
          getAdminRecordValue_(
            record,
            'Application Information Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      const documentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Document Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      const paymentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      if (
        informationStatus === 'confirmed' &&
        documentStatus === 'complete' &&
        (
          !paymentStatus ||
          paymentStatus === 'not available' ||
          paymentStatus === 'not submitted'
        )
      ) {
        setAdminRecordMemory_(
          record,
          'Payment Status',
          'Awaiting Payment'
        );

        paymentUnlocked =
          true;

        const invitationSentAt =
          String(
            getAdminRecordValue_(
              record,
              'Payment Invitation Sent At'
            ) || ''
          ).trim();

        if (!invitationSentAt) {
          const paymentPayload = {
            applicationId:
              applicationId,
            applicantName:
              String(
                getAdminRecordValue_(
                  record,
                  'Full Name'
                ) || 'Applicant'
              ).trim(),
            email:
              String(
                getAdminRecordValue_(
                  record,
                  'Email Address'
                ) || ''
              ).trim(),
            portalUrl:
              String(
                getAdminRecordValue_(
                  record,
                  'Portal URL'
                ) || ''
              ).trim()
          };

          queuedJob = {
            jobType:
              'PAYMENT_INVITATION_EMAIL',
            applicationId:
              applicationId,
            jobKey:
              makeAdminBackgroundJobKey_(
                applicationId,
                'PAYMENT_INVITATION_EMAIL',
                paymentPayload,
                now
              ),
            payload:
              paymentPayload
          };
        }
      }
    }

    if (
      !options.approved &&
      (
        stage === 'information' ||
        stage === 'documents'
      )
    ) {
      const paymentStatus =
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || ''
        )
          .trim()
          .toLowerCase();

      if (
        !paymentStatus ||
        paymentStatus === 'not available' ||
        paymentStatus === 'not submitted' ||
        paymentStatus === 'awaiting payment'
      ) {
        setAdminRecordMemory_(
          record,
          'Payment Status',
          'Not Available'
        );

        setAdminRecordMemory_(
          record,
          'Payment Invitation Sent At',
          ''
        );
      }
    }

    commitAdminRecord_(
      sheet,
      record
    );

    invalidateAdminCachesAfterWrite_(
      applicationId
    );

    result = {
      ok: true,
      applicationId:
        applicationId,
      stage:
        stage,
      status:
        status,
      informationStatus:
        String(
          getAdminRecordValue_(
            record,
            'Application Information Status'
          ) || 'Pending'
        ),
      documentStatus:
        String(
          getAdminRecordValue_(
            record,
            'Document Status'
          ) || 'Pending'
        ),
      paymentStatus:
        String(
          getAdminRecordValue_(
            record,
            'Payment Status'
          ) || 'Not Available'
        ),
      verificationStatus:
        String(
          getAdminRecordValue_(
            record,
            'Verification Status'
          ) || 'Pending'
        ),
      licenceStatus:
        String(
          getAdminRecordValue_(
            record,
            'Licence Status'
          ) || 'Not Generated'
        ),
      paymentUnlocked:
        paymentUnlocked,
      notificationQueued:
        false,
      message:
        getAdminStageSuccessMessage_(
          stage,
          status
        )
    };
  } finally {
    lock.releaseLock();
  }

  if (queuedJob) {
    try {
      const queueResult =
        queueSystemJob_(
          queuedJob
        );

      result.notificationQueued =
        Boolean(
          queueResult &&
          queueResult.ok
        );

      if (
        queuedJob.jobType ===
        'PAYMENT_INVITATION_EMAIL'
      ) {
        result.message +=
          ' Payment is now available and the payment invitation has been queued.';
      } else {
        result.message +=
          ' The applicant notification has been queued.';
      }
    } catch (queueError) {
      result.notificationQueued =
        false;

      result.notificationQueueError =
        queueError.message;

      result.message +=
        ' The review decision was saved, but the notification could not be queued.';
    }
  }

  return result;
}



/**
 * Mutates the already-loaded application row in memory.
 * No Sheet write happens here.
 */
function setAdminRecordMemory_(
  record,
  header,
  value
) {
  const column =
    record.headerMap[header];

  if (!column) {
    throw new Error(
      'Required column is missing: ' +
      header +
      '. Run setupAdminReviewColumns().'
    );
  }

  record.rowValues[
    column - 1
  ] = value;
}


/**
 * Commits the whole application row once.
 *
 * This is used only while the script lock is held, so the row
 * cannot be partially updated by another admin action.
 */
function commitAdminRecord_(
  sheet,
  record
) {
  sheet
    .getRange(
      record.rowNumber,
      1,
      1,
      record.rowValues.length
    )
    .setValues([
      record.rowValues
    ]);
}


/**
 * Central invalidation after an authoritative admin write.
 */
function invalidateAdminCachesAfterWrite_(
  applicationId
) {
  invalidateAdminDocumentReviewCache_(
    applicationId
  );

  if (
    typeof invalidateAdminApplicationCache_ ===
      'function'
  ) {
    invalidateAdminApplicationCache_(
      applicationId
    );
  }

  if (
    typeof invalidateApplicantCacheByApplicationId_ ===
      'function'
  ) {
    invalidateApplicantCacheByApplicationId_(
      applicationId
    );
  }
}


/**
 * Stable-enough unique key for one review outcome.
 * The timestamp is taken from the same authoritative action,
 * preventing duplicate queue creation from the same request.
 */
function makeAdminBackgroundJobKey_(
  applicationId,
  jobType,
  payload,
  actionDate
) {
  const signatureSource =
    JSON.stringify(
      payload || {}
    ) +
    '|' +
    String(
      actionDate instanceof Date
        ? actionDate.getTime()
        : Date.now()
    );

  const digest =
    Utilities
      .computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        signatureSource,
        Utilities.Charset.UTF_8
      )
      .slice(0, 8)
      .map(function(value) {
        const normalized =
          value < 0
            ? value + 256
            : value;

        return (
          '0' +
          normalized.toString(16)
        ).slice(-2);
      })
      .join('');

  return [
    String(applicationId || '').trim(),
    String(jobType || '').trim(),
    digest
  ].join(':');
}

/**
 * Finds an application row and creates
 * its header map.
 */
function findAdminApplicationRecord_(
  sheet,
  applicationId
) {
  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error(
      'No applications were found.'
    );
  }

  const headers = sheet
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
    getHeaderMap_(
      headers
    );

  const applicationIdColumn =
    headerMap[
      'Application ID'
    ];

  if (!applicationIdColumn) {
    throw new Error(
      'Required column is missing: Application ID'
    );
  }

  const applicationIds =
    sheet
      .getRange(
        2,
        applicationIdColumn,
        lastRow - 1,
        1
      )
      .getDisplayValues();

  let targetIndex =
    -1;

  /*
   * Search backwards so a resubmission/correction row is used
   * instead of the original row with the same Application ID.
   */
  for (
    let index =
      applicationIds.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        applicationIds[index][0] ||
        ''
      ).trim() ===
      applicationId
    ) {
      targetIndex =
        index;

      break;
    }
  }

  if (targetIndex === -1) {
    throw new Error(
      'Application not found: ' +
      applicationId
    );
  }

  const rowNumber =
    targetIndex + 2;

  const rowRange =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      );

  const rowValues =
    rowRange
      .getValues()[0];

  /*
   * Preserve existing formulas when a grouped full-row commit is used.
   * setValues() interprets strings beginning with "=" as formulas.
   */
  const rowFormulas =
    rowRange
      .getFormulas()[0];

  rowFormulas.forEach(
    function(formula, index) {
      if (formula) {
        rowValues[index] =
          formula;
      }
    }
  );

  return {
    rowNumber:
      rowNumber,
    headers:
      headers,
    headerMap:
      headerMap,
    rowValues:
      rowValues,
  };
}

/**
 * Writes one value into an application row.
 */
function setAdminRecordValue_(
  sheet,
  record,
  header,
  value
) {
  const column =
    record.headerMap[header];

  if (!column) {
    throw new Error(
      'Required column is missing: ' +
      header +
      '. Run setupAdminReviewColumns().'
    );
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

/**
 * Reads one value from the current record.
 */
function getAdminRecordValue_(
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

/**
 * Human-readable action result.
 */
function getAdminStageSuccessMessage_(
  stage,
  status
) {
  const labels = {
    payment:
      'Payment review',
    documents:
      'Document review',
    information:
      'Application information review',
    verification:
      'Final verification',
  };

  return (
    labels[stage] +
    ' updated to "' +
    status +
    '".'
  );
}
/**
 * Validates that an application has completed
 * every approval stage before licence generation.
 */
function validateLicenceGenerationReadiness_(
  record
) {
  const requirements = [
    {
      label:
        'Payment',
      header:
        'Payment Status',
      requiredValue:
        'Confirmed',
    },
    {
      label:
        'Documents',
      header:
        'Document Status',
      requiredValue:
        'Complete',
    },
    {
      label:
        'Application Information',
      header:
        'Application Information Status',
      requiredValue:
        'Confirmed',
    },
    {
      label:
        'Verification',
      header:
        'Verification Status',
      requiredValue:
        'Approved',
    },
  ];

  const missingRequirements =
    requirements
      .filter(requirement => {
        const currentValue =
          String(
            getAdminRecordValue_(
              record,
              requirement.header
            ) || ''
          ).trim();

        return (
          currentValue.toLowerCase() !==
          requirement.requiredValue
            .toLowerCase()
        );
      })
      .map(requirement => ({
        stage:
          requirement.label,
        required:
          requirement.requiredValue,
        current:
          String(
            getAdminRecordValue_(
              record,
              requirement.header
            ) || 'Pending'
          ),
      }));

  return {
    ready:
      missingRequirements.length === 0,

    missingRequirements:
      missingRequirements,
  };
}

/**
 * Final admin action.
 *
 * This function refuses PDF generation unless
 * all reviews have been approved.
 */
function approveAndGenerateLicence(
  applicationId,
  sessionToken
) {
  const adminAccess =
    requireAdminAccess_(
      sessionToken
    );

  const normalizedApplicationId =
    String(applicationId || '').trim();

  if (!normalizedApplicationId) {
    throw new Error(
      'Application ID is required.'
    );
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  try {
    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();

    const sheet =
      getResponseSheet_(spreadsheet);

    const record =
      findAdminApplicationRecord_(
        sheet,
        normalizedApplicationId
      );

    const readiness =
      validateLicenceGenerationReadiness_(
        record
      );

    if (!readiness.ready) {
      const problems =
        readiness
          .missingRequirements
          .map(item => {
            return (
              item.stage +
              ': expected "' +
              item.required +
              '", currently "' +
              item.current +
              '"'
            );
          })
          .join('; ');

      throw new Error(
        'Licence cannot be generated. ' +
        problems
      );
    }

    const existingLicenceStatus =
      String(
        getAdminRecordValue_(
          record,
          'Licence Status'
        ) || ''
      )
        .trim()
        .toLowerCase();

    const existingLicenceUrl =
      String(
        getAdminRecordValue_(
          record,
          'Licence PDF URL'
        ) || ''
      ).trim();

    if (
      (
        existingLicenceStatus === 'generated' ||
        existingLicenceStatus === 'released'
      ) &&
      existingLicenceUrl
    ) {
      return {
        ok: true,

        alreadyGenerated:
          true,

        licenceStatus:
          existingLicenceStatus === 'released'
            ? 'Released'
            : 'Generated',

        licenceNumber:
          getAdminRecordValue_(
            record,
            'Licence Number'
          ),

        documentReference:
          getAdminRecordValue_(
            record,
            'Document Reference'
          ),

        licenceDocumentUrl:
          getAdminRecordValue_(
            record,
            'Licence Document URL'
          ),

        licencePdfUrl:
          existingLicenceUrl,

        draftAvailable:
          true,

        message:
          existingLicenceStatus === 'released'
            ? 'This licence has already been released.'
            : 'The licence has already been generated.',
      };
    }

    /*
     * This calls the actual licence generator.
     * We will place that function in Licence.gs.
     */
    if (
      typeof generateLicenceForApplication_ !==
      'function'
    ) {
      throw new Error(
        'The application is fully approved, but the licence PDF generator has not yet been installed.'
      );
    }

    const result =
      generateLicenceForApplication_({
        applicationId:
          normalizedApplicationId,

        sheet:
          sheet,

        record:
          record,

        generatedBy:
          adminAccess.email,
      });

    invalidateAdminCachesAfterWrite_(
      normalizedApplicationId
    );

    return {
      ok: true,

      alreadyGenerated:
        Boolean(
          result.alreadyGenerated
        ),

      licenceStatus:
        result.licenceStatus ||
        'Generated',

      licenceNumber:
        result.licenceNumber,

      documentReference:
        result.documentReference,

      licenceDocumentUrl:
        result.licenceDocumentUrl,

      licencePdfUrl:
        result.licencePdfUrl,

      draftAvailable:
        Boolean(
          result.licencePdfUrl
        ),

      message:
        result.message ||
        'Unstamped licence generated successfully.',
    };
  } finally {
    lock.releaseLock();
  }
}
