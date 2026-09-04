const FORM_TITLE =
  'CRFFN LICENCE APPLICATION FORM';

const EMAIL_FIELD = 'Email Address';
const SETUP_SHEET_NAME = 'Setup';

/**
 * These status options are used for the Record Status dropdown.
 */
const STATUS_OPTIONS = [
  'New',
  'Resubmission',
  'Needs Review',
  'Duplicate',
  'Approved',
  'Rejected',
];

/**
 * These columns are added to the Google Form response sheet.
 */
const CUSTOM_COLUMNS = [
  'S/N',
  'Application ID',
  'Record Status',
  'Duplicate Flag',
  'Duplicate Reason',
  'Matched Existing ID',
  'Review Note',
  'Submission Count',

  // Applicant access
  'Secure Token',
  'Portal URL',

  // Application email delivery audit
  'Application Email Status',
  'Application Email Sent At',
  'Application Email Error',

  // Application workflow
  'Application Correction Fields JSON',
  'Payment Status',
  'Payment Reference',
  'Verification Status',
  'Licence Status',

  // Generated documents
  'Application PDF URL',
  'Receipt PDF URL',
  'Payment Proof File ID',
  'Payment Proof Uploaded At',
  'Payment Verified At',
  'Payment Verified By',
  'Payment Invitation Sent At',
  'Payment Proof Delete After',
  'Payment Proof Deletion Status',
  'Payment Proof Deleted At',
  'Licence PDF URL',
];

/**
 * Script property names.
 */
const PROP_KEYS = {
  FORM_ID: 'FORM_ID',
  APP_COUNTER: 'APP_COUNTER',
  SN_COUNTER: 'SN_COUNTER',
  WEB_APP_URL: 'WEB_APP_URL',
};

/**
 * Runs whenever the Google Sheet is opened.
 * Adds the custom Form Tools menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Form Tools')
    .addItem(
      'Create or Repair Form',
      'createFormAndSetup'
    )
    .addItem(
      'Show / Refresh Setup Tab',
      'refreshSetupTab'
    )
    .addSeparator()
    .addItem(
      'Install Submission Trigger',
      'installSubmissionTriggerForCurrentAccount'
    )
    .addItem(
      'Remove My Submission Trigger',
      'removeMySubmissionTrigger'
    )
    .addItem(
      'Reinstall Trigger After Transfer',
      'reinstallTriggerForCurrentAccount'
    )
    .addSeparator()
    .addItem(
      'Reset Test Application',
      'showResetTestApplicationPrompt'
    )
    .addSeparator()
    .addItem(
      'Go Live Reset & Set Numbers',
      'goLiveResetAndSetStartingNumbers'
    )
    .addToUi();
}

/**
 * Creates the Google Form if it does not exist.
 * Links the Form to the current Google Sheet.
 * Repairs missing custom columns.
 * Creates or refreshes the Setup sheet.
 */
/**
 * One-off repair for the Nationality / State of Origin layout.
 * Use this if the form already exists and you only want to remove the old
 * branching pages without running the entire Create / Repair process.
 */
function repairNationalityFormLayout() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const formUrl =
    spreadsheet.getFormUrl();

  if (!formUrl) {
    throw new Error(
      'This spreadsheet is not linked to a Google Form.'
    );
  }

  const form =
    FormApp.openByUrl(
      formUrl
    );

  ensureSimpleNationalityStateFields_(
    form
  );

  SpreadsheetApp.flush();

  Logger.log(
    'Nationality layout repaired successfully. The old routing pages were removed.'
  );
}


function createFormAndSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'Open this project from a Google Sheet through Extensions > Apps Script.'
    );
  }

  const docProps = PropertiesService.getDocumentProperties();
  const existingFormId = docProps.getProperty(PROP_KEYS.FORM_ID);

  let form;

  if (existingFormId) {
    try {
      form = FormApp.openById(existingFormId);
    } catch (error) {
      docProps.deleteProperty(PROP_KEYS.FORM_ID);
      form = null;
    }
  }

  if (!form) {
    form = FormApp.create(FORM_TITLE);

    buildForm_(form);

    form.setDescription(
      'Please complete this CRFFN licence application form. Ensure all information provided is accurate and complete. IMPORTANT: After submitting the form, check the email address you provided for your secure Practitioner Portal link. If it is not in your inbox, check your Spam or Junk folder.'
    );

    form.setConfirmationMessage(
      [
        'APPLICATION SUBMITTED SUCCESSFULLY',
        '',
        'Thank you. Your CRFFN Licence application has been received successfully.',
        '',
        'Your Application ID and secure Practitioner Portal link will be sent shortly to the email address provided in your application.',
        '',
        'IMPORTANT: Check the email address you provided for your secure Practitioner Portal link. Use that link to upload your required supporting documents and monitor your application.',
        '',
        'Payment is not required at this stage. You will be notified when payment becomes available after the required review and approval.',
        '',
        'If the Practitioner Portal email is not in your inbox, please check your Spam or Junk folder before contacting support.',
        '',
        'For assistance, contact licensing@crffn.gov.ng.'
      ].join('\n')
    );

    form.setShowLinkToRespondAgain(false);

    form.setDestination(
      FormApp.DestinationType.SPREADSHEET,
      ss.getId()
    );

    docProps.setProperty(PROP_KEYS.FORM_ID, form.getId());
  }
  /**
   * Adds required application fields that are missing
   * from an existing Google Form.
   */
  
  updateExistingCrffnFormFlow_(form);
  removeExpiryDateFromApplicationForm_(form);
  ensureRequiredFormFields_(form);
  ensureApplicantEmailNotice_(form);
  ensureSimpleNationalityStateFields_(form);

  SpreadsheetApp.flush();
  Utilities.sleep(1500);

  const responseSheet = getResponseSheet_(ss);

  ensureCustomColumns_(responseSheet);
  writeSetupTab_(ss, form);

  Logger.log('Form edit URL: ' + form.getEditUrl());
  Logger.log('Form live URL: ' + form.getPublishedUrl());
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
}
/**
 * Repairs the EXISTING Google Form to the current CRFFN flow.
 *
 * Payment must not be requested at form-submission stage.
 */
function updateExistingCrffnFormFlow_(
  form
) {
  form.setTitle(
    'CRFFN LICENCE APPLICATION FORM'
  );

  form.setDescription(
    'Please complete this CRFFN licence application form. IMPORTANT: After submitting the form, check the email address you provided for your secure Practitioner Portal link. If it is not in your inbox, check your Spam or Junk folder. Use the Practitioner Portal to upload all required supporting documents and monitor your application. Payment becomes available only after CRFFN approves your Application Information and Supporting Documents.'
  );

  form.setConfirmationMessage(
    [
      'APPLICATION SUBMITTED SUCCESSFULLY',
      '',
      'Thank you. Your CRFFN Licence application has been received successfully.',
      '',
      'Your Application ID and secure Practitioner Portal link will be sent shortly to the email address provided in your application.',
      '',
      'IMPORTANT: Check the email address you provided for your secure Practitioner Portal link. Use that link to upload your required supporting documents and monitor your application.',
      '',
      'Payment is not required at this stage. You will be notified when payment becomes available after the required review and approval.',
      '',
      'If the Practitioner Portal email is not in your inbox, please check your Spam or Junk folder before contacting support.',
      '',
      'For assistance, contact licensing@crffn.gov.ng.'
    ].join('\n')
  );

  /*
   * Hide Google's "Submit another response" link on the
   * confirmation page. This does not prevent CRFFN correction
   * links / prefilled resubmissions from being used later.
   */
  form.setShowLinkToRespondAgain(false);

  const items =
    form.getItems();

  for (
    let index =
      items.length - 1;
    index >= 0;
    index--
  ) {
    const item =
      items[index];

    const title =
      String(
        item.getTitle() || ''
      ).trim();

    if (
      title ===
      '5. PAYMENT INSTRUCTION'
    ) {
      form.deleteItem(
        item
      );
    }
  }

  form.getItems().forEach(
    function(item) {
      if (
        item.getType() !==
        FormApp.ItemType.SECTION_HEADER
      ) {
        return;
      }

      const title =
        String(
          item.getTitle() || ''
        ).trim();

      if (
        title ===
        '3. SUPPORTING DOCUMENTS'
      ) {
        item
          .asSectionHeaderItem()
          .setHelpText(
            [
              'After submitting this form, upload the following documents through your Practitioner Portal:',
              '',
              '• CAC Document',
              '• Passport Photograph of Director or Company Owner',
              '• Educational Certificates',
              '• Proof of Experience (CV of CEO)',
              '• Valid Means of Identification',
              '',
              'Payment is not required at this stage. CRFFN will notify you when payment becomes available.'
            ].join('\n')
          );
      }
    }
  );
}


/**
 * Removes the old Expiry Date question from the application form.
 *
 * Expiry dates are not universally applicable to accepted identification
 * documents (for example, NIN does not have an expiry date), so the
 * application form must not request or require this field.
 */
function removeExpiryDateFromApplicationForm_(form) {
  const items =
    form.getItems();

  for (
    let index =
      items.length - 1;
    index >= 0;
    index--
  ) {
    const title =
      String(
        items[index].getTitle() || ''
      ).trim();

    if (
      title ===
      'Expiry Date'
    ) {
      form.deleteItem(
        index
      );
    }
  }
}


// Ensure require fields
function ensureRequiredFormFields_(form) {
    const existingTitles = form
      .getItems()
      .map(item => String(item.getTitle() || '').trim());

    if (!existingTitles.includes('Company RC Number')) {
      form
        .addTextItem()
        .setTitle('Company RC Number')
        .setRequired(true);
    }

    if (
      !existingTitles.includes(
        'CRFFN Membership Number'
      )
    ) {
      form
        .addTextItem()
        .setTitle('CRFFN Membership Number')
        .setRequired(true);
    }
  }
/**
 * Creates all questions and sections in the application form.
 */
function buildForm_(form) {
  /*
   * Google Forms section-header titles are rendered prominently.
   * This is used to draw attention to the email/Practitioner Portal instruction.
   */
  form
    .addSectionHeaderItem()
    .setTitle('IMPORTANT — CHECK YOUR EMAIL AFTER SUBMISSION')
    .setHelpText(
      'After submitting this form, check the email address you provided for your secure Practitioner Portal link. If the email is not in your inbox, check your Spam or Junk folder.'
    );

  form
    .addSectionHeaderItem()
    .setTitle('1. APPLICANT INFORMATION');

  form
    .addTextItem()
    .setTitle('Full Name')
    .setRequired(true);

  form
    .addMultipleChoiceItem()
    .setTitle('Gender')
    .setChoiceValues(['Male', 'Female'])
    .setRequired(true);

  form
    .addDateItem()
    .setTitle('Date of Birth')
    .setRequired(true);

  /*
   * Nationality is a country dropdown with Nigeria listed first.
   * State of Origin stays directly underneath on the same page.
   * Google Forms cannot make a field conditionally required without
   * section branching, so State of Origin is optional and clearly marked
   * for Nigerian applicants only.
   */
  form
    .addListItem()
    .setTitle('Nationality')
    .setChoiceValues(getCountries_())
    .setRequired(true);

  form
    .addListItem()
    .setTitle('State of Origin — Nigerian Applicants Only')
    .setHelpText(
      'If your nationality is Nigeria, select your State of Origin. If you are not Nigerian, leave this field blank.'
    )
    .setChoiceValues(getNigerianStates_())
    .setRequired(false);


  form
    .addParagraphTextItem()
    .setTitle('Residential Address')
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('Phone Number')
    .setRequired(true);

  const emailItem = form
    .addTextItem()
    .setTitle(EMAIL_FIELD)
    .setRequired(true);

  emailItem.setValidation(
    FormApp
      .createTextValidation()
      .requireTextIsEmail()
      .build()
  );

  form
    .addMultipleChoiceItem()
    .setTitle('Means of Identification')
    .setChoiceValues([
      'National ID',
      'International Passport',
      'Driver’s License',
      'Voter’s Card',
    ])
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('ID Number')
    .setRequired(true);

  form
    .addSectionHeaderItem()
    .setTitle('2. PROFESSIONAL DETAILS');

  form
    .addCheckboxItem()
    .setTitle('Area of Practice')
    .setChoiceValues([
      'Clearing & Forwarding',
      'Haulage/Transportation Services',
      'Warehousing',
      'Courier Services',
      'Cold Chain',
      'Other',
    ])
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('If Other, please specify')
    .setRequired(false);

  form
    .addTextItem()
    .setTitle('Company Name')
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('Company RC Number')
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('Company TIN')
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('CRFFN Membership Number')
    .setRequired(true);

  form
    .addParagraphTextItem()
    .setTitle('Company Address')
    .setRequired(true);

  form
    .addTextItem()
    .setTitle('Position Held')
    .setRequired(true);

  form
    .addSectionHeaderItem()
    .setTitle('3. SUPPORTING DOCUMENTS')
    .setHelpText(
      [
        'After submitting this form, upload the following documents through your Practitioner Portal:',
        '',
        '• CAC Document',
        '• Passport Photograph of Director or Company Owner',
        '• Educational Certificates',
        '• Proof of Experience (CV of CEO)',
        '• Valid Means of Identification',
        '',
        'Documents are not uploaded through this Google Form.',
        'Payment is not required at this stage. CRFFN will notify you when payment becomes available.'
      ].join('\n')
    );

  form
    .addCheckboxItem()
    .setTitle('Supporting Documents Confirmation')
    .setChoiceValues([
      'I confirm that I have all applicable supporting documents listed above.',
    ])
    .setRequired(true);

  form
    .addCheckboxItem()
    .setTitle('4. DECLARATION')
    .setHelpText(
      'I hereby declare that the information provided in this application is true and correct. I understand that any false declaration may lead to disqualification or withdrawal of the licence.'
    )
    .setChoiceValues(['I Agree'])
    .setRequired(true);
}


/**
 * Countries for the Nationality dropdown.
 * Nigeria is deliberately first; the remaining countries are alphabetical.
 */
function getCountries_() {
  return [
    'Nigeria',
    'Afghanistan',
    'Albania',
    'Algeria',
    'Andorra',
    'Angola',
    'Antigua and Barbuda',
    'Argentina',
    'Armenia',
    'Australia',
    'Austria',
    'Azerbaijan',
    'Bahamas',
    'Bahrain',
    'Bangladesh',
    'Barbados',
    'Belarus',
    'Belgium',
    'Belize',
    'Benin',
    'Bhutan',
    'Bolivia',
    'Bosnia and Herzegovina',
    'Botswana',
    'Brazil',
    'Brunei',
    'Bulgaria',
    'Burkina Faso',
    'Burundi',
    'Cabo Verde',
    'Cambodia',
    'Cameroon',
    'Canada',
    'Central African Republic',
    'Chad',
    'Chile',
    'China',
    'Colombia',
    'Comoros',
    'Congo, Democratic Republic of the',
    'Congo, Republic of the',
    'Costa Rica',
    "Cote d'Ivoire",
    'Croatia',
    'Cuba',
    'Cyprus',
    'Czechia',
    'Denmark',
    'Djibouti',
    'Dominica',
    'Dominican Republic',
    'Ecuador',
    'Egypt',
    'El Salvador',
    'Equatorial Guinea',
    'Eritrea',
    'Estonia',
    'Eswatini',
    'Ethiopia',
    'Fiji',
    'Finland',
    'France',
    'Gabon',
    'Gambia',
    'Georgia',
    'Germany',
    'Ghana',
    'Greece',
    'Grenada',
    'Guatemala',
    'Guinea',
    'Guinea-Bissau',
    'Guyana',
    'Haiti',
    'Honduras',
    'Hungary',
    'Iceland',
    'India',
    'Indonesia',
    'Iran',
    'Iraq',
    'Ireland',
    'Israel',
    'Italy',
    'Jamaica',
    'Japan',
    'Jordan',
    'Kazakhstan',
    'Kenya',
    'Kiribati',
    'Korea, North',
    'Korea, South',
    'Kuwait',
    'Kyrgyzstan',
    'Laos',
    'Latvia',
    'Lebanon',
    'Lesotho',
    'Liberia',
    'Libya',
    'Liechtenstein',
    'Lithuania',
    'Luxembourg',
    'Madagascar',
    'Malawi',
    'Malaysia',
    'Maldives',
    'Mali',
    'Malta',
    'Marshall Islands',
    'Mauritania',
    'Mauritius',
    'Mexico',
    'Micronesia',
    'Moldova',
    'Monaco',
    'Mongolia',
    'Montenegro',
    'Morocco',
    'Mozambique',
    'Myanmar',
    'Namibia',
    'Nauru',
    'Nepal',
    'Netherlands',
    'New Zealand',
    'Nicaragua',
    'Niger',
    'North Macedonia',
    'Norway',
    'Oman',
    'Pakistan',
    'Palau',
    'Palestine',
    'Panama',
    'Papua New Guinea',
    'Paraguay',
    'Peru',
    'Philippines',
    'Poland',
    'Portugal',
    'Qatar',
    'Romania',
    'Russia',
    'Rwanda',
    'Saint Kitts and Nevis',
    'Saint Lucia',
    'Saint Vincent and the Grenadines',
    'Samoa',
    'San Marino',
    'Sao Tome and Principe',
    'Saudi Arabia',
    'Senegal',
    'Serbia',
    'Seychelles',
    'Sierra Leone',
    'Singapore',
    'Slovakia',
    'Slovenia',
    'Solomon Islands',
    'Somalia',
    'South Africa',
    'South Sudan',
    'Spain',
    'Sri Lanka',
    'Sudan',
    'Suriname',
    'Sweden',
    'Switzerland',
    'Syria',
    'Taiwan',
    'Tajikistan',
    'Tanzania',
    'Thailand',
    'Timor-Leste',
    'Togo',
    'Tonga',
    'Trinidad and Tobago',
    'Tunisia',
    'Turkey',
    'Turkmenistan',
    'Tuvalu',
    'Uganda',
    'Ukraine',
    'United Arab Emirates',
    'United Kingdom',
    'United States',
    'Uruguay',
    'Uzbekistan',
    'Vanuatu',
    'Vatican City',
    'Venezuela',
    'Vietnam',
    'Yemen',
    'Zambia',
    'Zimbabwe'
  ];
}


/**
 * Nigerian states used by the optional State of Origin dropdown.
 */
function getNigerianStates_() {
  return [
    'Abia',
    'Adamawa',
    'Akwa Ibom',
    'Anambra',
    'Bauchi',
    'Bayelsa',
    'Benue',
    'Borno',
    'Cross River',
    'Delta',
    'Ebonyi',
    'Edo',
    'Ekiti',
    'Enugu',
    'Gombe',
    'Imo',
    'Jigawa',
    'Kaduna',
    'Kano',
    'Katsina',
    'Kebbi',
    'Kogi',
    'Kwara',
    'Lagos',
    'Nasarawa',
    'Niger',
    'Ogun',
    'Ondo',
    'Osun',
    'Oyo',
    'Plateau',
    'Rivers',
    'Sokoto',
    'Taraba',
    'Yobe',
    'Zamfara',
    'Federal Capital Territory (FCT)'
  ];
}


/**
 * Moves a form item safely using the integer-index overload of Form.moveItem().
 *
 * Apps Script can reject concrete item subclasses such as SectionHeaderItem
 * when they are passed directly to Form.moveItem(item, toIndex), even though
 * the Forms reference also documents an Item overload. Resolving the current
 * item index and using moveItem(fromIndex, toIndex) avoids that runtime type
 * mismatch.
 */
function moveFormItemToIndex_(form, item, toIndex) {
  const itemId =
    item.getId();

  const items =
    form.getItems();

  const fromIndex =
    items.findIndex(
      function(candidate) {
        return (
          candidate.getId() ===
          itemId
        );
      }
    );

  if (
    fromIndex < 0
  ) {
    throw new Error(
      'Could not locate a Google Form item before moving it.'
    );
  }

  const maxIndex =
    items.length - 1;

  const safeToIndex =
    Math.max(
      0,
      Math.min(
        Number(toIndex),
        maxIndex
      )
    );

  if (
    fromIndex ===
    safeToIndex
  ) {
    return;
  }

  form.moveItem(
    fromIndex,
    safeToIndex
  );
}


/**
 * Adds/repairs the prominent Practitioner Portal email notice
 * on an existing Google Form.
 */
function ensureApplicantEmailNotice_(form) {
  const noticeTitle =
    'IMPORTANT — CHECK YOUR EMAIL AFTER SUBMISSION';

  const existingNotice =
    form
      .getItems()
      .find(function(item) {
        return (
          String(
            item.getTitle() || ''
          ).trim() ===
          noticeTitle
        );
      });

  if (existingNotice) {
    if (
      existingNotice.getType() ===
      FormApp.ItemType.SECTION_HEADER
    ) {
      existingNotice
        .asSectionHeaderItem()
        .setHelpText(
          'After submitting this form, check the email address you provided for your secure Practitioner Portal link. If the email is not in your inbox, check your Spam or Junk folder.'
        );
    }

    return;
  }

  const notice =
    form
      .addSectionHeaderItem()
      .setTitle(
        noticeTitle
      )
      .setHelpText(
        'After submitting this form, check the email address you provided for your secure Practitioner Portal link. If the email is not in your inbox, check your Spam or Junk folder.'
      );

  const applicantSectionIndex =
    form
      .getItems()
      .findIndex(function(item) {
        return (
          String(
            item.getTitle() || ''
          ).trim() ===
          '1. APPLICANT INFORMATION'
        );
      });

  if (
    applicantSectionIndex >= 0
  ) {
    moveFormItemToIndex_(
      form,
      notice,
      applicantSectionIndex
    );
  }
}


/**
 * Keeps Nationality and State of Origin on the same form page.
 *
 * This deliberately removes the temporary branching sections that were
 * introduced earlier. It reuses the existing questions instead of deleting
 * and recreating them, so the linked response-sheet columns remain stable.
 */
function ensureSimpleNationalityStateFields_(form) {
  const nationalityTitle =
    'Nationality';

  const oldStateTitle =
    'State of Origin';

  const stateTitle =
    'State of Origin — Nigerian Applicants Only';

  const generatedPageTitles = [
    'NIGERIAN APPLICANTS — STATE OF ORIGIN',
    'CONTACT AND IDENTIFICATION DETAILS'
  ];

  let items =
    form.getItems();

  let nationality =
    items.find(
      function(item) {
        return (
          String(
            item.getTitle() || ''
          ).trim() ===
          nationalityTitle
        );
      }
    );

  if (!nationality) {
    throw new Error(
      'Nationality question was not found in the Google Form.'
    );
  }

  if (
    nationality.getType() !==
    FormApp.ItemType.LIST
  ) {
    throw new Error(
      'Nationality exists but is not currently a dropdown.'
    );
  }

  /*
   * IMPORTANT:
   * Remove the old page-navigation choices FIRST.
   *
   * The earlier version tried to delete the page-break sections while
   * Nationality still contained choices pointing to those sections. Google
   * Forms rejects that state and throws "Invalid data updating form."
   *
   * Replacing the choices with ordinary values first removes all navigation
   * references safely.
   */
  nationality
    .asListItem()
    .setChoiceValues(
      getCountries_()
    )
    .setRequired(
      true
    );

  /*
   * Now that Nationality no longer points to the temporary routing sections,
   * remove the generated page breaks in dependency order.
   *
   * IMPORTANT:
   * The Nigerian page previously pointed to the Contact page. Deleting the
   * Contact page first makes the Form temporarily invalid. Therefore we must
   * delete the Nigerian page first, then the Contact page.
   *
   * We also use the numeric-index overload of deleteItem() rather than
   * passing a PageBreakItem object, which is more reliable in Apps Script.
   */
  generatedPageTitles.forEach(
    function(pageTitle) {
      const currentItems =
        form.getItems();

      const pageIndex =
        currentItems.findIndex(
          function(item) {
            return (
              item.getType() ===
                FormApp.ItemType.PAGE_BREAK &&
              String(
                item.getTitle() || ''
              ).trim() ===
                pageTitle
            );
          }
        );

      if (
        pageIndex >= 0
      ) {
        form.deleteItem(
          pageIndex
        );
      }
    }
  );

  items =
    form.getItems();

  let state =
    items.find(
      function(item) {
        const title =
          String(
            item.getTitle() || ''
          ).trim();

        return (
          title === oldStateTitle ||
          title === stateTitle
        );
      }
    );

  if (!state) {
    state =
      form
        .addListItem()
        .setTitle(
          stateTitle
        )
        .setHelpText(
          'If your nationality is Nigeria, select your State of Origin. If you are not Nigerian, leave this field blank.'
        )
        .setChoiceValues(
          getNigerianStates_()
        )
        .setRequired(
          false
        );
  } else {
    if (
      state.getType() !==
      FormApp.ItemType.LIST
    ) {
      throw new Error(
        'State of Origin exists but is not currently a dropdown.'
      );
    }

    state
      .asListItem()
      .setTitle(
        stateTitle
      )
      .setHelpText(
        'If your nationality is Nigeria, select your State of Origin. If you are not Nigerian, leave this field blank.'
      )
      .setChoiceValues(
        getNigerianStates_()
      )
      .setRequired(
        false
      );
  }

  /*
   * Keep State of Origin immediately under Nationality on the same page.
   */
  items =
    form.getItems();

  const nationalityIndex =
    items.findIndex(
      function(item) {
        return (
          item.getId() ===
          nationality.getId()
        );
      }
    );

  if (
    nationalityIndex < 0
  ) {
    throw new Error(
      'Could not locate the Nationality question while arranging the form.'
    );
  }

  moveFormItemToIndex_(
    form,
    state,
    nationalityIndex + 1
  );
}


/**
 * Read-only diagnostic for duplicate response-sheet headers.
 * This function does not modify the form or spreadsheet.
 */
function diagnoseDuplicateResponseHeaders() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    getResponseSheet_(
      ss
    );

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
      : [];

  const positions =
    {};

  headers.forEach(
    function(header, index) {
      const name =
        String(
          header || ''
        ).trim();

      if (!name) {
        return;
      }

      if (!positions[name]) {
        positions[name] =
          [];
      }

      positions[name].push(
        index + 1
      );
    }
  );

  const duplicateHeaders =
    Object.keys(
      positions
    )
      .filter(
        function(name) {
          return (
            positions[name].length > 1
          );
        }
      )
      .map(
        function(name) {
          return {
            columnName: name,
            columns: positions[name]
          };
        }
      );

  const result = {
    sheet: sheet.getName(),
    duplicateHeaders:
      duplicateHeaders
  };

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

  return result;
}


/**
 * Removes duplicate Nationality and State of Origin columns from the
 * linked Google Form response sheet.
 *
 * The rightmost duplicate is treated as the active current Form column.
 * Any value in an older duplicate is first copied into the active column
 * when the active cell is blank. The old duplicate column is then deleted.
 *
 * This is safe to rerun: if there are no duplicate headers, it does nothing.
 */
function repairNationalityResponseColumns() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName('Form responses 1') ||
    ss.getSheetByName('Form_Responses');

  if (!sheet) {
    throw new Error(
      'Form response sheet could not be found. Expected "Form responses 1" or "Form_Responses".'
    );
  }

  [
    'Nationality',
    'State of Origin'
  ].forEach(
    function(fieldName) {
      repairDuplicateResponseField_(
        sheet,
        fieldName
      );
    }
  );

  SpreadsheetApp.flush();

  Logger.log(
    'Nationality and State of Origin response columns are clean. Duplicate old columns were removed.'
  );
}


/**
 * Repairs one duplicated response-sheet field.
 */
function repairDuplicateResponseField_(
  sheet,
  fieldName
) {
  const lastColumn =
    sheet.getLastColumn();

  const lastRow =
    sheet.getLastRow();

  if (lastColumn < 1) {
    return;
  }

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const matchingColumns = [];

  headers.forEach(
    function(header, index) {
      if (
        String(header || '').trim() ===
        fieldName
      ) {
        matchingColumns.push(
          index + 1
        );
      }
    }
  );

  if (
    matchingColumns.length <= 1
  ) {
    return;
  }

  const activeColumn =
    matchingColumns[
      matchingColumns.length - 1
    ];

  const oldColumns =
    matchingColumns.slice(
      0,
      -1
    );

  if (
    lastRow > 1
  ) {
    const activeValues =
      sheet
        .getRange(
          2,
          activeColumn,
          lastRow - 1,
          1
        )
        .getValues();

    oldColumns.forEach(
      function(oldColumn) {
        const oldValues =
          sheet
            .getRange(
              2,
              oldColumn,
              lastRow - 1,
              1
            )
            .getValues();

        for (
          let rowIndex = 0;
          rowIndex < activeValues.length;
          rowIndex++
        ) {
          const activeValue =
            activeValues[rowIndex][0];

          const oldValue =
            oldValues[rowIndex][0];

          const activeBlank =
            activeValue === '' ||
            activeValue === null;

          const oldHasValue =
            oldValue !== '' &&
            oldValue !== null;

          if (
            activeBlank &&
            oldHasValue
          ) {
            activeValues[rowIndex][0] =
              oldValue;
          }
        }
      }
    );

    sheet
      .getRange(
        2,
        activeColumn,
        activeValues.length,
        1
      )
      .setValues(
        activeValues
      );
  }

  /*
   * Delete from right to left so shifting column numbers do not affect
   * the remaining deletions.
   */
  oldColumns
    .sort(
      function(a, b) {
        return b - a;
      }
    )
    .forEach(
      function(columnNumber) {
        sheet.deleteColumn(
          columnNumber
        );
      }
    );

  Logger.log(
    fieldName +
    ': removed ' +
    oldColumns.length +
    ' old duplicate column(s).'
  );
}


/**
 * Installs an on-form-submit trigger for the current Google account.
 */
function installSubmissionTriggerForCurrentAccount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Could not open the active spreadsheet.');
  }

  removeMySubmissionTriggerInternal_();

  ScriptApp
    .newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  refreshSetupTab();

  SpreadsheetApp.getUi().alert(
    'Submission trigger installed successfully.'
  );
}

/**
 * Removes the current account's on-form-submit trigger.
 */
function removeMySubmissionTrigger() {
  removeMySubmissionTriggerInternal_();
  refreshSetupTab();

  SpreadsheetApp.getUi().alert(
    'Your submission trigger has been removed.'
  );
}

/**
 * Reinstalls the trigger after ownership transfer.
 */
function reinstallTriggerForCurrentAccount() {
  removeMySubmissionTriggerInternal_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Could not open the active spreadsheet.');
  }

  ScriptApp
    .newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  refreshSetupTab();

  SpreadsheetApp.getUi().alert(
    'Submission trigger reinstalled successfully.'
  );
}

/**
 * Deletes submission triggers belonging to the current account.
 */
function removeMySubmissionTriggerInternal_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Refreshes the Setup sheet.
 */
function refreshSetupTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('Could not open the active spreadsheet.');
  }

  const formId = PropertiesService
    .getDocumentProperties()
    .getProperty(PROP_KEYS.FORM_ID);

  if (!formId) {
    throw new Error(
      'No form was found. Use Form Tools > Create or Repair Form first.'
    );
  }

  const form = FormApp.openById(formId);

  writeSetupTab_(ss, form);
}

/**
 * Creates or refreshes the Setup sheet.
 */
function writeSetupTab_(ss, form) {
  const sheet =
    getOrCreateSetupSheet_(ss);

  sheet.clear();

  const hasTrigger =
    hasCurrentAccountSubmissionTrigger_();

  const webAppUrl =
    APP_CONFIG.PRODUCTION_WEB_APP_URL;

  const properties =
    PropertiesService
      .getDocumentProperties();

  const currentApplicationCounter =
    Number(
      properties.getProperty(
        PROP_KEYS.APP_COUNTER
      ) || '0'
    );

  const currentSerialCounter =
    Number(
      properties.getProperty(
        PROP_KEYS.SN_COUNTER
      ) || '0'
    );

  const nextApplicationNumber =
    currentApplicationCounter + 1;

  const nextSerialNumber =
    currentSerialCounter + 1;

  const rows = [
    ['Item', 'Value'],

    [
      'Form Edit URL',
      form.getEditUrl()
    ],

    [
      'Form Live URL',
      form.getPublishedUrl()
    ],

    [
      'Spreadsheet URL',
      ss.getUrl()
    ],

    [
      'Current Account Trigger Installed',
      hasTrigger
        ? 'Yes'
        : 'No'
    ],

    [
      'Web App URL',
      webAppUrl ||
        'Not deployed yet'
    ],

    ['', ''],

    [
      'LIVE NUMBERING',
      ''
    ],

    [
      'Current Application Counter',
      currentApplicationCounter
    ],

    [
      'Next Application Number',
      'APP-' +
        String(
          nextApplicationNumber
        ).padStart(
          4,
          '0'
        )
    ],

    [
      'Current S/N Counter',
      currentSerialCounter
    ],

    [
      'Next S/N',
      nextSerialNumber
    ],

    ['', ''],

    [
      'How to use',
      ''
    ],

    [
      '1',
      'Use Form Tools > Create or Repair Form to create the form or reconnect to the existing one.'
    ],

    [
      '2',
      'Use Form Tools > Install Submission Trigger to activate automatic Application ID and secure token generation.'
    ],

    [
      '3',
      'Use the Form Live URL above to submit a test application.'
    ],

    [
      '4',
      'The Portal URL will show Pending web app deployment until the project is deployed as a web app.'
    ],

    [
      '5',
      'If ownership is transferred, the old owner should remove their trigger first.'
    ],

    [
      '6',
      'After transfer, the new owner should reload the Sheet and click Reinstall Trigger After Transfer.'
    ],

    [
      '7',
      'Custom columns may be moved, but their headings should not be renamed.'
    ],

    [
      '8',
      'After Go Live, do not reset APP_COUNTER or SN_COUNTER. Use the visible LIVE NUMBERING section to confirm the current and next values.'
    ],
  ];

  sheet
    .getRange(
      1,
      1,
      rows.length,
      2
    )
    .setValues(
      rows
    );

  sheet
    .getRange(
      'A1:B1'
    )
    .setFontWeight(
      'bold'
    );

  const liveNumberingRow =
    rows.findIndex(
      function(row) {
        return (
          row[0] ===
          'LIVE NUMBERING'
        );
      }
    ) + 1;

  if (liveNumberingRow > 0) {
    sheet
      .getRange(
        liveNumberingRow,
        1,
        1,
        2
      )
      .setFontWeight(
        'bold'
      );
  }

  sheet.setFrozenRows(
    1
  );

  sheet.autoResizeColumns(
    1,
    2
  );

  sheet
    .getRange(
      1,
      1,
      rows.length,
      2
    )
    .setWrap(
      true
    );
}

/**
 * Updates only the visible counter values in the Setup sheet.
 *
 * This is called after each new form submission so the Setup
 * sheet stays in sync with Document Properties.
 */
function syncLiveCountersToSetup_() {
  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!ss) {
    return;
  }

  const sheet =
    ss.getSheetByName(
      SETUP_SHEET_NAME
    );

  if (!sheet) {
    return;
  }

  const properties =
    PropertiesService
      .getDocumentProperties();

  const currentApplicationCounter =
    Number(
      properties.getProperty(
        PROP_KEYS.APP_COUNTER
      ) || '0'
    );

  const currentSerialCounter =
    Number(
      properties.getProperty(
        PROP_KEYS.SN_COUNTER
      ) || '0'
    );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 1) {
    return;
  }

  const labels =
    sheet
      .getRange(
        1,
        1,
        lastRow,
        1
      )
      .getDisplayValues()
      .map(function(row) {
        return String(
          row[0] || ''
        ).trim();
      });

  const valuesByLabel = {
    'Current Application Counter':
      currentApplicationCounter,

    'Next Application Number':
      'APP-' +
      String(
        currentApplicationCounter + 1
      ).padStart(
        4,
        '0'
      ),

    'Current S/N Counter':
      currentSerialCounter,

    'Next S/N':
      currentSerialCounter + 1,
  };

  Object.keys(
    valuesByLabel
  ).forEach(function(label) {
    const index =
      labels.indexOf(
        label
      );

    if (index === -1) {
      return;
    }

    sheet
      .getRange(
        index + 1,
        2
      )
      .setValue(
        valuesByLabel[label]
      );
  });
}

/**
 * Checks whether the current account has an onFormSubmit trigger.
 */
function hasCurrentAccountSubmissionTrigger_() {
  return ScriptApp
    .getProjectTriggers()
    .some(trigger =>
      trigger.getHandlerFunction() === 'onFormSubmit'
    );
}

/**
 * Finds the newest previous row for the same Application ID,
 * but never considers the newly submitted row itself.
 */
function findPreviousApplicationRow_(
  sheet,
  applicationId,
  currentRow
) {
  const normalizedId =
    String(
      applicationId || ''
    ).trim();

  if (
    !normalizedId ||
    currentRow <= 2
  ) {
    return null;
  }

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
    getHeaderMap_(
      headers
    );

  const applicationIdColumn =
    headerMap[
      'Application ID'
    ];

  if (!applicationIdColumn) {
    return null;
  }

  const numberOfRows =
    currentRow - 2;

  const values =
    sheet
      .getRange(
        2,
        1,
        numberOfRows,
        lastColumn
      )
      .getDisplayValues();

  const idIndex =
    applicationIdColumn - 1;

  for (
    let index =
      values.length - 1;
    index >= 0;
    index--
  ) {
    if (
      String(
        values[index][
          idIndex
        ] || ''
      ).trim() !==
      normalizedId
    ) {
      continue;
    }

    return rowToObject_(
      headers,
      values[index]
    );
  }

  return null;
}


/**
 * Copies workflow values that must survive an application
 * information correction/resubmission.
 *
 * Application Information and Final Verification themselves are
 * deliberately reset for a fresh review.
 */
function getPreservedResubmissionValues_(
  previousRow
) {
  if (!previousRow) {
    return {};
  }

  const headersToPreserve = [
    /*
     * Payment workflow
     */
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
    'Payment Review Notes',
    'Payment Rejection Reason',

    /*
     * Supporting-document workflow
     */
    'Document Status',
    'Document Review Notes',
    'Documents Verified At',
    'Documents Verified By',

    'CAC Document URL',
    'CAC Document File ID',
    'CAC Document Uploaded At',

    'Passport Photograph URL',
    'Passport Photograph File ID',
    'Passport Photograph Uploaded At',

    'Educational Certificates URL',
    'Educational Certificates File ID',
    'Educational Certificates Uploaded At',

    'Proof of Experience URL',
    'Proof of Experience File ID',
    'Proof of Experience Uploaded At',

    'Means of Identification URL',
    'Means of Identification File ID',
    'Means of Identification Uploaded At',

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
    'Document Correction Email Sent At',
    'Document Correction Email Sent By',
  ];

  const result =
    {};

  headersToPreserve.forEach(
    function(header) {
      if (
        Object.prototype
          .hasOwnProperty.call(
            previousRow,
            header
          )
      ) {
        result[header] =
          previousRow[header];
      }
    }
  );

  return result;
}


/**
 * Runs automatically whenever a new Google Form response is submitted.
 */
function onFormSubmit(e) {
  if (
    !e ||
    !e.range ||
    !e.namedValues
  ) {
    throw new Error(
      'This function must run through the spreadsheet form-submission trigger.'
    );
  }

  const lock =
    LockService
      .getDocumentLock();

  let emailPayload =
    null;

  lock.waitLock(
    30000
  );

  try {
    const sheet =
      e.range.getSheet();

    const row =
      e.range.getRow();

    ensureCustomColumns_(
      sheet
    );

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          sheet.getLastColumn()
        )
        .getValues()[0];

    const headerMap =
      getHeaderMap_(
        headers
      );

    const serialNumber =
      nextSerialNumber_();

    const email =
      normalizeEmail_(
        firstValue_(
          e.namedValues[
            EMAIL_FIELD
          ]
        )
      );

    const applicantName =
      String(
        firstValue_(
          e.namedValues[
            'Full Name'
          ]
        ) || ''
      ).trim();

    let applicationId =
      '';

    let secureToken =
      '';

    let portalUrl =
      '';

    let recordStatus =
      'Needs Review';

    let duplicateFlag =
      'No';

    let duplicateReason =
      '';

    let matchedExistingId =
      '';

    let reviewNote =
      '';

    let submissionCount =
      '';

    let isResubmission =
      false;

    let previousRow =
      null;

    if (!email) {
      reviewNote =
        'Missing email address on submission.';
    } else {
      const docProps =
        PropertiesService
          .getDocumentProperties();

      const idKey =
        'EMAIL_TO_APP::' +
        email;

      const tokenKey =
        'EMAIL_TO_TOKEN::' +
        email;

      const countKey =
        'EMAIL_SUBMISSION_COUNT::' +
        email;

      applicationId =
        docProps.getProperty(
          idKey
        );

      secureToken =
        docProps.getProperty(
          tokenKey
        );

      let count =
        Number(
          docProps.getProperty(
            countKey
          ) || '0'
        );

      if (applicationId) {
        isResubmission =
          true;

        previousRow =
          findPreviousApplicationRow_(
            sheet,
            applicationId,
            row
          );

        count += 1;

        recordStatus =
          'Resubmission';

        /*
         * This is an intentional correction/resubmission,
         * not a new Application ID.
         */
        duplicateFlag =
          'Yes';

        duplicateReason =
          'This email address already has an application number.';

        matchedExistingId =
          applicationId;

        reviewNote =
          'Correction/resubmission received from the same email address.';
      } else {
        applicationId =
          nextApplicationId_();

        secureToken =
          generateSecureToken_();

        count =
          1;

        recordStatus =
          'New';

        duplicateFlag =
          'No';

        docProps.setProperty(
          idKey,
          applicationId
        );

        docProps.setProperty(
          tokenKey,
          secureToken
        );
      }

      if (!secureToken) {
        secureToken =
          generateSecureToken_();

        docProps.setProperty(
          tokenKey,
          secureToken
        );
      }

      docProps.setProperty(
        countKey,
        String(
          count
        )
      );

      submissionCount =
        count;

      portalUrl =
        buildApplicantPortalUrl_(
          applicationId,
          secureToken
        );
    }

    /*
     * Default workflow for a completely new application.
     */
    const valuesToWrite = {
      'S/N':
        serialNumber,

      'Application ID':
        applicationId,

      'Record Status':
        recordStatus,

      'Duplicate Flag':
        duplicateFlag,

      'Duplicate Reason':
        duplicateReason,

      'Matched Existing ID':
        matchedExistingId,

      'Review Note':
        reviewNote,

      'Submission Count':
        submissionCount,

      'Secure Token':
        secureToken,

      'Portal URL':
        portalUrl,

      'Application Email Status':
        'Pending',

      'Application Email Sent At':
        '',

      'Application Email Error':
        '',

      'Payment Status':
        'Not Available',

      'Payment Reference':
        '',

      'Document Status':
        'Not Submitted',

      'Application Information Status':
        'Pending',

      'Application Information Review Notes':
        '',

      'Application Correction Fields JSON':
        '',

      'Application Information Verified At':
        '',

      'Application Information Verified By':
        '',

      'Verification Status':
        'Pending',

      'Verification Notes':
        '',

      'Verification Completed At':
        '',

      'Verification Completed By':
        '',

      'Licence Status':
        'Not Generated',

      'Application PDF URL':
        '',

      'Receipt PDF URL':
        '',

      'Licence PDF URL':
        '',
    };

    if (
      isResubmission &&
      previousRow
    ) {
      /*
       * Preserve unrelated approved/submitted workflow data.
       *
       * The corrected Google Form affects Application Information,
       * so ONLY that stage and Final Verification are reopened.
       */
      Object.assign(
        valuesToWrite,
        getPreservedResubmissionValues_(
          previousRow
        )
      );

      valuesToWrite[
        'Application Information Status'
      ] =
        'Pending';

      valuesToWrite[
        'Application Information Review Notes'
      ] =
        '';

      valuesToWrite[
        'Application Correction Fields JSON'
      ] =
        '';

      valuesToWrite[
        'Application Information Verified At'
      ] =
        '';

      valuesToWrite[
        'Application Information Verified By'
      ] =
        '';

      valuesToWrite[
        'Verification Status'
      ] =
        'Pending';

      valuesToWrite[
        'Verification Notes'
      ] =
        '';

      valuesToWrite[
        'Verification Completed At'
      ] =
        '';

      valuesToWrite[
        'Verification Completed By'
      ] =
        '';

      /*
       * Licence processing must start only after the corrected
       * application passes review again.
       */
      valuesToWrite[
        'Licence Status'
      ] =
        'Not Generated';

      valuesToWrite[
        'Licence PDF URL'
      ] =
        '';
    }

    writeRowValues_(
      sheet,
      row,
      headerMap,
      valuesToWrite
    );

    SpreadsheetApp.flush();

    syncLiveCountersToSetup_();

    if (
      email &&
      applicationId &&
      portalUrl &&
      portalUrl !==
        'Pending web app deployment'
    ) {
      emailPayload = {
        email:
          email,

        applicantName:
          applicantName,

        applicationId:
          applicationId,

        portalUrl:
          portalUrl,

        recordStatus:
          recordStatus,

        sheetName:
          sheet.getName(),

        row:
          row
      };
    } else {
      writeRowValues_(
        sheet,
        row,
        headerMap,
        {
          'Application Email Status':
            'Not Attempted',

          'Application Email Sent At':
            '',

          'Application Email Error':
            !email
              ? 'Email Address was missing from the submitted form event.'
              : (
                !applicationId
                  ? 'Application ID was not generated.'
                  : (
                    !portalUrl ||
                    portalUrl ===
                      'Pending web app deployment'
                      ? 'Practitioner Portal URL is not configured.'
                      : 'Application email prerequisites were incomplete.'
                  )
              )
        }
      );

      SpreadsheetApp.flush();
    }
  } catch (error) {
    console.error(
      error
    );

    throw error;
  } finally {
    lock.releaseLock();
  }

  if (emailPayload) {
    try {
      sendApplicantPortalEmail_(
        emailPayload.email,
        emailPayload.applicantName,
        emailPayload.applicationId,
        emailPayload.portalUrl,
        emailPayload.recordStatus
      );

      writeApplicationEmailAudit_(
        emailPayload,
        'Sent',
        ''
      );
    } catch (emailError) {
      const message =
        emailError &&
        emailError.message
          ? emailError.message
          : String(
              emailError
            );

      writeApplicationEmailAudit_(
        emailPayload,
        'Failed',
        message
      );

      console.error(
        'Application saved, but applicant email failed: ' +
        message
      );

      /*
       * Do not hide the email failure.
       *
       * The form application has already been saved safely.
       * Throwing here makes Apps Script > Executions show FAILED,
       * with the real MailApp error visible for diagnosis.
       */
      throw new Error(
        'Application saved, but applicant email failed: ' +
        message
      );
    }
  } else {
    console.warn(
      'Application email was not attempted because email, Application ID or Practitioner Portal URL was unavailable.'
    );
  }
}


/**
 * Generates the next Application ID.
 *
 * Example:
 * APP-0001
 * APP-0002
 */
function nextApplicationId_() {
  const docProps =
    PropertiesService.getDocumentProperties();

  const current =
    Number(
      docProps.getProperty(PROP_KEYS.APP_COUNTER) || '0'
    ) + 1;

  docProps.setProperty(
    PROP_KEYS.APP_COUNTER,
    String(current)
  );

  return 'APP-' + String(current).padStart(4, '0');
}

/**
 * Generates the next serial number.
 */
function nextSerialNumber_() {
  const docProps =
    PropertiesService.getDocumentProperties();

  const current =
    Number(
      docProps.getProperty(PROP_KEYS.SN_COUNTER) || '0'
    ) + 1;

  docProps.setProperty(
    PROP_KEYS.SN_COUNTER,
    String(current)
  );

  return current;
}

/**
 * Generates a long random secure token.
 */
function generateSecureToken_() {
  const randomValue = [
    Utilities.getUuid(),
    Utilities.getUuid(),
    String(Date.now()),
    String(Math.random()),
  ].join('|');

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    randomValue,
    Utilities.Charset.UTF_8
  );

  return digest
    .map(byte => {
      const unsignedByte =
        byte < 0 ? byte + 256 : byte;

      return unsignedByte
        .toString(16)
        .padStart(2, '0');
    })
    .join('');
}

/**
 * Builds the practitioner portal URL.
 *
 * Before web-app deployment, it returns:
 * Pending web app deployment
 */
// function buildApplicantPortalUrl_(
//   applicationId,
//   secureToken
// ) {
//   const webAppUrl =APP_CONFIG.PRODUCTION_WEB_APP_URL;
//     // PropertiesService
//     //   .getDocumentProperties()
//     //   .getProperty(PROP_KEYS.WEB_APP_URL) ||


//   if (!webAppUrl) {
//     return 'Pending web app deployment';
//   }

//   return (
//     webAppUrl +
//     '?view=applicant' +
//     '&ref=' +
//     encodeURIComponent(applicationId) +
//     '&token=' +
//     encodeURIComponent(secureToken)
//   );
// }

/**
 * Finds the Google Form response sheet.
 */
function getResponseSheet_(ss) {
  const sheets = ss.getSheets();

  const namedResponseSheet = sheets.find(sheet =>
    /^Form Responses\b/i.test(sheet.getName())
  );

  if (namedResponseSheet) {
    return namedResponseSheet;
  }

  for (const sheet of sheets) {
    const lastCol = sheet.getLastColumn();

    if (lastCol < 1) {
      continue;
    }

    const headers = sheet
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(String);

    if (
      headers.includes('Timestamp') &&
      headers.includes(EMAIL_FIELD)
    ) {
      return sheet;
    }
  }

  throw new Error(
    'Could not find the Google Form response sheet.'
  );
}

/**
 * Adds missing custom columns to the response sheet.
 */
function ensureCustomColumns_(sheet) {
  const currentLastColumn = Math.max(
    sheet.getLastColumn(),
    1
  );

  const existingHeaders = sheet
    .getRange(1, 1, 1, currentLastColumn)
    .getValues()[0]
    .map(header => String(header).trim());

  const missingColumns = CUSTOM_COLUMNS.filter(
    columnName => !existingHeaders.includes(columnName)
  );

  if (missingColumns.length > 0) {
    const startColumn = sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        startColumn,
        1,
        missingColumns.length
      )
      .setValues([missingColumns]);
  }

  const refreshedHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const headerMap = getHeaderMap_(refreshedHeaders);

  applyStatusDropdown_(
    sheet,
    headerMap['Record Status'],
    STATUS_OPTIONS
  );

  applyStatusDropdown_(
    sheet,
    headerMap['Payment Status'],
    [
      'Not Available',
      'Awaiting Payment',
      'Pending Verification',
      'Confirmed',
      'Rejected',
      'Duplicate Reference',
      'Under Review',
    ]
  );

  applyStatusDropdown_(
    sheet,
    headerMap['Verification Status'],
    [
      'Pending',
      'Under Review',
      'Approved',
      'Rejected',
      'Correction Required',
    ]
  );

  applyStatusDropdown_(
    sheet,
    headerMap['Licence Status'],
    [
      'Not Generated',
      'Generated',
      'Released',
      'Suspended',
      'Revoked',
    ]
  );

  formatCustomColumns_(sheet, headerMap);
}

/**
 * Adds a dropdown validation rule to a status column.
 */
function applyStatusDropdown_(
  sheet,
  statusColumn,
  options
) {
  if (!statusColumn) {
    return;
  }

  const rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();

  const rowsToApply = Math.max(
    sheet.getMaxRows() - 1,
    1
  );

  sheet
    .getRange(
      2,
      statusColumn,
      rowsToApply,
      1
    )
    .setDataValidation(rule);
}

/**
 * Applies basic formatting to important custom columns.
 */
function formatCustomColumns_(sheet, headerMap) {
  const headerNames = [
    'S/N',
    'Application ID',
    'Record Status',
    'Secure Token',
    'Portal URL',
    'Payment Status',
    'Verification Status',
    'Licence Status',
  ];

  headerNames.forEach(headerName => {
    const column = headerMap[headerName];

    if (!column) {
      return;
    }

    sheet
      .getRange(1, column)
      .setFontWeight('bold')
      .setWrap(true);
  });

  if (headerMap['Secure Token']) {
    sheet.setColumnWidth(
      headerMap['Secure Token'],
      260
    );
  }

  if (headerMap['Portal URL']) {
    sheet.setColumnWidth(
      headerMap['Portal URL'],
      320
    );
  }
}

/**
 * Converts the response-sheet headers into:
 *
 * {
 *   "Application ID": 15,
 *   "Secure Token": 23
 * }
 */
function getHeaderMap_(headers) {
  const map = {};

  headers.forEach((header, index) => {
    map[String(header).trim()] = index + 1;
  });

  return map;
}

/**
 * Writes multiple values to the submitted row.
 *
 * It uses one batch write instead of many individual setValue calls.
 */
function writeRowValues_(
  sheet,
  row,
  headerMap,
  valuesByHeader
) {
  const lastColumn = sheet.getLastColumn();

  const existingRow = sheet
    .getRange(row, 1, 1, lastColumn)
    .getValues()[0];

  Object.keys(valuesByHeader).forEach(columnName => {
    const columnNumber = headerMap[columnName];

    if (!columnNumber) {
      return;
    }

    existingRow[columnNumber - 1] =
      valuesByHeader[columnName];
  });

  sheet
    .getRange(row, 1, 1, lastColumn)
    .setValues([existingRow]);
}

/**
 * Gets the first value from a form response.
 */
function firstValue_(value) {
  if (!value) {
    return '';
  }

  if (Array.isArray(value)) {
    return String(value[0]).trim();
  }

  return String(value).trim();
}

/**
 * Normalises the applicant email address.
 */
function normalizeEmail_(email) {
  return String(email || '')
    .trim()
    .toLowerCase();
}

/**
 * Returns the Setup sheet, creating it when necessary.
 */
function getOrCreateSetupSheet_(ss) {
  let sheet = ss.getSheetByName(SETUP_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SETUP_SHEET_NAME);
  }

  return sheet;
}
/**
 * Writes the application-email result back to the exact submitted row.
 */
function writeApplicationEmailAudit_(
  payload,
  status,
  errorMessage
) {
  try {
    const spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();

    const sheet =
      spreadsheet.getSheetByName(
        payload.sheetName
      );

    if (!sheet) {
      console.error(
        'Could not write email audit: response sheet was not found.'
      );
      return;
    }

    ensureCustomColumns_(
      sheet
    );

    const headers =
      sheet
        .getRange(
          1,
          1,
          1,
          sheet.getLastColumn()
        )
        .getValues()[0];

    const headerMap =
      getHeaderMap_(
        headers
      );

    writeRowValues_(
      sheet,
      payload.row,
      headerMap,
      {
        'Application Email Status':
          status,

        'Application Email Sent At':
          status ===
            'Sent'
            ? new Date()
            : '',

        'Application Email Error':
          errorMessage || ''
      }
    );

    SpreadsheetApp.flush();
  } catch (auditError) {
    console.error(
      'Could not write application email audit: ' +
      (
        auditError &&
        auditError.message
          ? auditError.message
          : String(
              auditError
            )
      )
    );
  }
}


/**
 * Manual MailApp diagnostic.
 *
 * Run this from the Apps Script editor.
 * It sends a test email to the signed-in Google account.
 *
 * This isolates MailApp authorization/quota from the form trigger.
 */
function testApplicantEmailDelivery() {
  const recipient =
    String(
      Session
        .getActiveUser()
        .getEmail() ||
      Session
        .getEffectiveUser()
        .getEmail() ||
      ''
    ).trim();

  if (!recipient) {
    throw new Error(
      'Google could not determine the signed-in account email. Authorize the script and try again.'
    );
  }

  const remainingQuota =
    MailApp
      .getRemainingDailyQuota();

  if (
    remainingQuota <= 0
  ) {
    throw new Error(
      'MailApp daily sending quota has been exhausted for the account running this script.'
    );
  }

  MailApp.sendEmail({
    to:
      recipient,

    subject:
      'CRFFN Licensing Email Delivery Test',

    body:
      [
        'This is a CRFFN licensing email-delivery test.',
        '',
        'If you received this message, MailApp is authorized and able to send from the account running the script.',
        '',
        'Remaining daily MailApp quota before this send: ' +
          remainingQuota
      ].join('\n'),

    name:
      'CRFFN Licensing System',

    replyTo:
      'licensing@crffn.gov.ng'
  });

  return {
    ok: true,
    recipient:
      recipient,
    remainingDailyQuota:
      MailApp
        .getRemainingDailyQuota()
  };
}


/**
 * Diagnostic for the latest form-submission trigger.
 *
 * Shows whether an onFormSubmit installable trigger exists for
 * the account running this function.
 */



// Send Email
function escapeApplicationEmailHtml_(
  value
) {
  return String(
    value == null
      ? ''
      : value
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
      '&#39;'
    );
}


/**
 * Sends the initial Application Received email and the
 * Application Correction Received acknowledgement.
 *
 * Both messages use the same branded CRFFN email design.
 * The workflow, Application ID, secure portal link and
 * resubmission behaviour are unchanged.
 */
function sendApplicantPortalEmail_(
  email,
  applicantName,
  applicationId,
  portalUrl,
  recordStatus
) {
  if (
    !email ||
    !applicationId ||
    !portalUrl
  ) {
    throw new Error(
      'Email, Application ID and Practitioner Portal URL are required.'
    );
  }

  const isResubmission =
    recordStatus ===
    'Resubmission';

  const safeApplicantName =
    String(
      applicantName || ''
    ).trim();

  const greetingName =
    safeApplicantName ||
    'Applicant';

  const subject =
    isResubmission
      ? (
        'CRFFN Licence Application Correction Received — ' +
        applicationId
      )
      : (
        'CRFFN Licence Application Received — ' +
        applicationId
      );

  const emailTitle =
    isResubmission
      ? 'Application Correction Received'
      : 'Licence Application Received';

  const plainOpening =
    isResubmission
      ? [
          'Your corrected application information has been received successfully and linked to your existing CRFFN licence application.',
          'Your Application ID remains unchanged.'
        ].join('\n\n')
      : [
          'Your CRFFN Licence application has been received successfully.',
          'Your unique Application ID has been created.'
        ].join('\n\n');

  const plainNextSteps =
    isResubmission
      ? [
          'WHAT HAPPENS NEXT',
          '',
          '1. CRFFN will review the corrected application information you submitted.',
          '2. Any previously submitted supporting documents remain linked to this Application ID unless CRFFN specifically requests a replacement.',
          '3. Continue to use your Practitioner Portal to monitor the application.',
          '4. If another correction is required, CRFFN will notify you.'
        ].join('\n')
      : [
          'WHAT TO DO NEXT',
          '',
          '1. Open your Practitioner Portal using the secure link below.',
          '2. Upload all five required supporting documents.',
          '3. Wait for CRFFN to review your Application Information and Supporting Documents.',
          '4. Payment is not required yet. CRFFN will notify you when payment becomes available.',
          '5. Continue to use your Practitioner Portal to monitor the application.'
        ].join('\n');

  const message = [
    'Dear ' +
      greetingName +
      ',',
    '',
    plainOpening,
    '',
    'Application ID: ' +
      applicationId,
    '',
    plainNextSteps,
    '',
    'Practitioner Portal:',
    portalUrl,
    '',
    'Keep this secure portal link private.',
    '',
    'If you cannot find a CRFFN email in your inbox, check your Spam or Junk folder.',
    '',
    'NEED HELP?',
    'For enquiries, clarification or to report an issue, contact:',
    'licensing@crffn.gov.ng',
    '',
    'Please include your Application ID in your message.',
    '',
    'Regards,',
    'CRFFN Licensing Team'
  ].join('\n');

  const safeNameHtml =
    escapeApplicationEmailHtml_(
      greetingName
    );

  const safeApplicationIdHtml =
    escapeApplicationEmailHtml_(
      applicationId
    );

  const safePortalUrlHtml =
    escapeApplicationEmailHtml_(
      portalUrl
    );

  const safeTitleHtml =
    escapeApplicationEmailHtml_(
      emailTitle
    );

  const introHtml =
    isResubmission
      ? (
        '<p style="margin:0 0 18px;color:#344054;font-size:16px;line-height:1.65;">' +
          'Your corrected application information has been received successfully and linked to your existing CRFFN licence application.' +
        '</p>' +
        '<p style="margin:0;color:#344054;font-size:16px;line-height:1.65;">' +
          'Your Application ID remains unchanged.' +
        '</p>'
      )
      : (
        '<p style="margin:0 0 18px;color:#344054;font-size:16px;line-height:1.65;">' +
          'Your CRFFN Licence application has been received successfully.' +
        '</p>' +
        '<p style="margin:0;color:#344054;font-size:16px;line-height:1.65;">' +
          'Your unique Application ID has been created.' +
        '</p>'
      );

  const stepsHtml =
    isResubmission
      ? (
        '<div style="margin-top:30px;padding:24px;border:1px solid #e4e7ec;border-radius:14px;background:#f8fafc;">' +
          '<div style="margin-bottom:16px;color:#667085;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;">What happens next</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">1. CRFFN review</strong><br>Your corrected application information will be reviewed again.</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">2. Existing documents remain linked</strong><br>Previously submitted supporting documents remain linked to this Application ID unless CRFFN specifically requests a replacement.</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">3. Monitor your application</strong><br>Continue to use the Practitioner Portal to follow the review and processing status.</div>' +
          '<div style="color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">4. Further correction requests</strong><br>If another correction is required, CRFFN will notify you.</div>' +
        '</div>'
      )
      : (
        '<div style="margin-top:30px;padding:24px;border:1px solid #e4e7ec;border-radius:14px;background:#f8fafc;">' +
          '<div style="margin-bottom:16px;color:#667085;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;">What to do next</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">1. Open your Practitioner Portal</strong><br>Use the secure button below to continue your application and monitor its progress.</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">2. Upload supporting documents</strong><br>Upload the CAC Document, Passport Photograph, Educational Certificates, Proof of Experience and Valid Means of Identification.</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">3. Wait for CRFFN review</strong><br>CRFFN will review your Application Information and Supporting Documents.</div>' +
          '<div style="margin-bottom:14px;color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">4. Payment comes later</strong><br>Payment is not required at this stage. You will receive a separate notification when payment becomes available.</div>' +
          '<div style="color:#344054;font-size:14px;line-height:1.6;"><strong style="color:#17202a;">5. Monitor your application</strong><br>Return to the Practitioner Portal at any time to check your application status.</div>' +
        '</div>'
      );

  const htmlBody =
    '<!doctype html>' +
    '<html>' +
      '<body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;color:#17202a;">' +
        '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;background:#f4f6f8;margin:0;padding:0;">' +
          '<tr>' +
            '<td align="center" style="padding:28px 14px;">' +
              '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="width:100%;max-width:720px;background:#ffffff;border-collapse:separate;border-spacing:0;border-radius:18px;overflow:hidden;">' +

                '<tr>' +
                  '<td style="padding:34px 34px 38px;background:#183b56;color:#ffffff;">' +
                    '<div style="margin:0 0 14px;color:#d9e5ee;font-size:13px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;">CRFFN Licensing System</div>' +
                    '<div style="margin:0;color:#ffffff;font-size:30px;line-height:1.25;font-weight:800;">' +
                      safeTitleHtml +
                    '</div>' +
                  '</td>' +
                '</tr>' +

                '<tr>' +
                  '<td style="padding:34px;">' +
                    '<p style="margin:0 0 24px;color:#344054;font-size:18px;line-height:1.5;">Dear ' +
                      safeNameHtml +
                      ',</p>' +

                    introHtml +

                    '<div style="margin-top:28px;padding:22px 24px;border:1px solid #e4e7ec;border-radius:14px;background:#f8fafc;">' +
                      '<div style="margin-bottom:8px;color:#667085;font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;">Application ID</div>' +
                      '<div style="color:#183b56;font-size:26px;line-height:1.2;font-weight:800;">' +
                        safeApplicationIdHtml +
                      '</div>' +
                    '</div>' +

                    stepsHtml +

                    '<div style="margin-top:28px;text-align:center;">' +
                      '<a href="' +
                        safePortalUrlHtml +
                        '" style="display:inline-block;padding:14px 24px;border-radius:9px;background:#183b56;color:#ffffff;text-decoration:none;font-size:15px;font-weight:700;">Open Practitioner Portal</a>' +
                    '</div>' +

                    '<div style="margin-top:18px;padding:14px 16px;border-left:4px solid #183b56;background:#f8fafc;color:#667085;font-size:13px;line-height:1.6;">' +
                      '<strong style="color:#344054;">Keep this link private.</strong> The Practitioner Portal link provides secure access to your application. If you cannot find a CRFFN email in your inbox, check your Spam or Junk folder.' +
                    '</div>' +

                    '<div style="margin-top:30px;padding-top:24px;border-top:1px solid #e4e7ec;color:#667085;font-size:13px;line-height:1.65;">' +
                      '<strong style="color:#344054;">Need help?</strong><br>' +
                      'For enquiries, clarification or to report an issue, contact ' +
                      '<a href="mailto:licensing@crffn.gov.ng" style="color:#183b56;text-decoration:none;font-weight:700;">licensing@crffn.gov.ng</a> and include your Application ID.' +
                    '</div>' +

                    '<p style="margin:28px 0 0;color:#344054;font-size:14px;line-height:1.6;">Regards,<br><strong>CRFFN Licensing Team</strong></p>' +
                  '</td>' +
                '</tr>' +

              '</table>' +
            '</td>' +
          '</tr>' +
        '</table>' +
      '</body>' +
    '</html>';

  sendSystemEmail_({
  to:
    email,

  subject:
    subject,

  body:
    message,

  htmlBody:
    htmlBody
  });
}

function resetAllTestDataCompletely() {
  throw new Error(
    'This legacy full reset has been disabled. Use Form Tools > Reset Test Application for a single test record, or Form Tools > Go Live Reset & Set Numbers for the production cutover.'
  );
}

function testApplicantEmailPermission() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: 'CRFFN Licensing Email Test',
    body: 'CRFFN licensing email permission is working.',
    name: 'CRFFN Licensing System',
  });

  console.log('Test email sent.');
}
function diagnoseSubmissionEmailSetup() {
  const triggers =
    ScriptApp.getProjectTriggers();

  const triggerDetails =
    triggers.map(
      function(trigger) {
        return {
          handlerFunction:
            trigger.getHandlerFunction(),

          eventType:
            String(
              trigger.getEventType()
            ),

          triggerSource:
            String(
              trigger.getTriggerSource()
            ),

          triggerSourceId:
            String(
              trigger.getTriggerSourceId() ||
              ''
            ),

          uniqueId:
            String(
              trigger.getUniqueId() ||
              ''
            )
        };
      }
    );

  const spreadsheet =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const result = {
    currentAccount:
      String(
        Session
          .getEffectiveUser()
          .getEmail() ||
        ''
      ).trim(),

    currentSpreadsheetId:
      spreadsheet
        ? spreadsheet.getId()
        : '',

    currentSpreadsheetName:
      spreadsheet
        ? spreadsheet.getName()
        : '',

    submissionTriggerCount:
      triggerDetails.filter(
        function(trigger) {
          return (
            trigger.handlerFunction ===
            'onFormSubmit'
          );
        }
      ).length,

    triggers:
      triggerDetails,

    remainingDailyMailQuota:
      MailApp
        .getRemainingDailyQuota(),

    productionWebAppUrl:
      typeof APP_CONFIG !==
        'undefined'
        ? String(
            APP_CONFIG
              .PRODUCTION_WEB_APP_URL ||
            ''
          ).trim()
        : '',

    formId:
      String(
        PropertiesService
          .getDocumentProperties()
          .getProperty(
            PROP_KEYS.FORM_ID
          ) || ''
      ).trim()
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
function checkGmailAliases() {
  const aliases = GmailApp.getAliases();
  Logger.log(aliases);
}