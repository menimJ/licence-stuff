const LICENSING_SUPPORT_EMAIL =
  'licensing@crffn.gov.ng';


function escapeCrffnEmailHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


function buildCrffnEmailShell_(options) {
  const input = options || {};
  const title = escapeCrffnEmailHtml_(input.title || 'CRFFN Licensing Update');
  const applicantName = escapeCrffnEmailHtml_(input.applicantName || 'Applicant');
  const applicationId = escapeCrffnEmailHtml_(input.applicationId || '');
  const introHtml = String(input.introHtml || '');
  const contentHtml = String(input.contentHtml || '');
  const steps = Array.isArray(input.steps) ? input.steps : [];
  const portalUrl = String(input.portalUrl || '').trim();
  const buttonText = escapeCrffnEmailHtml_(input.buttonText || 'Open Practitioner Portal');

  const stepsHtml = steps.length
    ? [
        '<div style="margin:26px 0 0;">',
        '<div style="font-size:12px;font-weight:800;letter-spacing:.08em;color:#667085;text-transform:uppercase;margin-bottom:12px;">What to do next</div>',
        steps.map(function(step, index) {
          return [
            '<div style="display:table;width:100%;margin:0 0 12px;">',
            '<div style="display:table-cell;width:36px;vertical-align:top;">',
            '<span style="display:inline-block;width:28px;height:28px;line-height:28px;text-align:center;border-radius:50%;background:#183b56;color:#ffffff;font-size:12px;font-weight:800;">',
            String(index + 1),
            '</span>',
            '</div>',
            '<div style="display:table-cell;vertical-align:middle;color:#344054;font-size:14px;line-height:1.55;">',
            escapeCrffnEmailHtml_(step),
            '</div>',
            '</div>'
          ].join('');
        }).join(''),
        '</div>'
      ].join('')
    : '';

  const buttonHtml = portalUrl
    ? [
        '<div style="margin:26px 0 0;">',
        '<a href="', escapeCrffnEmailHtml_(portalUrl), '" target="_blank" style="display:inline-block;background:#183b56;color:#ffffff;text-decoration:none;font-size:14px;font-weight:800;padding:13px 20px;border-radius:8px;">',
        buttonText,
        '</a>',
        '</div>'
      ].join('')
    : '';

  return [
    '<!doctype html><html><body style="margin:0;padding:0;background:#f4f6f8;font-family:Arial,Helvetica,sans-serif;color:#17202a;">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background:#f4f6f8;padding:24px 12px;">',
    '<tr><td align="center">',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:640px;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e4e7ec;">',
    '<tr><td style="padding:26px 28px;background:#183b56;color:#ffffff;">',
    '<div style="font-size:11px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;opacity:.82;">CRFFN Licensing System</div>',
    '<div style="font-size:24px;font-weight:800;line-height:1.25;margin-top:7px;">', title, '</div>',
    '</td></tr>',
    '<tr><td style="padding:28px;">',
    '<div style="font-size:15px;line-height:1.65;color:#344054;">Dear ', applicantName, ',</div>',
    '<div style="font-size:14px;line-height:1.7;color:#475467;margin-top:12px;">', introHtml, '</div>',
    applicationId
      ? [
          '<div style="margin:22px 0;padding:15px 16px;background:#f8fafc;border:1px solid #e4e7ec;border-radius:9px;">',
          '<div style="font-size:10px;font-weight:800;letter-spacing:.08em;color:#667085;text-transform:uppercase;">Application ID</div>',
          '<div style="font-size:20px;font-weight:800;color:#183b56;margin-top:4px;">', applicationId, '</div>',
          '</div>'
        ].join('')
      : '',
    contentHtml,
    stepsHtml,
    buttonHtml,
    '<div style="margin-top:30px;padding-top:20px;border-top:1px solid #e4e7ec;color:#667085;font-size:12.5px;line-height:1.65;">',
    '<strong style="color:#344054;">Need clarification?</strong><br>',
    'Contact <a href="mailto:', escapeCrffnEmailHtml_(LICENSING_SUPPORT_EMAIL), '" style="color:#183b56;">', escapeCrffnEmailHtml_(LICENSING_SUPPORT_EMAIL), '</a> and include your Application ID.',
    '</div>',
    '<div style="margin-top:24px;color:#475467;font-size:13px;line-height:1.6;">Regards,<br><strong>CRFFN Licensing Team</strong></div>',
    '</td></tr>',
    '</table>',
    '<div style="max-width:640px;margin:12px auto 0;color:#98a2b3;font-size:11px;text-align:center;">This is an official CRFFN licensing communication.</div>',
    '</td></tr></table>',
    '</body></html>'
  ].join('');
}


function buildAdminReviewCorrectionEmail_(payload) {
  const input = payload || {};
  const email = String(input.email || '').trim();
  const applicantName = String(input.applicantName || 'Applicant').trim();
  const applicationId = String(input.applicationId || '').trim();
  const stage = String(input.stage || '').trim().toLowerCase();
  const reason = String(input.reason || '').trim();
  const portalUrl = String(input.portalUrl || '').trim();

  if (!applicationId) throw new Error('Application ID is required.');
  if (!reason) throw new Error('A correction/rejection reason is required.');

  const config = getAdminReviewEmailConfig_(stage);
  if (!config) throw new Error('Unsupported review-email stage: ' + stage);

  const reasonLines = reason.split('\n').map(function(line) {
    return String(line || '').trim();
  }).filter(Boolean);

  const formattedReason = reasonLines.length > 1
    ? reasonLines.map(function(line) { return '- ' + line; }).join('\n')
    : reasonLines[0];

  const subject = config.subject + ' — ' + applicationId;

  const body = [
    'Dear ' + applicantName + ',', '', config.introduction, '',
    'Application ID: ' + applicationId, '',
    'Affected Stage: ' + config.stageName, '',
    'REASON FOR CORRECTION', formattedReason, '',
    'WHAT TO DO NEXT', config.nextStep, '',
    'Open your secure Practitioner Portal:', portalUrl, '',
    'You do not need to start a new application or obtain a new Application ID. Continue using your existing Application ID and secure Practitioner Portal.', '',
    'After completing the requested correction, the affected stage will return for review before processing can continue.', '',
    'NEED CLARIFICATION?',
    'If you do not understand the correction requested, need more information, or wish to report an issue, contact:',
    LICENSING_SUPPORT_EMAIL, '',
    'Please include your Application ID in your message.', '',
    'Regards,', 'CRFFN Licensing Team'
  ].join('\n');

  const reasonHtml = reasonLines.map(function(line) {
    return '<li style="margin:0 0 7px;">' + escapeCrffnEmailHtml_(line) + '</li>';
  }).join('');

  const htmlBody = buildCrffnEmailShell_({
    title: config.subject,
    applicantName: applicantName,
    applicationId: applicationId,
    portalUrl: portalUrl,
    introHtml: escapeCrffnEmailHtml_(config.introduction),
    contentHtml: [
      '<div style="margin:18px 0;padding:16px;background:#fef3f2;border:1px solid #fecdca;border-radius:9px;">',
      '<div style="font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#b42318;margin-bottom:8px;">Reason for correction</div>',
      '<ul style="margin:0;padding-left:20px;color:#344054;font-size:14px;line-height:1.6;">', reasonHtml, '</ul>',
      '</div>'
    ].join(''),
    steps: [
      config.nextStep,
      'Use the same Application ID and secure Practitioner Portal.',
      'Submit the requested correction and wait for CRFFN review.'
    ]
  });

  return { ok: true, to: email, subject: subject, body: body, htmlBody: htmlBody };
}


function sendAdminReviewCorrectionEmail_(payload) {
  const built = buildAdminReviewCorrectionEmail_(payload);
  if (!built.to) throw new Error('The applicant email address is missing.');

  MailApp.sendEmail({
    to: built.to,
    subject: built.subject,
    body: built.body,
    htmlBody: built.htmlBody,
    name: 'CRFFN Licensing System',
    replyTo: LICENSING_SUPPORT_EMAIL
  });

  return { ok: true, email: built.to };
}


function buildSupportingDocumentsCorrectionSummaryEmail_(payload) {
  const input = payload || {};
  const email = String(input.email || '').trim();
  const applicantName = String(input.applicantName || 'Applicant').trim();
  const applicationId = String(input.applicationId || '').trim();
  const portalUrl = String(input.portalUrl || '').trim();
  const corrections = Array.isArray(input.corrections) ? input.corrections : [];
  const approvedDocuments = Array.isArray(input.approvedDocuments) ? input.approvedDocuments : [];

  if (!applicationId) throw new Error('Application ID is required.');
  if (!email) throw new Error('The applicant email address is missing.');
  if (!corrections.length) throw new Error('At least one supporting-document correction is required.');

  const correctionLines = [];
  corrections.forEach(function(item, index) {
    const label = String(item && item.label || 'Supporting Document').trim();
    const note = String(item && item.note || 'Please replace this document with the correct, complete and legible file.').trim();
    correctionLines.push((index + 1) + '. ' + label);
    correctionLines.push('   Correction required: ' + note);
    correctionLines.push('');
  });

  const approvedLines = approvedDocuments.length
    ? approvedDocuments.map(function(label) { return '✓ ' + String(label || '').trim(); })
    : ['No supporting documents have been approved yet.'];

  const subject = 'Supporting Documents Correction Required — ' + applicationId;
  const body = [
    'Dear ' + applicantName + ',', '',
    'CRFFN has completed the current review of your supporting documents.', '',
    'Application ID: ' + applicationId, '',
    'DOCUMENTS REQUIRING CORRECTION', '', correctionLines.join('\n').trim(), '',
    'DOCUMENTS ALREADY APPROVED', '', approvedLines.join('\n'), '',
    'You do not need to resubmit any document listed as Approved.', '',
    'WHAT TO DO NEXT',
    'Open the Supporting Documents section in your Practitioner Portal and resubmit ONLY the documents listed under Documents Requiring Correction.', '',
    'For each document requiring correction, the Practitioner Portal will show the administrator note and make only that document available for replacement.', '',
    'Open your secure Practitioner Portal:', portalUrl, '',
    'After you resubmit the corrected document(s), they will return to Pending Review. Documents already approved will remain approved and locked.', '',
    'You do not need to start a new application or obtain a new Application ID.', '',
    'NEED CLARIFICATION?', 'If you do not understand a correction request, contact:', LICENSING_SUPPORT_EMAIL, '',
    'Please include your Application ID in your message.', '',
    'Regards,', 'CRFFN Licensing Team'
  ].join('\n');

  const correctionsHtml = corrections.map(function(item, index) {
    const label = escapeCrffnEmailHtml_(item && item.label || 'Supporting Document');
    const note = escapeCrffnEmailHtml_(item && item.note || 'Please replace this document with the correct, complete and legible file.');
    return [
      '<div style="margin:0 0 12px;padding:13px 14px;border:1px solid #fecdca;border-radius:8px;background:#fff;">',
      '<div style="font-weight:800;color:#17202a;">', String(index + 1), '. ', label, '</div>',
      '<div style="margin-top:5px;color:#b42318;font-size:13px;line-height:1.55;">', note, '</div>',
      '</div>'
    ].join('');
  }).join('');

  const approvedHtml = approvedDocuments.length
    ? approvedDocuments.map(function(label) {
        return '<div style="margin:0 0 7px;color:#027a48;font-size:13px;font-weight:700;">✓ ' + escapeCrffnEmailHtml_(label) + '</div>';
      }).join('')
    : '<div style="color:#667085;font-size:13px;">No supporting documents have been approved yet.</div>';

  const htmlBody = buildCrffnEmailShell_({
    title: 'Supporting Documents Correction Required',
    applicantName: applicantName,
    applicationId: applicationId,
    portalUrl: portalUrl,
    introHtml: 'CRFFN has completed the current review of your supporting documents. Only the documents listed for correction need to be replaced.',
    contentHtml: [
      '<div style="margin-top:20px;font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#b42318;">Documents requiring correction</div>',
      '<div style="margin-top:10px;">', correctionsHtml, '</div>',
      '<div style="margin-top:22px;font-size:11px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#027a48;">Documents already approved</div>',
      '<div style="margin-top:10px;padding:14px 16px;background:#ecfdf3;border:1px solid #abefc6;border-radius:8px;">', approvedHtml, '</div>'
    ].join(''),
    steps: [
      'Open the Supporting Documents section in your Practitioner Portal.',
      'Replace only the documents listed under Documents Requiring Correction.',
      'Submit the corrected document(s). Approved documents will remain locked.',
      'Wait for CRFFN to review the corrected document(s).'
    ]
  });

  return { ok: true, to: email, subject: subject, body: body, htmlBody: htmlBody };
}


function sendSupportingDocumentsCorrectionSummaryEmail_(payload) {
  const built = buildSupportingDocumentsCorrectionSummaryEmail_(payload);
  MailApp.sendEmail({
    to: built.to,
    subject: built.subject,
    body: built.body,
    htmlBody: built.htmlBody,
    name: 'CRFFN Licensing System',
    replyTo: LICENSING_SUPPORT_EMAIL
  });
  return { ok: true, email: built.to };
}


function getAdminReviewEmailConfig_(stage) {
  if (stage === 'payment') {
    return {
      subject: 'Payment Correction Required',
      stageName: 'Payment Verification',
      introduction: 'Your payment submission requires correction before your application can proceed.',
      nextStep: 'Open the Payment section in your Practitioner Portal. Correct the payment reference and/or upload a valid replacement payment receipt, then submit it again for verification.'
    };
  }

  if (stage === 'documents') {
    return {
      subject: 'Supporting Document Correction Required',
      stageName: 'Supporting Documents',
      introduction: 'One or more supporting documents require correction before your application can proceed.',
      nextStep: 'Open the Supporting Documents section in your Practitioner Portal and replace or upload each document identified below with the correct, complete and legible file.'
    };
  }

  if (stage === 'information') {
    return {
      subject: 'Application Information Correction Required',
      stageName: 'Application Information',
      introduction: 'Your application information requires correction before your application can proceed.',
      nextStep: 'Open your Practitioner Portal and use the "Correct Application Information" button. Submit the correction using the SAME email address associated with your existing application so your current Application ID is retained.'
    };
  }

  if (stage === 'verification') {
    return {
      subject: 'Application Verification Requires Attention',
      stageName: 'Final Verification',
      introduction: 'Your application could not be approved at the final verification stage and requires further attention.',
      nextStep: "Open your Practitioner Portal to review the administrator's reason and follow the correction instruction shown there. If the required action is unclear, contact the Licensing Team before submitting anything new."
    };
  }

  return null;
}


function buildCrffnPaymentInvitationEmail_(payload) {
  const input = payload || {};
  const email = String(input.email || '').trim();
  const applicantName = String(input.applicantName || 'Applicant').trim();
  const applicationId = String(input.applicationId || '').trim();
  const portalUrl = String(input.portalUrl || '').trim();

  if (!email) throw new Error('The applicant email address is missing.');
  if (!applicationId) throw new Error('Application ID is required.');

  const paymentUrl = 'https://revop.gov.ng/payments/generate-bill?org=0229006001000';
  const subject = 'CRFFN Licence Payment Now Available — ' + applicationId;

  const body = [
    'Dear ' + applicantName + ',', '',
    'Your Application Information and Supporting Documents have been approved. Payment is now available for your CRFFN licence application.', '',
    'Application ID: ' + applicationId, '',
    'PAYMENT AMOUNT', '₦250,000', '',
    'HOW TO PAY',
    '1. Open the Revop payment portal using the link below.',
    '2. Select Licence Fee.',
    '3. If Licence Fee is unavailable, select Registration Fees.',
    '4. Complete the payment and keep your Revop payment reference and receipt.', '',
    'Revop Payment Portal:', paymentUrl, '',
    'AFTER PAYMENT',
    'Open your secure Practitioner Portal, enter your Revop payment reference and upload your payment receipt for verification.', '',
    'Practitioner Portal:', portalUrl, '',
    'Do not create a new application. Continue using your existing Application ID and secure Practitioner Portal.', '',
    'For assistance, contact:', LICENSING_SUPPORT_EMAIL, '',
    'Please include your Application ID in your message.', '',
    'Regards,', 'CRFFN Licensing Team'
  ].join('\n');

  const htmlBody = buildCrffnEmailShell_({
    title: 'Licence Payment Now Available',
    applicantName: applicantName,
    applicationId: applicationId,
    portalUrl: portalUrl,
    introHtml: 'Your Application Information and Supporting Documents have been approved. Payment is now available for your CRFFN licence application.',
    contentHtml: [
      '<div style="margin:20px 0;padding:18px;background:#f8fafc;border:1px solid #dbe6ef;border-radius:10px;">',
      '<div style="font-size:10px;font-weight:800;letter-spacing:.07em;text-transform:uppercase;color:#667085;">Payment amount</div>',
      '<div style="font-size:28px;font-weight:800;color:#183b56;margin-top:4px;">₦250,000</div>',
      '<div style="margin-top:12px;font-size:12px;color:#667085;">Payment provider: <strong style="color:#344054;">Revop</strong></div>',
      '</div>',
      '<div style="margin-top:12px;"><a href="', escapeCrffnEmailHtml_(paymentUrl), '" target="_blank" style="color:#183b56;font-size:13px;font-weight:700;">Open Revop Payment Portal</a></div>'
    ].join(''),
    steps: [
      'Open the Revop payment portal.',
      'Select Licence Fee. If it is unavailable, select Registration Fees.',
      'Pay ₦250,000 and keep your Revop payment reference and receipt.',
      'Return to the Practitioner Portal, enter the reference and upload the receipt for verification.'
    ]
  });

  return { ok: true, to: email, subject: subject, body: body, htmlBody: htmlBody };
}


function sendCrffnPaymentInvitationEmail_(payload) {
  const built = buildCrffnPaymentInvitationEmail_(payload);
  MailApp.sendEmail({
    to: built.to,
    subject: built.subject,
    body: built.body,
    htmlBody: built.htmlBody,
    name: 'CRFFN Licensing System',
    replyTo: LICENSING_SUPPORT_EMAIL
  });
  return { ok: true, email: built.to };
}


function getApplicationFormLiveUrl_() {
  const properties = PropertiesService.getDocumentProperties();
  const formId = String(properties.getProperty(PROP_KEYS.FORM_ID) || '').trim();

  if (!formId) throw new Error('The Google Form ID is not configured.');

  return FormApp.openById(formId).getPublishedUrl();
}


function testAdminReviewEmailPermission() {
  const email = String(Session.getActiveUser().getEmail() || '').trim();
  if (!email) throw new Error('The signed-in Google account email could not be identified.');

  MailApp.sendEmail({
    to: email,
    subject: 'Admin Review Email Permission Test',
    body: 'The review/correction email permission is working.',
    name: 'CRFFN Licensing System',
    replyTo: LICENSING_SUPPORT_EMAIL
  });
}
