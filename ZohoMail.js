function getZeptoMailConfig_() {
  const props =
    PropertiesService
      .getScriptProperties();

  return {
    apiKey:
      props.getProperty(
        'ZEPTOMAIL_API_KEY'
      ),

    fromEmail:
      props.getProperty(
        'ZEPTOMAIL_FROM_EMAIL'
      ) ||
      'licensing@crffn.gov.ng',

    fromName:
      props.getProperty(
        'ZEPTOMAIL_FROM_NAME'
      ) ||
      'CRFFN Licensing System',

    apiUrl:
      'https://api.zeptomail.com/v1.1/email'
  };
}

function sendSystemEmail_(options) {
  const input =
    options || {};

  const config =
    getZeptoMailConfig_();

  if (!config.apiKey) {
    throw new Error(
      'ZeptoMail API key is missing.'
    );
  }

  const to =
    String(
      input.to || ''
    ).trim();

  const subject =
    String(
      input.subject || ''
    );

  const body =
    String(
      input.body || ''
    );

  const htmlBody =
    String(
      input.htmlBody ||
      input.body ||
      ''
    );

  if (!to) {
    throw new Error(
      'Email recipient is required.'
    );
  }

  const payload = {
    from: {
      address:
        config.fromEmail,

      name:
        config.fromName
    },

    to: [
      {
        email_address: {
          address:
            to
        }
      }
    ],

    subject:
      subject,

    htmlbody:
      htmlBody,

    textbody:
      body
  };

  const response =
    UrlFetchApp.fetch(
      config.apiUrl,
      {
        method: 'post',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Zoho-enczapikey ' +
            config.apiKey
        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true
      }
    );

  const status =
    response.getResponseCode();

  const responseBody =
    response.getContentText();

  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      'ZeptoMail send failed: ' +
      responseBody
    );
  }

  return {
    ok: true,
    email: to
  };
}
function testZeptoMailEmail() {
  sendSystemEmail_({
    to: 'sammymenim@gmail.com',

    subject:
      'CRFFN ZeptoMail Test',

    body:
      'This is a test email from the CRFFN Licensing System through ZeptoMail.',

    htmlBody: [
      '<div style="font-family:Arial,sans-serif;">',
      '<h2>CRFFN ZeptoMail Test</h2>',
      '<p>This email was sent through the CRFFN transactional email service.</p>',
      '<p>If you received this, the ZeptoMail integration is working.</p>',
      '</div>'
    ].join('')
  });

  Logger.log(
    'ZeptoMail test email request completed.'
  );
}