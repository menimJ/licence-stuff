function getZohoMailConfig_() {
  const props =
    PropertiesService
      .getScriptProperties();

  return {
    clientId:
      props.getProperty(
        'ZOHO_CLIENT_ID'
      ),

    clientSecret:
      props.getProperty(
        'ZOHO_CLIENT_SECRET'
      ),

    refreshToken:
      props.getProperty(
        'ZOHO_REFRESH_TOKEN'
      ),

    accountId:
      props.getProperty(
        'ZOHO_ACCOUNT_ID'
      ),

    fromEmail:
      props.getProperty(
        'ZOHO_FROM_EMAIL'
      ) ||
      'licensing@crffn.gov.ng',

    accountsUrl:
      props.getProperty(
        'ZOHO_ACCOUNTS_URL'
      ),

    mailApiUrl:
      props.getProperty(
        'ZOHO_MAIL_API_URL'
      ) ||
      'https://mail.zoho.com',
  };
}


function getZohoAccessToken_() {
  const config =
    getZohoMailConfig_();

  if (
    !config.clientId ||
    !config.clientSecret ||
    !config.refreshToken ||
    !config.accountsUrl
  ) {
    throw new Error(
      'Zoho Mail OAuth configuration is incomplete.'
    );
  }

  const response =
    UrlFetchApp.fetch(
      config.accountsUrl +
      '/oauth/v2/token',
      {
        method: 'post',

        payload: {
          refresh_token:
            config.refreshToken,

          client_id:
            config.clientId,

          client_secret:
            config.clientSecret,

          grant_type:
            'refresh_token',
        },

        muteHttpExceptions: true,
      }
    );

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      'Zoho access-token request failed: ' +
      body
    );
  }

  const data =
    JSON.parse(body);

  if (!data.access_token) {
    throw new Error(
      'Zoho did not return an access token.'
    );
  }

  return data.access_token;
}


function sendSystemEmail_(options) {
  const input =
    options || {};

  const config =
    getZohoMailConfig_();

  if (
    !config.accountId ||
    !config.fromEmail
  ) {
    throw new Error(
      'Zoho Mail account configuration is incomplete.'
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

  const accessToken =
    getZohoAccessToken_();

  const url =
    config.mailApiUrl +
    '/api/accounts/' +
    encodeURIComponent(
      config.accountId
    ) +
    '/messages';

  const payload = {
    fromAddress:
      config.fromEmail,

    toAddress:
      to,

    subject:
      subject,

    content:
      htmlBody,

    mailFormat:
      'html',
  };

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',

        contentType:
          'application/json',

        headers: {
          Authorization:
            'Zoho-oauthtoken ' +
            accessToken,
        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions:
          true,
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
      'Zoho Mail send failed: ' +
      responseBody
    );
  }

  return {
    ok: true,
    email: to,
  };
}
function testGetZohoAccount_() {
  const config = getZohoMailConfig_();
  const accessToken = getZohoAccessToken_();

  const response = UrlFetchApp.fetch(
    config.mailApiUrl + '/api/accounts',
    {
      method: 'get',
      headers: {
        Authorization:
          'Zoho-oauthtoken ' + accessToken,
      },
      muteHttpExceptions: true,
    }
  );

  Logger.log(
    'HTTP Status: ' +
      response.getResponseCode()
  );

  Logger.log(
    response.getContentText()
  );
}