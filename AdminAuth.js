/**
 * ============================================================
 * CRFFN ADMIN AUTHENTICATION
 * ============================================================
 *
 * Deployment:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * Applicant access remains token-based and does not require Google login.
 * Admin access uses email + password.
 *
 * Security model:
 * - Passwords are NEVER stored in plain text in the spreadsheet.
 * - The Admin Users sheet stores a salted HMAC-SHA256 password hash.
 * - The HMAC pepper and the shared temporary/default password are stored
 *   only in Script Properties.
 * - Browser sessions use long random tokens.
 * - Only a SHA-256 hash of each session token is stored in the sheet.
 * - Admin accounts can be disabled and temporarily locked after failures.
 * - A user whose password is reset must change it before entering Admin.
 */

const ADMIN_AUTH_CONFIG = Object.freeze({
  USERS_SHEET:
    'Admin Users',

  SESSIONS_SHEET:
    'Admin Sessions',

  SESSION_IDLE_HOURS:
    12,

  MAX_FAILED_ATTEMPTS:
    5,

  LOCK_MINUTES:
    15,

  MIN_PASSWORD_LENGTH:
    10,

  PROP_PEPPER:
    'CRFFN_ADMIN_AUTH_PEPPER',

  PROP_DEFAULT_PASSWORD:
    'CRFFN_ADMIN_DEFAULT_PASSWORD'
});


const ADMIN_USER_HEADERS =
  Object.freeze([
    'Name',
    'Email',
    'Password Hash',
    'Salt',
    'Role',
    'Status',
    'Must Change Password',
    'Failed Attempts',
    'Locked Until',
    'Last Login',
    'Password Changed At',
    'Created At',
    'Created By'
  ]);


const ADMIN_SESSION_HEADERS =
  Object.freeze([
    'Session Token Hash',
    'Admin Email',
    'Role',
    'Created At',
    'Last Activity At',
    'Expires At',
    'Status'
  ]);


/**
 * Run ONCE from the Apps Script editor after adding this file.
 *
 * It:
 * 1. Creates the Admin Users and Admin Sessions sheets.
 * 2. Creates the server-side password pepper.
 * 3. Generates a temporary/default password if none exists.
 * 4. Creates accounts for the emails currently listed in
 *    ADMIN_PORTAL_CONFIG.ADMIN_EMAILS.
 *
 * The first configured admin becomes Super Admin.
 * Other configured admins become Admin.
 *
 * The generated temporary/default password is written to the execution log.
 * Keep it private.
 */
function setupAdminAuthentication() {
  assertAdminAuthEditorSetup_();

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  ensureAdminAuthSheet_(
    ss,
    ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
    ADMIN_SESSION_HEADERS
  );

  const properties =
    PropertiesService.getScriptProperties();

  let pepper =
    String(
      properties.getProperty(
        ADMIN_AUTH_CONFIG.PROP_PEPPER
      ) || ''
    ).trim();

  if (!pepper) {
    pepper =
      generateAdminSecret_(64);

    properties.setProperty(
      ADMIN_AUTH_CONFIG.PROP_PEPPER,
      pepper
    );
  }

  let defaultPassword =
    String(
      properties.getProperty(
        ADMIN_AUTH_CONFIG.PROP_DEFAULT_PASSWORD
      ) || ''
    );

  if (!defaultPassword) {
    defaultPassword =
      'CRFFN-' +
      generateAdminSecret_(18);

    properties.setProperty(
      ADMIN_AUTH_CONFIG.PROP_DEFAULT_PASSWORD,
      defaultPassword
    );
  }

  const configuredEmails =
    (
      ADMIN_PORTAL_CONFIG &&
      Array.isArray(
        ADMIN_PORTAL_CONFIG.ADMIN_EMAILS
      )
    )
      ? ADMIN_PORTAL_CONFIG.ADMIN_EMAILS
      : [];

  const created =
    [];

  configuredEmails
    .map(
      function(email) {
        return String(
          email || ''
        )
          .trim()
          .toLowerCase();
      }
    )
    .filter(Boolean)
    .forEach(
      function(email, index) {
        const existing =
          findAdminUserByEmail_(
            usersSheet,
            email
          );

        if (existing) {
          return;
        }

        const role =
          'Super Admin';

        createAdminUserRecord_(
          usersSheet,
          {
            name:
              email.split('@')[0],
            email:
              email,
            role:
              role,
            status:
              'Active',
            mustChangePassword:
              true,
            password:
              defaultPassword,
            createdBy:
              Session
                .getActiveUser()
                .getEmail() ||
              'System Setup'
          }
        );

        created.push(
          {
            email:
              email,
            role:
              role
          }
        );
      }
    );

  Logger.log(
    'CRFFN Admin Authentication setup complete.'
  );

  Logger.log(
    'TEMPORARY / DEFAULT ADMIN PASSWORD: ' +
    defaultPassword
  );

  Logger.log(
    'Keep this password private. Every admin is forced to change it before dashboard access.'
  );

  return {
    ok: true,
    created:
      created,
    defaultPassword:
      defaultPassword,
    message:
      'Admin authentication is ready.'
  };
}


/**
 * Login endpoint called by AdminPortalPage.html.
 */
function adminLogin(
  email,
  password
) {
  const normalizedEmail =
    normalizeAdminEmail_(
      email
    );

  const enteredPassword =
    String(
      password || ''
    );

  if (
    !normalizedEmail ||
    !enteredPassword
  ) {
    return {
      ok: false,
      message:
        'Enter your administrator email address and password.'
    };
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  ensureAdminAuthSheet_(
    ss,
    ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
    ADMIN_SESSION_HEADERS
  );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      normalizedEmail
    );

  if (
    !user ||
    String(
      user.values.Status || ''
    ).trim() !==
      'Active'
  ) {
    Utilities.sleep(
      350
    );

    return {
      ok: false,
      message:
        'Invalid administrator email or password.'
    };
  }

  const lockedUntil =
    parseAdminDate_(
      user.values[
        'Locked Until'
      ]
    );

  if (
    lockedUntil &&
    lockedUntil.getTime() >
      Date.now()
  ) {
    return {
      ok: false,
      locked: true,
      message:
        'This administrator account is temporarily locked. Try again later or contact the Super Admin.'
    };
  }

  const expectedHash =
    String(
      user.values[
        'Password Hash'
      ] || ''
    );

  const actualHash =
    hashAdminPassword_(
      enteredPassword,
      String(
        user.values.Salt || ''
      )
    );

  if (
    !constantTimeStringEquals_(
      expectedHash,
      actualHash
    )
  ) {
    registerAdminLoginFailure_(
      usersSheet,
      user
    );

    Utilities.sleep(
      350
    );

    return {
      ok: false,
      message:
        'Invalid administrator email or password.'
    };
  }

  clearAdminLoginFailures_(
    usersSheet,
    user
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Last Login',
    new Date()
  );

  const mustChangePassword =
    String(
      user.values[
        'Must Change Password'
      ] || ''
    )
      .trim()
      .toLowerCase() ===
      'yes';

  const session =
    createAdminSession_(
      normalizedEmail,
      String(
        user.values.Role ||
        'Admin'
      ),
      mustChangePassword
    );

  return {
    ok: true,
    sessionToken:
      session.token,
    mustChangePassword:
      mustChangePassword,
    admin: {
      name:
        String(
          user.values.Name || ''
        ),
      email:
        normalizedEmail,
      role:
        String(
          user.values.Role ||
          'Admin'
        )
    },
    message:
      mustChangePassword
        ? 'Password change required.'
        : 'Login successful.'
  };
}


/**
 * Forced password change after first login/reset.
 */
function changeAdminPassword(
  sessionToken,
  currentPassword,
  newPassword,
  confirmPassword
) {
  const session =
    validateAdminSession_(
      sessionToken,
      {
        allowPasswordChangeOnly:
          true
      }
    );

  const current =
    String(
      currentPassword || ''
    );

  const next =
    String(
      newPassword || ''
    );

  const confirmation =
    String(
      confirmPassword || ''
    );

  if (
    !current ||
    !next ||
    !confirmation
  ) {
    throw new Error(
      'Complete all password fields.'
    );
  }

  if (
    next !==
    confirmation
  ) {
    throw new Error(
      'The new passwords do not match.'
    );
  }

  validateNewAdminPassword_(
    next
  );

  if (
    current ===
    next
  ) {
    throw new Error(
      'Choose a new password that is different from the temporary password.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      session.email
    );

  if (!user) {
    throw new Error(
      'Administrator account was not found.'
    );
  }

  const currentHash =
    hashAdminPassword_(
      current,
      String(
        user.values.Salt || ''
      )
    );

  if (
    !constantTimeStringEquals_(
      String(
        user.values[
          'Password Hash'
        ] || ''
      ),
      currentHash
    )
  ) {
    throw new Error(
      'The current password is incorrect.'
    );
  }

  const salt =
    generateAdminSecret_(
      32
    );

  const passwordHash =
    hashAdminPassword_(
      next,
      salt
    );

  setAdminUserValue_(
    usersSheet,
    user,
    'Salt',
    salt
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Password Hash',
    passwordHash
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Must Change Password',
    'No'
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Password Changed At',
    new Date()
  );

  invalidateAdminSessionsForEmail_(
    session.email
  );

  return {
    ok: true,
    message:
      'Password changed successfully. Sign in again with your new password.'
  };
}


/**
 * Logout.
 */
function adminLogout(
  sessionToken
) {
  invalidateAdminSession_(
    sessionToken
  );

  return {
    ok: true,
    message:
      'You have been signed out.'
  };
}


/**
 * Super Admin reset.
 *
 * The account is reset to the system temporary/default password
 * and forced to change it on the next login.
 */
function resetAdminPassword(
  sessionToken,
  targetEmail
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can reset administrator passwords.'
    );
  }

  const normalizedTarget =
    normalizeAdminEmail_(
      targetEmail
    );

  if (!normalizedTarget) {
    throw new Error(
      'Administrator email is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      normalizedTarget
    );

  if (!user) {
    throw new Error(
      'Administrator account was not found.'
    );
  }

  const defaultPassword =
    getAdminDefaultPassword_();

  const salt =
    generateAdminSecret_(
      32
    );

  setAdminUserValue_(
    usersSheet,
    user,
    'Salt',
    salt
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Password Hash',
    hashAdminPassword_(
      defaultPassword,
      salt
    )
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Must Change Password',
    'Yes'
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Failed Attempts',
    0
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Locked Until',
    ''
  );

  invalidateAdminSessionsForEmail_(
    normalizedTarget
  );

  return {
    ok: true,
    email:
      normalizedTarget,
    temporaryPassword:
      defaultPassword,
    message:
      'Password reset successfully. The administrator must sign in with the temporary password and create a new password.'
  };
}



/**
 * Returns administrator accounts for the Super Admin settings screen.
 * Password hashes and salts are never returned to the browser.
 */
function getAdminUsers(
  sessionToken
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can manage administrator accounts.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return {
      ok: true,
      users: []
    };
  }

  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        ADMIN_USER_HEADERS.length
      )
      .getValues();

  const headerIndex =
    {};

  ADMIN_USER_HEADERS.forEach(
    function(header, index) {
      headerIndex[header] =
        index;
    }
  );

  const users =
    rows
      .filter(
        function(row) {
          return String(
            row[
              headerIndex.Email
            ] || ''
          ).trim();
        }
      )
      .map(
        function(row) {
          return {
            name:
              String(
                row[
                  headerIndex.Name
                ] || ''
              ),
            email:
              normalizeAdminEmail_(
                row[
                  headerIndex.Email
                ]
              ),
            role:
              String(
                row[
                  headerIndex.Role
                ] || 'Admin'
              ),
            status:
              String(
                row[
                  headerIndex.Status
                ] || 'Active'
              ),
            mustChangePassword:
              String(
                row[
                  headerIndex[
                    'Must Change Password'
                  ]
                ] || ''
              ),
            failedAttempts:
              Number(
                row[
                  headerIndex[
                    'Failed Attempts'
                  ]
                ] || 0
              ),
            lockedUntil:
              formatAdminManagementDate_(
                row[
                  headerIndex[
                    'Locked Until'
                  ]
                ]
              ),
            lastLogin:
              formatAdminManagementDate_(
                row[
                  headerIndex[
                    'Last Login'
                  ]
                ]
              ),
            passwordChangedAt:
              formatAdminManagementDate_(
                row[
                  headerIndex[
                    'Password Changed At'
                  ]
                ]
              ),
            createdAt:
              formatAdminManagementDate_(
                row[
                  headerIndex[
                    'Created At'
                  ]
                ]
              )
          };
        }
      );

  return {
    ok: true,
    users:
      users
  };
}


/**
 * Changes an administrator's role.
 */
function updateAdminUserRole(
  sessionToken,
  targetEmail,
  role
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can change administrator roles.'
    );
  }

  const normalizedEmail =
    normalizeAdminEmail_(
      targetEmail
    );

  const normalizedRole =
    String(
      role || ''
    ).trim();

  if (
    normalizedRole !==
      'Admin' &&
    normalizedRole !==
      'Super Admin'
  ) {
    throw new Error(
      'Role must be Admin or Super Admin.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      normalizedEmail
    );

  if (!user) {
    throw new Error(
      'Administrator account was not found.'
    );
  }

  if (
    normalizedEmail ===
      admin.email &&
    normalizedRole !==
      'Super Admin'
  ) {
    throw new Error(
      'You cannot remove your own Super Admin role while signed in.'
    );
  }

  setAdminUserValue_(
    usersSheet,
    user,
    'Role',
    normalizedRole
  );

  invalidateAdminSessionsForEmail_(
    normalizedEmail
  );

  return {
    ok: true,
    email:
      normalizedEmail,
    role:
      normalizedRole,
    message:
      'Administrator role updated. Existing sessions for this account were signed out.'
  };
}


/**
 * Updates an administrator's display name.
 */
function updateAdminUserName(
  sessionToken,
  targetEmail,
  name
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can update administrator accounts.'
    );
  }

  const normalizedEmail =
    normalizeAdminEmail_(
      targetEmail
    );

  const cleanName =
    String(
      name || ''
    ).trim();

  if (!cleanName) {
    throw new Error(
      'Administrator name is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      normalizedEmail
    );

  if (!user) {
    throw new Error(
      'Administrator account was not found.'
    );
  }

  setAdminUserValue_(
    usersSheet,
    user,
    'Name',
    cleanName
  );

  return {
    ok: true,
    email:
      normalizedEmail,
    name:
      cleanName,
    message:
      'Administrator name updated.'
  };
}


function formatAdminManagementDate_(
  value
) {
  if (!value) {
    return '';
  }

  const date =
    parseAdminDate_(
      value
    );

  if (!date) {
    return '';
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd MMM yyyy, HH:mm'
  );
}


/**
 * Optional Super Admin helper to add another admin account.
 */
function createAdminUser(
  sessionToken,
  name,
  email,
  role
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can create administrator accounts.'
    );
  }

  const normalizedEmail =
    normalizeAdminEmail_(
      email
    );

  if (!normalizedEmail) {
    throw new Error(
      'A valid administrator email is required.'
    );
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  if (
    findAdminUserByEmail_(
      usersSheet,
      normalizedEmail
    )
  ) {
    throw new Error(
      'An administrator account already exists for this email address.'
    );
  }

  const defaultPassword =
    getAdminDefaultPassword_();

  createAdminUserRecord_(
    usersSheet,
    {
      name:
        String(
          name || ''
        ).trim() ||
        normalizedEmail.split('@')[0],
      email:
        normalizedEmail,
      role:
        String(
          role || 'Admin'
        ).trim() ||
        'Admin',
      status:
        'Active',
      mustChangePassword:
        true,
      password:
        defaultPassword,
      createdBy:
        admin.email
    }
  );

  return {
    ok: true,
    email:
      normalizedEmail,
    temporaryPassword:
      defaultPassword,
    message:
      'Administrator created. They must change the temporary password on first login.'
  };
}


/**
 * Optional Super Admin helper to enable/disable an admin.
 */
function setAdminUserStatus(
  sessionToken,
  targetEmail,
  status
) {
  const admin =
    requireAdminSession_(
      sessionToken
    );

  if (
    String(
      admin.role || ''
    ).trim() !==
      'Super Admin'
  ) {
    throw new Error(
      'Only a Super Admin can change administrator account status.'
    );
  }

  const normalizedEmail =
    normalizeAdminEmail_(
      targetEmail
    );

  const normalizedStatus =
    String(
      status || ''
    )
      .trim()
      .toLowerCase() ===
      'active'
        ? 'Active'
        : 'Disabled';

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      normalizedEmail
    );

  if (!user) {
    throw new Error(
      'Administrator account was not found.'
    );
  }

  if (
    normalizedEmail ===
      admin.email &&
    normalizedStatus ===
      'Disabled'
  ) {
    throw new Error(
      'You cannot disable your own Super Admin account while signed in.'
    );
  }

  setAdminUserValue_(
    usersSheet,
    user,
    'Status',
    normalizedStatus
  );

  if (
    normalizedStatus !==
      'Active'
  ) {
    invalidateAdminSessionsForEmail_(
      normalizedEmail
    );
  }

  return {
    ok: true,
    email:
      normalizedEmail,
    status:
      normalizedStatus
  };
}



/* ============================================================
 * PERFORMANCE: ADMIN SESSION CACHE
 * ============================================================
 */

function adminSessionCacheKey_(
  sessionToken
) {
  return (
    'crffn:admin:session:' +
    hashAdminSessionToken_(
      sessionToken
    )
  );
}


function getCachedAdminSession_(
  sessionToken
) {
  const key =
    adminSessionCacheKey_(
      sessionToken
    );

  const cached =
    crffnGetCachedJson_(
      key
    );

  if (!cached) {
    return null;
  }

  const revokedAt =
    Number(
      PropertiesService
        .getScriptProperties()
        .getProperty(
          'CRFFN_ADMIN_REVOKED_AT_' +
          normalizeAdminEmail_(
            cached.email
          )
        ) ||
      0
    );

  const issuedAt =
    Number(
      cached.issuedAt ||
      0
    );

  if (
    revokedAt &&
    revokedAt >= issuedAt
  ) {
    crffnScriptCache_()
      .remove(key);

    return null;
  }

  return cached;
}


function putCachedAdminSession_(
  sessionToken,
  session
) {
  const value =
    Object.assign(
      {},
      session,
      {
        issuedAt:
          Number(
            session.issuedAt ||
            Date.now()
          )
      }
    );

  crffnPutCachedJson_(
    adminSessionCacheKey_(
      sessionToken
    ),
    value,
    CRFFN_PERFORMANCE_CONFIG
      .ADMIN_SESSION_CACHE_SECONDS
  );
}


function revokeCachedAdminSessionsForEmail_(
  email
) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      'CRFFN_ADMIN_REVOKED_AT_' +
      normalizeAdminEmail_(
        email
      ),
      String(Date.now())
    );
}


/**
 * Session access used by AdminPortal.gs.
 */
function getAdminSessionAccess_(
  sessionToken
) {
  try {
    const session =
      validateAdminSession_(
        sessionToken,
        {
          allowPasswordChangeOnly:
            true
        }
      );

    return {
      ok:
        !session.mustChangePassword,
      authenticated:
        true,
      mustChangePassword:
        session.mustChangePassword,
      email:
        session.email,
      role:
        session.role,
      name:
        session.name || '',
      errorCode:
        session.mustChangePassword
          ? 'ADMIN_PASSWORD_CHANGE_REQUIRED'
          : '',
      message:
        session.mustChangePassword
          ? 'You must change your temporary password before opening the Admin Portal.'
          : 'Administrator access granted.'
    };
  } catch (error) {
    return {
      ok: false,
      authenticated: false,
      mustChangePassword: false,
      email: '',
      role: '',
      name: '',
      errorCode:
        'ADMIN_LOGIN_REQUIRED',
      message:
        'Administrator login is required.'
    };
  }
}


/**
 * Required by every sensitive admin backend action.
 */
function requireAdminSession_(
  sessionToken
) {
  const session =
    validateAdminSession_(
      sessionToken,
      {
        allowPasswordChangeOnly:
          false
      }
    );

  return {
    email:
      session.email,
    role:
      session.role,
    name:
      session.name || ''
  };
}


/* ============================================================
 * Internal helpers
 * ============================================================
 */

function validateAdminSession_(
  sessionToken,
  options
) {
  const token =
    String(
      sessionToken || ''
    ).trim();

  if (!token) {
    throw new Error(
      'Administrator session has expired. Sign in again.'
    );
  }

  /*
   * Fast path: most admin actions never touch the spreadsheet.
   */
  const cached =
    getCachedAdminSession_(
      token
    );

  if (cached) {
    if (
      cached.mustChangePassword &&
      !(
        options &&
        options.allowPasswordChangeOnly
      )
    ) {
      throw new Error(
        'You must change your temporary password before using the Admin Portal.'
      );
    }

    putCachedAdminSession_(
      token,
      cached
    );

    return {
      email:
        cached.email,
      name:
        cached.name || '',
      role:
        cached.role || 'Admin',
      mustChangePassword:
        Boolean(
          cached.mustChangePassword
        )
    };
  }

  /*
   * Fallback path: cache miss, Apps Script restart, or cache eviction.
   * Read the Sheets once, validate, then repopulate cache.
   */
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sessionsSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
      ADMIN_SESSION_HEADERS
    );

  const tokenHash =
    hashAdminSessionToken_(
      token
    );

  const session =
    findAdminSessionByHash_(
      sessionsSheet,
      tokenHash
    );

  if (
    !session ||
    String(
      session.values.Status || ''
    ).trim() !==
      'Active'
  ) {
    throw new Error(
      'Administrator session has expired. Sign in again.'
    );
  }

  const expiresAt =
    parseAdminDate_(
      session.values[
        'Expires At'
      ]
    );

  if (
    !expiresAt ||
    expiresAt.getTime() <=
      Date.now()
  ) {
    setAdminSessionValue_(
      sessionsSheet,
      session,
      'Status',
      'Expired'
    );

    throw new Error(
      'Administrator session has expired. Sign in again.'
    );
  }

  const usersSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.USERS_SHEET,
      ADMIN_USER_HEADERS
    );

  const user =
    findAdminUserByEmail_(
      usersSheet,
      String(
        session.values[
          'Admin Email'
        ] || ''
      )
    );

  if (
    !user ||
    String(
      user.values.Status || ''
    ).trim() !==
      'Active'
  ) {
    setAdminSessionValue_(
      sessionsSheet,
      session,
      'Status',
      'Revoked'
    );

    throw new Error(
      'Administrator access is no longer active.'
    );
  }

  const mustChangePassword =
    String(
      user.values[
        'Must Change Password'
      ] || ''
    )
      .trim()
      .toLowerCase() ===
      'yes';

  if (
    mustChangePassword &&
    !(
      options &&
      options.allowPasswordChangeOnly
    )
  ) {
    throw new Error(
      'You must change your temporary password before using the Admin Portal.'
    );
  }

  const cachedSession = {
    email:
      normalizeAdminEmail_(
        user.values.Email
      ),
    name:
      String(
        user.values.Name || ''
      ),
    role:
      String(
        user.values.Role ||
        'Admin'
      ),
    mustChangePassword:
      mustChangePassword,
    issuedAt:
      Date.now()
  };

  putCachedAdminSession_(
    token,
    cachedSession
  );

  /*
   * Do not update Last Activity / Expires At on every click.
   * The browser session remains fast through CacheService.
   * The sheet continues to be the fallback/audit store.
   */
  return cachedSession;
}

function createAdminSession_(
  email,
  role,
  mustChangePassword
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sessionsSheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
      ADMIN_SESSION_HEADERS
    );

  const token =
    generateAdminSecret_(48) +
    generateAdminSecret_(48);

  const now =
    new Date();

  const expiresAt =
    new Date(
      now.getTime() +
      ADMIN_AUTH_CONFIG
        .SESSION_IDLE_HOURS *
      60 * 60 * 1000
    );

  const row = [
    hashAdminSessionToken_(
      token
    ),
    normalizeAdminEmail_(
      email
    ),
    String(
      role || 'Admin'
    ),
    now,
    now,
    expiresAt,
    'Active'
  ];

  sessionsSheet.appendRow(
    row
  );

  putCachedAdminSession_(
    token,
    {
      email:
        normalizeAdminEmail_(
          email
        ),
      name:
        '',
      role:
        String(
          role || 'Admin'
        ),
      mustChangePassword:
        Boolean(
          mustChangePassword
        ),
      issuedAt:
        now.getTime()
    }
  );

  return {
    token:
      token,
    mustChangePassword:
      Boolean(
        mustChangePassword
      )
  };
}


function invalidateAdminSession_(
  sessionToken
) {
  const token =
    String(
      sessionToken || ''
    ).trim();

  if (!token) {
    return;
  }

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
      ADMIN_SESSION_HEADERS
    );

  const session =
    findAdminSessionByHash_(
      sheet,
      hashAdminSessionToken_(
        token
      )
    );

  crffnScriptCache_()
    .remove(
      adminSessionCacheKey_(
        token
      )
    );

  if (session) {
    setAdminSessionValue_(
      sheet,
      session,
      'Status',
      'Logged Out'
    );
  }
}


function invalidateAdminSessionsForEmail_(
  email
) {
  const normalizedEmail =
    normalizeAdminEmail_(
      email
    );

  revokeCachedAdminSessionsForEmail_(
    normalizedEmail
  );

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ensureAdminAuthSheet_(
      ss,
      ADMIN_AUTH_CONFIG.SESSIONS_SHEET,
      ADMIN_SESSION_HEADERS
    );

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        ADMIN_SESSION_HEADERS.length
      )
      .getValues();

  const emailIndex =
    ADMIN_SESSION_HEADERS.indexOf(
      'Admin Email'
    );

  const statusIndex =
    ADMIN_SESSION_HEADERS.indexOf(
      'Status'
    );

  values.forEach(
    function(row, index) {
      if (
        normalizeAdminEmail_(
          row[emailIndex]
        ) ===
          normalizedEmail &&
        String(
          row[statusIndex] || ''
        ).trim() ===
          'Active'
      ) {
        sheet
          .getRange(
            index + 2,
            statusIndex + 1
          )
          .setValue(
            'Revoked'
          );
      }
    }
  );
}


function createAdminUserRecord_(
  usersSheet,
  options
) {
  const salt =
    generateAdminSecret_(
      32
    );

  const now =
    new Date();

  usersSheet.appendRow([
    String(
      options.name || ''
    ).trim(),
    normalizeAdminEmail_(
      options.email
    ),
    hashAdminPassword_(
      String(
        options.password || ''
      ),
      salt
    ),
    salt,
    String(
      options.role ||
      'Admin'
    ).trim(),
    String(
      options.status ||
      'Active'
    ).trim(),
    options.mustChangePassword
      ? 'Yes'
      : 'No',
    0,
    '',
    '',
    '',
    now,
    String(
      options.createdBy ||
      'System'
    ).trim()
  ]);
}


function registerAdminLoginFailure_(
  usersSheet,
  user
) {
  const current =
    Number(
      user.values[
        'Failed Attempts'
      ] || 0
    );

  const next =
    current + 1;

  setAdminUserValue_(
    usersSheet,
    user,
    'Failed Attempts',
    next
  );

  if (
    next >=
      ADMIN_AUTH_CONFIG
        .MAX_FAILED_ATTEMPTS
  ) {
    const lockedUntil =
      new Date(
        Date.now() +
        ADMIN_AUTH_CONFIG
          .LOCK_MINUTES *
        60 * 1000
      );

    setAdminUserValue_(
      usersSheet,
      user,
      'Locked Until',
      lockedUntil
    );

    setAdminUserValue_(
      usersSheet,
      user,
      'Failed Attempts',
      0
    );
  }
}


function clearAdminLoginFailures_(
  usersSheet,
  user
) {
  setAdminUserValue_(
    usersSheet,
    user,
    'Failed Attempts',
    0
  );

  setAdminUserValue_(
    usersSheet,
    user,
    'Locked Until',
    ''
  );
}


function hashAdminPassword_(
  password,
  salt
) {
  const pepper =
    String(
      PropertiesService
        .getScriptProperties()
        .getProperty(
          ADMIN_AUTH_CONFIG.PROP_PEPPER
        ) || ''
    );

  if (!pepper) {
    throw new Error(
      'Admin authentication has not been configured. Run setupAdminAuthentication() from the Apps Script editor.'
    );
  }

  const message =
    String(
      salt || ''
    ) +
    ':' +
    String(
      password || ''
    );

  const signature =
    Utilities
      .computeHmacSha256Signature(
        message,
        pepper
      );

  return bytesToHex_(
    signature
  );
}


function hashAdminSessionToken_(
  token
) {
  return bytesToHex_(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(
        token || ''
      )
    )
  );
}


function bytesToHex_(
  bytes
) {
  return bytes
    .map(
      function(value) {
        const normalized =
          (
            value < 0
              ? value + 256
              : value
          )
            .toString(16)
            .padStart(
              2,
              '0'
            );

        return normalized;
      }
    )
    .join('');
}


function constantTimeStringEquals_(
  left,
  right
) {
  const a =
    String(
      left || ''
    );

  const b =
    String(
      right || ''
    );

  let mismatch =
    a.length ^
    b.length;

  const maxLength =
    Math.max(
      a.length,
      b.length
    );

  for (
    let index = 0;
    index < maxLength;
    index++
  ) {
    const aCode =
      index < a.length
        ? a.charCodeAt(index)
        : 0;

    const bCode =
      index < b.length
        ? b.charCodeAt(index)
        : 0;

    mismatch |=
      aCode ^
      bCode;
  }

  return mismatch === 0;
}


function validateNewAdminPassword_(
  password
) {
  const value =
    String(
      password || ''
    );

  if (
    value.length <
      ADMIN_AUTH_CONFIG
        .MIN_PASSWORD_LENGTH
  ) {
    throw new Error(
      'Password must be at least ' +
      ADMIN_AUTH_CONFIG
        .MIN_PASSWORD_LENGTH +
      ' characters.'
    );
  }

  if (
    !/[A-Z]/.test(
      value
    ) ||
    !/[a-z]/.test(
      value
    ) ||
    !/[0-9]/.test(
      value
    ) ||
    !/[^A-Za-z0-9]/.test(
      value
    )
  ) {
    throw new Error(
      'Password must contain an uppercase letter, lowercase letter, number and special character.'
    );
  }
}


function getAdminDefaultPassword_() {
  const password =
    String(
      PropertiesService
        .getScriptProperties()
        .getProperty(
          ADMIN_AUTH_CONFIG
            .PROP_DEFAULT_PASSWORD
        ) || ''
    );

  if (!password) {
    throw new Error(
      'The admin temporary/default password has not been configured. Run setupAdminAuthentication() from the Apps Script editor.'
    );
  }

  return password;
}


function generateAdminSecret_(
  length
) {
  let value =
    '';

  while (
    value.length <
    length
  ) {
    value +=
      Utilities
        .getUuid()
        .replace(
          /-/g,
          ''
        );
  }

  return value.substring(
    0,
    length
  );
}


function normalizeAdminEmail_(
  email
) {
  return String(
    email || ''
  )
    .trim()
    .toLowerCase();
}


function ensureAdminAuthSheet_(
  ss,
  sheetName,
  headers
) {
  let sheet =
    ss.getSheetByName(
      sheetName
    );

  if (!sheet) {
    sheet =
      ss.insertSheet(
        sheetName
      );
  }

  const requiredColumns =
    headers.length;

  if (
    sheet.getMaxColumns() <
    requiredColumns
  ) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns -
      sheet.getMaxColumns()
    );
  }

  const existingHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        requiredColumns
      )
      .getDisplayValues()[0];

  const needsHeaders =
    headers.some(
      function(header, index) {
        return (
          String(
            existingHeaders[index] ||
            ''
          ).trim() !==
          header
        );
      }
    );

  if (
    sheet.getLastRow() === 0 ||
    needsHeaders
  ) {
    sheet
      .getRange(
        1,
        1,
        1,
        requiredColumns
      )
      .setValues([
        headers
      ]);

    sheet
      .getRange(
        1,
        1,
        1,
        requiredColumns
      )
      .setFontWeight(
        'bold'
      );

    sheet.setFrozenRows(
      1
    );
  }

  return sheet;
}


function findAdminUserByEmail_(
  sheet,
  email
) {
  const normalizedEmail =
    normalizeAdminEmail_(
      email
    );

  const lastRow =
    sheet.getLastRow();

  if (
    !normalizedEmail ||
    lastRow < 2
  ) {
    return null;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        ADMIN_USER_HEADERS.length
      )
      .getValues();

  const emailIndex =
    ADMIN_USER_HEADERS.indexOf(
      'Email'
    );

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    if (
      normalizeAdminEmail_(
        values[index][
          emailIndex
        ]
      ) ===
        normalizedEmail
    ) {
      return makeAdminRowRecord_(
        ADMIN_USER_HEADERS,
        values[index],
        index + 2
      );
    }
  }

  return null;
}


function findAdminSessionByHash_(
  sheet,
  tokenHash
) {
  const lastRow =
    sheet.getLastRow();

  if (
    !tokenHash ||
    lastRow < 2
  ) {
    return null;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        ADMIN_SESSION_HEADERS.length
      )
      .getValues();

  const hashIndex =
    ADMIN_SESSION_HEADERS.indexOf(
      'Session Token Hash'
    );

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    if (
      String(
        values[index][
          hashIndex
        ] || ''
      ) ===
        tokenHash
    ) {
      return makeAdminRowRecord_(
        ADMIN_SESSION_HEADERS,
        values[index],
        index + 2
      );
    }
  }

  return null;
}


function makeAdminRowRecord_(
  headers,
  row,
  rowNumber
) {
  const values =
    {};

  headers.forEach(
    function(header, index) {
      values[header] =
        row[index];
    }
  );

  return {
    rowNumber:
      rowNumber,
    values:
      values
  };
}


function setAdminUserValue_(
  sheet,
  record,
  header,
  value
) {
  const column =
    ADMIN_USER_HEADERS.indexOf(
      header
    );

  if (column < 0) {
    throw new Error(
      'Unknown Admin Users field: ' +
      header
    );
  }

  sheet
    .getRange(
      record.rowNumber,
      column + 1
    )
    .setValue(
      value
    );

  record.values[header] =
    value;
}


function setAdminSessionValue_(
  sheet,
  record,
  header,
  value
) {
  const column =
    ADMIN_SESSION_HEADERS.indexOf(
      header
    );

  if (column < 0) {
    throw new Error(
      'Unknown Admin Sessions field: ' +
      header
    );
  }

  sheet
    .getRange(
      record.rowNumber,
      column + 1
    )
    .setValue(
      value
    );

  record.values[header] =
    value;
}


function parseAdminDate_(
  value
) {
  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {
    return value;
  }

  if (!value) {
    return null;
  }

  const parsed =
    new Date(
      value
    );

  return isNaN(
    parsed.getTime()
  )
    ? null
    : parsed;
}


/**
 * Protects setup from anonymous web-app execution.
 */
function assertAdminAuthEditorSetup_() {
  const activeEmail =
    normalizeAdminEmail_(
      Session
        .getActiveUser()
        .getEmail()
    );

  const effectiveEmail =
    normalizeAdminEmail_(
      Session
        .getEffectiveUser()
        .getEmail()
    );

  if (
    !activeEmail ||
    !effectiveEmail ||
    activeEmail !==
      effectiveEmail
  ) {
    throw new Error(
      'Run setupAdminAuthentication() directly from the Apps Script editor using the project owner account.'
    );
  }
}
function showAdminDefaultPassword() {
  const password =
    getAdminDefaultPassword_();

  Logger.log(
    'DEFAULT ADMIN PASSWORD: ' +
    password
  );

  return password;
}