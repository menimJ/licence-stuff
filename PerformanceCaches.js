/**
 * ============================================================
 * CRFFN PERFORMANCE CACHE
 * ============================================================
 *
 * Adds short-lived server-side caching for read-heavy portal
 * operations. Sensitive values are never exposed to the browser
 * by this file.
 *
 * Cache is deliberately short-lived:
 * - Applicant portal summary: 45 seconds
 * - Supporting-document summary: 30 seconds
 * - Payment summary: 30 seconds
 *
 * Writes invalidate the relevant applicant cache immediately.
 */

const CRFFN_PERFORMANCE_CONFIG = Object.freeze({
  APPLICANT_PORTAL_TTL_SECONDS: 45,
  SUPPORTING_DOCUMENTS_TTL_SECONDS: 30,
  PAYMENT_TTL_SECONDS: 30,
  ADMIN_SESSION_CACHE_SECONDS: 21600,
  ADMIN_SESSION_SHEET_TOUCH_SECONDS: 600
});


function crffnScriptCache_() {
  return CacheService.getScriptCache();
}


function crffnCacheDigest_(value) {
  const bytes =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(value || ''),
      Utilities.Charset.UTF_8
    );

  return bytes
    .map(function(byte) {
      const unsigned =
        byte < 0
          ? byte + 256
          : byte;

      return (
        '0' +
        unsigned.toString(16)
      ).slice(-2);
    })
    .join('');
}


function crffnApplicantCacheBaseKey_(
  applicationId,
  secureToken
) {
  return [
    'crffn',
    'applicant',
    crffnCacheDigest_(
      String(applicationId || '') +
      '|' +
      String(secureToken || '')
    )
  ].join(':');
}


function crffnGetCachedJson_(
  key
) {
  const text =
    crffnScriptCache_()
      .get(key);

  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    crffnScriptCache_()
      .remove(key);

    return null;
  }
}


function crffnPutCachedJson_(
  key,
  value,
  ttlSeconds
) {
  if (
    value === undefined ||
    value === null
  ) {
    return;
  }

  try {
    crffnScriptCache_()
      .put(
        key,
        JSON.stringify(value),
        Math.max(
          1,
          Number(ttlSeconds) || 30
        )
      );
  } catch (error) {
    console.warn(
      'CRFFN cache write skipped: ' +
      error.message
    );
  }
}


function crffnRemoveCacheKeys_(
  keys
) {
  const validKeys =
    (keys || [])
      .filter(Boolean);

  if (!validKeys.length) {
    return;
  }

  crffnScriptCache_()
    .removeAll(validKeys);
}


function crffnApplicantPortalCacheKey_(
  applicationId,
  secureToken
) {
  return (
    crffnApplicantCacheBaseKey_(
      applicationId,
      secureToken
    ) +
    ':portal'
  );
}


function crffnSupportingDocumentsCacheKey_(
  applicationId,
  secureToken
) {
  return (
    crffnApplicantCacheBaseKey_(
      applicationId,
      secureToken
    ) +
    ':documents'
  );
}


function crffnPaymentCacheKey_(
  applicationId,
  secureToken
) {
  return (
    crffnApplicantCacheBaseKey_(
      applicationId,
      secureToken
    ) +
    ':payment'
  );
}


/**
 * Call after any applicant-side write that may change
 * visible portal data.
 */
function invalidateApplicantPerformanceCache_(
  applicationId,
  secureToken
) {
  crffnRemoveCacheKeys_([
    crffnApplicantPortalCacheKey_(
      applicationId,
      secureToken
    ),
    crffnSupportingDocumentsCacheKey_(
      applicationId,
      secureToken
    ),
    crffnPaymentCacheKey_(
      applicationId,
      secureToken
    )
  ]);
}


/**
 * Admin-side helper when only the Application ID is known.
 *
 * This bumps a generation number. Applicant cache wrappers include
 * that number, so cached portal data automatically becomes stale
 * after an admin action without needing the secure token.
 */
function invalidateApplicantCacheByApplicationId_(
  applicationId
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  if (!cleanId) {
    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      'CRFFN_APP_CACHE_VERSION_' +
      cleanId,
      String(Date.now())
    );
}


function getApplicantCacheVersion_(
  applicationId
) {
  return String(
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'CRFFN_APP_CACHE_VERSION_' +
        String(
          applicationId || ''
        ).trim()
      ) ||
      '0'
  );
}


function crffnVersionedApplicantKey_(
  baseKey,
  applicationId
) {
  return (
    baseKey +
    ':v:' +
    getApplicantCacheVersion_(
      applicationId
    )
  );
}


/**
 * Optional manual cache clear for testing.
 */
function clearCRFFNPerformanceCache() {
  CacheService
    .getScriptCache()
    .removeAll([]);

  console.log(
    'CRFFN performance cache clear requested. ' +
    'Versioned keys will naturally expire within 45 seconds.'
  );

  return true;
}


/* ============================================================
 * ADMIN APPLICATION DETAIL CACHE
 * ============================================================
 *
 * Display reads may use this short-lived cache.
 * Critical state transitions still validate the authoritative Sheet.
 */

const CRFFN_ADMIN_APPLICATION_CACHE_TTL_SECONDS =
  45;


function adminApplicationCacheVersionKey_(
  applicationId
) {
  return (
    'CRFFN_ADMIN_APP_CACHE_VERSION_' +
    String(
      applicationId || ''
    ).trim()
  );
}


function getAdminApplicationCacheVersion_(
  applicationId
) {
  return String(
    PropertiesService
      .getScriptProperties()
      .getProperty(
        adminApplicationCacheVersionKey_(
          applicationId
        )
      ) ||
      '0'
  );
}


function adminApplicationDetailCacheKey_(
  applicationId
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  return [
    'crffn',
    'admin',
    'application',
    crffnCacheDigest_(
      cleanId
    ),
    'v',
    getAdminApplicationCacheVersion_(
      cleanId
    )
  ].join(':');
}


function getCachedAdminApplicationDetail_(
  applicationId
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  if (!cleanId) {
    return null;
  }

  return crffnGetCachedJson_(
    adminApplicationDetailCacheKey_(
      cleanId
    )
  );
}


function putCachedAdminApplicationDetail_(
  applicationId,
  application
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  if (
    !cleanId ||
    !application
  ) {
    return;
  }

  crffnPutCachedJson_(
    adminApplicationDetailCacheKey_(
      cleanId
    ),
    application,
    CRFFN_ADMIN_APPLICATION_CACHE_TTL_SECONDS
  );
}


/**
 * Version bump is safer than trying to discover/remove every
 * previously generated cache key.
 */
function invalidateAdminApplicationCache_(
  applicationId
) {
  const cleanId =
    String(
      applicationId || ''
    ).trim();

  if (!cleanId) {
    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      adminApplicationCacheVersionKey_(
        cleanId
      ),
      String(
        Date.now()
      )
    );
}
