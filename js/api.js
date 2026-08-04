// Client ⇄ Apps Script Web App communication layer.
// All calls go through this thin wrapper for consistent error handling.

const Api = (function () {
  // IMPORTANT: Replace this with your deployed Apps Script web app URL
  // Get this URL after deploying your Apps Script as a Web App
  // Format: https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
  const BASE_URL = 'https://script.google.com/macros/s/AKfycbwprBU-at3dPlhQUP8QkiJFhrRLMFurW1ImrX0WNjTMmiqRAVWLPciB628TwNAidVl_KA/exec';

  // For testing, you can temporarily use a placeholder, but it won't work until you deploy:
  // const BASE_URL = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec';

  // Request timeout in milliseconds (2 minutes for field capture / uploads; avoid premature timeout on slow Apps Script backend)
  const REQUEST_TIMEOUT = 120000;

  // Simple cache for GET-like operations (formations, departments, etc.)
  const cache = new Map();
  const CACHE_TTL = 60000; // 1 minute cache

  function getCacheKey(action, payload) {
    // Only cache read operations
    const cacheableActions = ['listFormations', 'listDepartments', 'getRegistrationStatus', 'getAvailableModules'];
    if (!cacheableActions.includes(action)) return null;
    return `${action}_${JSON.stringify(payload || {})}`;
  }

  function getCached(key) {
    const cached = cache.get(key);
    if (!cached) return null;
    if (Date.now() - cached.timestamp > CACHE_TTL) {
      cache.delete(key);
      return null;
    }
    return cached.data;
  }

  function setCache(key, data) {
    if (key) {
      cache.set(key, { data, timestamp: Date.now() });
    }
  }

  const SENSITIVE_FIELDS = [
    'key', 'adminKey', 'token', 'adminToken', 'sessionToken', 'uploadToken', 'qrToken', 'password',
  ];

  /** Strip credentials so they never reach the browser console or an error report. */
  function redactSecrets(value) {
    if (!value || typeof value !== 'object') return value;
    if (Array.isArray(value)) return value.map(redactSecrets);
    const out = {};
    Object.keys(value).forEach(function (k) {
      const v = value[k];
      if (v && SENSITIVE_FIELDS.indexOf(k) !== -1) out[k] = '[redacted]';
      else if (v && typeof v === 'object') out[k] = redactSecrets(v);
      else out[k] = v;
    });
    return out;
  }

  // Public PRS read endpoints. Apps Script can turn a POST into a GET when it
  // redirects, which drops the body. These carry no credentials, so they are the
  // only actions safe to retry over GET with their params in the URL.
  const PRS_GET_FALLBACK_ACTIONS = [
    'prsValidateCampQr',
    'prsResolveAssignment',
    'prsListStaffAssignments',
    'prsGetStaffDashboard',
    'prsListStaffMaterials',
  ];

  /**
   * Apps Script sometimes answers a POST with a redirect that the browser follows
   * as a GET, discarding the body. For the public PRS reads we therefore repeat the
   * scalar params in the query string so doGet can still serve the request. Every
   * other action sends `action` only, so admin keys never reach a URL or server log.
   */
  function buildApiUrl(action, payload) {
    const params = new URLSearchParams();
    params.set('action', action);
    if (PRS_GET_FALLBACK_ACTIONS.indexOf(action) !== -1) {
      const p = payload || {};
      Object.keys(p).forEach(function (k) {
        if (k === 'action') return;
        const v = p[k];
        if (v == null || v === '' || typeof v === 'object') return;
        params.set(k, String(v));
      });
    }
    return BASE_URL + '?' + params.toString();
  }

  // Apps Script serves every POST result from a one-time script.googleusercontent.com
  // URL, and that handoff becomes unreliable when several requests are in flight at
  // once — it 404s, or the redirect arrives back as a GET with the body gone. Running
  // calls one at a time is what keeps the handoff stable.
  let requestChain = Promise.resolve();

  function queued(task) {
    const run = requestChain.then(task, task);
    requestChain = run.then(function () {}, function () {});
    return run;
  }

  // Observed in the field: the same call can lose its body on one attempt and lose its
  // result on the next, then succeed on a third. Each failure is independent, so a
  // couple of extra attempts converts a visible error into a slight delay.
  const MAX_ATTEMPTS = 3;

  const READ_VERBS = ['list', 'get', 'search', 'validate', 'resolve', 'export', 'download'];

  /** Leading verb of an action, e.g. prsListCamps -> "list", prsSignIn -> "sign". */
  function actionVerb(action) {
    const stripped = String(action || '').replace(/^(prs|hrm|admin)/, '');
    const lower = stripped.match(/^([a-z]+)/);
    if (lower) return lower[1].toLowerCase();
    const upper = stripped.match(/^([A-Z][a-z]+)/);
    return upper ? upper[1].toLowerCase() : '';
  }

  /**
   * Which transient failures are safe to repeat.
   *  - BODY_LOST: the redirect dropped the request body, so the server never ran the
   *    action. Safe to repeat for anything, writes included.
   *  - CONTENT_404: the script DID run and Google only failed to deliver the result,
   *    so repeating is safe for idempotent reads but not for writes.
   */
  /**
   * Writes whose result was lost in transit can be confirmed by simply repeating them:
   * the backend's own duplicate guard rejecting the second attempt is proof the first
   * one landed. Maps the action to the guard reason that confirms it.
   */
  const WRITE_CONFIRMATIONS = {
    prsSignIn: 'ALREADY_SIGNED_IN',
    prsSignOut: 'ALREADY_SIGNED_OUT',
  };

  function isRetryable(err, action) {
    if (!err) return false;
    if (err.transient === 'BODY_LOST') return true;
    if (err.transient === 'CONTENT_404') {
      return READ_VERBS.indexOf(actionVerb(action)) !== -1
        || Object.prototype.hasOwnProperty.call(WRITE_CONFIRMATIONS, action);
    }
    return false;
  }

  /**
   * Turns a duplicate-guard rejection on a retried write into the success it actually
   * represents, so a staff member is not told "already signed in" for a sign-in that
   * only failed because Google lost the response.
   */
  function confirmWriteFromGuard(action, firstErr, retryErr) {
    if (!firstErr || firstErr.transient !== 'CONTENT_404') return null;
    const expected = WRITE_CONFIRMATIONS[action];
    if (!expected || !retryErr || retryErr.reason !== expected) return null;
    return {
      success: true,
      data: (retryErr.raw && retryErr.raw.data) || {},
      message: 'Attendance recorded.',
      recoveredFromLostResponse: true,
    };
  }

  async function tryGetFallback(action, payload) {
    if (PRS_GET_FALLBACK_ACTIONS.indexOf(action) === -1) return null;
    const getRes = await queued(function () {
      return fetch(buildApiUrl(action, payload), { method: 'GET' });
    });
    if (!getRes.ok) return null;
    try {
      const data = JSON.parse(await getRes.text());
      return data && data.success === true ? data : null;
    } catch (e) {
      return null;
    }
  }

  async function postToGas(url, bodyString, options) {
    return queued(async function () {
      const timeout = options.timeout || REQUEST_TIMEOUT;
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), timeout);
      try {
        // 'follow' is required: Apps Script 302s POST responses to its content server.
        return await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: bodyString,
          redirect: 'follow',
          signal: controller.signal,
        });
      } finally {
        clearTimeout(timeoutId);
      }
    });
  }

  async function call(action, payload, options = {}) {
    // Check if BASE_URL is set
    if (!BASE_URL || BASE_URL === 'YOUR_DEPLOYED_WEB_APP_URL_HERE') {
      throw new Error('API URL not configured. Please update BASE_URL in api.js with your deployed Apps Script web app URL.');
    }

    // Check cache first (if not disabled)
    if (!options.skipCache) {
      const cacheKey = getCacheKey(action, payload);
      if (cacheKey) {
        const cached = getCached(cacheKey);
        if (cached) {
          return cached;
        }
      }
    }

    const url = buildApiUrl(action, payload);
    // Include action in the body too, so routing still works if query params are lost on redirect.
    const body = JSON.stringify(Object.assign({}, payload || {}, { action: action }));

    // Only log in development mode
    if (options.debug !== false) {
      console.log('API Call:', { action, url, payload: redactSecrets(payload) });
    }

    async function attempt() {
      const res = await postToGas(url, body, options);

      if (options.debug !== false) {
        console.log('API Response Status:', res.status, res.statusText);
      }

      if (res.status === 0) {
        throw new Error(
          'Could not reach the server (network blocked or CORS). Check the API URL in api.js and redeploy the Apps Script web app.'
        );
      }

      if (!res.ok) {
        let errorText = '';
        try { errorText = await res.text(); } catch (readErr) { /* ignore */ }

        // A 404 from the content server means the script ran but Google could not hand
        // back the result. Transient, and unrelated to anything in the request.
        if (res.status === 404 && /unable to open the file|Page not found/i.test(errorText)) {
          const contentErr = new Error(
            'Google could not return the result (404 from its content server, not from your script). '
            + 'This is usually transient — please try again in a moment.'
          );
          contentErr.transient = 'CONTENT_404';
          throw contentErr;
        }

        if (options.debug !== false) {
          console.error('API Error Response:', String(errorText).slice(0, 300));
        }
        throw new Error(`Network error: ${res.status} ${res.statusText || ''}`.trim());
      }

      const responseText = await res.text();

      let data;
      try {
        data = JSON.parse(responseText);
      } catch (parseError) {
        console.error('Failed to parse JSON response:', parseError);
        console.error('Response text:', String(responseText).slice(0, 300));
        throw new Error('Invalid response format from server. Please check server logs.');
      }

      if (options.debug !== false) {
        console.log('API Response Data (parsed):', redactSecrets(data));
      }

      if (!data || data.success !== true) {
        const message = (data && data.message) || 'Request failed.';
        const reason = (data && data.reason) || 'UNKNOWN';

        // The redirect dropped the POST body. Public PRS reads keep their params in the
        // URL so a GET can still answer them; anything else has to be sent again.
        if (/GET not supported/i.test(message) || reason === 'METHOD_NOT_ALLOWED') {
          const fallback = await tryGetFallback(action, payload);
          if (fallback) return fallback;
          const lostErr = new Error(message);
          lostErr.reason = reason;
          lostErr.raw = data;
          lostErr.transient = 'BODY_LOST';
          throw lostErr;
        }

        if (options.debug !== false) {
          console.error('API Error - Reason:', reason);
          console.error('API Error - Message:', message);
        }
        const error = new Error(message);
        error.reason = reason;
        error.raw = data;
        throw error;
      }

      return data;
    }

    try {
      let data = null;

      // Staff QR reads go over GET first. Apps Script's POST result handoff is the
      // unreliable part; doGet answers these straight from the query string, which is
      // why ?test=ping has been reliable throughout. Falls through to POST if it fails.
      if (PRS_GET_FALLBACK_ACTIONS.indexOf(action) !== -1) {
        data = await tryGetFallback(action, payload);
      }

      let firstTransient = null;
      for (let attemptNo = 1; attemptNo <= MAX_ATTEMPTS && !data; attemptNo++) {
        try {
          data = await attempt();
        } catch (err) {
          // A duplicate-guard rejection after we already lost a response proves the
          // earlier attempt landed, so report it as the success it actually was.
          const confirmed = confirmWriteFromGuard(action, firstTransient, err);
          if (confirmed) {
            data = confirmed;
            break;
          }

          if (!isRetryable(err, action) || attemptNo === MAX_ATTEMPTS) throw err;

          if (!firstTransient) firstTransient = err;
          if (options.debug !== false) {
            console.warn(
              'API transient failure (' + err.transient + ') on attempt ' + attemptNo
              + ' of ' + MAX_ATTEMPTS + ':', action
            );
          }
          await new Promise(function (resolve) { setTimeout(resolve, 400 * attemptNo); });
        }
      }

      // Cache successful responses
      const cacheKey = getCacheKey(action, payload);
      if (cacheKey && !options.skipCache) {
        setCache(cacheKey, data);
      }

      return data;
    } catch (err) {
      if (options.debug !== false) {
        console.error('API Call Error:', err);
      }

      // Handle timeout
      if (err.name === 'AbortError') {
        throw new Error('Request timeout. The server is taking too long to respond. Please try again.');
      }

      // Provide more helpful error messages
      if (err.message && (err.message.includes('Failed to fetch') || err.message.includes('NetworkError'))) {
        throw new Error('Cannot connect to server. Check:\n1. Internet connection\n2. API URL is correct\n3. Apps Script is deployed');
      }

      if (err.message && (/GET not supported/i.test(err.message) || err.reason === 'METHOD_NOT_ALLOWED')) {
        throw new Error(
          'Apps Script dropped the request body on redirect, and the retry failed too. '
          + 'Please try again in a moment.'
        );
      }

      throw err;
    }
  }

  // Clear cache function
  function clearCache() {
    cache.clear();
  }

  return {
    call,
    clearCache,
  };
})();

/**
 * Scroll every page/view to the top on open. Stops the browser from restoring
 * a previous scroll position when switching modules, tabs, or steps.
 */
const ScrollTop = (function () {
  'use strict';

  if (typeof history !== 'undefined' && 'scrollRestoration' in history) {
    history.scrollRestoration = 'manual';
  }

  function scrollRoots() {
    var seen = new Set();
    function reset(el) {
      if (!el || seen.has(el)) return;
      seen.add(el);
      el.scrollTop = 0;
      el.scrollLeft = 0;
    }
    [
      document.documentElement,
      document.body,
      document.querySelector('main'),
      document.querySelector('.page'),
      document.querySelector('.page-admin'),
      document.querySelector('.admin-layout'),
      document.querySelector('.admin-main-column'),
      document.querySelector('.admin-main-panel'),
      document.querySelector('.admin-workspace'),
      document.getElementById('moduleSelector'),
    ].forEach(reset);
    document.querySelectorAll('.workspace-content, [id$="Content"]').forEach(reset);
  }

  function toTop(options) {
    options = options || {};
    var behavior = options.behavior || 'auto';
    try {
      window.scrollTo({ top: 0, left: 0, behavior: behavior });
    } catch (e) {
      window.scrollTo(0, 0);
    }
    scrollRoots();
  }

  /** Call after async UI renders so layout shifts don't leave the page mid-scroll. */
  function afterRender(callback) {
    toTop({ behavior: 'auto' });
    requestAnimationFrame(function () {
      toTop({ behavior: 'auto' });
      requestAnimationFrame(function () {
        toTop({ behavior: 'auto' });
        setTimeout(function () {
          toTop({ behavior: 'auto' });
          if (typeof callback === 'function') callback();
        }, 50);
      });
    });
  }

  if (typeof document !== 'undefined') {
    var boot = function () { afterRender(); };
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', boot);
    } else {
      boot();
    }
    window.addEventListener('pageshow', function (ev) {
      if (ev.persisted) afterRender();
    });
    window.addEventListener('load', function () { afterRender(); });
  }

  return { toTop: toTop, afterRender: afterRender };
})();

