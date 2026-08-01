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

  // Apps Script may follow POST with GET and drop the body; keep scalar fields in
  // the query string so doGet + parseRequestBody can still route the request.
  const MAX_URL_QUERY_LEN = 1800;

  function buildApiUrl(action, payload) {
    const params = new URLSearchParams();
    params.set('action', action);
    const p = payload || {};
    let built = params.toString();
    Object.keys(p).forEach(function (k) {
      if (k === 'action') return;
      const v = p[k];
      if (v == null || v === '') return;
      if (typeof v === 'object') return;
      const next = built + (built ? '&' : '') + encodeURIComponent(k) + '=' + encodeURIComponent(String(v));
      if (next.length > MAX_URL_QUERY_LEN) return;
      params.set(k, String(v));
      built = params.toString();
    });
    return BASE_URL + '?' + params.toString();
  }

  async function tryGetFallback(action, payload) {
    const getUrl = buildApiUrl(action, payload);
    const getRes = await fetch(getUrl, { method: 'GET' });
    if (!getRes.ok) return null;
    const text = await getRes.text();
    try {
      const data = JSON.parse(text);
      return data && data.success === true ? data : null;
    } catch (e) {
      return null;
    }
  }

  async function postToGas(url, bodyString, options) {
    const timeout = options.timeout || REQUEST_TIMEOUT;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    const init = {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: bodyString,
      signal: controller.signal,
    };

    try {
      // Primary: follow redirects (required for adminLogin and most GAS calls).
      // Using redirect:'manual' alone returns status 0 (opaque) and breaks all API calls.
      let res = await fetch(url, Object.assign({}, init, { redirect: 'follow' }));

      // If GAS downgraded POST→GET, body is lost — try manual re-POST when Location is exposed.
      if (res.ok) {
        const peek = await res.clone().text();
        try {
          const peekData = JSON.parse(peek);
          if (peekData && /GET not supported/i.test(peekData.message || '')) {
            const manualRes = await fetch(url, Object.assign({}, init, { redirect: 'manual' }));
            if (manualRes.status >= 301 && manualRes.status <= 308) {
              let loc = manualRes.headers.get('Location');
              if (loc) {
                if (loc.indexOf('action=') === -1 && url.indexOf('action=') !== -1) {
                  const actionMatch = url.match(/[?&]action=([^&]+)/);
                  if (actionMatch) {
                    loc += (loc.indexOf('?') === -1 ? '?' : '&') + 'action=' + actionMatch[1];
                  }
                }
                res = await fetch(loc, init);
              }
            }
          }
        } catch (parsePeek) {
          /* use original res */
        }
      }

      clearTimeout(timeoutId);
      return res;
    } catch (err) {
      clearTimeout(timeoutId);
      throw err;
    }
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
    // Include action in body for doPost; query string carries scalars when POST→GET redirect drops body.
    const body = JSON.stringify(Object.assign({}, payload || {}, { action: action }));

    // Only log in development mode
    if (options.debug !== false) {
      console.log('API Call:', { action, url, payload });
    }

    try {
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
        if (options.debug !== false) {
          console.error('API Error Response:', errorText);
        }
        throw new Error(`Network error: ${res.status} ${res.statusText || ''}`.trim());
      }

      const responseText = await res.text();
      if (options.debug !== false) {
        console.log('API Response Text (raw):', responseText);
      }

      let data;
      try {
        data = JSON.parse(responseText);
      } catch (parseError) {
        console.error('Failed to parse JSON response:', parseError);
        console.error('Response text:', responseText);
        throw new Error('Invalid response format from server. Please check server logs.');
      }

      if (options.debug !== false) {
        console.log('API Response Data (parsed):', data);
      }

      if (!data || data.success !== true) {
        const message = (data && data.message) || 'Request failed.';
        const reason = (data && data.reason) || 'UNKNOWN';

        // GAS redirect dropped the POST body — retry as GET (params are in the URL).
        if (/GET not supported/i.test(message) || reason === 'METHOD_NOT_ALLOWED') {
          const fallback = await tryGetFallback(action, payload);
          if (fallback) return fallback;
        }

        if (options.debug !== false) {
          console.error('API Error - Reason:', reason);
          console.error('API Error - Message:', message);
          console.error('API Error - Full response:', data);
        }
        const error = new Error(message);
        error.reason = reason;
        error.raw = data;
        throw error;
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
          'Server received a GET request instead of POST. Hard-refresh this page (Ctrl+F5), ' +
          'then redeploy the Apps Script web app (Deploy → New version) if login still fails.'
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

