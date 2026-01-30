// Client ⇄ Apps Script Web App communication layer.
// All calls go through this thin wrapper for consistent error handling.

const Api = (function () {
  // IMPORTANT: Replace this with your deployed Apps Script web app URL
  // Get this URL after deploying your Apps Script as a Web App
  // Format: https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
  const BASE_URL = 'https://script.google.com/macros/s/AKfycbwprBU-at3dPlhQUP8QkiJFhrRLMFurW1ImrX0WNjTMmiqRAVWLPciB628TwNAidVl_KA/exec';

  // For testing, you can temporarily use a placeholder, but it won't work until you deploy:
  // const BASE_URL = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec';

  // Request timeout in milliseconds (30 seconds)
  const REQUEST_TIMEOUT = 30000;

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

    const url = BASE_URL + '?action=' + encodeURIComponent(action);
    const body = JSON.stringify(payload || {});

    // Only log in development mode
    if (options.debug !== false) {
      console.log('API Call:', { action, url, payload });
    }

    // Create AbortController for timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), options.timeout || REQUEST_TIMEOUT);

    try {
      // Use 'text/plain' instead of 'application/json' to avoid CORS preflight issues
      // Google Apps Script Web Apps handle CORS automatically, but preflight requests (OPTIONS) can fail
      // The backend will parse the JSON string from the body using JSON.parse(e.postData.contents)
      const res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain',
        },
        body,
        signal: controller.signal,
      });

      clearTimeout(timeoutId);

      if (options.debug !== false) {
        console.log('API Response Status:', res.status, res.statusText);
      }

      if (!res.ok) {
        const errorText = await res.text();
        if (options.debug !== false) {
          console.error('API Error Response:', errorText);
        }
        throw new Error(`Network error: ${res.status} ${res.statusText}`);
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
      clearTimeout(timeoutId);

      // Handle timeout
      if (err.name === 'AbortError') {
        throw new Error('Request timeout. The server is taking too long to respond. Please try again.');
      }

      if (options.debug !== false) {
        console.error('API Call Error:', err);
      }

      // Provide more helpful error messages
      if (err.message.includes('Failed to fetch') || err.message.includes('NetworkError')) {
        throw new Error('Cannot connect to server. Check:\n1. Internet connection\n2. API URL is correct\n3. Apps Script is deployed');
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

