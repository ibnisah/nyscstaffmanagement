// Device fingerprinting using only browser-accessible properties.
// No hardware identifiers, IMEI, or MAC are used.

const Fingerprint = (function () {
  function getCanvasFingerprint() {
    try {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) return 'no-canvas';

      canvas.width = 200;
      canvas.height = 50;
      ctx.textBaseline = 'top';
      ctx.font = "14px 'Arial'";
      ctx.textBaseline = 'alphabetic';
      ctx.fillStyle = '#f60';
      ctx.fillRect(0, 0, 200, 50);
      ctx.fillStyle = '#069';
      ctx.fillText('NYSC-ATTENDANCE-FP', 2, 15);
      ctx.strokeStyle = 'rgba(120, 186, 176, 0.5)';
      ctx.strokeRect(5, 5, 180, 40);
      return canvas.toDataURL();
    } catch (e) {
      return 'canvas-error';
    }
  }

  function getRawFingerprintString() {
    const nav = window.navigator || {};
    const screenObj = window.screen || {};
    const parts = [
      nav.userAgent || '',
      screenObj.width || '',
      screenObj.height || '',
      Intl.DateTimeFormat().resolvedOptions().timeZone || '',
      nav.language || nav.userLanguage || '',
      getCanvasFingerprint(),
    ];
    return parts.join('||');
  }

  async function sha256(text) {
    if (window.crypto && window.crypto.subtle) {
      const encoder = new TextEncoder();
      const data = encoder.encode(text);
      const hashBuffer = await window.crypto.subtle.digest('SHA-256', data);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
    }
    // Fallback: reject rather than weaken security
    throw new Error('Secure hashing is not supported on this device/browser.');
  }

  async function getHashedFingerprint() {
    const raw = getRawFingerprintString();
    return sha256(raw);
  }

  return {
    getHashedFingerprint,
  };
})();

