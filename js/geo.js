// Geolocation utilities with mandatory high-accuracy enforcement.

const Geo = (function () {
  const DEFAULT_MAX_AGE_MS = 30 * 1000;
  const DEFAULT_TIMEOUT_MS = 20 * 1000;

  function isIosSafari() {
    const ua = window.navigator.userAgent || '';
    const isIOS = /iP(hone|od|ad)/.test(ua);
    const isSafari = /Safari/.test(ua) && !/CriOS|FxiOS|OPiOS|EdgiOS/.test(ua);
    return isIOS && isSafari;
  }

  function getLocation(options = {}) {
    const {
      enableHighAccuracy = true,
      timeout = DEFAULT_TIMEOUT_MS,
      maximumAge = DEFAULT_MAX_AGE_MS,
      requiredAccuracyMeters = 50,
    } = options;

    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) {
        return reject({
          code: 'UNSUPPORTED',
          message: 'Geolocation is not supported on this device.',
        });
      }

      navigator.geolocation.getCurrentPosition(
        (pos) => {
          const { latitude, longitude, accuracy } = pos.coords;
          if (
            latitude == null ||
            longitude == null ||
            accuracy == null ||
            accuracy > requiredAccuracyMeters
          ) {
            return reject({
              code: 'ACCURACY_TOO_LOW',
              message:
                'Location accuracy is too low. Move closer to a window and try again.',
              details: { accuracy },
            });
          }
          resolve({
            lat: latitude,
            lng: longitude,
            accuracy,
          });
        },
        (err) => {
          let code = 'UNKNOWN';
          if (err.code === err.PERMISSION_DENIED) code = 'DENIED';
          if (err.code === err.POSITION_UNAVAILABLE) code = 'UNAVAILABLE';
          if (err.code === err.TIMEOUT) code = 'TIMEOUT';
          reject({
            code,
            message: err.message || 'Failed to get location.',
            isIosSafari: isIosSafari(),
          });
        },
        {
          enableHighAccuracy,
          timeout,
          maximumAge,
        }
      );
    });
  }

  return {
    getLocation,
    isIosSafari,
  };
})();

