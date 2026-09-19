/* Nebras Warehouse — isolated robust camera scanner.
   Scanner-only enhancement: inventory, sync, Excel bridge and warehouse operations are untouched. */
(() => {
  'use strict';

  const coreOpenScanner = window.openScanner;
  const coreCloseScanner = window.closeScanner;
  if (typeof coreOpenScanner !== 'function' || typeof coreCloseScanner !== 'function') return;

  let sessionId = 0;
  let handled = false;
  let robustRaf = null;
  let fallbackTimer = null;
  let zxingControls = null;
  let zxingLoadPromise = null;

  const desiredFormats = ['qr_code', 'code_128', 'code_39', 'ean_13', 'ean_8'];

  function ensureOverlay() {
    if (!document.querySelector('#scannerRobustStyle')) {
      const style = document.createElement('style');
      style.id = 'scannerRobustStyle';
      style.textContent = [
        '#scanStatusOverlay{position:absolute;left:10px;right:10px;bottom:10px;z-index:4;',
        'background:rgba(16,34,47,.82);color:#fff;border-radius:10px;padding:8px 10px;',
        'font-size:.78rem;line-height:1.55;text-align:center;pointer-events:none}',
        '#scanner .scanline{z-index:3}'
      ].join('');
      document.head.appendChild(style);
    }
    let overlay = document.querySelector('#scanStatusOverlay');
    if (!overlay) {
      const camera = document.querySelector('#scanner .camera');
      if (camera) {
        overlay = document.createElement('div');
        overlay.id = 'scanStatusOverlay';
        overlay.textContent = 'در حال آماده‌سازی دوربین…';
        camera.appendChild(overlay);
      }
    }
    return overlay;
  }

  function setStatus(message) {
    const text = String(message || '');
    const p = document.querySelector('#scanStatus');
    if (p) p.textContent = text;
    const overlay = ensureOverlay();
    if (overlay) overlay.textContent = text;
  }

  function stopRobustEngines() {
    clearTimeout(fallbackTimer);
    fallbackTimer = null;
    if (robustRaf) cancelAnimationFrame(robustRaf);
    robustRaf = null;
    try { zxingControls?.stop?.(); } catch {}
    zxingControls = null;
  }

  function acceptCode(id, raw) {
    if (id !== sessionId || handled || !raw) return;
    handled = true;
    stopRobustEngines();
    applyScan(String(raw));
  }

  async function improveFocus() {
    try {
      const track = stream?.getVideoTracks?.()[0];
      const caps = track?.getCapabilities?.();
      if (track && Array.isArray(caps?.focusMode) && caps.focusMode.includes('continuous')) {
        await track.applyConstraints({advanced:[{focusMode:'continuous'}]});
      }
    } catch {}
  }

  async function createAdaptiveDetector() {
    if (!('BarcodeDetector' in window)) return null;
    let formats = desiredFormats.slice();
    try {
      if (typeof BarcodeDetector.getSupportedFormats === 'function') {
        const supported = await BarcodeDetector.getSupportedFormats();
        const allowed = desiredFormats.filter(x => supported.includes(x));
        if (allowed.length) formats = allowed;
      }
    } catch {}
    try {
      return new BarcodeDetector({formats});
    } catch {
      try { return new BarcodeDetector(); } catch { return null; }
    }
  }

  async function startAdaptiveNative(id) {
    const detector = await createAdaptiveDetector();
    if (!detector || id !== sessionId || !stream || handled) return false;

    let busy = false;
    let failures = 0;
    const video = document.querySelector('#video');

    const tick = async () => {
      if (id !== sessionId || handled || !stream) return;
      if (busy) {
        robustRaf = requestAnimationFrame(tick);
        return;
      }
      busy = true;
      try {
        let codes = [];
        try {
          codes = await detector.detect(video);
        } catch (directError) {
          if (!('createImageBitmap' in window) || !video || video.readyState < 2) throw directError;
          let bitmap = null;
          try {
            bitmap = await createImageBitmap(video);
            codes = await detector.detect(bitmap);
          } finally {
            try { bitmap?.close?.(); } catch {}
          }
        }
        failures = 0;
        if (codes?.length) {
          acceptCode(id, codes[0].rawValue);
          return;
        }
      } catch {
        failures++;
        if (failures === 1) setStatus('دوربین فعال است؛ موتور اصلی پاسخ نداد، در حال فعال‌کردن پشتیبان…');
      } finally {
        busy = false;
      }
      if (id === sessionId && !handled) robustRaf = requestAnimationFrame(tick);
    };

    tick();
    return true;
  }

  function loadExternalScript(url) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url;
      s.async = true;
      s.crossOrigin = 'anonymous';
      s.onload = () => resolve(true);
      s.onerror = () => { try { s.remove(); } catch {} reject(new Error('load failed')); };
      document.head.appendChild(s);
    });
  }

  function loadZxing() {
    if (window.ZXingBrowser?.BrowserMultiFormatReader) return Promise.resolve(true);
    if (zxingLoadPromise) return zxingLoadPromise;
    zxingLoadPromise = (async () => {
      const urls = [
        'https://cdn.jsdelivr.net/npm/@zxing/browser@0.1.4/umd/zxing-browser.min.js',
        'https://unpkg.com/@zxing/browser@0.1.4/umd/zxing-browser.min.js'
      ];
      for (const url of urls) {
        try {
          await loadExternalScript(url);
          if (window.ZXingBrowser?.BrowserMultiFormatReader) return true;
        } catch {}
      }
      return false;
    })();
    return zxingLoadPromise;
  }

  async function startZxingFallback(id) {
    if (id !== sessionId || handled || !stream) return;
    setStatus('در حال راه‌اندازی موتور پشتیبان بارکد…');
    const loaded = await loadZxing();
    if (id !== sessionId || handled || !stream) return;
    if (!loaded) {
      setStatus('موتور پشتیبان بارگذاری نشد؛ اینترنت را بررسی کنید یا کد را دستی وارد کنید.');
      return;
    }

    try {
      const Reader = window.ZXingBrowser.BrowserMultiFormatReader;
      const reader = new Reader(undefined, {
        delayBetweenScanAttempts: 120,
        delayBetweenScanSuccess: 500
      });
      const video = document.querySelector('#video');
      const controls = await reader.decodeFromStream(stream, video, (result) => {
        if (!result || id !== sessionId || handled) return;
        const raw = typeof result.getText === 'function' ? result.getText() : (result.text || String(result));
        acceptCode(id, raw);
      });
      if (id !== sessionId || handled) {
        try { controls?.stop?.(); } catch {}
        return;
      }
      zxingControls = controls;
      setStatus('اسکن پشتیبان فعال است؛ QR یا بارکد را ثابت مقابل دوربین نگه دارید.');
    } catch {
      setStatus('دوربین فعال است اما موتور پشتیبان شروع نشد؛ صفحه را یک‌بار تازه‌سازی کنید.');
    }
  }

  window.closeScanner = function () {
    sessionId++;
    handled = true;
    stopRobustEngines();
    coreCloseScanner();
  };

  window.openScanner = async function (target) {
    const id = ++sessionId;
    handled = false;
    stopRobustEngines();
    ensureOverlay();
    setStatus('در حال آماده‌سازی دوربین…');

    await coreOpenScanner(target);

    if (id !== sessionId || handled) return;
    if (!stream) {
      const current = document.querySelector('#scanStatus')?.textContent || 'دوربین در دسترس نیست.';
      setStatus(current);
      return;
    }

    await improveFocus();
    if (id !== sessionId || handled || !stream) return;

    setStatus('دوربین فعال است؛ در حال شناسایی QR یا بارکد…');
    startAdaptiveNative(id).catch(() => {});

    // If the browser's native detector stopped working, start an isolated multi-format fallback.
    fallbackTimer = setTimeout(() => {
      if (id === sessionId && !handled && stream) startZxingFallback(id);
    }, 1400);
  };
})();