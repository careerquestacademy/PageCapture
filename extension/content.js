// content.js — PageCapture Firefox Extension
// One job: extract full page content and send to desktop app.
// No scrolling, no selection, no system hooks. Just read and send.

(function () {
  if (window.__pageCaptureV3) return;
  window.__pageCaptureV3 = true;

  const SKIP_TAGS = new Set([
    'script', 'style', 'noscript', 'iframe',
    'link', 'meta', 'head', 'svg', 'canvas'
  ]);

  const NOISE_SELECTORS = [
    'nav', 'header', 'footer',
    '[role="navigation"]', '[role="banner"]',
    '.nav', '.navbar', '.menu', '.sidebar',
    '.ad', '.ads', '.advertisement',
    '.cookie-banner', '.cookie-notice',
    '.popup', '.modal', '.overlay'
  ];

  // Convert an image to base64 data URI
  async function imageToBase64(src) {
    return new Promise(resolve => {
      try {
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            canvas.width  = img.naturalWidth  || img.width  || 100;
            canvas.height = img.naturalHeight || img.height || 100;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            resolve(canvas.toDataURL('image/png'));
          } catch (e) {
            resolve(null);
          }
        };
        img.onerror = () => resolve(null);
        img.src = src;
        setTimeout(() => resolve(null), 3000);
      } catch (e) {
        resolve(null);
      }
    });
  }

  // Extract all content blocks from the page
  async function extractPage() {
    const blocks = [];
    const seen   = new Set();

    const clone = document.body.cloneNode(true);
    NOISE_SELECTORS.forEach(sel => {
      try {
        clone.querySelectorAll(sel).forEach(el => el.remove());
      } catch (e) {}
    });

    const imgElements = Array.from(clone.querySelectorAll('img'));
    const imgSrcs     = imgElements.map(img =>
      img.currentSrc || img.src || img.dataset.src || '');

    const imgData = {};
    await Promise.all(
      imgSrcs.map(async (src, i) => {
        if (src && src.startsWith('http')) {
          imgData[src] = await imageToBase64(src);
        }
      })
    );

    function walk(node) {
      if (!node) return;

      const tag = node.nodeName
        ? node.nodeName.toLowerCase() : '';

      if (SKIP_TAGS.has(tag)) return;

      if (/^h[1-6]$/.test(tag)) {
        const text = (node.innerText || '').trim();
        if (text && !seen.has('h' + text)) {
          seen.add('h' + text);
          blocks.push({ type: 'heading', level: parseInt(tag[1]), text });
        }
        return;
      }

      if (tag === 'img') {
        const src = node.currentSrc || node.src || node.dataset.src || '';
        const alt = node.alt || '';
        if (src && !seen.has(src)) {
          seen.add(src);
          blocks.push({ type: 'image', src, alt, img_data: imgData[src] || null });
        }
        return;
      }

      if (tag === 'a') {
        const text = (node.innerText || '').trim();
        let   href = node.getAttribute('href') || '';
        if (href && !href.startsWith('http') && !href.startsWith('mailto')) {
          try { href = new URL(href, window.location.href).href; } catch (e) {}
        }
        if (text && href && !seen.has(text + href)) {
          seen.add(text + href);
          blocks.push({ type: 'link', text, href });
        }
        return;
      }

      if (tag === 'li') {
        const text = (node.innerText || '').trim();
        if (text && !seen.has('li' + text)) {
          seen.add('li' + text);
          blocks.push({ type: 'listitem', text });
        }
        return;
      }

      if (['p', 'blockquote', 'figcaption', 'td', 'th', 'label'].includes(tag)) {
        const text = (node.innerText || '').trim();
        if (text && text.length > 1 && !seen.has(text)) {
          seen.add(text);
          blocks.push({ type: 'text', text });
        }
        return;
      }

      for (const child of Array.from(node.childNodes)) {
        walk(child);
      }
    }

    walk(clone);
    return blocks;
  }

  // Send content to desktop app via background script (Native Messaging)
  async function sendToApp() {
    const blocks = await extractPage();
    const payload = {
      title:  document.title || window.location.hostname,
      url:    window.location.href,
      blocks
    };

    try {
      const response = await browser.runtime.sendMessage({
        action: 'sendToApp',
        payload
      });

      if (response && response.success) {
        showToast('Sent to PageCapture!', '#a6e3a1');
      } else {
        const msg = (response && response.error) || 'Unknown error';
        showToast('PageCapture app not running. Please open it first.', '#f38ba8');
        console.error('PageCapture error:', msg);
      }
    } catch (e) {
      showToast('PageCapture app not running. Please open it first.', '#f38ba8');
    }
  }

  function showToast(message, color) {
    const toast = document.createElement('div');
    Object.assign(toast.style, {
      position:     'fixed',
      bottom:       '24px',
      right:        '24px',
      background:   '#1e1e2e',
      color:        color || '#cdd6f4',
      padding:      '12px 20px',
      borderRadius: '8px',
      fontSize:     '14px',
      fontFamily:   'Arial, sans-serif',
      zIndex:       '2147483647',
      boxShadow:    '0 4px 16px rgba(0,0,0,0.4)',
      maxWidth:     '320px',
      lineHeight:   '1.4'
    });
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 4000);
  }

  browser.runtime.onMessage.addListener((msg) => {
    if (msg.action === 'capture') {
      sendToApp();
    }
  });

})();
