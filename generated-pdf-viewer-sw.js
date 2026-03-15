const GENERATED_PDF_CACHE = 'csd1-generated-pdf-v1';
const GENERATED_PDF_SEGMENT = '/__generated_pdf__/';

function encodeDispositionFilename(fileName) {
  return encodeURIComponent(String(fileName || 'document.pdf'))
    .replace(/['()]/g, escape)
    .replace(/\*/g, '%2A');
}

function asciiFallbackFileName(fileName) {
  const sanitized = String(fileName || 'document.pdf')
    .replace(/[^\x20-\x7E]+/g, '_')
    .replace(/["\\]/g, '_')
    .trim();
  return sanitized || 'document.pdf';
}

function buildPdfResponse(buffer, fileName) {
  const headers = new Headers({
    'Content-Type': 'application/pdf',
    'Cache-Control': 'no-store, max-age=0',
    'Content-Disposition': `inline; filename="${asciiFallbackFileName(fileName)}"; filename*=UTF-8''${encodeDispositionFilename(fileName)}`,
  });
  return new Response(buffer, { headers });
}

self.addEventListener('install', (event) => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener('activate', (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('message', (event) => {
  const data = event.data || {};
  const port = event.ports && event.ports[0];
  if (!port) return;

  if (data.type === 'STORE_GENERATED_PDF') {
    event.waitUntil((async () => {
      try {
        const cache = await caches.open(GENERATED_PDF_CACHE);
        await cache.put(data.url, buildPdfResponse(data.buffer, data.fileName));
        port.postMessage({ ok:true, url:data.url });
      } catch (error) {
        port.postMessage({ ok:false, error:error && error.message ? error.message : String(error) });
      }
    })());
    return;
  }

  if (data.type === 'REMOVE_GENERATED_PDF') {
    event.waitUntil((async () => {
      try {
        const cache = await caches.open(GENERATED_PDF_CACHE);
        await cache.delete(data.url);
        port.postMessage({ ok:true });
      } catch (error) {
        port.postMessage({ ok:false, error:error && error.message ? error.message : String(error) });
      }
    })());
  }
});

self.addEventListener('fetch', (event) => {
  const requestUrl = new URL(event.request.url);
  if (!requestUrl.pathname.includes(GENERATED_PDF_SEGMENT)) {
    return;
  }
  event.respondWith((async () => {
    const cache = await caches.open(GENERATED_PDF_CACHE);
    const cached = await cache.match(event.request, { ignoreSearch:false });
    if (cached) return cached;
    return new Response('Not found', { status:404, statusText:'Generated PDF Not Found' });
  })());
});
