/* ══════════════════════════════════════════════════════════
   시애틀지구촌교회 PWA — Service Worker
   오프라인 캐싱 + 백그라운드 갱신

   전략
     · 앱 셸(HTML)   : network-first  → 오프라인이면 캐시
     · 폰트/아이콘    : cache-first    (거의 안 바뀜)
     · 유튜브 썸네일  : stale-while-revalidate
     · 폼 전송(POST) : 캐시하지 않음 (항상 네트워크)

   배포 후 내용을 바꾸면 CACHE_VERSION 을 올려주세요.
   ══════════════════════════════════════════════════════════ */

const CACHE_VERSION = 'gmc-v1';
const SHELL_CACHE = CACHE_VERSION + '-shell';
const ASSET_CACHE = CACHE_VERSION + '-assets';
const IMG_CACHE   = CACHE_VERSION + '-img';

/* 앱 셸 — 개별 실패해도 설치는 계속됩니다 (파일명이 다를 수 있으므로) */
const SHELL_URLS = [
  './',
  './index.html',
  './GMC_PWA.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

const FONT_HOSTS = ['fonts.googleapis.com', 'fonts.gstatic.com'];
const IMG_HOSTS  = ['i.ytimg.com', 'img.youtube.com'];

/* ── install: 셸 선캐시 ── */
self.addEventListener('install', (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(SHELL_CACHE);
    await Promise.allSettled(
      SHELL_URLS.map((url) =>
        cache.add(new Request(url, { cache: 'reload' })).catch(() => {})
      )
    );
    self.skipWaiting();
  })());
});

/* ── activate: 옛 캐시 정리 ── */
self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(
      keys.filter((k) => !k.startsWith(CACHE_VERSION)).map((k) => caches.delete(k))
    );
    await self.clients.claim();
  })());
});

/* ── 메시지: 앱에서 즉시 업데이트 요청 ── */
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

/* ── helpers ── */
async function networkFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  try {
    const fresh = await fetch(request);
    if (fresh && fresh.ok) cache.put(request, fresh.clone());
    return fresh;
  } catch (_) {
    const cached = await cache.match(request);
    if (cached) return cached;
    // 내비게이션이면 셸이라도 돌려줍니다
    if (request.mode === 'navigate') {
      const shell =
        (await cache.match('./')) ||
        (await cache.match('./index.html')) ||
        (await cache.match('./GMC_PWA.html'));
      if (shell) return shell;
    }
    throw _;
  }
}

async function cacheFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  if (cached) return cached;
  const fresh = await fetch(request);
  if (fresh && (fresh.ok || fresh.type === 'opaque')) cache.put(request, fresh.clone());
  return fresh;
}

async function staleWhileRevalidate(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  const fetching = fetch(request)
    .then((res) => {
      if (res && (res.ok || res.type === 'opaque')) cache.put(request, res.clone());
      return res;
    })
    .catch(() => null);
  return cached || (await fetching) || Response.error();
}

/* ── fetch ── */
self.addEventListener('fetch', (event) => {
  const { request } = event;

  // 폼 전송 등 GET 이 아닌 요청은 절대 건드리지 않습니다
  if (request.method !== 'GET') return;

  let url;
  try { url = new URL(request.url); } catch (_) { return; }
  if (!url.protocol.startsWith('http')) return;

  // FormSubmit 은 항상 네트워크로
  if (url.hostname.endsWith('formsubmit.co')) return;

  // 유튜브 썸네일
  if (IMG_HOSTS.includes(url.hostname)) {
    event.respondWith(staleWhileRevalidate(request, IMG_CACHE));
    return;
  }

  // 구글 폰트
  if (FONT_HOSTS.includes(url.hostname)) {
    event.respondWith(cacheFirst(request, ASSET_CACHE));
    return;
  }

  // 같은 출처
  if (url.origin === self.location.origin) {
    if (request.mode === 'navigate' || url.pathname.endsWith('.html')) {
      event.respondWith(networkFirst(request, SHELL_CACHE));
    } else {
      event.respondWith(cacheFirst(request, ASSET_CACHE));
    }
    return;
  }

  // 그 외 외부 요청(교회 홈페이지, 유튜브 등)은 그대로 통과
});
