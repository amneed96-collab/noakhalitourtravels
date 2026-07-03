// ============================================================
// API হেল্পার - Google Apps Script ব্যাকএন্ডের সাথে যোগাযোগ (ক্যাশিং সহ, ফাস্ট)
// ============================================================

const _cache = new Map();
const CACHE_TTL_MS = 20000; // ২০ সেকেন্ড ক্যাশ - বারবার একই ডেটা আনতে দেরি হবে না

async function apiGet(action, params, opts) {
  const url = new URL(API_URL);
  url.searchParams.set('action', action);
  if (params) {
    Object.keys(params).forEach(k => {
      if (params[k] !== undefined && params[k] !== null) url.searchParams.set(k, params[k]);
    });
  }
  const cacheKey = url.toString();
  const useCache = !(opts && opts.noCache);

  if (useCache && _cache.has(cacheKey)) {
    const cached = _cache.get(cacheKey);
    if (Date.now() - cached.time < CACHE_TTL_MS) return cached.data;
  }

  const res = await fetch(cacheKey);
  const data = await res.json();
  if (useCache) _cache.set(cacheKey, { data, time: Date.now() });
  return data;
}

async function apiPost(action, body) {
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...body })
  });
  const data = await res.json();
  clearApiCache(); // ডেটা পরিবর্তন হলে ক্যাশ বাতিল
  return data;
}

function clearApiCache() {
  _cache.clear();
}

// ---------- সেশন ম্যানেজমেন্ট ----------

const Session = {
  KEY: 'tour_admin_session',
  set(data) { sessionStorage.setItem(this.KEY, JSON.stringify(data)); },
  get() { const raw = sessionStorage.getItem(this.KEY); return raw ? JSON.parse(raw) : null; },
  clear() { sessionStorage.removeItem(this.KEY); },
  isLoggedIn() { return !!this.get(); }
};

// ---------- সাধারণ ইউটিলিটি ----------

function formatTaka(amount) {
  const n = Number(amount) || 0;
  return '৳' + n.toLocaleString('bn-BD');
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('bn-BD', { year: 'numeric', month: 'long', day: 'numeric' });
}

function toBengaliNumber(num) {
  const bn = ['০','১','২','৩','৪','৫','৬','৭','৮','৯'];
  return String(num).replace(/[0-9]/g, d => bn[d]);
}

function qs(sel, root) { return (root || document).querySelector(sel); }
function qsa(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

function showToast(msg, type) {
  let container = qs('#toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = 'toast toast-' + (type || 'info');
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => toast.classList.add('show'), 10);
  setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 3500);
}

// ---------- শেয়ার্ড হেডার/ফুটার (সাইট জুড়ে একই তথ্য দেখানোর জন্য) ----------

function renderSiteHeaderFooter() {
  const adminLineEls = qsa('.js-header-admin-line');
  adminLineEls.forEach(el => el.textContent = SITE_CONFIG.headerAdminLine);

  const groupNameEls = qsa('.js-group-name');
  groupNameEls.forEach(el => el.textContent = SITE_CONFIG.groupName);

  const footerLineEls = qsa('.js-footer-line');
  footerLineEls.forEach(el => el.textContent = SITE_CONFIG.footerLine);

  const taglineEls = qsa('.js-tagline');
  taglineEls.forEach(el => el.textContent = SITE_CONFIG.tagline);
}

document.addEventListener('DOMContentLoaded', renderSiteHeaderFooter);
