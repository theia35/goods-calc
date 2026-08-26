/**
 * 韓國周邊成本計算器 — Google Apps Script 後端
 * 部署方式：請看「部署說明.md」
 */

// ====== 設定：你的試算表 ID ======
const SHEET_ID = '1o1MefZ4TsWgk55DyCbe6PtHjmEeO-8qTkmPHtXvEOFA';
const DATA_SHEET = '批次資料';   // 儲存批次的 tab 名稱

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    let result;
    switch (req.action) {
      case 'scrape':    result = scrape(req.url); break;
      case 'scrapeAll': result = scrapeAll(req.url); break;
      case 'save':      result = saveBatch(req); break;
      case 'list':      result = listBatches(); break;
      case 'get':       result = getBatch(req.id); break;
      case 'delProduct':result = delProduct(req.id, req.index); break;
      case 'delBatch':  result = delBatch(req.id); break;
      default: throw new Error('unknown action');
    }
    return json({ ok: true, ...result });
  } catch (err) {
    return json({ error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.appendRow(['id', 'name', 'rate', 'date', 'data']); // header
  }
  return sh;
}

/* ---------------- 抓取分類頁：全部商品 ---------------- */
// 支援 Cafe24 商城（如 shop-t1.gg）的分類頁，自動翻頁抓完所有商品
function scrapeAll(url) {
  const sep = url.indexOf('?') > -1 ? '&' : '?';
  const seen = {};
  const products = [];
  const MAX_PAGES = 40;

  for (let page = 1; page <= MAX_PAGES; page++) {
    let html;
    if (page === 1) {
      html = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText();
    } else {
      const res = UrlFetchApp.fetch(url + sep + 'page=' + page, { muteHttpExceptions: true });
      html = res.getContentText();
    }
    const items = parseListPage(html);
    if (!items.length) break; // 沒有更多商品 → 翻頁結束
    let fresh = 0;
    for (const it of items) {
      if (!seen[it.id]) {
        seen[it.id] = true;
        products.push(it);
        fresh++;
      }
    }
    if (fresh === 0) break; // 整頁都是重複的 → 停止
    Utilities.sleep(300);   // 禮貌性延遲
  }

  // 批次翻譯（逐筆，Google 免費端點）
  products.forEach(p => {
    p.title_ko = p.title;
    p.title_zh = translate(p.title) || p.title;
  });

  return { count: products.length, products: products };
}

// 解析 Cafe24 商品列表：li#anchorBoxId_NNN data-price=...
function parseListPage(html) {
  const items = [];
  const re = /<li id="anchorBoxId_(\d+)"[^>]*data-price="(\d+)"([\s\S]*?)<\/li>/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    const id = m[1], price = parseInt(m[2], 10), body = m[3];
    const nameM = body.match(/alt="([^"]+)"/);
    const imgM = body.match(/data-src="(\/\/[^"]+)"/) || body.match(/src="(\/\/[^"]+web\/product[^"]+)"/);
    const linkM = body.match(/href="(\/product\/[^"]+)"/);
    items.push({
      id: id,
      title: nameM ? decodeEntities(nameM[1]).replace(/\s+/g, ' ').trim() : '',
      image: imgM ? (imgM[1].indexOf('//') === 0 ? 'https:' + imgM[1] : imgM[1]) : '',
      krw: price,
      url: linkM ? linkM[1] : ''
    });
  }
  return items;
}

/* ---------------- 抓取單一商品頁 ---------------- */
function scrape(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const html = res.getContentText();

  // 標題：og:title 或 <title>
  let title = pick(html, /property=["']og:title["']\s+content=["']([^"']+)["']/i)
           || pick(html, /<title>([^<]+)<\/title>/i) || '';
  title = decodeEntities(title).trim();

  // 商品圖：og:image
  const image = pick(html, /property=["']og:image["']\s+content=["']([^"']+)["']/i) || '';

  // 價格：優先 meta itemprop/og，再找「xx,xxx원」
  let krw =
    num(pick(html, /property=["']product:price:amount["']\s+content=["']([\d,]+)["']/i)) ||
    num(pick(html, /itemprop=["']price["'][^>]*content=["']([\d,]+)["']/i)) ||
    firstKwon(html);

  return { title_ko: title, title_zh: translate(title), image: image, krw: krw };
}

function pick(s, re) { const m = s.match(re); return m ? m[1] : null; }
function num(s) { if (!s) return 0; const n = parseInt(String(s).replace(/[^\d]/g, ''), 10); return isNaN(n) ? 0 : n; }

function firstKwon(html) {
  const re = /([\d]{1,3}(?:,[\d]{3})+|\d{4,})\s*(?:원|KRW)/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    const v = num(m[1]);
    if (v >= 100) return v; // 避免抓到無關數字
  }
  return 0;
}

function decodeEntities(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&nbsp;/g, ' ');
}

/* ---------------- 免費翻譯（Google translate 端點）---------------- */
function translate(text) {
  if (!text) return '';
  try {
    const url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=ko&tl=zh-TW&dt=t&q='
      + encodeURIComponent(text);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const arr = JSON.parse(res.getContentText());
    return (arr[0] || []).map(x => x[0]).join('').trim();
  } catch (e) {
    return ''; // 失敗就回空字串，前端顯示原文
  }
}

/* ---------------- 批次 CRUD ---------------- */
function saveBatch(req) {
  const sh = sheet();
  const id = Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, req.name, req.rate, new Date(), JSON.stringify(req.products)]);
  return { id };
}

function rows() {
  const sh = sheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues();
}

function listBatches() {
  const batches = rows().map(r => ({
    id: r[0], name: r[1], rate: r[2],
    date: r[3] ? new Date(r[3]).toLocaleDateString('zh-TW') : '',
    products: safeParse(r[4])
  })).reverse();
  return { batches };
}

function getBatch(id) {
  const b = rows().find(r => r[0] === id);
  if (!b) throw new Error('找不到批次');
  return { batch: { id: b[0], name: b[1], rate: b[2], date: b[3], products: safeParse(b[4]) } };
}

function delProduct(id, index) {
  const sh = sheet();
  const all = rows();
  for (let i = 0; i < all.length; i++) {
    if (all[i][0] === id) {
      const products = safeParse(all[i][4]);
      products.splice(index, 1);
      sh.getRange(i + 2, 5).setValue(JSON.stringify(products));
      return {};
    }
  }
  throw new Error('找不到批次');
}

function delBatch(id) {
  const sh = sheet();
  const all = rows();
  for (let i = 0; i < all.length; i++) {
    if (all[i][0] === id) { sh.deleteRow(i + 2); return {}; }
  }
  throw new Error('找不到批次');
}

function safeParse(s) { try { return JSON.parse(s) || []; } catch (e) { return []; } }
