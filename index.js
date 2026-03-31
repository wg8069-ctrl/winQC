const express = require('express');
const crypto = require('crypto');
const axios = require('axios');

const app = express();
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.json({ limit: '10mb', verify: (req, res, buf) => { req.rawBody = buf; } }));

// ── 環境變數（原有的，不用動）──
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID = process.env.NOTION_DATABASE_ID;
// ── 新增：LIFF 通知對象（Render 環境變數加這個）──
const NOTIFY_USERS = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);

const sessions = {};

// ════════════════════════════════════════
//  原有的 helper 函式（完全不動）
// ════════════════════════════════════════

function verifySignature(req) {
  const sig = req.headers['x-line-signature'];
  const hash = crypto.createHmac('sha256', LINE_CHANNEL_SECRET).update(req.rawBody).digest('base64');
  return hash === sig;
}

async function getDisplayName(userId) {
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    return r.data.displayName || '用戶';
  } catch { return '用戶'; }
}

async function getImageUrl(messageId) {
  try {
    const r = await axios.get(
      `https://api-data.line.me/v2/bot/message/${messageId}/content`,
      { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }, responseType: 'arraybuffer' }
    );
    return `data:image/jpeg;base64,${Buffer.from(r.data).toString('base64')}`;
  } catch { return null; }
}

async function replyText(replyToken, text) {
  await axios.post('https://api.line.me/v2/bot/message/reply',
    { replyToken, messages: [{ type: 'text', text }] },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

async function pushText(userId, text) {
  await axios.post('https://api.line.me/v2/bot/message/push',
    { to: userId, messages: [{ type: 'text', text }] },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

async function createNotionPage(data, senderName) {
  const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
  const statusName = data.caseNumber ? '處理中' : '未開始';
  const caseNum = data.caseNumber ? parseInt(data.caseNumber.replace(/[^0-9]/g, '')) || null : null;

  const properties = {
    '發生地':                    { title: [{ text: { content: data.location || '(未填)' } }] },
    '產品編號':                  { rich_text: toText(data.productId) },
    '品名':                      { rich_text: toText(data.itemName) },
    '異常狀況':                  { rich_text: toText(data.issue) },
    '異常廠商':                  { rich_text: toText(data.vendor) },
    '客戶':                      { rich_text: toText(data.customer) },
    '處理方式':                  { rich_text: toText(data.solution) },
    '數量':                      data.quantity ? { number: parseInt(data.quantity) } : { number: null },
    '發生日期':                  { date: { start: new Date().toISOString().split('T')[0] } },
    '已開立異常單(請輸入單號)':   caseNum ? { number: caseNum } : { number: null },
    '免開異常(請輸入原因)':       { rich_text: toText(data.skipReason) },
    '目前處理狀態':              { rich_text: toText(statusName) },
    '回報人':                    { rich_text: toText(senderName) },
  };

  const pageBody = { parent: { database_id: NOTION_DATABASE_ID }, properties };
  if (data.photoUrl) {
    pageBody.children = [{
      object: 'block', type: 'image',
      image: { type: 'external', external: { url: data.photoUrl } }
    }];
  }
  const res = await axios.post('https://api.notion.com/v1/pages',
    pageBody,
    { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
  );
  return res.data;
}

async function searchNotion(keyword) {
  const filters = ['產品編號','品名','異常狀況','異常廠商','免開異常(請輸入原因)'].map(field => ({
    property: field, rich_text: { contains: keyword }
  }));
  if (!isNaN(keyword)) {
    filters.push({ property: '已開立異常單(請輸入單號)', number: { equals: parseInt(keyword) } });
  }
  const res = await axios.post(
    `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`,
    { filter: { or: filters }, sorts: [{ property: '發生日期', direction: 'descending' }] },
    { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
  );
  return res.data.results.map(p => {
    const props = p.properties;
    const getText = (k) => props[k]?.rich_text?.[0]?.text?.content || '';
    const getNum  = (k) => props[k]?.number ?? '';
    const getDate = (k) => props[k]?.date?.start?.slice(0,10) || '';
    return {
      date:       getDate('發生日期'),
      productId:  getText('產品編號'),
      issue:      getText('異常狀況'),
      quantity:   getNum('數量'),
      status:     getText('目前處理狀態'),
      caseNumber: getNum('已開立異常單(請輸入單號)'),
    };
  });
}

// ════════════════════════════════════════
//  原有的 Bot 對話邏輯（完全不動）
// ════════════════════════════════════════

const MAIN_MENU =
  '📋 WinGun 異常回報系統\n\n' +
  '請選擇功能：\n\n' +
  '1️⃣  回報異常\n' +
  '2️⃣  查詢紀錄\n' +
  '0️⃣  顯示此選單\n\n' +
  '（直接輸入數字選擇）';

const BASE_STEPS = [
  { key: 'location',   required: true,  ask: '📍 請輸入發生地點\n（例如：本廠／二廠／廠商地）' },
  { key: 'productId',  required: true,  ask: '📦 請輸入產品編號\n（例如：WCB4-215B-CR）\n隨時輸入「0」回主選單' },
  { key: 'itemName',   required: true,  ask: '🏷️ 請輸入品名（物料名稱）\n（例如：O環／滑套）' },
  { key: 'issue',      required: true,  ask: '⚠️ 請描述異常狀況' },
  { key: 'photo',      required: false, ask: '📸 請上傳異常照片\n（可直接拍照傳送，或輸入「無」跳過）', isPhoto: true },
  { key: 'quantity',   required: true,  ask: '🔢 請輸入異常數量\n（只要輸入數字不用單位，例如：150）',
    validate: (v) => isNaN(v) ? '請輸入純數字！' : null },
  { key: 'solution',   required: true,  ask: '🔧 請輸入目前處理方式' },
  { key: 'vendor',     required: true,  ask: '🏭 請輸入異常廠商名稱' },
  { key: 'customer',   required: false, ask: '👥 請輸入客戶名稱\n（不知道可輸入「無」跳過）' },
  { key: 'caseNumber', required: false, ask: '📝 已開立異常單號？\n（有開立請輸入單號數字）\n（沒有請輸入「無」，之後需填免開原因）\n⚠️ 有填單號→狀態自動設為「處理中」' },
];

const SKIP_REASON_STEP = { key: 'skipReason', required: true, ask: '🚫 請輸入為何免開異常原因' };

function buildSummary(d) {
  const status = d.caseNumber ? '處理中' : '未開始';
  return (
    `📋 請確認以下資料：\n\n` +
    `📍 發生地：${d.location}\n` +
    `📦 產品編號：${d.productId}\n` +
    `🏷️ 品名：${d.itemName}\n` +
    `⚠️ 異常狀況：${d.issue}\n` +
    `🔢 數量：${d.quantity}\n` +
    `🔧 處理方式：${d.solution}\n` +
    `🏭 異常廠商：${d.vendor}\n` +
    (d.customer   ? `👥 客戶：${d.customer}\n`       : '') +
    (d.caseNumber ? `📝 異常單號：${d.caseNumber}\n` : '') +
    (d.skipReason ? `🚫 免開原因：${d.skipReason}\n` : '') +
    `\n🔘 處理狀態：${status}\n\n` +
    `輸入「確認」送出\n輸入「重填」重新開始\n輸入「0」回主選單`
  );
}

async function handleMessage(event) {
  const userId = event.source?.userId;
  const replyToken = event.replyToken;
  const text = event.message?.type === 'text' ? event.message.text.trim() : null;
  const imageId = event.message?.type === 'image' ? event.message.id : null;
  if (!userId) return;

  let session = sessions[userId] || { step: 'idle', data: {} };

  if (text === '0' || text === '選單' || text === 'menu') {
    delete sessions[userId];
    await replyText(replyToken, MAIN_MENU);
    return;
  }

  if (text === '重填' || text === '取消' || text === '2') {
    delete sessions[userId];
    await replyText(replyToken, '已取消，重新開始。\n\n' + MAIN_MENU);
    return;
  }

  if (session.step === 'idle') {
    if (text === '1' || text === '回報異常') {
      sessions[userId] = { step: 0, data: {} };
      await replyText(replyToken, '📋 開始填寫異常回報！\n隨時輸入「0」回主選單\n\n' + BASE_STEPS[0].ask);
    } else if (text === '2' || text === '查詢紀錄' || text === '查詢') {
      sessions[userId] = { step: 'searching', data: {} };
      await replyText(replyToken, '🔍 請輸入查詢關鍵字\n\n可查詢：產品編號、品名、異常狀況、異常廠商、異常單號\n\n輸入「0」回主選單');
    } else {
      await replyText(replyToken, MAIN_MENU);
    }
    return;
  }

  if (session.step === 'searching') {
    if (!text) { await replyText(replyToken, '請輸入查詢關鍵字'); return; }
    try {
      const results = await searchNotion(text);
      if (results.length === 0) {
        await replyText(replyToken, `🔍 查無「${text}」相關紀錄\n\n輸入「0」回主選單`);
      } else {
        const lines = [`🔍 找到 ${results.length} 筆「${text}」相關紀錄：\n`];
        results.slice(0, 5).forEach((r, i) => {
          lines.push(
            `${i + 1}. ${r.date} ${r.productId}\n` +
            `   📋 ${r.issue}\n` +
            `   🔢 ${r.quantity} pcs｜🔘 ${r.status}` +
            (r.caseNumber ? `\n   📝 單號：${r.caseNumber}` : '')
          );
        });
        if (results.length > 5) lines.push(`\n...共 ${results.length} 筆，僅顯示前 5 筆`);
        lines.push('\n輸入「0」回主選單');
        await replyText(replyToken, lines.join('\n'));
      }
    } catch (err) {
      console.error(err.response?.data || err.message);
      await replyText(replyToken, '❌ 查詢失敗，請通知管理員');
    }
    delete sessions[userId];
    return;
  }

  if (typeof session.step === 'number') {
    const cur = BASE_STEPS[session.step];
    if (cur.isPhoto) {
      if (imageId) {
        const url = await getImageUrl(imageId);
        session.data.photoUrl = url;
      } else if (text === '無' || text === '略過') {
        session.data.photoUrl = null;
      } else {
        await replyText(replyToken, '請上傳照片，或輸入「無」跳過\n\n' + cur.ask);
        return;
      }
      const next = session.step + 1;
      session.step = next;
      sessions[userId] = session;
      await replyText(replyToken, '✅ 已記錄！\n\n' + BASE_STEPS[next].ask);
      return;
    }

    if (!text) { await replyText(replyToken, '請輸入文字\n\n' + cur.ask); return; }
    if (cur.validate) {
      const err = cur.validate(text);
      if (err) { await replyText(replyToken, err + '\n\n' + cur.ask); return; }
    }
    session.data[cur.key] = (!cur.required && (text === '無' || text === '略過')) ? '' : text;
    const next = session.step + 1;

    if (next >= BASE_STEPS.length) {
      if (!session.data.caseNumber) {
        session.step = 'ask_skip_reason';
        sessions[userId] = session;
        await replyText(replyToken, '✅ 已記錄！\n\n' + SKIP_REASON_STEP.ask);
      } else {
        session.step = 'confirm';
        sessions[userId] = session;
        await replyText(replyToken, buildSummary(session.data));
      }
    } else {
      session.step = next;
      sessions[userId] = session;
      await replyText(replyToken, '✅ 已記錄！\n\n' + BASE_STEPS[next].ask);
    }
    return;
  }

  if (session.step === 'ask_skip_reason') {
    if (!text) { await replyText(replyToken, '請輸入免開異常原因'); return; }
    session.data.skipReason = text;
    session.step = 'confirm';
    sessions[userId] = session;
    await replyText(replyToken, buildSummary(session.data));
    return;
  }

  if (session.step === 'confirm') {
    if (text !== '1') {
      await replyText(replyToken, '請輸入「1」送出\n或輸入「2」重新開始\n或輸入「0」回主選單');
      return;
    }
    try {
      const name = await getDisplayName(userId);
      await createNotionPage(session.data, name);
      const status = session.data.caseNumber ? '處理中' : '未開始';
      delete sessions[userId];
      await replyText(replyToken,
        `✅ 已成功寫入 Notion！\n\n` +
        `📦 ${session.data.productId}\n` +
        `⚠️ ${session.data.issue}\n` +
        `🔢 ${session.data.quantity} pcs\n` +
        `🔘 狀態：${status}\n` +
        `👤 回報人：${name}\n\n感謝回報！\n\n` +
        `輸入「0」回主選單`
      );
    } catch (err) {
      console.error(err.response?.data || err.message);
      await replyText(replyToken, '❌ 寫入失敗，請通知管理員\n' + (err.response?.data?.message || err.message));
    }
    return;
  }

  await replyText(replyToken, MAIN_MENU);
}

// ════════════════════════════════════════
//  流水號產生器（WG+年+月+日+兩位流水）
// ════════════════════════════════════════
const dailyCounters = {};

function genWGNumber() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const key = `${y}${m}${d}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  const seq = String(dailyCounters[key]).padStart(2, '0');
  return `WG${y}${m}${d}${seq}`;
}
// ════════════════════════════════════════
//  Cloudinary 照片上傳
// ════════════════════════════════════════
const CLOUDINARY_CLOUD = 'dlpxz4qlh';
const CLOUDINARY_KEY   = '953226455671951';
const CLOUDINARY_SECRET = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';

async function uploadToCloudinary(base64Data) {
  try {
    const timestamp = Math.floor(Date.now() / 1000);
    const sigStr = `timestamp=${timestamp}${CLOUDINARY_SECRET}`;
    const signature = crypto.createHash('sha1').update(sigStr).digest('hex');
    const formData = new URLSearchParams();
    formData.append('file', base64Data);
    formData.append('timestamp', timestamp);
    formData.append('api_key', CLOUDINARY_KEY);
    formData.append('signature', signature);
    const r = await axios.post(
      `https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/image/upload`,
      formData.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, maxContentLength: Infinity, maxBodyLength: Infinity }
    );
    return r.data.secure_url;
  } catch (e) {
    console.error('Cloudinary upload failed:', e.response?.data || e.message);
    return null;
  }
}

// ════════════════════════════════════════
//  新增：LIFF 表單 API
// ════════════════════════════════════════

app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;

    // 1. 從 LINE 取得回報人名稱
    let reporterName = '(未知)';
    if (d.userId) {
      reporterName = await getDisplayName(d.userId);
    }

    // 2. 產生流水單號
    const wgNumber = genWGNumber();

    // 3. 上傳照片到 Cloudinary
    let photoUrl = null;
    if (d.photoData) {
      photoUrl = await uploadToCloudinary(d.photoData);
    }

    // 4. 組 Notion properties（對應你的實際欄位）
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    const properties = {
      '異常單號':     { title: [{ text: { content: wgNumber } }] },
      '發生日期':     { date: { start: new Date().toISOString().split('T')[0] } },
      '發生地':       { rich_text: toText(d.unit || '') },
      '發生單位':     { rich_text: toText(d.unit || '') },
      '責任單位':     { rich_text: toText(d.resp || '') },
      '客戶':         { rich_text: toText('') },
      '零件編號':     { rich_text: toText('') },
      '零件名稱':     { rich_text: toText(d.product || '') },
      '異常狀況':     { rich_text: toText(d.anomaly || '') },
      '處理方式':     { rich_text: toText(d.judge || '') },
      '判定':         { rich_text: toText(d.judge || '') },
      '訂單數量':     { number: parseInt(d.qty) || null },
      '異常比例':     { rich_text: toText(d.ratio || '') },
      '目前處理狀態': { rich_text: toText('未開始') },
      '回報人':       { rich_text: toText(reporterName) },
      '異常照片':     photoUrl ? { url: photoUrl } : { url: null },
    };

    const pageBody = { parent: { database_id: NOTION_DATABASE_ID }, properties };

    // 5. 如果有照片也加到頁面內容
    if (photoUrl) {
      pageBody.children = [{
        object: 'block', type: 'image',
        image: { type: 'external', external: { url: photoUrl } }
      }];
    }

    await axios.post('https://api.notion.com/v1/pages', pageBody, {
      headers: {
        Authorization: `Bearer ${NOTION_TOKEN}`,
        'Notion-Version': '2022-06-28',
        'Content-Type': 'application/json'
      }
    });

    // 6. 推播通知給主管
    const judgeEmoji = d.judge === '驗退X' ? '❌' : d.judge === '特採△' ? '⚠️' : '🔧';
    const msg =
      `【異常通報 ${wgNumber}】\n` +
      `👤 回報人：${reporterName}\n` +
      `📦 品名：${d.product || '(未填)'}\n` +
      `📍 發生單位：${d.unit}\n` +
      `🏭 責任單位：${d.resp}\n` +
      `⚠️ 異常：${d.anomaly}\n` +
      `🔢 訂單數量：${d.qty}　比例：${d.ratio}\n` +
      `${judgeEmoji} 判定：${d.judge}\n` +
      `📅 日期：${d.date}` +
      (photoUrl ? `\n📷 照片：${photoUrl}` : '');

    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg).catch(e => console.error('push failed:', e.message));
    }

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ════════════════════════════════════════
//  原有的 webhook + 健康檢查（不動）
// ════════════════════════════════════════

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message') await handleMessage(event).catch(console.error);
    if (event.type === 'follow') {
      await replyText(event.replyToken,
        '👋 歡迎使用 WinGun 異常回報系統！\n\n' + MAIN_MENU
      ).catch(console.error);
    }
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅' }));


// ════════════════════════════════════════
//  新增：Make.com 呼叫 → 產生 Excel → LINE 傳送
// ════════════════════════════════════════

const ExcelJS = require('exceljs');
const https   = require('https');
const http    = require('http');

// 下載圖片回傳 Buffer
function fetchImageBuffer(url) {
  return new Promise((resolve) => {
    const mod = url.startsWith('https') ? https : http;
    mod.get(url, (res) => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', () => resolve(null));
    }).on('error', () => resolve(null));
  });
}

// 從 Google Drive 下載範本（用公開連結或服務帳戶）
async function fetchTemplateBuffer() {
  const fileId = process.env.GDRIVE_TEMPLATE_ID;
  // 使用 Google Drive 直接下載連結（範本需設為「知道連結的人可以查看」）
  const url = `https://drive.google.com/uc?export=download&id=${fileId}`;
  return fetchImageBuffer(url);
}

// 用 LINE Bot 傳送檔案（使用 multipart upload）
async function sendExcelViaLine(buffer, filename, toUserId) {
  try {
    // LINE 不直接支援傳 xlsx，改傳為訊息說明 + 圖片預覽
    // 這裡改用：上傳到 Cloudinary 取得下載連結，再傳給 LINE
    const timestamp = Math.floor(Date.now() / 1000);
    const sigStr = `timestamp=${timestamp}${CLOUDINARY_SECRET}`;
    const signature = crypto.createHash('sha1').update(sigStr).digest('hex');

    const FormData = require('form-data');
    const form = new FormData();
    form.append('file', buffer, { filename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    form.append('timestamp', timestamp);
    form.append('api_key', CLOUDINARY_KEY);
    form.append('signature', signature);
    form.append('resource_type', 'raw');

    const uploadRes = await axios.post(
      `https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/raw/upload`,
      form,
      { headers: form.getHeaders(), maxContentLength: Infinity, maxBodyLength: Infinity }
    );

    const downloadUrl = uploadRes.data.secure_url;

    // 傳送 LINE 訊息（含下載連結）
    await axios.post('https://api.line.me/v2/bot/message/push',
      {
        to: toUserId,
        messages: [{
          type: 'text',
          text: `📋 品質異常通知單已產生！\n\n點擊下載 Excel：\n${downloadUrl}`
        }]
      },
      { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
    );

    return { success: true, url: downloadUrl };
  } catch (e) {
    console.error('sendExcelViaLine error:', e.response?.data || e.message);
    throw e;
  }
}

app.post('/api/generate-excel', async (req, res) => {
  try {
    const d = req.body;
    // d 裡面包含 Make.com 從 Notion 傳來的欄位
    // 必填：notionPageId 或直接傳欄位資料

    // ── 欄位對應 ──
    const data = {
      異常單號:  d.anomalyNo   || d['異常單號']  || '',
      發生日期:  d.date        || d['發生日期']  || '',
      發生單位:  d.unit        || d['發生單位']  || '',
      責任單位:  d.resp        || d['責任單位']  || '',
      客戶:      d.customer    || d['客戶']      || '',
      零件名稱:  d.product     || d['零件名稱']  || '',
      系列別:    d.series      || d['系列別']    || '',
      異常狀況:  d.anomaly     || d['異常狀況']  || '',
      訂單數量:  d.qty         || d['訂單數量']  || '',
      異常比例:  d.ratio       || d['異常比例']  || '',
      判定:      d.judge       || d['判定']      || '',
      回報人:    d.reporter    || d['回報人']    || '',
      photo1Url: d.photo1Url   || d['異常照片']  || '',
      photo2Url: d.photo2Url   || d['異常照片2'] || '',
    };

    const toUserId = d.toUserId || process.env.NOTIFY_USERS?.split(',')[0] || '';

    // ── 載入範本 ──
    const tmplBuffer = await fetchTemplateBuffer();
    const workbook   = new ExcelJS.Workbook();
    await workbook.xlsx.load(tmplBuffer);
    const ws = workbook.worksheets[0];

    // ── 填入儲存格 ──
    const CELL_MAP = {
      '異常單號': 'C2',
      '發生日期': 'A4',
      '發生單位': 'B4',
      '責任單位': 'C4',
      '客戶':     'D4',
      '零件名稱': 'E4',
      '系列別':   'F4',
      '異常狀況': 'G4',
      '訂單數量': 'H4',
      '異常比例': 'J4',
      '判定':     'K4',
      '回報人':   'L4',
    };
    for (const [key, coord] of Object.entries(CELL_MAP)) {
      try { ws.getCell(coord).value = data[key]; } catch(e) {}
    }

    // ── 嵌入照片 ──
    for (const [urlKey, cellCoord] of [['photo1Url','A5'],['photo2Url','G5']]) {
      if (data[urlKey]) {
        const imgBuf = await fetchImageBuffer(data[urlKey]);
        if (imgBuf) {
          const imgId = workbook.addImage({ buffer: imgBuf, extension: 'jpeg' });
          ws.addImage(imgId, {
            tl: { col: cellCoord === 'A5' ? 0 : 6, row: 4 },
            ext: { width: 360, height: 360 }
          });
        }
      }
    }

    // ── 輸出為 Buffer ──
    const outBuffer = await workbook.xlsx.writeBuffer();
    const filename  = `${data['異常單號'] || 'anomaly'}.xlsx`;

    // ── 上傳 Cloudinary 並傳 LINE ──
    const result = await sendExcelViaLine(outBuffer, filename, toUserId);

    res.json({ success: true, url: result.url, filename });
  } catch (err) {
    console.error('/api/generate-excel error:', err.response?.data || err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
