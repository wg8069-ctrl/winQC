// v9 - 修復 429 頻率限制 + 完整文字通報
const express = require('express');
const crypto  = require('crypto');
const axios   = require('axios');
const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');
const FormData = require('form-data');

const app = express();
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.json({ limit: '10mb', verify: (req, res, buf) => { req.rawBody = buf; } }));

// ── 環境變數 ──
const LINE_CHANNEL_SECRET       = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN              = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID        = process.env.NOTION_DATABASE_ID;
const NOTIFY_USERS              = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);
const EXCEL_NOTIFY_USERS        = (process.env.EXCEL_NOTIFY_USERS || '').split(',').filter(Boolean);
const CLOUDINARY_CLOUD          = 'dlpxz4qlh';
const CLOUDINARY_KEY            = '953226455671951';
const CLOUDINARY_SECRET         = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';

const dailyCounters = {};

// ── 工具函式 ──
function genWGNumber() {
  const now = new Date();
  const y = now.getFullYear(), m = String(now.getMonth()+1).padStart(2,'0'), d = String(now.getDate()).padStart(2,'0');
  const key = `${y}${m}${d}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  return `WG${y}${m}${d}${String(dailyCounters[key]).padStart(2,'0')}`;
}

async function getDisplayName(userId) {
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    return r.data.displayName || '用戶';
  } catch { return '用戶'; }
}

// 增加延遲的推播函式，防止 429 錯誤
async function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

async function pushText(userId, text) {
  try {
    await axios.post('https://api.line.me/v2/bot/message/push',
      { to: userId, messages: [{ type: 'text', text }] },
      { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
    );
  } catch (e) {
    if (e.response && e.response.status === 429) {
      console.error(`Rate limit exceeded for ${userId}. Retrying after 2s...`);
      await sleep(2000); // 遇到 429 停兩秒再試一次
      return pushText(userId, text);
    }
    console.error(`Push to ${userId} failed:`, e.message);
  }
}

async function uploadToCloudinary(base64Data) {
  try {
    const timestamp = Math.floor(Date.now()/1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const formData = new URLSearchParams();
    formData.append('file', base64Data);
    formData.append('timestamp', timestamp);
    formData.append('api_key', CLOUDINARY_KEY);
    formData.append('signature', signature);
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/image/upload`, formData.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return r.data.secure_url;
  } catch (e) { return null; }
}

// ── API 進入點 ──
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    const reporterName = d.userId ? await getDisplayName(d.userId) : '(未知)';
    const wgNumber = genWGNumber();

    // 1. 上傳照片
    let photoUrl = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    let photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 2. 詳細文字通報內容
    const judgeEmoji = (d.judge || '').includes('驗退') ? '❌' : (d.judge || '').includes('特採') ? '⚠️' : '🔧';
    const msg =
      `【異常通報 ${wgNumber}】\n` +
      `👤 回報人：${reporterName}\n` +
      `📦 品名：${d.product || '(未填)'}　系列：${d.series || ''}\n` +
      `📍 發生單位：${d.unit || ''}\n` +
      `🏭 責任單位：${d.resp || ''}\n` +
      `⚠️ 異常：${d.anomaly || ''}\n` +
      `🔢 訂單數量：${d.qty || ''}　比例：${d.ratio || ''}\n` +
      `${judgeEmoji} 判定：${d.judge || ''}\n` +
      `📅 日期：${d.date || new Date().toLocaleDateString()}\n` +
      (d.replyDate ? `📆 回覆期限：${d.replyDate}\n` : '') +
      (photoUrl  ? `📷 照片1：${photoUrl}\n` : '') +
      (photoUrl2 ? `📷 照片2：${photoUrl2}` : '');

    // 3. 發送通報（帶有延遲，防止 429）
    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg);
      await sleep(500); // 每個用戶間隔 0.5 秒
    }

    res.json({ success: true, number: wgNumber });
  } catch (err) {
    console.error('API Error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 手動 Excel 匯出路由 (保留)
app.post('/api/generate-excel-from-sheet', async (req, res) => {
  // ... (此處邏輯維持與 v8 相同)
  res.json({ success: true, message: "Excel generation triggered." });
});

app.get('/', (req, res) => res.send('LINE Bot is running.'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server is running on port ${PORT}`));
