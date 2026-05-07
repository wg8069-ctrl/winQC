// v10 - 強化型發送邏輯，解決 429 持續報錯問題
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

const LINE_CHANNEL_SECRET       = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN              = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID        = process.env.NOTION_DATABASE_ID;
const NOTIFY_USERS              = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);
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

// 延遲函式
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// 強化型推播：遇到 429 會等待更久，並有最大重試次數
async function pushTextSafe(userId, text, retryCount = 0) {
  try {
    await axios.post('https://api.line.me/v2/bot/message/push',
      { to: userId, messages: [{ type: 'text', text }] },
      { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
    );
    console.log(`Successfully notified ${userId}`);
  } catch (e) {
    if (e.response && e.response.status === 429 && retryCount < 3) {
      const waitTime = (retryCount + 1) * 5000; // 第一次等5秒，第二次10秒...
      console.warn(`[429] Rate limit for ${userId}. Retrying in ${waitTime/1000}s...`);
      await sleep(waitTime);
      return pushTextSafe(userId, text, retryCount + 1);
    }
    console.error(`Final push failure for ${userId}:`, e.message);
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

// ── API ──
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    const reporterName = d.userId ? await getDisplayName(d.userId) : '(未知)';
    const wgNumber = genWGNumber();

    // 先回應前端，避免 LIFF 超時
    res.json({ success: true, number: wgNumber });

    // 背景處理後續動作
    let photoUrl = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    let photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 構建通報訊息
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

    // 依序發送通報，每個人之間強制間隔 1 秒
    for (const uid of NOTIFY_USERS) {
      await pushTextSafe(uid, msg);
      await sleep(1000); 
    }

  } catch (err) {
    console.error('Outer Error:', err.message);
  }
});

app.get('/', (req, res) => res.send('System Active. Messaging service reinforced.'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
