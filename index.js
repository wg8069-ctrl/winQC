// v11 - 暫停 LINE 通知，優先確保 Sheets 寫入與名稱抓取
const express = require('express');
const crypto  = require('crypto');
const axios   = require('axios');
const path    = require('path');
const fs      = require('fs');
const { google } = require('googleapis');

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
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || '1PDzqFlsjPgHJBhDsB8gwuu3_eRwJ6GS_uV2Sc6ZanXM';
const CLOUDINARY_SECRET = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';
const CLOUDINARY_CLOUD  = 'dlpxz4qlh';
const CLOUDINARY_KEY    = '953226455671951';

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
  if (!userId) return '未知用戶';
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    return r.data.displayName || 'LINE用戶';
  } catch (e) {
    console.error('抓取名稱失敗:', e.message);
    return '抓取失敗(請檢查TOKEN)';
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
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/image/upload`, formData.toString());
    return r.data.secure_url;
  } catch (e) { return null; }
}

// ── Google Sheets 寫入函式 ──
async function appendToSheet(dataMap) {
  try {
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    const sheets = google.sheets({ version: 'v4', auth });
    
    // 取得第一個工作表名稱
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetName = meta.data.sheets[0].properties.title;

    const row = [
      dataMap['異常單號'], dataMap['發生日期'], dataMap['需求回覆時間'],
      dataMap['發生單位'], dataMap['責任單位'], dataMap['系列別'],
      dataMap['單號'], dataMap['零件名稱'], dataMap['異常狀況'],
      dataMap['訂單數量'], dataMap['異常比例'], dataMap['目前處理狀態'],
      dataMap['判定'], dataMap['回報人'], '', '', '', '', '', '', 
      dataMap['異常照片'], dataMap['異常照片2']
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
    console.log('Google Sheets 寫入成功');
  } catch (e) {
    console.error('Google Sheets 寫入失敗:', e.message);
  }
}

// ── API ──
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    
    // 1. 優先處理名稱抓取
    const reporterName = await getDisplayName(d.userId);
    const wgNumber = genWGNumber();

    // 2. 上傳照片
    let photoUrl = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    let photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 3. 執行 Google Sheets 寫入
    await appendToSheet({
      '異常單號': wgNumber,
      '發生日期': d.date || new Date().toISOString().split('T')[0],
      '需求回覆時間': d.replyDate || '',
      '發生單位': d.unit || '',
      '責任單位': d.resp || '',
      '系列別': d.series || '',
      '單號': d.orderNo || '',
      '零件名稱': d.product || '',
      '異常狀況': d.anomaly || '',
      '訂單數量': d.qty || '',
      '異常比例': d.ratio || '',
      '目前處理狀態': '未開始',
      '判定': d.judge || '',
      '回報人': reporterName,
      '異常照片': photoUrl,
      '異常照片2': photoUrl2
    });

    // 4. LINE 通報目前停止 (已註解)
    /* const msg = `【測試】通報已收到，單號: ${wgNumber}`;
    for (const uid of (process.env.NOTIFY_USERS || '').split(',')) {
       // pushText(uid, msg); 
    }
    */

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    console.error('API Error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/', (req, res) => res.send('System running. LINE service paused. Sheets priority mode.'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
