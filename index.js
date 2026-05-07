// v12 - 徹底修復 Google Sheets 欄位錯位問題
const express = require('express');
const crypto  = require('crypto');
const axios   = require('axios');
const path    = require('path');
const { google } = require('googleapis');

const app = express();
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.json({ limit: '10mb' }));

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
  if (!userId) return '未知(無ID)';
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    return r.data.displayName || 'LINE用戶';
  } catch (e) {
    return '未知(抓取失敗)';
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

// ── 核心修復：自動對應標題的寫入函式 ──
async function appendToSheet(dataMap) {
  try {
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    const sheets = google.sheets({ version: 'v4', auth });
    
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheetName = meta.data.sheets[0].properties.title;

    // 1. 先抓取第一列的標題 (Headers)
    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!1:1`
    });
    const headers = headerRes.data.values ? headerRes.data.values[0] : [];

    // 2. 根據標題順序組成資料列
    // 如果標題是「零件名稱」，程式就會去找 dataMap['零件名稱']
    const row = headers.map(header => {
      const val = dataMap[header.trim()];
      return val !== undefined ? val : '';
    });

    // 3. 寫入資料
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
    console.log('Google Sheets 對應寫入成功');
  } catch (e) {
    console.error('Sheets 寫入失敗:', e.message);
  }
}

// ── API ──
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    
    // 取得名稱與編號
    const reporterName = await getDisplayName(d.userId);
    const wgNumber = genWGNumber();

    // 上傳照片
    let photoUrl = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    let photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 根據你的 Excel 標題準備資料
    // 請確保這裡的 Key 跟 Excel 第一列的文字完全一樣
    const dataToSheet = {
      '異常單號': wgNumber,
      '發生日期': d.date || new Date().toISOString().split('T')[0],
      '需求回覆時間': d.replyDate || '',
      '發生單位': d.unit || '',
      '責任單位': d.resp || '',
      '槍型號': d.series || '', // 對應你圖中的 E 欄
      '單號': d.orderNo || '',   // 對應你圖中的 F 欄
      '訂單數量': d.qty || '',
      '異常比例': d.ratio || '',
      '零件名稱': d.product || '',
      '異常狀況': d.anomaly || '',
      '判定': d.judge || '',
      '目前處理狀態': '未開始',
      '回報人': reporterName,
      '異常照片': photoUrl || '',
      '異常照片2': photoUrl2 || ''
    };

    // 執行寫入
    await appendToSheet(dataToSheet);

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/', (req, res) => res.send('Corrected Mapping Engine Running.'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
