// v5.2 - 優化 LINE 通知格式，移除自動 Excel 下載
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
const CLOUDINARY_CLOUD          = 'dlpxz4qlh';
const CLOUDINARY_KEY            = '953226455671951';
const CLOUDINARY_SECRET         = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';

const sessions = {};

// ── Google Sheets ──
const { google } = require('googleapis');
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || '1PDzqFlsjPgHJBhDsB8gwuu3_eRwJ6GS_uV2Sc6ZanXM';
const SHEET_HEADERS = ['異常單號','發生日期','狀態','需求回覆時間','發生單位','責任單位','槍型號','單號','零件名稱','異常狀況','訂單數量','異常比例','目前處理狀態','判定','回報人','人工成本(人)','人工成本(時)','行政成本(人)','行政成本(時)','所耗人力成本','異常標註內容','異常照片','異常照片2'];

async function getSheets(){
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
  });
  return google.sheets({ version: 'v4', auth });
}

let SHEET_NAME = null;
async function getSheetName(){
  if(SHEET_NAME) return SHEET_NAME;
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  SHEET_NAME = meta.data.sheets[0].properties.title;
  return SHEET_NAME;
}

async function ensureSheetHeader(){
  try {
    const sheets = await getSheets();
    const name = await getSheetName();
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${name}!A1:Z1` });
    if(!res.data.values || !res.data.values[0] || res.data.values[0].length===0){
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [SHEET_HEADERS] }
      });
    }
  } catch(e){ console.error('ensureSheetHeader failed:', e.message); }
}

async function appendToSheet(dataMap){
  try {
    const sheets = await getSheets();
    const name = await getSheetName();
    const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${name}!1:1` });
    const headers = (headerRes.data.values && headerRes.data.values[0]) ? headerRes.data.values[0] : SHEET_HEADERS;
    const row = headers.map(h => dataMap[h] !== undefined ? dataMap[h] : '');
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
  } catch(e){ console.error('appendToSheet failed:', e.message); }
}

ensureSheetHeader();

// ── Helper 函式 ──
function verifySignature(req) {
  const sig  = req.headers['x-line-signature'];
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

async function replyFlex(replyToken) {
  await axios.post('https://api.line.me/v2/bot/message/reply',
    {
      replyToken,
      messages: [{
        type: 'text',
        text: 'WinGun 異常通報 👇',
        quickReply: {
          items: [
            { type: 'action', action: { type: 'uri', label: '📋 建立異常單', uri: 'https://liff.line.me/2009600334-UpN6esDu' } },
            { type: 'action', action: { type: 'message', label: '🔍 查詢紀錄', text: '查詢' } }
          ]
        }
      }]
    },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` } }
  );
}

// ── 流水號與上傳 ──
const dailyCounters = {};
function genWGNumber() {
  const now = new Date();
  const key = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  return `WG${key}${String(dailyCounters[key]).padStart(2,'0')}`;
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

// ── API Routes ──

// LIFF 表單提交
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    let reporterName = d.userId ? await getDisplayName(d.userId) : '(未知)';
    const wgNumber = genWGNumber();
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '/');

    const photoUrl  = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    const photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 1. 寫入 Notion
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    await axios.post(`https://api.notion.com/v1/pages`, {
      parent: { database_id: NOTION_DATABASE_ID },
      properties: {
        '異常單號': { title: [{ text: { content: wgNumber } }] },
        '發生日期': { date: { start: new Date().toISOString().split('T')[0] } },
        '零件名稱': { rich_text: toText(d.product) },
        '系列別':   { rich_text: toText(d.series) },
        '發生單位': { rich_text: toText(d.unit) },
        '責任單位': { rich_text: toText(d.resp) },
        '異常狀況': { rich_text: toText(d.anomaly) },
        '判定':     { rich_text: toText(d.judge) },
        '回報人':   { rich_text: toText(reporterName) }
      }
    }, { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28' } });

    // 2. 寫入 Google Sheets
    await appendToSheet({
      '異常單號': wgNumber,
      '發生日期': today,
      '狀態': '未處理',
      '發生單位': d.unit || '',
      '責任單位': d.resp || '',
      '零件名稱': d.product || '',
      '槍型號': d.series || '',
      '異常狀況': d.anomaly || '',
      '訂單數量': d.qty || '',
      '異常比例': d.ratio || '',
      '回報人': reporterName,
      '目前處理狀態': d.status || '未處理',
      '異常照片': photoUrl || '',
      '異常照片2': photoUrl2 || ''
    });

    // 3. 推播通知 (格式更新)
    const msg = 
      `【異常通報 ${wgNumber}】\n` +
      `回報人：${reporterName}\n` +
      `📌品名：${d.product || ''}　系列：${d.series || ''}\n` +
      `📍 發生單位：${d.unit || ''}\n` +
      `🏭 責任單位：${d.resp || ''}\n` +
      `⚠️ 異常：${d.anomaly || ''}\n` +
      `📃 訂單數量：${d.qty || ''}　比例：${d.ratio || ''}\n` +
      `🔥 目前處理狀態：${d.status || '目前處理狀態'}\n` +
      `📅 日期：${today}`;

    for (const uid of NOTIFY_USERS) {
      await axios.post('https://api.line.me/v2/bot/message/push',
        { to: uid, messages: [{ type: 'text', text: msg }] },
        { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` } }
      ).catch(e => console.error('Push failed:', e.message));
    }

    res.json({ success: true, number: wgNumber });
  } catch (err) {
    console.error('API Error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message' || event.type === 'follow') await replyFlex(event.replyToken);
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
