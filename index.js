// v7 - 恢復詳細文字通報 + 手動觸發 Excel 模式
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

const sessions = {};

// ── Google Sheets ──
const { google } = require('googleapis');
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || '1PDzqFlsjPgHJBhDsB8gwuu3_eRwJ6GS_uV2Sc6ZanXM';
const SHEET_HEADERS = ['異常單號','發生日期','需求回覆時間','發生單位','責任單位','系列別','單號','零件名稱','異常狀況','訂單數量','異常比例','目前處理狀態','判定','回報人','人工成本(人)','人工成本(時)','行政成本(人)','行政成本(時)','所耗人力成本','異常標註內容','異常照片','異常照片2'];

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

// ── 流水號 ──
const dailyCounters = {};
function genWGNumber() {
  const now = new Date();
  const y = now.getFullYear(), m = String(now.getMonth()+1).padStart(2,'0'), d = String(now.getDate()).padStart(2,'0');
  const key = `${y}${m}${d}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  return `WG${y}${m}${d}${String(dailyCounters[key]).padStart(2,'0')}`;
}

// ── 工具函式 ──
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

async function pushText(userId, text) {
  await axios.post('https://api.line.me/v2/bot/message/push',
    { to: userId, messages: [{ type: 'text', text }] },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

// ── Cloudinary ──
async function uploadToCloudinary(base64Data) {
  try {
    const timestamp = Math.floor(Date.now()/1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const formData  = new URLSearchParams();
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

async function uploadExcelToCloudinary(buffer, filename) {
  try {
    const timestamp = Math.floor(Date.now()/1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const form = new FormData();
    form.append('file', buffer, { filename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    form.append('timestamp', String(timestamp));
    form.append('api_key', CLOUDINARY_KEY);
    form.append('signature', signature);
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/raw/upload`, form, { headers: form.getHeaders() });
    return r.data.secure_url;
  } catch (e) { return null; }
}

// ── Excel 產生邏輯 ──
async function generateAndSendExcel(data, wgNumber, reporterName, photoUrl, photoUrl2) {
  try {
    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const ws = workbook.worksheets[0];

    const setCell = (addr, val) => { ws.getCell(addr).value = val || ''; };
    setCell('C2', wgNumber);
    setCell('D2', data.replyDate);
    setCell('A4', data.date);
    setCell('B4', data.unit);
    setCell('C4', data.resp);
    setCell('D4', data.customer);
    setCell('E4', data.product);
    setCell('F4', data.series);
    setCell('G4', data.anomaly);
    setCell('H4', data.qty);
    setCell('I4', data.orderNo);
    setCell('J4', data.ratio);
    setCell('K4', data.judge);
    setCell('L4', reporterName);
    setCell('M8', data.laborPeople);
    setCell('O8', data.laborHours);
    setCell('R8', data.adminPeople);
    setCell('T8', data.adminHours);
    setCell('X8', data.laborCost);
    setCell('M10', data.remark);

    const fetchBuf = async (url) => {
        try { const r = await axios.get(url, { responseType: 'arraybuffer' }); return Buffer.from(r.data); } 
        catch { return null; }
    };

    if (photoUrl) {
      const buf = await fetchBuf(photoUrl);
      if (buf) ws.addImage(workbook.addImage({ buffer: buf, extension: 'jpeg' }), { tl: { col: 0, row: 4 }, ext: { width: 360, height: 360 } });
    }
    if (photoUrl2) {
      const buf = await fetchBuf(photoUrl2);
      if (buf) ws.addImage(workbook.addImage({ buffer: buf, extension: 'jpeg' }), { tl: { col: 6, row: 4 }, ext: { width: 360, height: 360 } });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const downloadUrl = await uploadExcelToCloudinary(buffer, `${wgNumber}.xlsx`);

    if (downloadUrl) {
      const targets = EXCEL_NOTIFY_USERS.length > 0 ? EXCEL_NOTIFY_USERS : NOTIFY_USERS;
      for (const uid of targets) {
        await pushText(uid, `📋 品質異常通知單已產生！\n單號：${wgNumber}\n下載連結：${downloadUrl}`);
      }
    }
    return { downloadUrl };
  } catch (e) { console.error(e); return { downloadUrl: null }; }
}

// ── API Routes ──

app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    const reporterName = d.userId ? await getDisplayName(d.userId) : '(未知)';
    const wgNumber = genWGNumber();

    let photoUrl = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    let photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 1. 寫入 Notion
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    const properties = {
      '異常單號': { title: [{ text: { content: wgNumber } }] },
      '發生日期': { date: { start: new Date().toISOString().split('T')[0] } },
      '發生單位': { rich_text: toText(d.unit) },
      '責任單位': { rich_text: toText(d.resp) },
      '零件名稱': { rich_text: toText(d.product) },
      '系列別':   { rich_text: toText(d.series) },
      '異常狀況': { rich_text: toText(d.anomaly) },
      '判定':     { rich_text: toText(d.judge) },
      '訂單數量': { number: parseInt(d.qty) || null },
      '異常比例': { rich_text: toText(d.ratio) },
      '回報人':   { rich_text: toText(reporterName) }
    };
    if (photoUrl) properties['異常照片'] = { url: photoUrl };
    if (photoUrl2) properties['異常照片2'] = { url: photoUrl2 };
    
    await axios.post('https://api.notion.com/v1/pages', { parent: { database_id: NOTION_DATABASE_ID }, properties }, {
      headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28' }
    });

    // 2. 寫入 Google Sheets
    await appendToSheet({
      '異常單號': wgNumber, '發生日期': d.date || new Date().toISOString().split('T')[0],
      '需求回覆時間': d.replyDate || '', '發生單位': d.unit, '責任單位': d.resp, 
      '系列別': d.series, '零件名稱': d.product, '異常狀況': d.anomaly,
      '訂單數量': d.qty, '異常比例': d.ratio, '判定': d.judge, '回報人': reporterName,
      '異常照片': photoUrl, '異常照片2': photoUrl2
    });

    // 3. 發送詳細文字通報 (你要求恢復的部分)
    const judgeEmoji = d.judge.includes('驗退') ? '❌' : d.judge.includes('特採') ? '⚠️' : '🔧';
    const msg =
      `【異常通報 ${wgNumber}】\n` +
      `👤 回報人：${reporterName}\n` +
      `📦 品名：${d.product || '(未填)'}　系列：${d.series || ''}\n` +
      `📍 發生單位：${d.unit}\n` +
      `🏭 責任單位：${d.resp}\n` +
      `⚠️ 異常：${d.anomaly}\n` +
      `🔢 訂單數量：${d.qty}　比例：${d.ratio}\n` +
      `${judgeEmoji} 判定：${d.judge}\n` +
      `📅 日期：${d.date || new Date().toLocaleDateString()}` +
      (d.replyDate ? `\n📆 回覆期限：${d.replyDate}` : '') +
      (photoUrl  ? `\n📷 照片1：${photoUrl}`  : '') +
      (photoUrl2 ? `\n📷 照片2：${photoUrl2}` : '');

    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg).catch(e => console.error('push text failed:', e.message));
    }

    res.json({ success: true, number: wgNumber });
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.post('/api/generate-excel-from-sheet', async (req, res) => {
  try {
    const { data } = req.body;
    const mapped = {
      date: data['發生日期'], replyDate: data['需求回覆時間'], unit: data['發生單位'],
      resp: data['責任單位'], customer: data['客戶'], product: data['零件名稱'],
      series: data['系列別'], orderNo: data['單號'], anomaly: data['異常狀況'],
      qty: data['訂單數量'], ratio: data['異常比例'], judge: data['判定'],
      laborPeople: data['人工成本(人)'], laborHours: data['人工成本(時)'],
      adminPeople: data['行政成本(人)'], adminHours: data['行政成本(時)'],
      laborCost: data['所耗人力成本'], remark: data['異常標註內容']
    };
    const { downloadUrl } = await generateAndSendExcel(mapped, data['異常單號'], data['回報人'], data['異常照片'], data['異常照片2']);
    res.json({ success: !!downloadUrl, url: downloadUrl });
  } catch (err) { res.status(500).json({ success: false, error: err.message }); }
});

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server started on ${PORT}`));
