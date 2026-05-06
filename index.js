// v4 - final
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
const SHEET_HEADERS = ['異常單號','發生日期','需求回覆時間','發生單位','責任單位','系列別','單號','零件名稱','異常狀況','訂單數量','異常比例','目前處理狀態','判定','回報人','異常照片','異常照片2'];

async function getSheets(){
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ]
  });
  return google.sheets({ version: 'v4', auth });
}

let SHEET_NAME = null;

async function getSheetName(){
  if(SHEET_NAME) return SHEET_NAME;
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  SHEET_NAME = meta.data.sheets[0].properties.title;
  console.log('Sheet name:', SHEET_NAME);
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
      console.log('Sheet header created');
    }
  } catch(e){ console.error('ensureSheetHeader failed:', e.message); }
}

async function appendToSheet(dataMap){
  try {
    const sheets = await getSheets();
    const name = await getSheetName();
    // 讀取第一行標題
    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID, range: `${name}!1:1`
    });
    const headers = (headerRes.data.values && headerRes.data.values[0]) ? headerRes.data.values[0] : [];
    if(headers.length === 0){
      // 沒有標題就先建立
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [SHEET_HEADERS] }
      });
      headers.push(...SHEET_HEADERS);
    }
    // 依照標題順序填入資料
    const row = headers.map(function(h){ return dataMap[h] !== undefined ? dataMap[h] : ''; });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
    console.log('Sheet row appended, columns:', headers.length);
  } catch(e){ console.error('appendToSheet failed:', e.message); }
}

ensureSheetHeader();

// ════════════════════════════════════════
//  Helper 函式
// ════════════════════════════════════════
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

async function pushMessage(userId, messages) {
  await axios.post('https://api.line.me/v2/bot/message/push',
    { to: userId, messages },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
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
            { type: 'action', action: { type: 'message', label: '🔍 查詢紀錄', text: '查詢' } },
            { type: 'action', action: { type: 'uri', label: '📊 異常總表', uri: 'https://cream-scilla-479.notion.site/3361694680a780818f79cc3b72653beb?v=3361694680a78063aba2000c3ab24962' } }
          ]
        }
      }]
    },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

// ════════════════════════════════════════
//  Bot 對話邏輯
// ════════════════════════════════════════
async function handleMessage(event) {
  const userId     = event.source?.userId;
  const replyToken = event.replyToken;
  const text       = event.message?.type === 'text' ? event.message.text.trim() : null;
  const imageId    = event.message?.type === 'image' ? event.message.id : null;
  if (!userId) return;

  let session = sessions[userId] || { step: 'idle', data: {} };

  if (text === '0' || text === '選單' || text === 'menu') {
    delete sessions[userId]; await replyFlex(replyToken); return;
  }
  if (text === '重填' || text === '取消') {
    delete sessions[userId]; await replyFlex(replyToken); return;
  }

  if (session.step === 'idle') {
    if (text === '查詢' || text === '查詢紀錄') {
      sessions[userId] = { step: 'search_pick', data: {} };
      await axios.post('https://api.line.me/v2/bot/message/reply',
        {
          replyToken,
          messages: [{
            type: 'text',
            text: '🔍 請選擇查詢方式：',
            quickReply: {
              items: [
                { type: 'action', action: { type: 'message', label: '🔩 零件名稱', text: 'search:零件名稱' } },
                { type: 'action', action: { type: 'message', label: '📋 異常單號', text: 'search:異常單號' } },
                { type: 'action', action: { type: 'message', label: '🏭 發生單位', text: 'search:發生單位' } },
                { type: 'action', action: { type: 'message', label: '👤 回報人',   text: 'search:回報人'   } },
                { type: 'action', action: { type: 'message', label: '⚠️ 異常狀況', text: 'search:異常狀況' } },
              ]
            }
          }]
        },
        { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
      );
    } else {
      await replyFlex(replyToken);
    }
    return;
  }

  if (session.step === 'search_pick') {
    if (text && text.startsWith('search:')) {
      const field = text.replace('search:', '');
      sessions[userId] = { step: 'search_keyword', data: { field } };
      const hints = {
        '零件名稱': '例如：WC4、601、307。部分文字即可',
        '異常單號': '例如：2026040101。部分文字即可',
        '發生單位': '例如：一廠、品保、大元。部分文字即可',
        '回報人':   '例如：琛。部分文字即可',
        '異常狀況': '例如：外觀、尺寸、來料。部分文字即可',
      };
      await replyText(replyToken, `🔍 查詢【${field}】

輸入關鍵字（${hints[field] || '部分文字即可'}）：

輸入「0」回主選單`);
    } else {
      await replyText(replyToken, '請點選上方按鈕選擇查詢方式\n\n輸入「0」回主選單');
    }
    return;
  }

  if (session.step === 'search_keyword') {
    if (!text) { await replyText(replyToken, '請輸入關鍵字'); return; }
    const field = session.data.field;
    try {
      // 組 filter
      let filter;
      if (field === '異常單號') {
        filter = { property: '異常單號', title: { contains: text } };
      } else {
        filter = { property: field, rich_text: { contains: text } };
      }

      // DB ID 加上 hyphen 格式
      const dbId = NOTION_DATABASE_ID.includes('-')
        ? NOTION_DATABASE_ID
        : NOTION_DATABASE_ID.replace(/^(.{8})(.{4})(.{4})(.{4})(.{12})$/, '$1-$2-$3-$4-$5');

      const res = await axios.post(
        `https://api.notion.com/v1/databases/${dbId}/query`,
        { filter, page_size: 5 },
        { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
      );

      const results = res.data.results.map(p => {
        const props = p.properties;
        const t  = (k) => props[k]?.title?.[0]?.plain_text || props[k]?.title?.[0]?.text?.content || '';
        const rt = (k) => props[k]?.rich_text?.[0]?.plain_text || props[k]?.rich_text?.[0]?.text?.content || '';
        const dt = (k) => props[k]?.date?.start?.slice(0,10) || '';
        const st = (k) => props[k]?.status?.name || props[k]?.select?.name || rt(k) || '';
        const nm = (k) => props[k]?.number != null ? String(props[k].number) : '';
        return {
          num:      t('異常單號'),
          date:     dt('發生日期'),
          unit:     rt('發生單位'),
          resp:     rt('責任單位'),
          part:     rt('零件名稱'),
          series:   rt('系列別'),
          issue:    rt('異常狀況'),
          qty:      nm('訂單數量'),
          ratio:    rt('異常比例'),
          judge:    rt('判定'),
          status:   st('狀態') || rt('目前處理狀態'),
          reporter: rt('回報人'),
        };
      });

      if (results.length === 0) {
        await replyText(replyToken, `🔍 查無「${text}」相關紀錄

輸入「0」回主選單`);
      } else {
        const lines = [`🔍 找到 ${results.length} 筆「${text}」紀錄：
`];
        results.forEach((r, i) => {
          const je = r.judge.includes('驗退') ? '❌' : r.judge.includes('特採') ? '⚠️' : r.judge.includes('加工') ? '🔧' : '🔘';
          lines.push(
            `${i+1}. ${r.num}　${r.date}
` +
            `   📍 ${r.unit}｜🏭 ${r.resp}
` +
            `   🔩 ${r.part}　📂 ${r.series}
` +
            `   ⚠️ ${r.issue}　🔢 ${r.qty}
` +
            `   📊 ${r.ratio}　${je} ${r.judge}
` +
            `   🔘 ${r.status}　👤 ${r.reporter}`
          );
        });
        if (res.data.has_more) lines.push('\n...僅顯示前 5 筆');
        lines.push('\n輸入「0」回主選單');
        await replyText(replyToken, lines.join('\n'));
      }
    } catch (err) {
      console.error('search error full:', JSON.stringify(err.response?.data), 'url:', err.config?.url);
      console.error('search error:', err.response?.data || err.message);
      await replyText(replyToken, '❌ 查詢失敗，請通知管理員');
    }
    delete sessions[userId];
    return;
  }

  await replyFlex(replyToken);
}

// ════════════════════════════════════════
//  流水號
// ════════════════════════════════════════
const dailyCounters = {};
function genWGNumber() {
  const now = new Date();
  const y = now.getFullYear(), m = String(now.getMonth()+1).padStart(2,'0'), d = String(now.getDate()).padStart(2,'0');
  const key = `${y}${m}${d}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  return `WG${y}${m}${d}${String(dailyCounters[key]).padStart(2,'0')}`;
}

// ════════════════════════════════════════
//  Cloudinary 上傳圖片
// ════════════════════════════════════════
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
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, maxContentLength: Infinity, maxBodyLength: Infinity
    });
    return r.data.secure_url;
  } catch (e) { console.error('Cloudinary image upload failed:', e.response?.data || e.message); return null; }
}

// ════════════════════════════════════════
//  Cloudinary 上傳 Excel (raw)
// ════════════════════════════════════════
async function uploadExcelToCloudinary(buffer, filename) {
  try {
    const timestamp = Math.floor(Date.now()/1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const form = new FormData();
    form.append('file', buffer, { filename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    form.append('timestamp', String(timestamp));
    form.append('api_key', CLOUDINARY_KEY);
    form.append('signature', signature);
    form.append('resource_type', 'raw');
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/raw/upload`, form, {
      headers: form.getHeaders(), maxContentLength: Infinity, maxBodyLength: Infinity
    });
    return r.data.secure_url;
  } catch (e) { console.error('Cloudinary excel upload failed:', e.response?.data || e.message); return null; }
}

// ════════════════════════════════════════
//  產生 Excel 並傳送 LINE
// ════════════════════════════════════════
async function generateAndSendExcel(data, wgNumber, reporterName, photoUrl, photoUrl2) {
  try {
    const sheets = await getSheets();
    const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    });
    const drive = google.drive({ version: 'v3', auth });

    // 1. 找到「異常單範本」分頁的 sheetId
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const templateSheet = meta.data.sheets.find(s => s.properties.title === '異常單範本');
    if (!templateSheet) throw new Error('找不到「異常單範本」分頁');
    const templateSheetId = templateSheet.properties.sheetId;

    // 2. 複製「異常單範本」到全新的試算表
    const copyRes = await drive.files.copy({
      fileId: SPREADSHEET_ID,
      requestBody: { name: `output_${wgNumber}` }
    });
    const newSpreadsheetId = copyRes.data.id;

    // 3. 在新試算表裡只保留「異常單範本」分頁，刪除其他分頁
    const newMeta = await sheets.spreadsheets.get({ spreadsheetId: newSpreadsheetId });
    const newTemplateSheet = newMeta.data.sheets.find(s => s.properties.title === '異常單範本');
    const otherSheets = newMeta.data.sheets.filter(s => s.properties.title !== '異常單範本');
    if (otherSheets.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: newSpreadsheetId,
        requestBody: { requests: otherSheets.map(s => ({ deleteSheet: { sheetId: s.properties.sheetId } })) }
      });
    }

    // 4. 替換 {{標記}} 填入資料
    const dataMap = {
      '{{異常單號}}':     wgNumber,
      '{{需求回覆時間}}': data.replyDate || '',
      '{{發生日期}}':     data.date || new Date().toISOString().split('T')[0],
      '{{發生單位}}':     data.unit || '',
      '{{責任單位}}':     data.resp || '',
      '{{客戶}}':         data.customer || '',
      '{{零件名稱}}':     data.product || '',
      '{{系列別}}':       data.series || '',
      '{{單號}}':         data.orderNo || '',
      '{{異常狀況}}':     data.anomaly || '',
      '{{訂單數量}}':     data.qty || '',
      '{{異常比例}}':     data.ratio || '',
      '{{判定}}':         data.judge || '',
      '{{回報人}}':       reporterName,
      '{{人工成本人}}':   data.laborPeople || '',
      '{{人工成本時}}':   data.laborHours || '',
      '{{行政成本人}}':   data.adminPeople || '',
      '{{行政成本時}}':   data.adminHours || '',
      '{{所耗人力成本}}': data.laborCost || '',
      '{{異常照片}}':     photoUrl || '',
      '{{異常照片2}}':    photoUrl2 || '',
    };

    const cellRes = await sheets.spreadsheets.values.get({
      spreadsheetId: newSpreadsheetId, range: '異常單範本!A1:Z200'
    });
    const rows = cellRes.data.values || [];
    const updatedRows = rows.map(row => row.map(cell => {
      let val = String(cell || '');
      Object.entries(dataMap).forEach(([k, v]) => { val = val.replace(k, v); });
      return val;
    }));

    await sheets.spreadsheets.values.update({
      spreadsheetId: newSpreadsheetId, range: '異常單範本!A1',
      valueInputOption: 'RAW', requestBody: { values: updatedRows }
    });

    // 5. 匯出成 xlsx（只有一個分頁所以就是範本那頁）
    const exportRes = await drive.files.export({
      fileId: newSpreadsheetId,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }, { responseType: 'arraybuffer' });

    const buffer = Buffer.from(exportRes.data);
    const filename = `${wgNumber}.xlsx`;

    // 6. 刪除暫時試算表
    await drive.files.delete({ fileId: newSpreadsheetId })
      .catch(e => console.error('delete temp sheet failed:', e.message));

    // 7. 上傳到 Cloudinary
    const downloadUrl = await uploadExcelToCloudinary(buffer, filename);

    // 8. 傳送 LINE 訊息
    if (downloadUrl) {
      const targets = EXCEL_NOTIFY_USERS.length > 0 ? EXCEL_NOTIFY_USERS : NOTIFY_USERS;
      for (const uid of targets) {
        await pushText(uid, `📋 品質異常通知單已產生！\n\n異常單號：${wgNumber}\n\n點擊下載 Excel：\n${downloadUrl}`)
          .catch(e => console.error('push excel link failed:', e.message));
      }
    }

    return { buffer, downloadUrl };
  } catch (e) {
    console.error('generateAndSendExcel failed:', e.message);
    return { buffer: null, downloadUrl: null };
  }
}

// ════════════════════════════════════════
//  API Routes
// ════════════════════════════════════════

// LIFF 表單提交
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    let reporterName = '(未知)';
    console.log('anomaly submitted by userId:', d.userId || '(no userId)');
    if (d.userId) reporterName = await getDisplayName(d.userId);

    const wgNumber = genWGNumber();

    // 上傳照片
    let photoUrl = null, photoUrl2 = null;
    if (d.photoData)  photoUrl  = await uploadToCloudinary(d.photoData);
    if (d.photoData2) photoUrl2 = await uploadToCloudinary(d.photoData2);

    // 寫入 Notion
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    const properties = {
      '異常單號':     { title: [{ text: { content: wgNumber } }] },
      '發生日期':     { date: { start: new Date().toISOString().split('T')[0] } },
      '發生單位':     { rich_text: toText(d.unit || '') },
      '責任單位':     { rich_text: toText(d.resp || '') },
      '客戶':         { rich_text: toText(d.customer || '') },
      '單號':           { rich_text: toText(d.orderNo || '') },
      '系列別':       { rich_text: toText(d.series || '') },
      '零件名稱':     { rich_text: toText(d.product || '') },
      '異常狀況':     { rich_text: toText(d.anomaly || '') },
      '判定':         { rich_text: toText(d.judge || '') },
      '訂單數量':     { number: parseInt(d.qty) || null },
      '異常比例':     { rich_text: toText(d.ratioText || '') },
      '不良數 / 抽驗數': { rich_text: toText(d.bad && d.samp ? `${d.bad} / ${d.samp}` : '') },
      '目前處理狀態': { rich_text: toText(d.status || '未開始') },
      '回報人':       { rich_text: toText(reporterName) },
    };
    if (d.replyDate) properties['需求回覆時間'] = { rich_text: [{ text: { content: d.replyDate } }] };
    if (photoUrl)    properties['異常照片']  = { url: photoUrl };
    if (photoUrl2)   properties['異常照片2'] = { url: photoUrl2 };

    const pageBody = { parent: { database_id: NOTION_DATABASE_ID }, properties };
    if (photoUrl || photoUrl2) {
      pageBody.children = [];
      if (photoUrl)  pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl  } } });
      if (photoUrl2) pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl2 } } });
    }
    await axios.post('https://api.notion.com/v1/pages', pageBody, {
      headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' }
    });

    // 寫入 Google Sheets（同步）
    appendToSheet({
      '異常單號':     wgNumber,
      '發生日期':     new Date().toISOString().split('T')[0],
      '需求回覆時間': d.replyDate || '',
      '發生單位':     d.unit || '',
      '責任單位':     d.resp || '',
      '系列別':       d.series || '',
      '單號':         d.orderNo || '',
      '零件名稱':     d.product || '',
      '異常狀況':     d.anomaly || '',
      '訂單數量':     parseInt(d.qty) || '',
      '異常比例':     d.ratio || '',
      '目前處理狀態': d.status || '未開始',
      '判定':         d.judge || '',
      '回報人':       reporterName,
      '人工成本(人)': d.laborPeople || '',
      '人工成本(時)': d.laborHours || '',
      '行政成本(人)': d.adminPeople || '',
      '行政成本(時)': d.adminHours || '',
      '所耗人力成本': d.laborCost || '',
      '異常照片':     photoUrl || '',
      '異常照片2':    photoUrl2 || '',
    });

    // 推播異常通知
    const judgeEmoji = d.judge === '驗退X' ? '❌' : d.judge === '特採△' ? '⚠️' : '🔧';
    const msg =
      `【異常通報 ${wgNumber}】\n` +
      `👤 回報人：${reporterName}\n` +
      `📦 品名：${d.product || '(未填)'}　系列：${d.series || ''}\n` +
      `📍 發生單位：${d.unit}\n` +
      `🏭 責任單位：${d.resp}\n` +
      `⚠️ 異常：${d.anomaly}\n` +
      `🔢 訂單數量：${d.qty}　比例：${d.ratio}\n` +
      `${judgeEmoji} 判定：${d.judge}\n` +
      `📅 日期：${d.date}` +
      (d.replyDate ? `\n📆 回覆期限：${d.replyDate}` : '') +
      (photoUrl  ? `\n📷 照片1：${photoUrl}`  : '') +
      (photoUrl2 ? `\n📷 照片2：${photoUrl2}` : '');

    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg).catch(e => console.error('push failed:', e.message));
    }

    // 非同步產生 Excel 並傳送下載連結
    generateAndSendExcel(d, wgNumber, reporterName, photoUrl, photoUrl2)
      .then(({ downloadUrl }) => {
        if (downloadUrl) console.log('Excel uploaded:', downloadUrl);
      })
      .catch(e => console.error('Excel background task failed:', e.message));

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Webhook
app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message') await handleMessage(event).catch(console.error);
    if (event.type === 'follow')  await replyFlex(event.replyToken).catch(console.error);
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅', routes: ['/api/anomaly', '/webhook'] }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
