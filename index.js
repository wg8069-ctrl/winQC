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
            { type: 'action', action: { type: 'uri', label: '📊 異常總表', uri: 'https://cream-scilla-479.notion.site/3281694680a7800f984dd246bd4e7904?v=3281694680a780e3b597000c7979345a' } }
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
      const fieldLabels = {
        '零件名稱': '零件名稱（例如：WC4-795B）',
        '異常單號': '異常單號（例如：WG20260326）',
        '發生單位': '發生單位（例如：本廠）',
        '回報人':   '回報人姓名',
        '異常狀況': '異常狀況關鍵字',
      };
      await replyText(replyToken, `🔍 查詢 ${field}\n\n請輸入${fieldLabels[field] || '關鍵字'}：\n\n輸入「0」回主選單`);
    } else {
      await replyText(replyToken, '請點選上方按鈕選擇查詢方式\n\n輸入「0」回主選單');
    }
    return;
  }

  if (session.step === 'search_keyword') {
    if (!text) { await replyText(replyToken, '請輸入關鍵字'); return; }
    const field = session.data.field;
    try {
      const filter = field === '異常單號'
        ? { property: '異常單號', title: { contains: text } }
        : { property: field, rich_text: { contains: text } };
      console.log('Query DB:', NOTION_DATABASE_ID);
      const res = await axios.post(
        `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`,
        { filter, sorts: [{ property: '發生日期', direction: 'descending' }], page_size: 5 },
        { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2025-09-03', 'Content-Type': 'application/json' } }
      );
      const results = res.data.results.map(p => {
        const props   = p.properties;
        const getText = (k) => props[k]?.rich_text?.[0]?.text?.content || props[k]?.title?.[0]?.text?.content || '';
        const getDate = (k) => props[k]?.date?.start?.slice(0,10) || '';
        const getUrl  = (k) => props[k]?.url || '';
        return {
          num:      getText('異常單號'),
          date:     getDate('發生日期'),
          unit:     getText('發生單位'),
          part:     getText('零件名稱'),
          series:   getText('系列別'),
          issue:    getText('異常狀況'),
          ratio:    getText('異常比例'),
          judge:    getText('判定'),
          status:   getText('目前處理狀態'),
          reporter: getText('回報人'),
          photo:    getUrl('異常照片'),
        };
      });
      if (results.length === 0) {
        await replyText(replyToken, `🔍 查無「${text}」相關紀錄\n\n輸入「0」回主選單`);
      } else {
        const lines = [`🔍 找到 ${res.data.results.length} 筆「${text}」紀錄：\n`];
        results.forEach((r, i) => {
          const judgeEmoji = r.judge.includes('驗退') ? '❌' : r.judge.includes('特採') ? '⚠️' : r.judge.includes('加工') ? '🔧' : '🔘';
          lines.push(
            `${i+1}. ${r.num} ${r.date}\n` +
            `   🏭 ${r.unit}　👤 ${r.reporter}\n` +
            `   🔩 ${r.part}　📂 ${r.series}\n` +
            `   ⚠️ ${r.issue}\n` +
            `   📊 ${r.ratio}　${judgeEmoji} ${r.judge}` +
            (r.photo ? `\n   📷 ${r.photo}` : '')
          );
        });
        if (res.data.has_more) lines.push(`\n...僅顯示前 5 筆`);
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
    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const ws = workbook.worksheets[0];

    // 填入儲存格
    const setCell = (addr, val) => { try { ws.getCell(addr).value = val; } catch(e) {} };
    setCell('C2', wgNumber);
    setCell('D2', data.replyDate || '');
    setCell('A4', data.date || new Date().toISOString().split('T')[0]);
    setCell('B4', data.unit || '');
    setCell('C4', data.resp || '');
    setCell('D4', data.customer || '');
    setCell('E4', data.product || '');
    setCell('F4', data.series || '');
    setCell('G4', data.anomaly || '');
    setCell('H4', parseInt(data.qty) || null);
    setCell('J4', data.ratio || '');
    setCell('K4', data.judge || '');
    setCell('L4', reporterName);

    // 嵌入照片
    const fetchBuf = (url) => new Promise((resolve) => {
      const mod = url.startsWith('https') ? require('https') : require('http');
      mod.get(url, (res) => {
        const chunks = [];
        res.on('data', c => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
        res.on('error', () => resolve(null));
      }).on('error', () => resolve(null));
    });

    for (const [url, col, row] of [[photoUrl, 0, 4], [photoUrl2, 6, 4]]) {
      if (url) {
        const imgBuf = await fetchBuf(url);
        if (imgBuf) {
          const imgId = workbook.addImage({ buffer: imgBuf, extension: 'jpeg' });
          ws.addImage(imgId, { tl: { col, row }, ext: { width: 360, height: 360 } });
        }
      }
    }

    const buffer   = await workbook.xlsx.writeBuffer();
    const filename = `${wgNumber}.xlsx`;

    // 上傳到 Cloudinary
    const downloadUrl = await uploadExcelToCloudinary(buffer, filename);

    // 傳送 LINE 訊息給通知對象
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
      '異常比例':     { rich_text: toText(d.ratio || '') },
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
    console.log('Writing to DB:', NOTION_DATABASE_ID);
    await axios.post('https://api.notion.com/v1/pages', pageBody, {
      headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' }
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
