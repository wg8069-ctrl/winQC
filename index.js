const express = require('express');
const crypto = require('crypto');
const axios = require('axios');
const htmlPdf = require('html-pdf-node');

const app = express();
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.json({ limit: '10mb', verify: (req, res, buf) => { req.rawBody = buf; } }));

const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID = process.env.NOTION_DATABASE_ID;
const NOTIFY_USERS = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);

const sessions = {};

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
    '異常單號':                  { title: [{ text: { content: data.location || '(未填)' } }] },
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
  const filters = ['零件名稱','異常狀況','發生單位','回報人'].map(field => ({
    property: field, rich_text: { contains: keyword }
  }));
  const res = await axios.post(
    `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`,
    { filter: { or: filters }, sorts: [{ property: '發生日期', direction: 'descending' }] },
    { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
  );
  return res.data.results.map(p => {
    const props = p.properties;
    const getText = (k) => props[k]?.rich_text?.[0]?.text?.content || props[k]?.title?.[0]?.text?.content || '';
    const getNum  = (k) => props[k]?.number ?? '';
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
}

const MAIN_MENU = '請使用下方選單進行操作 👇';

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
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

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

  console.log('userId:', userId);

  let session = sessions[userId] || { step: 'idle', data: {} };

  if (text === '0' || text === '選單' || text === 'menu') {
    delete sessions[userId];
    await replyFlex(replyToken);
    return;
  }

  if (text === '重填' || text === '取消') {
    delete sessions[userId];
    await replyFlex(replyToken);
    return;
  }

  if (session.step === 'idle') {
    if (text === '1' || text === '回報異常') {
      sessions[userId] = { step: 0, data: {} };
      await replyText(replyToken, '📋 開始填寫異常回報！\n隨時輸入「0」回主選單\n\n' + BASE_STEPS[0].ask);
    } else if (text === '2' || text === '查詢紀錄' || text === '查詢') {
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
        '異常狀況': '異常狀況關鍵字（例如：斷差）',
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

      const res = await axios.post(
        `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`,
        { filter, sorts: [{ property: '發生日期', direction: 'descending' }], page_size: 5 },
        { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
      );
      const results = res.data.results.map(p => {
        const props = p.properties;
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
            `   📊 ${r.ratio}　${judgeEmoji} ${r.judge}\n` +
            (r.photo ? `   📷 ${r.photo}` : '')
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

  await replyFlex(replyToken);
}

// ════════════════════════════════════════
//  流水號產生器
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
//  PDF 產生（HTML → PDF）+ 上傳 Cloudinary
// ════════════════════════════════════════
function buildPdfHtml(data) {
  const j = data.judge || '';
  const isVerify  = j.includes('驗退');
  const isSpecial = j.includes('特採');
  const isProcess = j.includes('加工');
  const chk = (v) => v ? '☑' : '☐';
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Microsoft JhengHei','PingFang TC','Noto Sans CJK TC',Arial,sans-serif;font-size:11px;padding:8px;width:100%;}
table{width:100%;border-collapse:collapse;table-layout:fixed;}
td{border:1px solid #000;padding:2px 4px;font-size:10px;vertical-align:middle;overflow:hidden;}
.nb{border:none;}
.bold{font-weight:bold;}
.center{text-align:center;}
.bg{background:#e0e0e0;}
.big{font-size:16px;font-weight:bold;}
.med{font-size:13px;font-weight:bold;}
.tall{height:180px;vertical-align:top;padding:4px;}
.sig{height:60px;}
.note{font-size:8px;margin-top:6px;line-height:1.6;}
.right-panel{font-size:10px;}
</style>
</head>
<body>
<table>
  <!-- 標題列 -->
  <tr>
    <td colspan="2" rowspan="2" class="nb" style="width:90px;font-size:13px;font-weight:bold;border:2px solid #000;text-align:center;">
      偉剛科技<br>WinGun
    </td>
    <td colspan="8" rowspan="2" class="big center" style="border:2px solid #000;">
      品質異常通知單&nbsp;&nbsp;&nbsp;品質判定：驗退X. 特採△. 加工○
    </td>
    <td colspan="3" class="center bold bg" style="font-size:10px;">外發加工</td>
    <td colspan="2" class="center"></td>
    <td colspan="2" class="center bold bg">試組</td>
    <td colspan="2" class="center bold bg">線前加工</td>
    <td colspan="2" class="center bold bg">上線中</td>
  </tr>
  <tr>
    <td colspan="2" class="center bold bg">特採</td>
    <td colspan="2" class="center bold bg">退料</td>
    <td colspan="5" class="center bold bg">預計回廠時間</td>
    <td colspan="2" class="center"></td>
  </tr>
  <!-- 欄位標題列 -->
  <tr style="height:28px;">
    <td class="center bg bold" style="width:55px;">發生日期</td>
    <td class="center bg bold" style="width:55px;">發生單位</td>
    <td class="center bg bold" style="width:55px;">責任單位</td>
    <td class="center bg bold" style="width:45px;">客戶</td>
    <td class="center bg bold" style="width:80px;">零件名稱</td>
    <td class="center bg bold" style="width:60px;">零件編號</td>
    <td class="center bg bold" style="width:60px;">異常狀況</td>
    <td class="center bg bold" style="width:50px;">訂單數量</td>
    <td class="center bg bold" style="width:50px;">異常數量</td>
    <td class="center bg bold" style="width:50px;">異常比例</td>
    <td class="center bg bold" style="width:40px;">判定</td>
    <td class="center bg bold" style="width:40px;">確認</td>
    <td colspan="3" class="center bold bg" style="font-size:10px;">上線日期</td>
    <td colspan="4" class="center bold bg" style="font-size:10px;">出貨日/待驗日</td>
    <td colspan="3" class="center bold bg right-panel">
      產線處理 &nbsp; 來廠處理 &nbsp; 外廠處理
    </td>
  </tr>
  <!-- 資料列 -->
  <tr style="height:30px;">
    <td class="center">${data.date || ''}</td>
    <td class="center">${data.unit || ''}</td>
    <td class="center">${data.resp || ''}</td>
    <td class="center">${data.customer || ''}</td>
    <td class="center">${data.product || ''}</td>
    <td class="center">${data.series || ''}</td>
    <td class="center">${data.anomaly || ''}</td>
    <td class="center">${data.qty || ''}</td>
    <td class="center"></td>
    <td class="center">${data.ratio || ''}</td>
    <td class="center bold">${j}</td>
    <td class="center">${data.reporter || ''}</td>
    <td colspan="3" class="center"></td>
    <td colspan="4" class="center"></td>
    <td colspan="3" class="right-panel center">
      ${chk(false)} 產線 &nbsp; ${chk(false)} 來廠 &nbsp; ${chk(false)} 外廠
    </td>
  </tr>
  <!-- 異常狀況 / 處理方式 標題 -->
  <tr>
    <td colspan="6" class="center bold bg">異常狀況</td>
    <td colspan="6" class="center bold bg">處理方式</td>
    <td colspan="9" class="right-panel" style="font-size:9px;">
      補料狀況：${chk(false)} 廠商補 &nbsp; ${chk(false)} 庫出 &nbsp; 補料時間：____<br>
      ${chk(false)} 停線 &nbsp; ${chk(false)} 無法一次完線 &nbsp; ${chk(false)} 再上線<br>
      庫存數量：____ &nbsp; 庫存處理：${chk(false)} 加工 &nbsp; ${chk(false)} 報廢
    </td>
  </tr>
  <!-- 異常狀況 / 處理方式 內容 -->
  <tr>
    <td colspan="6" class="tall">${(data.anomaly || '').replace(/、/g,'<br>')}</td>
    <td colspan="6" class="tall">${j}</td>
    <td colspan="9" class="tall" style="font-size:9px;vertical-align:top;">
      <strong>${data.resp || ''}</strong>
    </td>
  </tr>
  <!-- 廠商異常處理 -->
  <tr>
    <td colspan="3" class="bold bg">廠商異常處理</td>
    <td colspan="9"></td>
    <td colspan="9" class="center bold bg">廠商簽回</td>
  </tr>
  <tr class="sig">
    <td colspan="12"></td>
    <td colspan="9" class="center bold bg">簽核</td>
  </tr>
</table>
<div class="note">
  1.請於通知單到後3日內完成問題回覆並回傳，否則視同確認並以我司處理方式處理；無不可抗力因素且未回傳者則當月票期加開乙個月。如2個月未改善則終止合作。<br>
  2.若於次月無異常通知則票期可提前一個月；若連續2個月無異常則以現金票支付貨款。<br>
  3.生產前務必比對成品與樣品無誤；如有不符樣品需告知本司進行處理；未告知而逕行交貨者由製造者負責後續發生所有費用。
</div>
</body>
</html>`;
}

async function generateAndUploadPDF(data) {
  try {
    const html = buildPdfHtml(data);
    const file = { content: html };
    const options = { format: 'A4', margin: { top: '15mm', bottom: '15mm', left: '15mm', right: '15mm' } };
    const pdfBuffer = await htmlPdf.generatePdf(file, options);
    const base64 = 'data:application/pdf;base64,' + pdfBuffer.toString('base64');

    const formData = new URLSearchParams();
    formData.append('file', base64);
    formData.append('api_key', CLOUDINARY_KEY);
    formData.append('upload_preset', 'ml_default');

    const r = await axios.post(
      `https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/raw/upload`,
      formData.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, maxContentLength: Infinity, maxBodyLength: Infinity }
    );
    return r.data.secure_url;
  } catch(e) {
    console.error('PDF gen/upload failed:', e.message);
    return null;
  }
}

// ════════════════════════════════════════
//  LIFF 表單 API
// ════════════════════════════════════════
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;

    let reporterName = '(未知)';
    if (d.userId) {
      reporterName = await getDisplayName(d.userId);
    }

    const wgNumber = genWGNumber();

    let photoUrl = null;
    let photoUrl2 = null;
    if (d.photoData)  photoUrl  = await uploadToCloudinary(d.photoData);
    if (d.photoData2) photoUrl2 = await uploadToCloudinary(d.photoData2);

    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    const properties = {
      '異常單號':     { title: [{ text: { content: wgNumber } }] },
      '發生日期':     { date: { start: new Date().toISOString().split('T')[0] } },
      '發生單位':     { rich_text: toText(d.unit || '') },
      '責任單位':     { rich_text: toText(d.resp || '') },
      '客戶':         { rich_text: toText('') },
      '系列別':       { rich_text: toText(d.series || '') },
      '零件名稱':     { rich_text: toText(d.product || '') },
      '異常狀況':     { rich_text: toText(d.anomaly || '') },
      '處理方式':     { rich_text: toText(d.judge || '') },
      '判定':         { rich_text: toText(d.judge || '') },
      '訂單數量':     { number: parseInt(d.qty) || null },
      '異常比例':     { rich_text: toText(d.ratio || '') },
      '目前處理狀態': { rich_text: toText('未開始') },
      '回報人':       { rich_text: toText(reporterName) },
      '異常照片':     photoUrl  ? { url: photoUrl  } : { url: null },
      '異常照片2':    photoUrl2 ? { url: photoUrl2 } : { url: null },
    };

    const pageBody = { parent: { database_id: NOTION_DATABASE_ID }, properties };

    if (photoUrl || photoUrl2) {
      pageBody.children = [];
      if (photoUrl)  pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl  } } });
      if (photoUrl2) pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl2 } } });
    }

    await axios.post('https://api.notion.com/v1/pages', pageBody, {
      headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' }
    });

    const judgeEmoji = d.judge === '驗退X' ? '❌' : d.judge === '特採△' ? '⚠️' : '🔧';

    // 產生 PDF
    const pdfUrl = await generateAndUploadPDF({
      wgNumber, date: d.date || new Date().toISOString().split('T')[0],
      unit: d.unit, resp: d.resp, series: d.series,
      product: d.product, qty: d.qty, ratio: d.ratio,
      anomaly: d.anomaly, judge: d.judge, reporter: reporterName,
    }).catch(e => { console.error('PDF gen failed:', e.message); return null; });

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
      (photoUrl  ? `\n📷 照片1：${photoUrl}`  : '') +
      (photoUrl2 ? `\n📷 照片2：${photoUrl2}` : '') +
      (pdfUrl    ? `\n📄 異常單PDF：${pdfUrl}` : '');

    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg).catch(e => console.error('push failed:', e.message));
    }

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message') await handleMessage(event).catch(console.error);
    if (event.type === 'follow') await replyFlex(event.replyToken).catch(console.error);
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
