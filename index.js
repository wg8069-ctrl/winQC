const express = require('express');
const crypto = require('crypto');
const axios = require('axios');

const app = express();
app.use(express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }));

const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID = process.env.NOTION_DATABASE_ID;

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

async function replyText(replyToken, text) {
  await axios.post('https://api.line.me/v2/bot/message/reply',
    { replyToken, messages: [{ type: 'text', text }] },
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

  await axios.post('https://api.notion.com/v1/pages',
    { parent: { database_id: NOTION_DATABASE_ID }, properties },
    { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
  );
}

// 基本步驟（不含動態步驟）
const BASE_STEPS = [
  { key: 'location',   required: true,  ask: '📍 請輸入發生地點\n（例如：組裝線A）' },
  { key: 'productId',  required: true,  ask: '📦 請輸入產品編號\n（例如：WCB4-215B-CR）' },
  { key: 'itemName',   required: true,  ask: '🏷️ 請輸入品名' },
  { key: 'issue',      required: true,  ask: '⚠️ 請描述異常狀況' },
  { key: 'quantity',   required: true,  ask: '🔢 請輸入異常數量\n（純數字，例如：150）',
    validate: (v) => isNaN(v) ? '請輸入純數字！' : null },
  { key: 'solution',   required: true,  ask: '🔧 請輸入處理方式' },
  { key: 'vendor',     required: true,  ask: '🏭 請輸入異常廠商名稱' },
  { key: 'customer',   required: false, ask: '👥 請輸入客戶名稱\n（可輸入「略過」或「無」跳過）' },
  { key: 'caseNumber', required: false, ask: '📝 已開立異常單號？\n（有請輸入單號數字）\n（沒有請輸入「無」，之後需填免開原因）\n⚠️ 有填單號→狀態自動設為「處理中」' },
  // skipReason 是動態步驟，只在 caseNumber = 略過 時才問
];

const SKIP_REASON_STEP = {
  key: 'skipReason',
  required: true,
  ask: '🚫 請輸入免開異常原因'
};

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
    (d.customer   ? `👥 客戶：${d.customer}\n`        : '') +
    (d.caseNumber ? `📝 異常單號：${d.caseNumber}\n`  : '') +
    (d.skipReason ? `🚫 免開原因：${d.skipReason}\n`  : '') +
    `\n🔘 處理狀態：${status}\n` +
    `\n輸入「確認」送出\n輸入「重填」重新開始`
  );
}


async function searchNotion(keyword) {
  const filters = ['產品編號','品名','異常狀況','異常廠商','免開異常(請輸入原因)'].map(field => ({
    property: field,
    rich_text: { contains: keyword }
  }));

  // Also search by case number if keyword is numeric
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

async function handleMessage(event) {
  const userId = event.source?.userId;
  const replyToken = event.replyToken;
  const text = event.message?.type === 'text' ? event.message.text.trim() : null;
  if (!userId) return;

  let session = sessions[userId] || { step: 'idle', data: {} };

  // 查詢功能（任何步驟都可觸發）
  if (text && text.startsWith('查詢')) {
    const keyword = text.replace(/^查詢\s*/, '').trim();
    if (!keyword) {
      await replyText(replyToken, '請輸入查詢關鍵字\n例如：查詢 WCB4-215B-CR');
      return;
    }
    try {
      const results = await searchNotion(keyword);
      if (results.length === 0) {
        await replyText(replyToken, `🔍 查無「${keyword}」相關紀錄`);
      } else {
        const lines = [`🔍 找到 ${results.length} 筆「${keyword}」相關紀錄：\n`];
        results.slice(0, 5).forEach((r, i) => {
          lines.push(
            `${i + 1}. ${r.date} ${r.productId}\n` +
            `   📋 ${r.issue}\n` +
            `   🔢 ${r.quantity} pcs｜🔘 ${r.status}` +
            (r.caseNumber ? `\n   📝 單號：${r.caseNumber}` : '')
          );
        });
        if (results.length > 5) lines.push(`\n...共 ${results.length} 筆，僅顯示前 5 筆`);
        await replyText(replyToken, lines.join('\n'));
      }
    } catch (err) {
      console.error(err.response?.data || err.message);
      await replyText(replyToken, '❌ 查詢失敗，請通知管理員');
    }
    return;
  }

  // 重置
  if (text === '重填' || text === '取消') {
    delete sessions[userId];
    await replyText(replyToken, '已取消。\n\n輸入「回報異常」重新開始。');
    return;
  }

  // 開始
  if (session.step === 'idle' || text === '回報異常') {
    sessions[userId] = { step: 0, data: {} };
    await replyText(replyToken,
      '📋 開始填寫異常回報！\n\n隨時輸入「取消」結束\n\n' + BASE_STEPS[0].ask
    );
    return;
  }

  // 填寫基本步驟
  if (typeof session.step === 'number') {
    const cur = BASE_STEPS[session.step];
    if (!text) { await replyText(replyToken, '請輸入文字\n\n' + cur.ask); return; }
    if (cur.validate) {
      const err = cur.validate(text);
      if (err) { await replyText(replyToken, err + '\n\n' + cur.ask); return; }
    }
    session.data[cur.key] = (!cur.required && (text === '略過' || text === '無')) ? '' : text;
    const next = session.step + 1;

    if (next >= BASE_STEPS.length) {
      // 基本步驟填完，判斷是否需要問免開原因
      if (!session.data.caseNumber) {
        // 沒有填異常單號 → 問免開原因
        session.step = 'ask_skip_reason';
        sessions[userId] = session;
        await replyText(replyToken, '✅ 已記錄！\n\n' + SKIP_REASON_STEP.ask);
      } else {
        // 有填異常單號 → 直接跳到確認
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

  // 填寫免開原因（動態步驟）
  if (session.step === 'ask_skip_reason') {
    if (!text) { await replyText(replyToken, '請輸入免開異常原因\n\n' + SKIP_REASON_STEP.ask); return; }
    session.data.skipReason = text;
    session.step = 'confirm';
    sessions[userId] = session;
    await replyText(replyToken, buildSummary(session.data));
    return;
  }

  // 確認送出
  if (session.step === 'confirm') {
    if (text !== '確認') {
      await replyText(replyToken, '請輸入「確認」送出\n或輸入「重填」重新開始');
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
        `👤 回報人：${name}\n\n感謝回報！`
      );
    } catch (err) {
      console.error(err.response?.data || err.message);
      await replyText(replyToken, '❌ 寫入失敗，請通知管理員\n' + (err.response?.data?.message || err.message));
    }
    return;
  }

  await replyText(replyToken, '輸入「回報異常」開始回報\n輸入「查詢 關鍵字」查詢紀錄\n輸入「取消」結束目前流程');
}

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message') await handleMessage(event).catch(console.error);
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
