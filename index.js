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

  // 處理狀態邏輯：有填異常單號 → 處理中，否則 → 未開始
  const status = data.caseNumber ? '處理中' : '未開始';

  const properties = {
    '發生地':                    { title: [{ text: { content: data.location || '(未填)' } }] },
    '產品編號':                  { rich_text: toText(data.productId) },
    '品名':                      { rich_text: toText(data.itemName) },
    '異常狀況':                  { rich_text: toText(data.issue) },
    '廠商':                      { rich_text: toText(data.vendor) },
    '客戶':                      { rich_text: toText(data.customer) },
    '處理方式':                  { rich_text: toText(data.solution) },
    '數量':                      data.quantity ? { number: parseInt(data.quantity) } : { number: null },
    '發生日期':                  { date: { start: new Date().toISOString().split('T')[0] } },
    '已開立異常單(請輸入單號)':   { rich_text: toText(data.caseNumber) },
    '免開異常(請輸入原因)':       { rich_text: toText(data.skipReason) },
    '處理狀態':                  { select: { name: status } },
  };

  await axios.post('https://api.notion.com/v1/pages',
    { parent: { database_id: NOTION_DATABASE_ID }, properties },
    { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
  );
}

// 對話步驟
const STEPS = [
  { key: 'location',   required: true,  ask: '📍 請輸入發生地點\n（例如：組裝線A）' },
  { key: 'productId',  required: true,  ask: '📦 請輸入產品編號\n（例如：WCB4-215B-CR）' },
  { key: 'itemName',   required: true,  ask: '🏷️ 請輸入品名' },
  { key: 'issue',      required: true,  ask: '⚠️ 請描述異常狀況' },
  { key: 'quantity',   required: true,  ask: '🔢 請輸入異常數量\n（純數字，例如：150）',
    validate: (v) => isNaN(v) ? '請輸入純數字！' : null },
  { key: 'solution',   required: true,  ask: '🔧 請輸入處理方式' },
  { key: 'vendor',     required: true,  ask: '🏭 請輸入廠商名稱' },
  { key: 'customer',   required: false, ask: '👥 請輸入客戶名稱\n（可輸入「略過」跳過）' },
  { key: 'caseNumber', required: false, ask: '📝 已開立異常單號？\n（有請輸入單號，沒有請輸入「略過」）\n⚠️ 有填單號→狀態自動設為「處理中」' },
  { key: 'skipReason', required: false, ask: '🚫 免開異常原因？\n（有請輸入原因，沒有請輸入「略過」）' },
];

async function handleMessage(event) {
  const userId = event.source?.userId;
  const replyToken = event.replyToken;
  const text = event.message?.type === 'text' ? event.message.text.trim() : null;
  if (!userId) return;

  let session = sessions[userId] || { step: 'idle', data: {} };

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
      '📋 開始填寫異常回報！\n必填 7 項＋選填 3 項\n\n隨時輸入「取消」結束\n\n' + STEPS[0].ask
    );
    return;
  }

  // 填寫中
  if (typeof session.step === 'number') {
    const cur = STEPS[session.step];

    if (!text) {
      await replyText(replyToken, '請輸入文字\n\n' + cur.ask);
      return;
    }

    if (cur.validate) {
      const err = cur.validate(text);
      if (err) { await replyText(replyToken, err + '\n\n' + cur.ask); return; }
    }

    session.data[cur.key] = (!cur.required && text === '略過') ? '' : text;
    const next = session.step + 1;

    if (next >= STEPS.length) {
      session.step = 'confirm';
      sessions[userId] = session;
      const d = session.data;
      const status = d.caseNumber ? '處理中' : '未開始';
      const summary =
        `📋 請確認以下資料：\n\n` +
        `📍 發生地：${d.location}\n` +
        `📦 產品編號：${d.productId}\n` +
        `🏷️ 品名：${d.itemName}\n` +
        `⚠️ 異常狀況：${d.issue}\n` +
        `🔢 數量：${d.quantity}\n` +
        `🔧 處理方式：${d.solution}\n` +
        `🏭 廠商：${d.vendor}\n` +
        (d.customer   ? `👥 客戶：${d.customer}\n`        : '') +
        (d.caseNumber ? `📝 異常單號：${d.caseNumber}\n`  : '') +
        (d.skipReason ? `🚫 免開原因：${d.skipReason}\n`  : '') +
        `\n🔘 處理狀態：${status}\n` +
        `\n輸入「確認」送出\n輸入「重填」重新開始`;
      await replyText(replyToken, summary);
    } else {
      session.step = next;
      sessions[userId] = session;
      await replyText(replyToken, '✅ 已記錄！\n\n' + STEPS[next].ask);
    }
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

  await replyText(replyToken, '輸入「回報異常」開始\n輸入「取消」結束目前流程');
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
