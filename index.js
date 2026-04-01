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

// â”€â”€ ç’°å¢ƒè®Šæ•¸ â”€â”€
const LINE_CHANNEL_SECRET       = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN              = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID        = process.env.NOTION_DATABASE_ID;
const NOTIFY_USERS              = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);
const CLOUDINARY_CLOUD          = 'dlpxz4qlh';
const CLOUDINARY_KEY            = '953226455671951';
const CLOUDINARY_SECRET         = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';

const sessions = {};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  Helper å‡½å¼
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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
    return r.data.displayName || 'ç”¨æˆ¶';
  } catch { return 'ç”¨æˆ¶'; }
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
        text: 'WinGun ç•°å¸¸é€šå ± ðŸ‘‡',
        quickReply: {
          items: [
            { type: 'action', action: { type: 'uri', label: 'ðŸ“‹ å»ºç«‹ç•°å¸¸å–®', uri: 'https://liff.line.me/2009600334-UpN6esDu' } },
            { type: 'action', action: { type: 'message', label: 'ðŸ” æŸ¥è©¢ç´€éŒ„', text: 'æŸ¥è©¢' } },
            { type: 'action', action: { type: 'uri', label: 'ðŸ“Š ç•°å¸¸ç¸½è¡¨', uri: 'https://cream-scilla-479.notion.site/3281694680a7800f984dd246bd4e7904?v=3281694680a780e3b597000c7979345a' } }
          ]
        }
      }]
    },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`, 'Content-Type': 'application/json' } }
  );
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  Bot å°è©±é‚è¼¯
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function handleMessage(event) {
  const userId     = event.source?.userId;
  const replyToken = event.replyToken;
  const text       = event.message?.type === 'text' ? event.message.text.trim() : null;
  const imageId    = event.message?.type === 'image' ? event.message.id : null;
  if (!userId) return;

  let session = sessions[userId] || { step: 'idle', data: {} };

  if (text === '0' || text === 'é¸å–®' || text === 'menu') {
    delete sessions[userId]; await replyFlex(replyToken); return;
  }
  if (text === 'é‡å¡«' || text === 'å–æ¶ˆ') {
    delete sessions[userId]; await replyFlex(replyToken); return;
  }

  if (session.step === 'idle') {
    if (text === 'æŸ¥è©¢' || text === 'æŸ¥è©¢ç´€éŒ„') {
      sessions[userId] = { step: 'search_pick', data: {} };
      await axios.post('https://api.line.me/v2/bot/message/reply',
        {
          replyToken,
          messages: [{
            type: 'text',
            text: 'ðŸ” è«‹é¸æ“‡æŸ¥è©¢æ–¹å¼ï¼š',
            quickReply: {
              items: [
                { type: 'action', action: { type: 'message', label: 'ðŸ”© é›¶ä»¶åç¨±', text: 'search:é›¶ä»¶åç¨±' } },
                { type: 'action', action: { type: 'message', label: 'ðŸ“‹ ç•°å¸¸å–®è™Ÿ', text: 'search:ç•°å¸¸å–®è™Ÿ' } },
                { type: 'action', action: { type: 'message', label: 'ðŸ­ ç™¼ç”Ÿå–®ä½', text: 'search:ç™¼ç”Ÿå–®ä½' } },
                { type: 'action', action: { type: 'message', label: 'ðŸ‘¤ å›žå ±äºº',   text: 'search:å›žå ±äºº'   } },
                { type: 'action', action: { type: 'message', label: 'âš ï¸ ç•°å¸¸ç‹€æ³', text: 'search:ç•°å¸¸ç‹€æ³' } },
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
        'é›¶ä»¶åç¨±': 'é›¶ä»¶åç¨±ï¼ˆä¾‹å¦‚ï¼šWC4-795Bï¼‰',
        'ç•°å¸¸å–®è™Ÿ': 'ç•°å¸¸å–®è™Ÿï¼ˆä¾‹å¦‚ï¼šWG20260326ï¼‰',
        'ç™¼ç”Ÿå–®ä½': 'ç™¼ç”Ÿå–®ä½ï¼ˆä¾‹å¦‚ï¼šæœ¬å» ï¼‰',
        'å›žå ±äºº':   'å›žå ±äººå§“å',
        'ç•°å¸¸ç‹€æ³': 'ç•°å¸¸ç‹€æ³é—œéµå­—',
      };
      await replyText(replyToken, `ðŸ” æŸ¥è©¢ ${field}\n\nè«‹è¼¸å…¥${fieldLabels[field] || 'é—œéµå­—'}ï¼š\n\nè¼¸å…¥ã€Œ0ã€å›žä¸»é¸å–®`);
    } else {
      await replyText(replyToken, 'è«‹é»žé¸ä¸Šæ–¹æŒ‰éˆ•é¸æ“‡æŸ¥è©¢æ–¹å¼\n\nè¼¸å…¥ã€Œ0ã€å›žä¸»é¸å–®');
    }
    return;
  }

  if (session.step === 'search_keyword') {
    if (!text) { await replyText(replyToken, 'è«‹è¼¸å…¥é—œéµå­—'); return; }
    const field = session.data.field;
    try {
      const filter = field === 'ç•°å¸¸å–®è™Ÿ'
        ? { property: 'ç•°å¸¸å–®è™Ÿ', title: { contains: text } }
        : { property: field, rich_text: { contains: text } };
      const res = await axios.post(
        `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`,
        { filter, sorts: [{ property: 'ç™¼ç”Ÿæ—¥æœŸ', direction: 'descending' }], page_size: 5 },
        { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' } }
      );
      const results = res.data.results.map(p => {
        const props   = p.properties;
        const getText = (k) => props[k]?.rich_text?.[0]?.text?.content || props[k]?.title?.[0]?.text?.content || '';
        const getDate = (k) => props[k]?.date?.start?.slice(0,10) || '';
        const getUrl  = (k) => props[k]?.url || '';
        return {
          num:      getText('ç•°å¸¸å–®è™Ÿ'),
          date:     getDate('ç™¼ç”Ÿæ—¥æœŸ'),
          unit:     getText('ç™¼ç”Ÿå–®ä½'),
          part:     getText('é›¶ä»¶åç¨±'),
          series:   getText('ç³»åˆ—åˆ¥'),
          issue:    getText('ç•°å¸¸ç‹€æ³'),
          ratio:    getText('ç•°å¸¸æ¯”ä¾‹'),
          judge:    getText('åˆ¤å®š'),
          status:   getText('ç›®å‰è™•ç†ç‹€æ…‹'),
          reporter: getText('å›žå ±äºº'),
          photo:    getUrl('ç•°å¸¸ç…§ç‰‡'),
        };
      });
      if (results.length === 0) {
        await replyText(replyToken, `ðŸ” æŸ¥ç„¡ã€Œ${text}ã€ç›¸é—œç´€éŒ„\n\nè¼¸å…¥ã€Œ0ã€å›žä¸»é¸å–®`);
      } else {
        const lines = [`ðŸ” æ‰¾åˆ° ${res.data.results.length} ç­†ã€Œ${text}ã€ç´€éŒ„ï¼š\n`];
        results.forEach((r, i) => {
          const judgeEmoji = r.judge.includes('é©—é€€') ? 'âŒ' : r.judge.includes('ç‰¹æŽ¡') ? 'âš ï¸' : r.judge.includes('åŠ å·¥') ? 'ðŸ”§' : 'ðŸ”˜';
          lines.push(
            `${i+1}. ${r.num} ${r.date}\n` +
            `   ðŸ­ ${r.unit}ã€€ðŸ‘¤ ${r.reporter}\n` +
            `   ðŸ”© ${r.part}ã€€ðŸ“‚ ${r.series}\n` +
            `   âš ï¸ ${r.issue}\n` +
            `   ðŸ“Š ${r.ratio}ã€€${judgeEmoji} ${r.judge}` +
            (r.photo ? `\n   ðŸ“· ${r.photo}` : '')
          );
        });
        if (res.data.has_more) lines.push(`\n...åƒ…é¡¯ç¤ºå‰ 5 ç­†`);
        lines.push('\nè¼¸å…¥ã€Œ0ã€å›žä¸»é¸å–®');
        await replyText(replyToken, lines.join('\n'));
      }
    } catch (err) {
      console.error(err.response?.data || err.message);
      await replyText(replyToken, 'âŒ æŸ¥è©¢å¤±æ•—ï¼Œè«‹é€šçŸ¥ç®¡ç†å“¡');
    }
    delete sessions[userId];
    return;
  }

  await replyFlex(replyToken);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  æµæ°´è™Ÿ
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const dailyCounters = {};
function genWGNumber() {
  const now = new Date();
  const y = now.getFullYear(), m = String(now.getMonth()+1).padStart(2,'0'), d = String(now.getDate()).padStart(2,'0');
  const key = `${y}${m}${d}`;
  dailyCounters[key] = (dailyCounters[key] || 0) + 1;
  return `WG${y}${m}${d}${String(dailyCounters[key]).padStart(2,'0')}`;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  Cloudinary ä¸Šå‚³åœ–ç‰‡
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  Cloudinary ä¸Šå‚³ Excel (raw)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  ç”¢ç”Ÿ Excel ä¸¦å‚³é€ LINE
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
async function generateAndSendExcel(data, wgNumber, reporterName, photoUrl, photoUrl2) {
  try {
    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const ws = workbook.worksheets[0];

    // å¡«å…¥å„²å­˜æ ¼
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

    // åµŒå…¥ç…§ç‰‡
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

    // ä¸Šå‚³åˆ° Cloudinary
    const downloadUrl = await uploadExcelToCloudinary(buffer, filename);

    // å‚³é€ LINE è¨Šæ¯çµ¦é€šçŸ¥å°è±¡
    if (downloadUrl) {
      for (const uid of NOTIFY_USERS) {
        await pushText(uid, `ðŸ“‹ å“è³ªç•°å¸¸é€šçŸ¥å–®å·²ç”¢ç”Ÿï¼\n\nç•°å¸¸å–®è™Ÿï¼š${wgNumber}\n\né»žæ“Šä¸‹è¼‰ Excelï¼š\n${downloadUrl}`)
          .catch(e => console.error('push excel link failed:', e.message));
      }
    }

    return { buffer, downloadUrl };
  } catch (e) {
    console.error('generateAndSendExcel failed:', e.message);
    return { buffer: null, downloadUrl: null };
  }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  API Routes
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// LIFF è¡¨å–®æäº¤
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    let reporterName = '(æœªçŸ¥)';
    console.log('anomaly submitted by userId:', d.userId || '(no userId)');
    if (d.userId) reporterName = await getDisplayName(d.userId);

    const wgNumber = genWGNumber();

    // ä¸Šå‚³ç…§ç‰‡
    let photoUrl = null, photoUrl2 = null;
    if (d.photoData)  photoUrl  = await uploadToCloudinary(d.photoData);
    if (d.photoData2) photoUrl2 = await uploadToCloudinary(d.photoData2);

    // å¯«å…¥ Notion
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    const properties = {
      'ç•°å¸¸å–®è™Ÿ':     { title: [{ text: { content: wgNumber } }] },
      'ç™¼ç”Ÿæ—¥æœŸ':     { date: { start: new Date().toISOString().split('T')[0] } },
      'ç™¼ç”Ÿå–®ä½':     { rich_text: toText(d.unit || '') },
      'è²¬ä»»å–®ä½':     { rich_text: toText(d.resp || '') },
      'å®¢æˆ¶':         { rich_text: toText(d.customer || '') },
      'ç³»åˆ—åˆ¥':       { rich_text: toText(d.series || '') },
      'é›¶ä»¶åç¨±':     { rich_text: toText(d.product || '') },
      'ç•°å¸¸ç‹€æ³':     { rich_text: toText(d.anomaly || '') },
      'è™•ç†æ–¹å¼':     { rich_text: toText(d.judge || '') },
      'åˆ¤å®š':         { rich_text: toText(d.judge || '') },
      'è¨‚å–®æ•¸é‡':     { number: parseInt(d.qty) || null },
      'ç•°å¸¸æ¯”ä¾‹':     { rich_text: toText(d.ratio || '') },
      'ç›®å‰è™•ç†ç‹€æ…‹': { rich_text: toText('æœªé–‹å§‹') },
      'å›žå ±äºº':       { rich_text: toText(reporterName) },
    };
    if (d.replyDate) properties['éœ€æ±‚å›žè¦†æ™‚é–“'] = { date: { start: d.replyDate } };
    if (photoUrl)    properties['ç•°å¸¸ç…§ç‰‡']  = { url: photoUrl };
    if (photoUrl2)   properties['ç•°å¸¸ç…§ç‰‡2'] = { url: photoUrl2 };

    const pageBody = { parent: { database_id: NOTION_DATABASE_ID }, properties };
    if (photoUrl || photoUrl2) {
      pageBody.children = [];
      if (photoUrl)  pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl  } } });
      if (photoUrl2) pageBody.children.push({ object:'block', type:'image', image:{ type:'external', external:{ url: photoUrl2 } } });
    }
    await axios.post('https://api.notion.com/v1/pages', pageBody, {
      headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28', 'Content-Type': 'application/json' }
    });

    // æŽ¨æ’­ç•°å¸¸é€šçŸ¥
    const judgeEmoji = d.judge === 'é©—é€€X' ? 'âŒ' : d.judge === 'ç‰¹æŽ¡â–³' ? 'âš ï¸' : 'ðŸ”§';
    const msg =
      `ã€ç•°å¸¸é€šå ± ${wgNumber}ã€‘\n` +
      `ðŸ‘¤ å›žå ±äººï¼š${reporterName}\n` +
      `ðŸ“¦ å“åï¼š${d.product || '(æœªå¡«)'}ã€€ç³»åˆ—ï¼š${d.series || ''}\n` +
      `ðŸ“ ç™¼ç”Ÿå–®ä½ï¼š${d.unit}\n` +
      `ðŸ­ è²¬ä»»å–®ä½ï¼š${d.resp}\n` +
      `âš ï¸ ç•°å¸¸ï¼š${d.anomaly}\n` +
      `ðŸ”¢ è¨‚å–®æ•¸é‡ï¼š${d.qty}ã€€æ¯”ä¾‹ï¼š${d.ratio}\n` +
      `${judgeEmoji} åˆ¤å®šï¼š${d.judge}\n` +
      `ðŸ“… æ—¥æœŸï¼š${d.date}` +
      (d.replyDate ? `\nðŸ“† å›žè¦†æœŸé™ï¼š${d.replyDate}` : '') +
      (photoUrl  ? `\nðŸ“· ç…§ç‰‡1ï¼š${photoUrl}`  : '') +
      (photoUrl2 ? `\nðŸ“· ç…§ç‰‡2ï¼š${photoUrl2}` : '');

    for (const uid of NOTIFY_USERS) {
      await pushText(uid, msg).catch(e => console.error('push failed:', e.message));
    }

    // éžåŒæ­¥ç”¢ç”Ÿ Excel ä¸¦å‚³é€ä¸‹è¼‰é€£çµ
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

app.get('/', (req, res) => res.json({ status: 'LINE Bot running âœ…', routes: ['/api/anomaly', '/webhook'] }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
