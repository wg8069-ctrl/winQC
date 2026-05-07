// v13 - 強化 Excel 標題自動追蹤，解決對位壞掉問題
const express = require('express');
const crypto  = require('crypto');
const axios   = require('axios');
const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');
const FormData = require('form-data');

const app = express();
app.use(express.json({ limit: '10mb' }));

// ── 環境變數 (請確保你的環境變數已設定) ──
const CLOUDINARY_CLOUD  = 'dlpxz4qlh';
const CLOUDINARY_KEY    = '953226455671951';
const CLOUDINARY_SECRET = process.env.CLOUDINARY_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;

// ── Excel 核心填寫邏輯 ──
async function generateAndSendExcel(data, wgNumber, reporterName) {
  try {
    const templatePath = path.join(__dirname, 'template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    const ws = workbook.worksheets[0];

    // 建立一個對照表：將後端欄位對應到你的 Excel 標題文字
    const mapping = {
      '異常單號': wgNumber,
      '發生日期': data.date,
      '需求回覆時間': data.replyDate,
      '發生單位': data.unit,
      '責任單位': data.resp,
      '零件名稱': data.product,
      '系列別': data.series, // 或者是 '槍型號'
      '單號': data.orderNo,
      '異常狀況': data.anomaly,
      '訂單數量': data.qty,
      '異常比例': data.ratio,
      '判定': data.judge,
      '回報人': reporterName,
      '人工成本(人)': data.laborPeople,
      '人工成本(時)': data.laborHours,
      '行政成本(人)': data.adminPeople,
      '行政成本(時)': data.adminHours,
      '所耗人力成本': data.laborCost,
      '異常標註內容': data.remark
    };

    // 💡 自動掃描 Template 內容並替換 {{標籤}}
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === 'string') {
          let text = cell.value;
          Object.keys(mapping).forEach(key => {
            const placeholder = `{{${key}}}`;
            if (text.includes(placeholder)) {
              text = text.replace(placeholder, mapping[key] || '');
            }
          });
          cell.value = text;
        }
      });
    });

    // 處理圖片嵌入 (根據你的 template 位置)
    // ... (圖片處理程式碼保持不變)

    const buffer = await workbook.xlsx.writeBuffer();
    // 上傳到 Cloudinary... (省略重複代碼)
    return { success: true, url: "Cloudinary 網址" };
  } catch (e) {
    console.error('Excel 生成失敗:', e.message);
    return { success: false };
  }
}

// ── 手動觸發 API ──
app.post('/api/generate-excel-from-sheet', async (req, res) => {
  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ error: '沒有收到資料' });

    // 重新對齊來自 Google Sheets 的欄位
    const mappedData = {
      date: data['發生日期'],
      replyDate: data['需求回覆時間'],
      unit: data['發生單位'],
      resp: data['責任單位'],
      product: data['零件名稱'],
      series: data['系列別'] || data['槍型號'], // 兼顧兩種標題
      orderNo: data['單號'],
      anomaly: data['異常狀況'],
      qty: data['訂單數量'],
      ratio: data['異常比例'],
      judge: data['判定'],
      laborPeople: data['人工成本(人)'],
      laborHours: data['人工成本(時)'],
      adminPeople: data['行政成本(人)'],
      adminHours: data['行政成本(時)'],
      laborCost: data['所耗人力成本'],
      remark: data['異常標註內容']
    };

    const result = await generateAndSendExcel(mappedData, data['異常單號'], data['回報人']);
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});
