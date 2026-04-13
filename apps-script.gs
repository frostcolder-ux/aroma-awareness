// ══════════════════════════════════════════════════════════════
//  精油香氛覺察表 — Google Apps Script 後端
//  使用方式：
//    1. 開啟 Google 試算表 → 擴充功能 → Apps Script
//    2. 貼上此程式碼，儲存
//    3. 部署 → 新增部署 → 類型選「網頁應用程式」
//       執行身分：「我自己」
//       誰可以存取：「所有人」
//    4. 複製部署後的 URL（格式：https://script.google.com/macros/s/...）
//    5. 貼入 awareness-form-v2.html 的 APPS_SCRIPT_URL 變數
// ══════════════════════════════════════════════════════════════

const SHEET_NAME = '覺察表資料';

const OIL_NAMES = [
  '依蘭','真正薰衣草','芳香萬壽菊',
  '甜橙','佛手柑','萊姆','桔子',
  '胡椒薄荷','迷迭香','羅勒',
  '雪松','廣藿香','檸檬香茅'
];

// ── 接收表單送出（POST）────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let sheet  = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        '填答時間','姓名','Email','職業','年齡','居住地區','工作環境',
        '壓力指數','疲憊指數',
        '困擾第1名','困擾第2名','困擾第3名',
        '偏好香調','已選精油',
        ...OIL_NAMES,
        '原始JSON'
      ];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight('bold')
           .setBackground('#0a2218')
           .setFontColor('#ffffff');
    }

    const concerns = data.concerns || (data.concern ? [data.concern] : []);
    const oils     = data.oils     || new Array(13).fill(0);

    sheet.appendRow([
      new Date(data.ts || Date.now()).toLocaleString('zh-TW'),
      data.name        || '',
      data.email       || '',
      data.job         || '',
      data.age         || '',
      data.residence   || '',
      data.environment || '',
      data.stress      || 0,
      data.fatigue     || 0,
      concerns[0] || '', concerns[1] || '', concerns[2] || '',
      (data.preferredCats || []).join('、'),
      (data.selectedOils  || []).join('、'),
      ...oils.slice(0, 13),
      JSON.stringify(data)
    ]);

    // 寄送報告 Email
    if (data.email) {
      sendResultEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 寄送報告 Email ────────────────────────────────────────
function sendResultEmail(data) {
  try {
    const name    = data.name    || '您';
    const email   = data.email;
    const stress  = data.stress  || 0;
    const fatigue = data.fatigue || 0;
    const oils    = data.oils    || [];
    const concerns = data.concerns || [];

    // Top 3 精油
    const topOils = oils
      .map((s, i) => ({ name: OIL_NAMES[i] || '', s }))
      .filter(o => o.s > 0)
      .sort((a, b) => b.s - a.s)
      .slice(0, 3);

    // 香氛人格判斷
    const OIL_CATS = [
      'flower','flower','flower',
      'citrus','citrus','citrus','citrus',
      'herb','herb','herb',
      'wood','wood','fresh'
    ];
    const catScore = { flower:0, citrus:0, herb:0, wood:0, fresh:0 };
    oils.forEach((s, i) => { if (s > 0 && OIL_CATS[i]) catScore[OIL_CATS[i]] += s; });
    (data.preferredCats || []).forEach(k => { if (catScore[k] !== undefined) catScore[k] += 5; });
    const dominant = Object.entries(catScore).sort((a, b) => b[1] - a[1])[0][0];

    const PERS = {
      flower: { type:'浪漫滋養型', emoji:'🌹', color:'#d4307a',
        suggest:'建議每日以依蘭或真正薰衣草進行睡前儀式，入浴或睡前五分鐘的香氛冥想，重建與自己的深層連結。' },
      citrus: { type:'活力陽光型', emoji:'🍊', color:'#e8850a',
        suggest:'工作前擴香甜橙或佛手柑，注入正向能量。情緒低落時直接深吸精油瓶，清新的柑橘氣息能快速轉換心情。' },
      herb:   { type:'清晰理性型', emoji:'🌿', color:'#2d8a50',
        suggest:'午後疲勞時塗抹胡椒薄荷或迷迭香於太陽穴，提振專注力，效果比咖啡更持久且溫和。' },
      wood:   { type:'深根沉穩型', emoji:'🪵', color:'#8a5a2a',
        suggest:'睡前以雪松進行深呼吸冥想練習，有助降低皮質醇、穩定自律神經，改善睡眠品質。' },
      fresh:  { type:'自由清明型', emoji:'🌬️', color:'#1a7aa8',
        suggest:'在密閉空間工作時使用檸檬香茅擴香，提升空間清新感與專注力，減少悶塞與壓迫感。' },
    };
    const pers = PERS[dominant] || PERS.flower;

    const coupon = 'AROMA-' + (name.charAt(0) || 'X').toUpperCase() + (stress + fatigue);

    const stressColor  = stress  >= 7 ? '#e85252' : stress  >= 4 ? '#e8a020' : '#52c989';
    const fatigueColor = fatigue >= 7 ? '#e85252' : fatigue >= 4 ? '#e8a020' : '#52c989';

    const topOilsHtml = topOils.length
      ? topOils.map((o, i) =>
          `<span style="margin-right:10px"><strong style="color:#1a5c3a">第${i+1}名</strong>　${o.name}　<strong style="color:#52c989">${o.s} 分</strong></span>`)
          .join('')
      : '（尚未評分）';

    const concernsHtml = concerns.length
      ? concerns.map((c, i) =>
          `<span style="margin-right:8px">${['🥇','🥈','🥉'][i] || '·'} ${c}</span>`).join('')
      : '（未填寫）';

    const html = `<!DOCTYPE html>
<html lang="zh-TW">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<style>
  body { margin:0; padding:0; background:#f0f3f0; font-family: Arial, sans-serif; color:#1c1c1e; }
  .wrap { max-width:560px; margin:28px auto; background:#fff; border-radius:18px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,.09); }
  .row2 { display:flex; border-bottom:1px solid #eee; }
  .cell { flex:1; text-align:center; padding:18px 12px; border-right:1px solid #eee; }
  .cell:last-child { border-right:none; }
  .sec { padding:0 28px 20px; }
  .sec-title { font-size:11px; font-weight:700; color:#1a5c3a; letter-spacing:.1em; text-transform:uppercase; margin:20px 0 8px; }
  .box { background:#f4f8f5; border-radius:10px; padding:12px 16px; font-size:14px; line-height:1.8; }
  .box-warm { background:#faf5e8; color:#5a4a20; }
</style>
</head>
<body>
<div class="wrap">

  <div style="background:#0a2218;padding:28px 32px;text-align:center">
    <div style="font-size:2rem;margin-bottom:6px">${pers.emoji}</div>
    <div style="color:#52c989;font-size:11px;font-weight:700;letter-spacing:.15em;margin-bottom:6px">AROMA AWARENESS REPORT</div>
    <div style="color:#fff;font-size:18px;font-weight:900;margin-bottom:4px">${name} 的香氛覺察報告</div>
    <div style="color:rgba(255,255,255,.55);font-size:13px">香氛人格：<strong style="color:#e8c56a">${pers.type}</strong></div>
  </div>

  <div class="row2">
    <div class="cell">
      <div style="font-size:11px;color:#6b7a72;margin-bottom:4px">壓力指數</div>
      <div style="font-size:2rem;font-weight:900;color:${stressColor}">${stress}</div>
      <div style="font-size:11px;color:#aaa">/ 10</div>
    </div>
    <div class="cell">
      <div style="font-size:11px;color:#6b7a72;margin-bottom:4px">疲憊指數</div>
      <div style="font-size:2rem;font-weight:900;color:${fatigueColor}">${fatigue}</div>
      <div style="font-size:11px;color:#aaa">/ 10</div>
    </div>
  </div>

  <div class="sec">
    <div class="sec-title">最愛精油 Top 3</div>
    <div class="box">${topOilsHtml}</div>

    <div class="sec-title">主要困擾</div>
    <div class="box">${concernsHtml}</div>

    <div class="sec-title">香氛人格解析</div>
    <div class="box box-warm">${pers.suggest}</div>

    <div style="background:#0a2218;border-radius:14px;padding:18px;text-align:center;margin-top:20px">
      <div style="color:rgba(255,255,255,.5);font-size:11px;margin-bottom:6px">🎁 您的專屬折扣碼</div>
      <div style="color:#e8c56a;font-size:20px;font-weight:900;letter-spacing:.12em">${coupon}</div>
      <div style="color:rgba(255,255,255,.4);font-size:11px;margin-top:6px">報名課程折 NT$200 · 購買精油折 NT$100</div>
    </div>
  </div>

  <div style="background:#f8faf8;padding:14px 28px;text-align:center;border-top:1px solid #eee">
    <div style="font-size:11px;color:#aaa;line-height:1.9">此報告由香氛覺察體驗系統自動產生<br>感謝您的參與 ✦</div>
  </div>

</div>
</body>
</html>`;

    MailApp.sendEmail({
      to:       email,
      subject:  `【香氛覺察報告】${name}｜${pers.type}`,
      htmlBody: html,
    });

  } catch (err) {
    console.error('Email 寄送失敗：', err.toString());
  }
}

// ── 提供後台讀取資料（GET）────────────────────────────────
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const jsonCol = sheet.getRange(2, lastCol, lastRow - 1, 1).getValues();

    const records = jsonCol
      .map(row => { try { return JSON.parse(row[0]); } catch { return null; } })
      .filter(Boolean);

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
