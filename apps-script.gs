// ══════════════════════════════════════════════════════════════
//  精油香氛覺察表 — Google Apps Script 後端
// ══════════════════════════════════════════════════════════════

const SHEET_NAME = '覺察表資料';

const OIL_NAMES = [
  '依蘭','真正薰衣草','芳香萬壽菊',
  '甜橙','佛手柑','萊姆','桔子',
  '胡椒薄荷','迷迭香','羅勒',
  '雪松','廣藿香','檸檬香茅'
];

// ── 測試函式（在編輯器選此函式執行，確認授權與寄信正常）──────────
function testEmail() {
  sendResultEmail({
    name:         '測試姓名',
    email:        Session.getActiveUser().getEmail(),
    stress:       7,
    fatigue:      6,
    oils:         [8,0,0,9,7,0,0,5,0,0,6,0,4],
    concerns:     ['睡眠品質差','工作壓力大','情緒起伏'],
    preferredCats:['citrus','flower'],
    selectedOils: ['甜橙','佛手柑','依蘭'],
  });
  Logger.log('測試 Email 已寄出');
}

// ── 接收表單送出（POST）────────────────────────────────────────
function doPost(e) {
  try {
    // hidden iframe form POST 的資料在 e.parameter.data
    const rawJson = (e.parameter && e.parameter.data)
                 || (e.postData  && e.postData.contents)
                 || '{}';

    if (rawJson === '{}') {
      Logger.log('doPost 收到空資料，e.parameter=' + JSON.stringify(e.parameter));
    }

    const data = JSON.parse(rawJson);
    Logger.log('doPost 收到資料：name=' + data.name + ', email=' + data.email);

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET_NAME);

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
      sheet.getRange(1,1,1,headers.length)
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
      concerns[0]||'', concerns[1]||'', concerns[2]||'',
      (data.preferredCats||[]).join('、'),
      (data.selectedOils ||[]).join('、'),
      ...oils.slice(0,13),
      JSON.stringify(data)
    ]);

    if (data.email) sendResultEmail(data);

    return ContentService
      .createTextOutput(JSON.stringify({ status:'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('doPost 錯誤：' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status:'error', message:err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 後台讀取資料（GET）────────────────────────────────────────
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
    const jsonCol = sheet.getRange(2, lastCol, lastRow-1, 1).getValues();

    const records = jsonCol
      .map(row => { try { return JSON.parse(row[0]); } catch { return null; } })
      .filter(Boolean);

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error:err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 寄送報告 Email ────────────────────────────────────────────
function sendResultEmail(data) {
  if (!data || !data.email) return;
  try {
    const name     = data.name    || '您';
    const email    = data.email;
    const stress   = data.stress  || 0;
    const fatigue  = data.fatigue || 0;
    const oils     = data.oils    || [];
    const concerns = data.concerns|| [];

    const topOils = oils
      .map((s,i) => ({ name: OIL_NAMES[i]||'', s }))
      .filter(o => o.s > 0)
      .sort((a,b) => b.s - a.s)
      .slice(0,3);

    const OIL_CATS = [
      'flower','flower','flower',
      'citrus','citrus','citrus','citrus',
      'herb','herb','herb',
      'wood','wood','fresh'
    ];
    const catScore = { flower:0, citrus:0, herb:0, wood:0, fresh:0 };
    oils.forEach((s,i) => { if (s>0 && OIL_CATS[i]) catScore[OIL_CATS[i]] += s; });
    (data.preferredCats||[]).forEach(k => { if (catScore[k]!==undefined) catScore[k]+=5; });
    const dominant = Object.entries(catScore).sort((a,b)=>b[1]-a[1])[0][0];

    const PERS = {
      flower:{ type:'浪漫滋養型', emoji:'🌹',
        suggest:'建議每日以依蘭或真正薰衣草進行睡前香氛儀式，重建與自己的深層連結。' },
      citrus:{ type:'活力陽光型', emoji:'🍊',
        suggest:'工作前擴香甜橙或佛手柑，注入正向能量，清新柑橘氣息能快速轉換心情。' },
      herb:  { type:'清晰理性型', emoji:'🌿',
        suggest:'午後疲勞時塗抹胡椒薄荷或迷迭香於太陽穴，提振專注力，比咖啡更持久溫和。' },
      wood:  { type:'深根沉穩型', emoji:'🪵',
        suggest:'睡前以雪松進行深呼吸冥想，降低皮質醇，穩定自律神經，改善睡眠品質。' },
      fresh: { type:'自由清明型', emoji:'🌬️',
        suggest:'密閉空間工作時使用檸檬香茅擴香，提升清新感與專注力，減少悶塞感。' },
    };
    const pers = PERS[dominant] || PERS.flower;
    const coupon = 'AROMA-' + (name.charAt(0)||'X').toUpperCase() + (stress+fatigue);

    const sc = stress  >=7?'#e85252':stress  >=4?'#e8a020':'#52c989';
    const fc = fatigue >=7?'#e85252':fatigue >=4?'#e8a020':'#52c989';

    const topHtml = topOils.length
      ? topOils.map((o,i)=>`<span style="margin-right:12px"><b style="color:#1a5c3a">第${i+1}名</b> ${o.name} <b style="color:#52c989">${o.s}分</b></span>`).join('')
      : '（尚未評分）';
    const conHtml = concerns.length
      ? concerns.map((c,i)=>`<span style="margin-right:8px">${['🥇','🥈','🥉'][i]||'·'} ${c}</span>`).join('')
      : '（未填寫）';

    const html = `<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">
<style>
body{margin:0;padding:0;background:#f0f3f0;font-family:Arial,sans-serif;color:#1c1c1e}
.w{max-width:560px;margin:28px auto;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.09)}
.r2{display:flex;border-bottom:1px solid #eee}
.cell{flex:1;text-align:center;padding:18px 12px;border-right:1px solid #eee}
.cell:last-child{border-right:none}
.s{padding:0 28px 20px}
.st{font-size:11px;font-weight:700;color:#1a5c3a;letter-spacing:.1em;text-transform:uppercase;margin:20px 0 8px}
.bx{background:#f4f8f5;border-radius:10px;padding:12px 16px;font-size:14px;line-height:1.8}
</style></head><body>
<div class="w">
  <div style="background:#0a2218;padding:28px 32px;text-align:center">
    <div style="font-size:2rem;margin-bottom:6px">${pers.emoji}</div>
    <div style="color:#52c989;font-size:11px;font-weight:700;letter-spacing:.15em;margin-bottom:6px">AROMA AWARENESS REPORT</div>
    <div style="color:#fff;font-size:18px;font-weight:900;margin-bottom:4px">${name} 的香氛覺察報告</div>
    <div style="color:rgba(255,255,255,.55);font-size:13px">香氛人格：<strong style="color:#e8c56a">${pers.type}</strong></div>
  </div>
  <div class="r2">
    <div class="cell"><div style="font-size:11px;color:#6b7a72;margin-bottom:4px">壓力指數</div><div style="font-size:2rem;font-weight:900;color:${sc}">${stress}</div><div style="font-size:11px;color:#aaa">/ 10</div></div>
    <div class="cell"><div style="font-size:11px;color:#6b7a72;margin-bottom:4px">疲憊指數</div><div style="font-size:2rem;font-weight:900;color:${fc}">${fatigue}</div><div style="font-size:11px;color:#aaa">/ 10</div></div>
  </div>
  <div class="s">
    <div class="st">最愛精油 Top 3</div><div class="bx">${topHtml}</div>
    <div class="st">主要困擾</div><div class="bx">${conHtml}</div>
    <div class="st">香氛人格解析</div><div class="bx" style="background:#faf5e8;color:#5a4a20">${pers.suggest}</div>
    <div style="background:#0a2218;border-radius:14px;padding:18px;text-align:center;margin-top:20px">
      <div style="color:rgba(255,255,255,.5);font-size:11px;margin-bottom:6px">🎁 您的專屬折扣碼</div>
      <div style="color:#e8c56a;font-size:20px;font-weight:900;letter-spacing:.12em">${coupon}</div>
      <div style="color:rgba(255,255,255,.4);font-size:11px;margin-top:6px">報名課程折 NT$200 · 購買精油折 NT$100</div>
    </div>
  </div>
  <div style="background:#f8faf8;padding:14px 28px;text-align:center;border-top:1px solid #eee">
    <div style="font-size:11px;color:#aaa;line-height:1.9">此報告由香氛覺察體驗系統自動產生<br>感謝您的參與 ✦</div>
  </div>
</div></body></html>`;

    MailApp.sendEmail({ to:email, subject:`【香氛覺察報告】${name}｜${pers.type}`, htmlBody:html });
    Logger.log('Email 寄出成功：' + email);

  } catch(err) {
    Logger.log('sendResultEmail 錯誤：' + err.toString());
  }
}
