/** @OnlyCurrentDoc
 *  物件レコメンド + URL特徴抽出（マルチシート横断・UniLife最適化・賃料スコア抽出）
 *
 *  主なポイント:
 *   - 全シート横断、1枚目は自社扱い（IsInHouse の既存列があれば優先）
 *   - フリーワード検索：家賃/徒歩/築年/階/キーワード/フラグ/自社/住所優先などでスコアリング
 *   - 優先ワード（例: 天久保）でまず上位を埋め、残りを他で補充
 *   - URL本文から Name/RentYen/Address/BuiltYear/Layout/Floor/Station/WalkMin/向き/管理会社/Notes/各フラグを抽出
 *   - 賃料は「JSON-LD > 賃料語近傍のスコア」で決定。費用系ワードは減点。レンジは下限を採用
 *   - 既存値がプレースホルダ（「最寄り駅」「〒だけ」「家賃が1万円未満」）ならスマート上書き
 *
 *  推奨ヘッダー:
 *   Name, URL, RentYen, Layout, Station, WalkMin, BuiltYear, Address,
 *   ManagementCompany, IsInHouse, Floor,
 *   食事付き, 家具付き, ペット可, ネット無料, 駐車場あり, オートロック,
 *   バス・トイレ別, エアコン, 角部屋, 南向き, Notes
 */

// ====== メニュー / サイドバー ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('物件レコメンド')
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('URLから特徴抽出（全シート）', 'menuExtractAll')
    .addItem('URLから特徴抽出（選択行のみ）', 'menuExtractSelection')
    .addToUi();
}
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar').setTitle('物件レコメンド');
  SpreadsheetApp.getUi().showSidebar(html);
}
function menuExtractAll() {
  const res = extractFromUrlsAllSheets(false);
  SpreadsheetApp.getUi().alert(res);
}
function menuExtractSelection() {
  const res = extractFromUrlsSelection(false);
  SpreadsheetApp.getUi().alert(res);
}

// ====== 優先設定・語彙 ======
const PREFERRED_COMPANIES = ['自社', '自社管理'];
const PREFERRED_AREAS = ['天久保1','天久保2','天久保3','天久保4','春日1','春日2','春日3','春日4','桜1','桜2'];
const AREA_KEYWORDS = ['天久保', '春日', '桜', '吾妻'];
const STOPWORDS = ['希望','以内','以上','以下','まで','徒歩','分','家賃','築','年','駅','strict','オンリー','のみ','だけ'];

const FLAG_COLUMNS = [
  { col: '食事付き',   synonyms: ['食事付','朝夕食付','寮食','食事あり'] },
  { col: '家具付き',   synonyms: ['家具家電付き','家具家電','家電付き','家電付'] },
  { col: 'ペット可',   synonyms: ['ペット相談','ペットok','ペットＯＫ'] },
  { col: 'ネット無料', synonyms: ['wi-fi無料','インターネット無料','ネット込み','ネット使い放題'] },
  { col: '駐車場あり', synonyms: ['駐車場有','pあり','パーキング有'] },
  { col: 'オートロック', synonyms: ['自動ロック'] },
  { col: 'バス・トイレ別', synonyms: ['バストイレ別','bt別','セパレート'] },
  { col: 'エアコン',   synonyms: ['ac','冷暖房','エアコン有'] },
  { col: '角部屋',     synonyms: ['角','角住戸'] },
  { col: '南向き',     synonyms: ['南向','南向き'] },
];

const SCRAPE_PATTERNS = {
  '食事付き':     [ /食事付/, /食事あり/, /朝夕食/, /寮食/ ],
  '家具付き':     [ /家具家電(付|付き)/, /家具付/, /家電付/ ],
  'ペット可':     [ /ペット(可|相談|Ｏ?K|OK)/i ],
  'ネット無料':   [ /(ネット|インターネット|Wi-?Fi).{0,6}無料/i, /ネット使い放題/ ],
  '駐車場あり':   [ /駐車場(有|あり)/ ],
  'オートロック': [ /オートロック|自動ロック/ ],
  'バス・トイレ別':[ /バス.?トイレ.?別|セパレート/ ],
  'エアコン':     [ /エアコン|AC|冷暖房/ ],
  '角部屋':       [ /角部屋|角住戸/ ],
  '南向き':       [ /南向(き)?/ ],
};

// ====== ユーティリティ ======
function toHalfWidth(s){ return String(s||'').replace(/[０-９Ａ-Ｚａ-ｚ]/g,ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)); }
function boolFromCell(v){ const s=String(v).trim().toLowerCase(); return ['true','1','yes','y','t','on','はい','有','あり'].includes(s); }

// ====== 要件パース ======
function parseRequirements(queryRaw){
  const q0 = toHalfWidth(queryRaw||'');
  const q = q0.toLowerCase();

  const requiredTerms = Array.from(q.matchAll(/\+(\S+)/g)).map(m=>m[1]);
  const excludedTerms = Array.from(q.matchAll(/-(\S+)/g)).map(m=>m[1]);

  const rentMax = (()=>{ const m1=q.match(/家賃[^0-9]{0,3}([0-9]+(?:\.[0-9])?)\s*(万|万円)?/); const m2=q.match(/([0-9]{4,7})\s*(円|yen)/); if(m1)return Math.round(parseFloat(m1[1])*10000); if(m2)return parseInt(m2[1],10); return null; })();
  const walkMax = (()=>{ const m=q.match(/(徒歩|歩)[^0-9]{0,3}([0-9]{1,2})\s*分/); return m?parseInt(m[2],10):null; })();
  const yearsWithin = (()=>{ const m=q.match(/(築|ちく)[^0-9]{0,3}([0-9]{1,2})\s*年/); return m?parseInt(m[2],10):null; })();
  const floorMin = (()=>{ const m=q.match(/([0-9]{1,2})\s*階以上/); return m?parseInt(m[1],10):null; })();

  const tokens = q.replace(/[,\u3001]/g,' ').split(/\s+/).filter(Boolean);
  const priorityTerms = tokens.filter(t=>t && !/^[0-9]+$/.test(t) && !STOPWORDS.some(sw=>t.includes(sw)));

  const areaPicks = [];
  AREA_KEYWORDS.forEach(base=>{
    const re=new RegExp(`${base}\\s*([0-9]{1,2})?`,'gi');
    const mAll=q0.match(re);
    if(mAll) mAll.forEach(s=>areaPicks.push(s.replace(/\s+/g,'')));
  });
  const strictArea = /\b(限定|のみ|だけ|strict|オンリー)\b/.test(q) ||
    requiredTerms.some(t=>AREA_KEYWORDS.some(a=>t.includes(a.toLowerCase())));

  return { rentMax, walkMax, yearsWithin, floorMin, tokens, requiredTerms, excludedTerms, areaPicks, strictArea, priorityTerms };
}

// ====== 全シート横断読込（1枚目=自社扱い） ======
function getAllRows_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const rows = [];

  sheets.forEach((sh, idx)=>{
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const header = values[0].map(h=>String(h).trim());
    const idxMap = Object.fromEntries(header.map((h,i)=>[h,i]));
    const gv = (row, key)=> row[idxMap[key]];
    if (idxMap['Name']==null || idxMap['URL']==null) return;

    const isInHouseBySheet = (idx === 0); // 1枚目=自社

    for (let r=1; r<values.length; r++){
      const row = values[r];
      const rec = {
        Name: String(gv(row,'Name')||''),
        URL: String(gv(row,'URL')||''),
        RentYen: Number(gv(row,'RentYen')||0),
        Layout: String(gv(row,'Layout')||''),
        Station: String(gv(row,'Station')||''),
        WalkMin: Number(gv(row,'WalkMin')||999),
        BuiltYear: Number(gv(row,'BuiltYear')||0),
        Address: String(gv(row,'Address')||''),
        ManagementCompany: String(gv(row,'ManagementCompany')||''),
        Floor: Number(gv(row,'Floor')||0),
        Flags: {
          '食事付き': boolFromCell(gv(row,'食事付き')),
          '家具付き': boolFromCell(gv(row,'家具付き')),
          'ペット可': boolFromCell(gv(row,'ペット可')),
          'ネット無料': boolFromCell(gv(row,'ネット無料')),
          '駐車場あり': boolFromCell(gv(row,'駐車場あり')),
          'オートロック': boolFromCell(gv(row,'オートロック')),
          'バス・トイレ別': boolFromCell(gv(row,'バス・トイレ別')),
          'エアコン': boolFromCell(gv(row,'エアコン')),
          '角部屋': boolFromCell(gv(row,'角部屋')),
          '南向き': boolFromCell(gv(row,'南向き')),
        },
        Notes: String(gv(row,'Notes')||'')
      };

      const hasIsInHouseCol = idxMap['IsInHouse'] != null;
      const isInHouseCell = hasIsInHouseCol ? boolFromCell(gv(row,'IsInHouse')) : null;
      rec.IsInHouse = (isInHouseCell !== null && isInHouseCell !== undefined) ? isInHouseCell : isInHouseBySheet;
      if (rec.IsInHouse && !rec.ManagementCompany) rec.ManagementCompany = '自社';

      rows.push(rec);
    }
  });

  return rows;
}

// ====== レコメンド本体 ======
function recommendProperties(query, options){
  const topN = (options && options.topN) || 10;
  const req = parseRequirements(query||'');
  const rows = getAllRows_();
  const thisYear = new Date().getFullYear();

  const scored = rows.map(r=>{
    let score = 0;
    if (!r.Name || !r.URL) score -= 1000;

    if (req.rentMax != null) {
      if (r.RentYen <= req.rentMax) score += 50;
      else score -= Math.min(60, Math.ceil((r.RentYen - req.rentMax)/5000));
    }
    if (req.walkMax != null) {
      if (r.WalkMin <= req.walkMax) score += 15;
      else score -= Math.min(15, r.WalkMin - req.walkMax);
    }
    if (req.yearsWithin != null && r.BuiltYear) {
      const age = thisYear - r.BuiltYear;
      if (age <= req.yearsWithin) score += 12;
      else score -= Math.min(12, age - req.yearsWithin);
    }
    if (req.floorMin != null && r.Floor) {
      if (r.Floor >= req.floorMin) score += 15;
      else score -= 10;
    }

    if (r.Address) {
      PREFERRED_AREAS.forEach(a=>{ if (r.Address.includes(a)) score += 8; });
      if (req.areaPicks && req.areaPicks.length) {
        req.areaPicks.forEach(ap=>{ if (ap && r.Address.includes(ap)) score += 30; });
      }
    }

    if (r.IsInHouse) score += 60;
    if (r.ManagementCompany) {
      const low=r.ManagementCompany.toLowerCase();
      PREFERRED_COMPANIES.forEach(pc=>{ if (pc && low.includes(pc.toLowerCase())) score += 35; });
    }

    const hay = (r.Name+' '+r.Layout+' '+r.Station+' '+r.Address+' '+r.ManagementCompany+' '+r.Notes).toLowerCase();
    req.tokens.forEach(tk=>{ if (tk && hay.includes(tk)) score += 8; });

    req.tokens.forEach(tk=>{
      FLAG_COLUMNS.forEach(fc=>{
        const hitByName = tk.includes(fc.col.toLowerCase());
        const hitBySyn  = fc.synonyms.some(s=>tk.includes(String(s).toLowerCase()));
        if ((hitByName||hitBySyn) && r.Flags[fc.col]) score += 10;
      });
    });

    return { ...r, score };
  });

  const scored2 = scored.filter(r=>{
    const hay=(r.Name+' '+r.Address+' '+r.Notes).toLowerCase();
    return !req.excludedTerms.some(t=>t && hay.includes(String(t).toLowerCase()));
  });

  const filtered = scored2.filter(r=>{
    if (!r.Name || !r.URL) return false;
    if (req.strictArea && req.areaPicks && req.areaPicks.length) {
      const okArea = req.areaPicks.some(ap=>ap && r.Address.includes(ap));
      if (!okArea) return false;
    }
    if (req.requiredTerms && req.requiredTerms.length) {
      const hay=(r.Name+' '+r.Layout+' '+r.Station+' '+r.Address+' '+r.ManagementCompany+' '+r.Notes).toLowerCase();
      const okAll=req.requiredTerms.every(t=>t && hay.includes(String(t).toLowerCase()));
      if (!okAll) return false;
    }
    return true;
  });

  const sorted = filtered.slice().sort((a,b)=>{
    if (b.score !== a.score) return b.score - a.score;
    return a.RentYen - b.RentYen;
  });

  const matchHay = r => (r.Name+' '+r.Layout+' '+r.Station+' '+r.Address+' '+r.ManagementCompany+' '+r.Notes).toLowerCase();
  const preferred=[], others=[];
  sorted.forEach(r=>{
    if (req.priorityTerms && req.priorityTerms.length) {
      const hay=matchHay(r);
      const hit=req.priorityTerms.some(t=>hay.includes(String(t).toLowerCase()));
      if (hit) preferred.push(r); else others.push(r);
    } else {
      others.push(r);
    }
  });
  const picked = preferred.concat(others).slice(0, topN);

  const emailLines=[];
  picked.forEach(p=>{ emailLines.push(p.Name); emailLines.push(p.URL); });

  return {
    items: picked.map(p=>({
      Name:p.Name, URL:p.URL, RentYen:p.RentYen, Layout:p.Layout,
      Station:p.Station, WalkMin:p.WalkMin, BuiltYear:p.BuiltYear,
      Address:p.Address, ManagementCompany:p.ManagementCompany,
      IsInHouse:p.IsInHouse, Floor:p.Floor, Flags:p.Flags, Notes:p.Notes,
      score: Math.round(p.score*10)/10
    })),
    emailText: emailLines.join('\n')
  };
}

// ====== URL → タイトル＋本文テキスト取得（htmlRaw付き） ======
function fetchTextFromUrl_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  const code = res.getResponseCode();
  if (code < 200 || code >= 400) throw new Error('HTTP ' + code);

  const htmlRaw = res.getContentText(); // 生HTML
  const titleMatch = htmlRaw.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  const title = titleMatch ? titleMatch[1].replace(/\s+/g,' ').trim() : '';

  let html = htmlRaw.replace(/<script[\s\S]*?<\/script>/gi, '')
                    .replace(/<style[\s\S]*?<\/style>/gi, '')
                    .replace(/<[^>]+>/g, ' ');
  const text = html.replace(/\s+/g, ' ').trim();
  return { title, text, htmlRaw };
}

// ====== 賃料抽出ユーティリティ（JSON-LD優先 + スコア方式） ======
function toYenLowerFromRange_(s) {
  const t = toHalfWidth(s);
  let m = t.match(/([0-9]+(?:\.[0-9]+)?)\s*万\s*[~〜\-]\s*([0-9]+(?:\.[0-9]+)?)\s*万/);
  if (m) return Math.round(parseFloat(m[1]) * 10000);
  m = t.match(/([0-9]+(?:\.[0-9]+)?)\s*万\s*円?/);
  if (m) return Math.round(parseFloat(m[1]) * 10000);
  m = t.match(/([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,7})\s*円/);
  if (m) return parseInt(m[1].replace(/,/g,''), 10);
  return null;
}
function tryJsonLdRent_(htmlRaw) {
  const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = re.exec(htmlRaw)) !== null) {
    try {
      const data = JSON.parse(m[1]);
      const arr = Array.isArray(data) ? data : [data];
      for (const obj of arr) {
        const offers = obj.offers || obj.Offer || obj.offer;
        const price = offers?.price || offers?.lowPrice || offers?.priceSpecification?.price;
        const currency = offers?.priceCurrency || offers?.priceSpecification?.priceCurrency || '';
        if (price && (!currency || /JPY|¥|円/i.test(currency))) {
          const n = Number(String(price).replace(/,/g,''));
          if (isFinite(n) && n >= 10000 && n <= 300000) return n;
        }
      }
    } catch (_) {}
  }
  return null;
}
function pickRentFromText_(text) {
  const feeWords = /(管理費|共益費|更新料|清掃|除菌|鍵交換|保証|手数料|敷金|礼金|駐車|消火|補償)/;
  const rentWords = /(賃料|家賃|月額)/;

  const windows = [];
  let m;
  const reWin = new RegExp(`(.{0,120}${rentWords.source}.{0,120})`, 'g');
  while ((m = reWin.exec(text)) !== null) windows.push(m[1]);
  const block = text.match(/賃料情報[^。]{0,300}/);
  if (block) windows.push(block[0]);
  if (windows.length === 0) windows.push(text.slice(0, 600));

  const cand = [];
  windows.forEach((w, wi) => {
    const yenLower = toYenLowerFromRange_(w);
    if (yenLower != null) {
      let score = 0;
      if (rentWords.test(w)) score += 20;
      if (feeWords.test(w)) score -= 40;
      if (/万/.test(w)) score += 10;
      score += Math.max(0, 5 - wi);
      if (yenLower < 10000 || yenLower > 300000) score -= 100;
      cand.push({ value: yenLower, score });
    }
  });
  cand.sort((a,b)=> b.score - a.score);
  return cand.length ? cand[0].value : null;
}

// ====== タイトル/本文 → 特徴抽出（賃料スコア方式・UniLife最適化） ======
function extractFeaturesFromText_(payload) {
  const title = toHalfWidth(payload.title || '');
  const text  = toHalfWidth(payload.text  || '').replace(/\s+/g, ' ');
  const htmlRaw = payload.htmlRaw || '';

  // フラグ
  const flags = {};
  Object.keys(SCRAPE_PATTERNS).forEach(key => {
    flags[key] = SCRAPE_PATTERNS[key].some(re => re.test(text));
  });

  // Name
  let name = '';
  if (title) {
    name = title.split(/[\|\-｜–—]/)[0].trim();
    if (name.length < 2 || name.length > 60) name = '';
  }
  if (!name) {
    const mName = text.match(/([^\s　]{2,30}田川|パレーシャル[^\s　]{1,15})/);
    if (mName) name = mName[1];
  }

  // RentYen（JSON-LD > スコア方式）
  let rentYen = tryJsonLdRent_(htmlRaw);
  if (rentYen == null) rentYen = pickRentFromText_(text);

  // Address（〒… or 所在地/住所:）
  let address = '';
  let m = text.match(/〒\s*\d{3}-\d{4}\s*([^\s　]*?県[^\s　]{0,80}?)(?=\s(?:地図|MAP|周辺|アクセス)|\s|$)/i);
  if (m) address = m[1].replace(/　/g,' ').trim();
  if (!address) {
    m = text.match(/(所在地|住所)\s*[:：]?\s*([^\s　].{6,160}?)(?=\s(?:地図|MAP|周辺|アクセス)|\s|$)/i);
    if (m) address = m[2].replace(/　/g,' ').trim();
  }

  // BuiltYear
  let builtYear = null;
  m = text.match(/(?:完成年月|築年|築年数|建築年)[^\d]{0,6}(\d{4})\s*年/);
  if (m) builtYear = parseInt(m[1], 10);
  if (builtYear == null) { m = text.match(/\((\d{4})年\)/); if (m) builtYear = parseInt(m[1], 10); }

  // Layout
  let layout = null;
  m = text.match(/\b([1-4]R|[1-4]K|[1-4]DK|[1-4]LDK|ワンルーム|1ルーム)\b/i);
  if (m) layout = m[1].toUpperCase().replace('ワンルーム','1R').replace('1ルーム','1R');

  // Floor（◯階／◯階建は除外）
  let floor = null;
  m = text.match(/(^|[^建])(\d{1,2})\s*階(?!建)/);
  if (m) floor = parseInt(m[2], 10);

  // Station / WalkMin（駅 or バス停）
  let station = '';
  let walkMin = null;
  m = text.match(/最寄り駅\s*[:：]?\s*([一-龥ぁ-んァ-ヶA-Za-z0-9・\-]+駅)[^0-9]{0,6}(?:徒歩|歩)\s*([0-9]{1,2})\s*分/);
  if (m) { station = m[1]; walkMin = parseInt(m[2], 10); }
  if (!station) {
    m = text.match(/「?([一-龥ぁ-んァ-ヶA-Za-z0-9・\-]{1,20})」?\s*停\s*徒歩\s*([0-9]{1,2})\s*分/);
    if (m) { station = m[1] + '(バス)'; walkMin = parseInt(m[2], 10); }
  }
  if (!station) {
    m = text.match(/(?!最寄り)([一-龥ぁ-んァ-ヶA-Za-z0-9・\-]{1,20}駅)[^0-9]{0,6}(?:徒歩|歩)\s*([0-9]{1,2})\s*分/);
    if (m) { station = m[1]; walkMin = parseInt(m[2], 10); }
  }
  if (!station) {
    m = text.match(/(?!最寄り)([一-龥ぁ-んァ-ヶA-Za-z0-9・\-]{1,20}駅)/);
    if (m) station = m[1];
  }
  if (walkMin == null) {
    m = text.match(/(?:徒歩|歩)\s*([0-9]{1,2})\s*分/);
    if (m) walkMin = parseInt(m[1], 10);
  }

  // 向き補強
  if (/向き\s*南/.test(text)) flags['南向き'] = true;

  // 管理会社
  let managementCompany = '';
  m = text.match(/(株式会社[^\s　]{2,30}|[^\s　]{2,30}不動産[^\s　]{0,6}|[^\s　]{2,30}管理[^\s　]{0,6})/);
  if (m) managementCompany = m[1];

  // Notes
  const notesHits = [];
  if (/日当たり|日当り|陽当り/.test(text)) notesHits.push('日当たり良好');
  if (/女性専用/.test(text)) notesHits.push('女性専用');
  if (/学生向け|学生専用/.test(text)) notesHits.push('学生向け');
  if (/ファミリー向け|家族/.test(text)) notesHits.push('ファミリー向け');
  if (/ロフト/.test(text)) notesHits.push('ロフト付き');
  if (/自転車\s*[0-9]{1,2}\s*分/.test(text)) notesHits.push('自転車距離あり');

  return {
    flags, floor, layout, notesHits,
    name, rentYen, address, builtYear, station, walkMin,
    managementCompany
  };
}

// ====== 抽出結果を1行に反映（スマート上書き） ======
function applyExtractedToRow_(sheet, rowIdx, headerIdx, ext, overwrite=false) {
  function shouldOverwrite(colName, currentVal, newVal) {
    if (overwrite) return true;
    if (currentVal === '' || currentVal === null) return true;
    const cur = String(currentVal).trim();
    if (colName === 'Station' && /^最寄り駅$/.test(cur)) return true; // プレースホルダ
    if (colName === 'Address' && /^〒\s*\d{3}-\d{4}\s*$/.test(cur)) return true; // 郵便番号のみ
    if (colName === 'RentYen') {
      const n = Number(currentVal);
      if (!isFinite(n) || n < 10000) return true; // 700/18 など異常値
    }
    return false;
  }
  function setCellSmart(colName, value) {
    if (!(colName in headerIdx)) return;
    if (value === undefined || value === null || value === '') return;
    const col = headerIdx[colName] + 1, r = rowIdx + 1;
    const current = sheet.getRange(r, col).getValue();
    if (shouldOverwrite(colName, current, value)) sheet.getRange(r, col).setValue(value);
  }

  // フラグ
  Object.entries(ext.flags || {}).forEach(([key, val])=>{
    if (!(key in headerIdx)) return;
    const col = headerIdx[key] + 1, r = rowIdx + 1;
    const cur = sheet.getRange(r, col).getValue();
    if (overwrite || cur === '' || cur === null) sheet.getRange(r, col).setValue(val ? 'TRUE' : 'FALSE');
  });

  // 基本属性
  setCellSmart('Name', ext.name);
  if (ext.rentYen != null) setCellSmart('RentYen', ext.rentYen);
  setCellSmart('Address', ext.address);
  if (ext.builtYear != null) setCellSmart('BuiltYear', ext.builtYear);
  setCellSmart('Station', ext.station);
  if (ext.walkMin != null) setCellSmart('WalkMin', ext.walkMin);
  if (ext.floor != null) setCellSmart('Floor', ext.floor);
  if (ext.layout) setCellSmart('Layout', ext.layout);
  setCellSmart('ManagementCompany', ext.managementCompany);

  // Notes 追記
  const noteCol = headerIdx['Notes'];
  if (noteCol != null && ext.notesHits && ext.notesHits.length) {
    const rng = sheet.getRange(rowIdx + 1, noteCol + 1);
    const cur = String(rng.getValue() || '');
    const add = Array.from(new Set(cur.split(/[、,\s・]+/).concat(ext.notesHits))).filter(Boolean).join('・');
    rng.setValue(add);
  }
}

// ====== シート1枚ぶん抽出 ======
function extractFromUrlsInSheet_(sheet, { startRow=2, endRow=null, overwrite=false } = {}) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const header = values[0].map(h => String(h).trim());
  const headerIdx = Object.fromEntries(header.map((h,i)=>[h,i]));
  const urlCol = headerIdx['URL'];
  if (urlCol == null) throw new Error('URL列が見つかりません: ' + sheet.getName());

  const last = endRow ? Math.min(endRow, values.length) : values.length;
  let updated = 0;

  for (let r = startRow; r <= last; r++) {
    const row = values[r-1];
    const url = row[urlCol];
    if (!url) continue;

    try {
      const payload = fetchTextFromUrl_(String(url));
      const ext = extractFeaturesFromText_(payload);
      applyExtractedToRow_(sheet, r-1, headerIdx, ext, overwrite);
      Utilities.sleep(1200); // 連続アクセス抑制
      updated++;
    } catch (e) {
      const noteIdx = headerIdx['Notes'];
      if (noteIdx != null) {
        const rng = sheet.getRange(r, noteIdx + 1);
        const cur = String(rng.getValue() || '');
        rng.setValue(cur + (cur ? ' / ' : '') + '抽出失敗:' + (e.message || e));
      }
    }
  }
  return updated;
}

// ====== 全シート/選択行で抽出 ======
function extractFromUrlsAllSheets(overwrite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let total = 0;
  sheets.forEach(sh => { total += extractFromUrlsInSheet_(sh, { overwrite: !!overwrite }); });
  return '更新行数: ' + total;
}
function extractFromUrlsSelection(overwrite) {
  const sh = SpreadsheetApp.getActiveSheet();
  const sel = sh.getActiveRange();
  if (!sel) throw new Error('選択範囲がありません');
  const startRow = sel.getRow();
  const endRow = startRow + sel.getNumRows() - 1;
  const updated = extractFromUrlsInSheet_(sh, { startRow, endRow, overwrite: !!overwrite });
  return '更新行数: ' + updated;
}
