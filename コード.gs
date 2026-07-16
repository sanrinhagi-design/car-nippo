// ============================================================
// 自動車運転日報 — GAS バックエンド v1
// 重機稼働日報システムと同一アーキテクチャ／別スプレッドシート
// ============================================================

const SPREADSHEET_ID = 'YOUR_NEW_SPREADSHEET_ID_HERE'; // ← 新規スプレッドシートのIDに差し替えてください
const SHEET_REPORT     = '日報';
const SHEET_CAR        = '車両マスタ';
const SHEET_OPERATOR   = '運転者マスタ';
const ALERT_EMAIL      = 'm.nakamoto@sanyochip.com';
const APP_TOKEN = 'sanyochip-car-2026-7q2m'; // index.html の APP_TOKEN と同じ文字列にする

function checkToken(token) {
  return token === APP_TOKEN;
}

// ---------- ルーティング ----------

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  if (!checkToken(e.parameter && e.parameter.token)) {
    return json({error: 'unauthorized'});
  }
  let result;
  try {
    if      (action === 'getCars')        result = getCars();
    else if (action === 'getReports')     result = getReports(e.parameter.month || '', e.parameter.carId || '');
    else if (action === 'getOperators')   result = getOperators();
    else if (action === 'ping')           result = {ok: true};
    else result = {error: 'Unknown action: ' + action};
  } catch(err) {
    result = {error: err.message};
  }
  return json(result);
}

function doPost(e) {
  let data, result;
  try {
    data = JSON.parse(e.postData.contents);
    if (!checkToken(data.token)) {
      return json({error: 'unauthorized'});
    }
    const a = data.action || '';
    if      (a === 'batchAddReports') result = batchAddReports(data.reports || []);
    else if (a === 'addReport')       result = addReport(data.report);
    else if (a === 'editReport')      result = editReport(data.report);
    else if (a === 'deleteReport')    result = deleteReport(data.id);
    else if (a === 'addCar')          result = addCar(data.car);
    else if (a === 'deleteCar')       result = deleteCar(data.carId);
    else if (a === 'editCar')         result = editCar(data.car);
    else if (a === 'addOperator')     result = addOperator(data.name);
    else if (a === 'deleteOperator')  result = deleteOperator(data.name);
    else result = {error: 'Unknown action: ' + a};
  } catch(err) {
    result = {error: err.message};
  }
  return json(result);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- キャッシュ ----------
const CACHE_TTL = 60;

function cacheGet(key, builder) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const val = builder();
  try { cache.put(key, JSON.stringify(val), CACHE_TTL); } catch(e){}
  return val;
}

function cacheClear() {
  const cache = CacheService.getScriptCache();
  const keys = ['cars','operators'];
  cache.removeAll(keys);
  const now = new Date();
  const rm = [];
  for (let i=0;i<3;i++){
    const d = new Date(now.getFullYear(), now.getMonth()-i, 1);
    const ym = Utilities.formatDate(d,'Asia/Tokyo','yyyy-MM');
    rm.push('reports_'+ym);
  }
  cache.removeAll(rm);
}

// ---------- シート取得 ----------

function getOrCreate(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1,1,1,headers.length)
         .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getReportSheet() {
  return getOrCreate(SHEET_REPORT, REPORT_HEADERS);
}

function getOperatorSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_OPERATOR);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OPERATOR);
    sheet.appendRow(OPERATOR_HEADERS);
    sheet.getRange(1,1,1,OPERATOR_HEADERS.length)
         .setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------- 運転者マスタ ----------

const OPERATOR_HEADERS = ['氏名','登録日','状態'];

function getOperators() {
  return cacheGet('operators', getOperatorsRaw);
}
function getOperatorsRaw() {
  const s = getOperatorSheet();
  const rows = s.getDataRange().getValues().slice(1);
  return {
    operators: rows.filter(r => r[0] && r[2] !== '削除').map(r => String(r[0]))
  };
}

function addOperator(name) {
  if (!name || !String(name).trim()) return {error:'氏名を入力してください'};
  name = String(name).trim();
  const s = getOperatorSheet();
  const rows = s.getDataRange().getValues().slice(1);
  if (rows.some(r => String(r[0]) === name && r[2] !== '削除')) return {error: name + ' は既に登録されています'};
  s.appendRow([name, new Date(), '有効']);
  cacheClear();
  return {success:true};
}

function deleteOperator(name) {
  const s = getOperatorSheet();
  const rows = s.getDataRange().getValues();
  for (let i=1;i<rows.length;i++){
    if (String(rows[i][0]) === String(name)){ s.deleteRow(i+1); cacheClear(); return {success:true}; }
  }
  return {error: name + ' が見つかりません'};
}

// ---------- 車両マスタ ----------
// type: 'passenger'（乗用車）| 'forest'（山林部トラック）

const CAR_HEADERS = ['車両番号','車種名','拠点','車両タイプ','登録日','状態'];

function getCars() {
  return cacheGet('cars', getCarsRaw);
}
function getCarsRaw() {
  const s = getOrCreate(SHEET_CAR, CAR_HEADERS);
  const rows = s.getDataRange().getValues().slice(1);
  return {
    cars: rows.map(r => ({
      id: r[0], model: r[1], yard: r[2], type: r[3],
      registeredAt: r[4] ? Utilities.formatDate(new Date(r[4]),'Asia/Tokyo','yyyy-MM-dd') : '',
      status: r[5]
    })).filter(c => c.id)
  };
}

function addCar(c) {
  if (!c || !c.id || !c.model) return {error:'車両番号と車種名は必須です'};
  const s = getOrCreate(SHEET_CAR, CAR_HEADERS);
  const rows = s.getDataRange().getValues().slice(1);
  if (rows.some(r => r[0] === c.id)) return {error: c.id + ' は既に登録されています'};
  s.appendRow([c.id, c.model, c.yard||'', c.type||'passenger', new Date(), '稼働中']);
  cacheClear();
  return {success:true};
}

function deleteCar(id) {
  const s = getOrCreate(SHEET_CAR, CAR_HEADERS);
  const rows = s.getDataRange().getValues();
  for (let i=1;i<rows.length;i++){
    if (rows[i][0]===id){ s.deleteRow(i+1); cacheClear(); return {success:true}; }
  }
  return {error: id + ' が見つかりません'};
}

function editCar(c) {
  const s = getOrCreate(SHEET_CAR, CAR_HEADERS);
  const rows = s.getDataRange().getValues();
  const targetId = c.oldId || c.id;
  for (let i=1;i<rows.length;i++){
    if (rows[i][0]===targetId){
      if (c.id !== targetId && rows.some((r,idx)=>idx>0 && idx!==i && r[0]===c.id)) {
        return {error: c.id + ' は既に使用されています'};
      }
      s.getRange(i+1,1,1,4).setValues([[c.id, c.model, c.yard||'', c.type||'passenger']]);
      cacheClear();
      return {success:true};
    }
  }
  return {error: targetId + ' が見つかりません'};
}

// ---------- 日報 ----------
// 乗用車: date, carId, yard, operator, location(現場名・営業先名), meter, fuelLocation, fuelAmount, notes
// 山林部: 上記に加え route(経路), loadCount(材の台数)

const REPORT_HEADERS = [
  '受信日時','日付','車両番号','拠点','車両タイプ','運転者','現場名・営業先名','経路',
  'メーター(km)','走行距離(km)','給油場所','給油量(L)','材の台数',
  '緯度','経度','精度(m)','備考','デバイスID','端末保存日時','報告ID'
];

function addReport(r) {
  if (!r) return {error:'データなし'};
  const s = getReportSheet();
  if (r.id) {
    const lastCol = s.getLastColumn();
    const headers = s.getRange(1,1,1,lastCol).getValues()[0];
    const idCol = headers.indexOf('報告ID') + 1;
    if (idCol > 0) {
      const ids = s.getRange(2, idCol, Math.max(s.getLastRow()-1,0), 1).getValues().flat();
      if (ids.indexOf(r.id) !== -1) {
        return {success:true, duplicate:true};
      }
    }
  }
  s.appendRow([
    new Date(), r.date||'', r.carId||'', r.yard||'', r.carType||'', r.operator||'', r.location||'', r.route||'',
    r.meter!=='' && r.meter!==undefined ? Number(r.meter):'',
    r.distance!=='' && r.distance!==undefined ? Number(r.distance):'',
    r.fuelLocation||'',
    r.fuelAmount!=='' && r.fuelAmount!==undefined ? Number(r.fuelAmount):'',
    r.loadCount!=='' && r.loadCount!==undefined ? Number(r.loadCount):'',
    r.lat||'', r.lng||'', r.accuracy||'', r.notes||'',
    r.deviceId||'', r.savedAt||'', r.id||''
  ]);
  cacheClear();
  return {success:true};
}

function batchAddReports(reports) {
  if (!reports || reports.length===0) return {success:true, count:0};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    reports.forEach(r => addReport(r));
  } finally {
    lock.releaseLock();
  }
  return {success:true, count:reports.length};
}

function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v||'');
}

function getReports(month, carId) {
  const all = month ? cacheGet('reports_'+month, function(){ return getReportsRaw(month); })
                    : getReportsRaw('');
  let reports = all.reports;
  if (carId) reports = reports.filter(function(r){ return String(r.carId) === String(carId); });
  return {reports: reports};
}
function getReportsRaw(month) {
  const s = getReportSheet();
  const rows = s.getDataRange().getValues().slice(1);
  let reports = rows.map(r => ({
    receivedAt: r[0]?Utilities.formatDate(new Date(r[0]),'Asia/Tokyo','yyyy-MM-dd HH:mm'):'',
    date: normDate(r[1]), carId:r[2], yard:r[3], carType:r[4], operator:r[5], location:r[6], route:r[7],
    meter:r[8], distance:r[9], fuelLocation:r[10], fuelAmount:r[11], loadCount:r[12],
    lat:r[13], lng:r[14], accuracy:r[15], notes:r[16],
    savedAt:r[18]||'', id:r[19]||''
  }));
  if (month) reports = reports.filter(r => r.date.startsWith(month));
  return {reports};
}

function editReport(r) {
  if (!r || !r.id) return {error:'報告IDがありません（旧データは編集できません）'};
  const s = getReportSheet();
  const lastCol = s.getLastColumn();
  const headers = s.getRange(1,1,1,lastCol).getValues()[0];
  const idCol = headers.indexOf('報告ID') + 1;
  if (idCol === 0) return {error:'報告ID列がありません'};
  const rows = s.getDataRange().getValues();
  for (let i=1;i<rows.length;i++){
    if (String(rows[i][idCol-1]) === String(r.id)) {
      const row = i+1;
      if (r.date) s.getRange(row, 2).setValue(r.date);
      if (r.carId) {
        s.getRange(row, 3).setValue(r.carId);
        const cc = getCarsRaw().cars.find(function(x){ return x.id === r.carId; });
        if (cc) {
          s.getRange(row, 4).setValue(cc.yard || '');
          s.getRange(row, 5).setValue(cc.type || '');
        }
      }
      s.getRange(row, 6, 1, 7).setValues([[
        r.operator||'', r.location||'', r.route||'',
        r.meter!=='' && r.meter!==undefined ? Number(r.meter):'',
        r.distance!=='' && r.distance!==undefined ? Number(r.distance):'',
        r.fuelLocation||'',
        r.fuelAmount!=='' && r.fuelAmount!==undefined ? Number(r.fuelAmount):''
      ]]);
      const loadCol = headers.indexOf('材の台数') + 1;
      if (loadCol > 0) s.getRange(row, loadCol).setValue(r.loadCount!=='' && r.loadCount!==undefined ? Number(r.loadCount) : '');
      const noteCol = headers.indexOf('備考') + 1;
      if (noteCol > 0) s.getRange(row, noteCol).setValue(r.notes||'');
      cacheClear();
      return {success:true};
    }
  }
  return {error:'該当する日報が見つかりません'};
}

function deleteReport(id) {
  if (!id) return {error:'報告IDがありません（旧データは削除できません）'};
  const s = getReportSheet();
  const lastCol = s.getLastColumn();
  const headers = s.getRange(1,1,1,lastCol).getValues()[0];
  const idCol = headers.indexOf('報告ID') + 1;
  if (idCol === 0) return {error:'報告ID列がありません'};
  const rows = s.getDataRange().getValues();
  for (let i=1;i<rows.length;i++){
    if (String(rows[i][idCol-1]) === String(id)) {
      s.deleteRow(i+1);
      cacheClear();
      return {success:true};
    }
  }
  return {error:'該当する日報が見つかりません'};
}

// ============================================================
// 初期データ投入（GASエディタから1回だけ手動実行してください）
// 車両番号は空欄のまま。CARS配列にご自身で車番・車種・拠点・タイプを追加して実行してください。
// type: 'passenger'（乗用車）or 'forest'（山林部トラック）
// ============================================================
function seedDefaultCars() {
  const CARS = [
    // ['車番','車種名','拠点','type'],
    // 例: ['4750','トヨタ','菊川ヤード','passenger'],
    // 例: ['211','三菱・ダンプ','萩ヤード','forest'],
  ];
  if (CARS.length === 0) {
    Logger.log('CARS配列が空です。車両情報を追加してから実行してください。');
    return;
  }
  const s = getOrCreate(SHEET_CAR, CAR_HEADERS);
  const existing = s.getDataRange().getValues().slice(1).map(function(r){return String(r[0]);});
  let added = 0;
  CARS.forEach(function(c){
    if (existing.indexOf(c[0]) === -1) {
      s.appendRow([c[0], c[1], c[2], c[3], new Date(), '稼働中']);
      added++;
    }
  });
  cacheClear();
  Logger.log('登録完了: ' + added + '台追加（既存 ' + existing.length + '台はスキップ）');
}
