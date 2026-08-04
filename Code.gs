// 改善報告アプリ GAS
//
// このファイルはGoogle Apps Scriptプロジェクトの内容をバージョン管理用に
// リポジトリへコピーしたものです。実際の反映は Apps Script エディタ
// (script.google.com) 側でこの内容を貼り付けて保存・再デプロイする必要があります。
// このリポジトリを更新するだけではアプリの動作は変わりません。

const SHEET   = '報告データ';
const MASTER  = 'マスタ';
const SETTING = '設定';
const HEADERS = ['id','datetime','customer','unit','staff','improvement',
  'reductionTime','reductionCost',
  'beforePhoto','beforePhoto2','beforePhoto3','beforePhoto4','beforePhoto5',
  'afterPhoto','afterPhoto2','afterPhoto3','afterPhoto4','afterPhoto5',
  'status','updatedAt','bCount','aCount'];
const META_COLS = ['id','datetime','customer','unit','staff','improvement','reductionTime','reductionCost'];
const BK = ['beforePhoto','beforePhoto2','beforePhoto3','beforePhoto4','beforePhoto5'];
const AK = ['afterPhoto','afterPhoto2','afterPhoto3','afterPhoto4','afterPhoto5'];

function doGet(e){
  const a=e.parameter.action||'';
  try{
    if(a==='getReports')    return ok(getReports());
    if(a==='getReportFull') return ok(getReportFull(e.parameter));
    if(a==='getSettings')   return ok(getSettings());
    return ok({status:'error',message:'unknown:'+a});
  }catch(ex){return ok({status:'error',message:ex.message});}
}
function doPost(e){
  let p;
  try{
    // フォーム送信(e.parameter.payload)と生POST(e.postData.contents)の両方に対応
    var raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : (e.postData ? e.postData.contents : '');
    p = JSON.parse(raw);
  }
  catch(ex){ return ok({status:'error',message:'JSON parse error'}); }
  try{
    const a=p.action||'';
    if(a==='addReport')    return ok(addReport(p));
    if(a==='updateReport') return ok(updateReport(p));
    if(a==='deleteReport') return ok(deleteReport(p));
    if(a==='saveSettings') return ok(saveSettings(p));
    return ok({status:'error',message:'unknown:'+a});
  }catch(ex){ return ok({status:'error',message:ex.message}); }
}
function ok(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- sheet helpers ----------
function getOrCreate(name,hdrs){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);if(hdrs)sh.appendRow(hdrs);}
  return sh;
}
function reportSheet(){
  const sh=getOrCreate(SHEET,HEADERS);
  ensureCols(sh,HEADERS);
  return sh;
}
function ensureCols(sh,req){
  const lc=sh.getLastColumn();
  const ex=lc>0?sh.getRange(1,1,1,lc).getValues()[0].map(String):[];
  req.forEach(h=>{if(!ex.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);ex.push(h);}});
}
function settingSheet(){return getOrCreate(SETTING,['key','value']);}

// ---------- getReports ----------
// 一覧用：写真は1枚目のみ送信し、2〜5枚目は枚数カウントのみ（大幅軽量化）
// 一覧用：写真列は一切読み込まず、メタデータと保存済み枚数のみ返す（件数が増えても高速）
function getReports(){
  const sh=reportSheet();
  const lr=sh.getLastRow();
  if(lr<=1)return{status:'ok',reports:[]};
  const lc=sh.getLastColumn();
  const hdrs=sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const n=lr-1;

  // メタ列はヘッダー名で実際の列位置を特定してから読む（列の並び替え・列挿入があっても値がずれないように）。
  // 通常はMETA_COLSが連続した列にまとまっているため、その範囲だけを1回のAPI呼び出しで読み込む。
  const metaIdx=META_COLS.map(name=>hdrs.indexOf(name)).filter(c=>c>=0);
  const minC=Math.min(...metaIdx), maxC=Math.max(...metaIdx);
  const metaSpan=sh.getRange(2,minC+1,n,maxC-minC+1).getValues();

  // status / bCount / aCount を列単位で読む（写真は読まない）
  function colVals(name){
    const c=hdrs.indexOf(name);
    return c>=0 ? sh.getRange(2,c+1,n,1).getValues() : null;
  }
  const stData=colVals('status');
  const bcData=colVals('bCount');
  const acData=colVals('aCount');

  const out=[];
  for(let i=n-1;i>=0;i--){ // 新しい順
    const r={};
    META_COLS.forEach(name=>{
      const c=hdrs.indexOf(name);
      r[name] = (c>=0 && metaSpan[i][c-minC]!=null) ? String(metaSpan[i][c-minC]) : '';
    });
    r.status = stData ? (String(stData[i][0])||'completed') : 'completed';
    if(!r.status) r.status='completed';
    r.bCount = bcData ? (Number(bcData[i][0])||0) : 0;
    r.aCount = acData ? (Number(acData[i][0])||0) : 0;
    out.push(r);
  }
  return{status:'ok',reports:out};
}

// 既存データの写真枚数を一括計算して保存（GASエディターで1回だけ実行）
function backfillCounts(){
  const sh=reportSheet();
  const lr=sh.getLastRow();
  if(lr<=1) return;
  const lc=sh.getLastColumn();
  const data=sh.getRange(1,1,lr,lc).getValues();
  const hdrs=data[0].map(String);
  const bcCol=hdrs.indexOf('bCount'), acCol=hdrs.indexOf('aCount');
  if(bcCol<0||acCol<0){ Logger.log('bCount/aCount列がありません'); return; }
  const bIdx=BK.map(k=>hdrs.indexOf(k));
  const aIdx=AK.map(k=>hdrs.indexOf(k));
  const bcOut=[], acOut=[];
  for(let i=1;i<data.length;i++){
    let bc=0,ac=0;
    bIdx.forEach(j=>{if(j>=0&&data[i][j])bc++;});
    aIdx.forEach(j=>{if(j>=0&&data[i][j])ac++;});
    bcOut.push([bc]); acOut.push([ac]);
  }
  sh.getRange(2,bcCol+1,bcOut.length,1).setValues(bcOut);
  sh.getRange(2,acCol+1,acOut.length,1).setValues(acOut);
  Logger.log('backfillCounts完了: '+bcOut.length+'行');
}

// datetime列がGoogle Sheetsの自動書式判定でDate型になっているのを修正する（GASエディターで1回だけ実行）。
// 自動判定のままだと"2026-08-04"のような文字列を書き込んでもDate型に変換され、
// 読み出し時にString(Date)がロケール依存の文字列（"Sun Aug 04 2026 ..."等）になり、
// フロント側のparseD()やSafari等でのDate解析が崩れる恐れがある。
// 列の書式をプレーンテキストに固定し、既存セルもyyyy-MM-dd文字列に正規化する。
function fixDatetimeColumn(){
  const sh=reportSheet();
  const hdrs=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const dc=hdrs.indexOf('datetime');
  if(dc<0){ Logger.log('datetime列がありません'); return; }

  // 列全体（未使用の将来行も含む）をプレーンテキスト書式に固定し、以後の自動変換を防ぐ
  sh.getRange(1,dc+1,sh.getMaxRows(),1).setNumberFormat('@');

  const lr=sh.getLastRow();
  if(lr<=1) return;
  const tz=Session.getScriptTimeZone();
  const vals=sh.getRange(2,dc+1,lr-1,1).getValues();
  const fixed=vals.map(row=>{
    const v=row[0];
    return [v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v)];
  });
  sh.getRange(2,dc+1,fixed.length,1).setValues(fixed);
  Logger.log('fixDatetimeColumn完了: '+fixed.length+'行');
}

// 詳細・編集用：1件の全写真を含めて返す（モーダルを開いた時のみ呼ぶ）
function getReportFull(param){
  const id=param.id;
  const {sh,tr,hdrs}=findRow(id);
  const row=sh.getRange(tr,1,1,hdrs.length).getValues()[0];
  const r={};
  hdrs.forEach((h,j)=>{r[h]=row[j]!=null?String(row[j]):'';});
  if(!r.status)r.status='completed';
  return{status:'ok',report:r};
}

// ---------- addReport (header-based row build) ----------
function addReport(p){
  const sh=reportSheet();
  const lc=sh.getLastColumn();
  const hdrs=sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const id=Utilities.getUuid();
  const now=new Date().toISOString();
  const status=AK.some(k=>p[k]&&p[k].length>0)?'completed':'in-progress';
  const bc=BK.filter(k=>p[k]&&p[k].length>0).length;
  const ac=AK.filter(k=>p[k]&&p[k].length>0).length;
  const vals={id,datetime:p.datetime||'',customer:p.customer||'',unit:p.unit||'',
    staff:p.staff||'',improvement:p.improvement||'',
    reductionTime:p.reductionTime||'',reductionCost:p.reductionCost||'',
    beforePhoto:p.beforePhoto||'',beforePhoto2:p.beforePhoto2||'',
    beforePhoto3:p.beforePhoto3||'',beforePhoto4:p.beforePhoto4||'',
    beforePhoto5:p.beforePhoto5||'',afterPhoto:p.afterPhoto||'',
    afterPhoto2:p.afterPhoto2||'',afterPhoto3:p.afterPhoto3||'',
    afterPhoto4:p.afterPhoto4||'',afterPhoto5:p.afterPhoto5||'',
    status,updatedAt:now,bCount:bc,aCount:ac};
  const row=hdrs.map(h=>vals[h]!==undefined?vals[h]:'');
  sh.appendRow(row);
  return{status:'ok',id};
}

// ---------- updateReport (header+ID only reads, batch photo write) ----------
// ---------- 共通ヘルパー：ID から対象行を特定 ----------
function findRow(id){
  const sh=reportSheet();
  const lc=sh.getLastColumn(), lr=sh.getLastRow();
  if(lr<=1) throw new Error('データがありません');
  const hdrs=sh.getRange(1,1,1,lc).getValues()[0].map(String);
  const idCol=hdrs.indexOf('id');
  if(idCol<0) throw new Error('idカラム不在');
  const ids=sh.getRange(2,idCol+1,lr-1,1).getValues();
  let tr=-1;
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0])===String(id)){tr=i+2;break;}
  }
  if(tr<0) throw new Error('対象レコード不在: '+id);
  return {sh, tr, hdrs};
}

// ---------- updateReport：対象行を1回読み込み→メモリ更新→1回書き戻し ----------
//    列の並びが連続でも非連続でも常に「読み取り1回＋書き込み1回」で完結する。
//    ensureColsで後から追加された列（非連続レイアウト）にも完全対応。
//    LockServiceで排他制御し、複数人が同じ報告を同時編集した際の上書き事故（ロストアップデート）を防ぐ。
function updateReport(p){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const {sh,tr,hdrs}=findRow(p.id);
    const lastCol=hdrs.length;

    // 対象行を1回だけ読み込む（写真データを含むが「1行分」のみ・高速）
    const row=sh.getRange(tr,1,1,lastCol).getValues()[0];

    // メモリ上で該当フィールドのみ更新（ペイロードにあるものだけ）
    [...BK,...AK,'status','improvement'].forEach(k=>{
      if(p[k]!==undefined){
        const c=hdrs.indexOf(k);
        if(c>=0) row[c]=p[k];
      }
    });
    const uc=hdrs.indexOf('updatedAt');
    if(uc>=0) row[uc]=new Date().toISOString();

    // 写真枚数を再計算して保存（一覧を写真なしで高速表示するため）
    const bcCol=hdrs.indexOf('bCount'), acCol=hdrs.indexOf('aCount');
    if(bcCol>=0){ let bc=0; BK.forEach(k=>{const c=hdrs.indexOf(k); if(c>=0&&row[c])bc++;}); row[bcCol]=bc; }
    if(acCol>=0){ let ac=0; AK.forEach(k=>{const c=hdrs.indexOf(k); if(c>=0&&row[c])ac++;}); row[acCol]=ac; }

    // 1回のsetValuesで書き戻す（列レイアウトに関係なく常に1回のAPIコール）
    sh.getRange(tr,1,1,lastCol).setValues([row]);
    return{status:'ok'};
  } finally {
    lock.releaseLock();
  }
}

// ---------- deleteReport ----------
// findRowを使い、idカラムのみを読んで対象行を特定する（写真データを含むシート全体の読み込みを避ける）。
// LockServiceで排他制御し、削除中に行番号がずれるのを防ぐ。
function deleteReport(p){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const {sh,tr}=findRow(p.id);
    sh.deleteRow(tr);
    return{status:'ok'};
  } finally {
    lock.releaseLock();
  }
}

// ---------- getSettings ----------
function getSettings(){
  const st=settingSheet();
  const d=st.getDataRange().getValues();
  const m={};
  for(let i=1;i<d.length;i++)m[String(d[i][0])]=String(d[i][1]);
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const ms=ss.getSheetByName(MASTER);
  let staffList=[],unitList=[];
  if(ms&&ms.getLastRow()>=2){
    ms.getRange(2,1,ms.getLastRow()-1,2).getValues().forEach(r=>{
      const s=String(r[0]).trim(),u=String(r[1]).trim();
      if(s&&!staffList.includes(s))staffList.push(s);
      if(u&&!unitList.includes(u))unitList.push(u);
    });
  }
  return{status:'ok',staffList,unitList,
    monthlyGoal:m.monthlyGoal||'4',annualGoal:m.annualGoal||'100',pinHash:m.pinHash||''};
}

// ---------- saveSettings ----------
// LockServiceで排他制御し、キー未存在時の追記が同時実行で重複行になるのを防ぐ。
function saveSettings(p){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const sh=settingSheet();
    const d=sh.getDataRange().getValues();
    const m={};
    for(let i=1;i<d.length;i++)m[String(d[i][0])]=i+1;
    function up(k,v){
      if(v===undefined)return;
      if(m[k])sh.getRange(m[k],2).setValue(v);
      else{const r=sh.getLastRow()+1;sh.getRange(r,1).setValue(k);sh.getRange(r,2).setValue(v);m[k]=r;}
    }
    up('monthlyGoal',p.monthlyGoal);up('annualGoal',p.annualGoal);up('pinHash',p.pinHash);
    return{status:'ok'};
  } finally {
    lock.releaseLock();
  }
}

// ---------- utilities ----------
function initializeSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ['報告データ','設定'].forEach(n=>{const s=ss.getSheetByName(n);if(s)ss.deleteSheet(s);});
  const rs=ss.insertSheet('報告データ');rs.appendRow(HEADERS);
  // datetime列は自動書式判定でDate型化されるとフロント側の日付処理が崩れるため、
  // 最初からプレーンテキスト書式に固定しておく（既存シートの移行はfixDatetimeColumnを参照）
  const dc=HEADERS.indexOf('datetime');
  if(dc>=0) rs.getRange(1,dc+1,rs.getMaxRows(),1).setNumberFormat('@');
  const st=ss.insertSheet('設定');st.appendRow(['key','value']);
  ['monthlyGoal|4','annualGoal|100','pinHash|'].forEach(s=>{const[k,v]=s.split('|');st.appendRow([k,v]);});
  if(!ss.getSheetByName('マスタ')){
    const ms=ss.insertSheet('マスタ');ms.appendRow(['担当者','部署']);
    ms.appendRow(['山田 太郎','回収部']);ms.appendRow(['鈴木 花子','管理部']);
  }
  Logger.log('初期化完了');
}
function backfillIds(){
  const sh=reportSheet();
  const data=sh.getDataRange().getValues();
  const hdrs=data[0].map(String);
  const ic=hdrs.indexOf('id');
  let n=0;
  for(let i=1;i<data.length;i++){
    if(!data[i][ic]){sh.getRange(i+1,ic+1).setValue(Utilities.getUuid());n++;}
  }
  Logger.log('backfill: '+n);
}
