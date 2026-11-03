/**
 * Lingel · DISPATCH PLANNING — standalone Apps Script (own sheet, own deployment)
 * App: https://vishalsaini-afk.github.io/dispatchplanning/   ·   also feeds Team Availability → To Assign → "Ready & paid"
 *
 * SET-UP (once, ~5 min)
 *  1. Open the "Lingel Dispatch Planning" Google Sheet
 *     (https://docs.google.com/spreadsheets/d/1sZydo7x63z_vQutJQHKi4dG1f7vLy2tPWpsRSvlMXik) → Extensions → Apps Script.
 *  2. Delete the sample code, paste this whole file, Save.
 *  3. Set the passwords below: DP_KEY (planning) and DP_DISPATCH_KEY (dispatch). Leave TEAM_KEY as it is.
 *  4. Choose setupDispatchPlanning in the function menu → ▶ Run → allow access (reads ORDER PLANNING, writes "Ready Orders").
 *  5. Deploy → New deployment → type Web app → Execute as: Me · Who has access: Anyone → Deploy → copy the /exec URL.
 *  Later changes: Deploy → Manage deployments → ✏ → Version: New version → Deploy (the URL stays the same).
 *
 * The order filter (DT_SKIP_CODES, DT_SKIP_CLIENT) is a COPY of the one in the invoice script's
 * Dispatch Tracker — if you change the code list there, change it here too.
 */

/* ---------------- settings ---------------- */
var DT_PLAN_SS = '1QqxejwkcLgptXDppqaLQYYcjXVRYBOyxX6Y--AnWAmg';   // ORDER PLANNING (Aluminum / UPVC tabs)
var DT_TZ      = 'Asia/Kolkata';
var TEAM_KEY   = 'lingel-kyc-7Q2';     // Team Availability's key — lets its "Ready & paid" tab read paid orders and confirm loading
var DP_SS_ID   = '1sZydo7x63z_vQutJQHKi4dG1f7vLy2tPWpsRSvlMXik';   // the "Lingel Dispatch Planning" sheet (works bound or standalone)

/* order filter — same as To Assign (a copy; keep both in step) */
var DT_SKIP_CODES = ['C','S','B','A','H','J','V','R','M'];
function dt_codeSkipped_(code){ return DT_SKIP_CODES.indexOf(String(code||'').trim().toUpperCase())>=0; }
var DT_SKIP_CLIENT = /\([^)]*\b(hardware|glass|mesh)\b[^)]*\)/i;   // e.g. (Hardware) (Glass) ( GLASS ) (Mesh) ( Pleated mesh ) ( Fabric Mesh Net )

/* ---------------- web app ---------------- */
function doGet(e){ return dp_http_(e); }
function doPost(e){ return dp_http_(e); }
function dp_http_(e){
  var p = (e && e.parameter) || {}, a = String(p.action||'');
  var out = a==='ping' ? {ok:true, app:'dispatch-planning', at:new Date().toISOString()} : (readyRoute_(a, p) || {ok:false, error:'Unknown action: '+a});
  var s = JSON.stringify(out);
  if (p.callback) return ContentService.createTextOutput(p.callback+'('+s+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function dp_ss_(){ return DP_SS_ID ? SpreadsheetApp.openById(DP_SS_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

/* run once from the editor: creates the tab, reads the plan, asks for permissions */
function setupDispatchPlanning(){
  dp_sheet_(); dp_sync_();
  Logger.log('Ready Orders: '+(dp_sheet_().getLastRow()-1)+' orders. Now deploy as a Web app (Execute as Me, access Anyone).');
}

/* ---------------- helpers (copied from the invoice script) ---------------- */
function dt_col_(h, names){ for(var i=0;i<names.length;i++){ var j=h.indexOf(names[i]); if(j>=0) return j; } return -1; }
var DT_MON_={jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
function dt_toDate_(v){
  if(v instanceof Date) return isNaN(v) ? null : v;
  var s=String(v||'').trim(); if(!s) return null; var m;
  if((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1],+m[2]-1,+m[3]);
  if((m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/))){ var y=+m[3]; if(y<100) y+=2000; return new Date(y,+m[1]-1,+m[2]); }       // M/D/YYYY (sheet format)
  if((m=s.match(/^(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s,]*(\d{2,4})/))){ var y2=+m[3]; if(y2<100) y2+=2000; var mo=DT_MON_[m[2].toLowerCase()]; if(mo!=null) return new Date(y2,mo,+m[1]); }
  if((m=s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/))) return new Date(+m[3],+m[2]-1,+m[1]);                                         // DD-MM-YYYY
  var d=new Date(s); return isNaN(d) ? null : d;
}
/* =====================================================================================
   READY → PAYMENT → DISPATCH  (Dispatch Planning app · Team Availability · Dispatch Team app)
   1) Every order with status "Ready" in ORDER PLANNING (same filter as To Assign: codes in
      DT_SKIP_CODES and Hardware/Glass/Mesh supply orders left out) is listed by REGION, read
      from its address (the planner can set it by hand).
   2) Dispatch Planning taps "Payment received" (office password DP_KEY).
   3) Team Availability → To Assign → "Ready & paid": confirm loading + installation needed.
   4) Dispatch Team app shows what was confirmed.
   State lives in the "Ready Orders" tab of the Installation invoice sheet.
   ===================================================================================== */
var DP_KEY          = 'lingel-dispatch';   // PLANNING login: sees everything, approves payment, plans dates — CHANGE IT
var DP_DISPATCH_KEY = 'lingel-loading';    // DISPATCH login: sees only the dispatch plan (loading confirmed) — CHANGE IT
/* who is asking: planning / dispatch (Dispatch Planning app logins) or team (Team Availability, via TEAM_KEY) */
function dp_role_(key){ key=String(key||''); if(!key) return ''; if(key===DP_KEY) return 'plan'; if(key===DP_DISPATCH_KEY) return 'disp'; if(key===TEAM_KEY) return 'team'; return ''; }
var DP_TAB       = 'Ready Orders';
var DP_KEEP_DAYS = 30;                  // a paid order that is no longer Ready (e.g. dispatched) stays listed this long
var DP_HEAD = ['Offer','Client','Sales Person','Upvc /Alu','Code','Units','Address','Region','Region (manual)','Status',
               'Ready Since','Payment','Approved By','Approved On','Loading Confirmed','Installation Needed','Confirmed By','Confirmed On',
               'Dispatch Date','Planned By','Planned On'];

function readyRoute_(a, p){
  if(['readyLogin','readyList','readyApprove','readyUnapprove','readyRegion','readyConfirm','readyUnconfirm','readyPlan','readyUnplan'].indexOf(a)<0) return null;
  try{
    var role = dp_role_(p.key);
    if(a==='readyLogin') return role==='plan'||role==='disp' ? {ok:true, role:role} : {ok:false, error:'Wrong password'};
    if(!role) return {ok:false, error:'Wrong password — please sign in again'};
    if(a==='readyList'){            // each login sees only its part
      var L = dp_list_(String(p.refresh||'')==='1' && role==='plan');
      if(role==='disp') L.rows = L.rows.filter(function(o){ return o.stage==='dispatch'; });
      if(role==='team') L.rows = L.rows.filter(function(o){ return o.paid; });
      L.role = role; return L;
    }
    if(a==='readyApprove' || a==='readyUnapprove' || a==='readyRegion' || a==='readyPlan' || a==='readyUnplan'){
      if(role!=='plan') return {ok:false, error:'Only the planning login can do this'};
    }
    if((a==='readyConfirm' || a==='readyUnconfirm') && role!=='team' && role!=='plan') return {ok:false, error:'Loading is confirmed from Team Availability'};
    var lock = LockService.getScriptLock(); if(!lock.tryLock(20000)) return {ok:false, error:'Sheet busy — try again'};
    try{
      var sh = dp_sheet_(), row = dp_find_(sh, p.offer); if(row<0) return {ok:false, error:'Offer '+(p.offer||'')+' is not in the Ready list'};
      var now = Utilities.formatDate(new Date(), DT_TZ, 'dd-MMM-yyyy HH:mm'), by = String(p.by||'').trim();
      var set = function(col, v){ sh.getRange(row, DP_HEAD.indexOf(col)+1).setValue(v); };
      if(a==='readyApprove'){ set('Payment','Received'); set('Approved By', by); set('Approved On', now); }
      if(a==='readyUnapprove'){ set('Payment','Pending'); set('Approved By',''); set('Approved On',''); set('Loading Confirmed',''); set('Installation Needed',''); set('Confirmed By',''); set('Confirmed On',''); set('Dispatch Date',''); set('Planned By',''); set('Planned On',''); }
      if(a==='readyPlan'){
        if(String(sh.getRange(row, DP_HEAD.indexOf('Payment')+1).getValue())!=='Received') return {ok:false, error:'Mark payment received first'};
        var dd = String(p.date||'').trim(); if(!/^\d{4}-\d{2}-\d{2}$/.test(dd)) return {ok:false, error:'Pick a dispatch date'};
        set('Dispatch Date', "'"+dd); set('Installation Needed', /^n/i.test(String(p.install||'')) ? 'No' : 'Yes'); set('Planned By', by); set('Planned On', now);
      }
      if(a==='readyUnplan'){ set('Dispatch Date',''); set('Planned By',''); set('Planned On',''); }
      if(a==='readyRegion'){ set('Region (manual)', String(p.region||'').trim()); }
      if(a==='readyConfirm'){
        if(String(sh.getRange(row, DP_HEAD.indexOf('Payment')+1).getValue())!=='Received') return {ok:false, error:'Payment not marked received yet'};
        set('Loading Confirmed','Yes'); set('Installation Needed', /^n/i.test(String(p.install||'')) ? 'No' : 'Yes'); set('Confirmed By', by); set('Confirmed On', now);
      }
      if(a==='readyUnconfirm'){ set('Loading Confirmed',''); set('Confirmed By',''); set('Confirmed On',''); }   // keeps the planner's installation choice
      return {ok:true, offer:p.offer, row: dp_rowObj_(sh.getRange(row,1,1,DP_HEAD.length).getValues()[0])};
    } finally { lock.releaseLock(); }
  }catch(err){ return {ok:false, error:String(err && err.message || err)}; }
}

function dp_sheet_(){
  var ss = dp_ss_(), sh = ss.getSheetByName(DP_TAB);
  if(!sh){ sh = ss.insertSheet(DP_TAB); sh.setFrozenRows(1); }
  if(sh.getLastColumn() < DP_HEAD.length || String(sh.getRange(1,1).getValue())!=='Offer')
    sh.getRange(1,1,1,DP_HEAD.length).setValues([DP_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f2937');
  return sh;
}
function dp_find_(sh, offer){
  var o = String(offer||'').trim(); if(!o || sh.getLastRow()<2) return -1;
  var v = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(var i=0;i<v.length;i++) if(String(v[i][0]).trim()===o) return i+2;
  return -1;
}
function dp_rowObj_(r){
  var o = {}; DP_HEAD.forEach(function(h,i){ o[h]=r[i]; });
  var paid = String(o['Payment'])==='Received', conf = String(o['Loading Confirmed'])==='Yes';
  var dd = o['Dispatch Date'] instanceof Date ? Utilities.formatDate(o['Dispatch Date'], DT_TZ, 'yyyy-MM-dd') : String(o['Dispatch Date']||'').replace(/^'/,'');
  return { offer:String(o['Offer']), client:String(o['Client']), sales:String(o['Sales Person']), dept:String(o['Upvc /Alu']), code:String(o['Code']),
    units:String(o['Units']), address:String(o['Address']), region:String(o['Region (manual)']||o['Region']||'Address missing'), regionAuto:String(o['Region']),
    regionManual:String(o['Region (manual)']), status:String(o['Status']), readySince:String(o['Ready Since']),
    paid:paid, approvedBy:String(o['Approved By']), approvedOn:String(o['Approved On']),
    confirmed:conf, install:String(o['Installation Needed']), confirmedBy:String(o['Confirmed By']), confirmedOn:String(o['Confirmed On']),
    dispatchDate:dd, plannedBy:String(o['Planned By']), plannedOn:String(o['Planned On']),
    /* ready = waiting payment · paid = needs a dispatch date · planned = dated, waiting loading confirmation · dispatch = dated + loading confirmed */
    stage: !paid ? 'ready' : (!dd ? 'paid' : (conf ? 'dispatch' : 'planned')) };
}

/* re-read ORDER PLANNING at most every 2 minutes (or on refresh), keep approvals */
function dp_list_(refresh){
  var props = PropertiesService.getScriptProperties(), last = +props.getProperty('dpSync')||0;
  if(refresh || Date.now()-last > 120000) dp_sync_();
  var sh = dp_sheet_(), rows = [];
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,DP_HEAD.length).getValues().forEach(function(r){ if(String(r[0]).trim()) rows.push(dp_rowObj_(r)); });
  return {ok:true, lastSync: new Date(+props.getProperty('dpSync')||Date.now()).toISOString(), rows:rows};
}
function dp_sync_(){
  var lock = LockService.getScriptLock(); if(!lock.tryLock(20000)) return;
  try{
    var ready = dp_readReady_(), sh = dp_sheet_(), W = DP_HEAD.length, H = function(n){ return DP_HEAD.indexOf(n); };
    var old = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,W).getValues() : [];
    var today = Utilities.formatDate(new Date(), DT_TZ, 'dd-MMM-yyyy'), out = [], seen = {};
    old.forEach(function(r){
      var k = String(r[0]).trim(); if(!k) return; seen[k]=1;
      var o = ready[k];
      if(o){   // still Ready → refresh the order details, keep approval / confirmation / manual region
        r[H('Client')]=o.client; r[H('Sales Person')]=o.sales; r[H('Upvc /Alu')]=o.dept; r[H('Code')]=o.code; r[H('Units')]=o.units;
        r[H('Address')]=o.address; r[H('Region')]=dp_region_(o.address); r[H('Status')]=o.status; out.push(r); return;
      }
      // no longer Ready: keep only paid orders, for DP_KEEP_DAYS, with the status they moved to
      if(String(r[H('Payment')])!=='Received') return;
      var on = dt_toDate_(String(r[H('Approved On')]).split(' ')[0]);
      if(on && (Date.now()-on.getTime()) > DP_KEEP_DAYS*86400000) return;
      if(!/^left ready/i.test(String(r[H('Status')]))) r[H('Status')] = 'Left Ready list';
      out.push(r);
    });
    Object.keys(ready).forEach(function(k){ if(seen[k]) return; var o = ready[k];
      var r = DP_HEAD.map(function(){ return ''; });
      r[H('Offer')]=k; r[H('Client')]=o.client; r[H('Sales Person')]=o.sales; r[H('Upvc /Alu')]=o.dept; r[H('Code')]=o.code; r[H('Units')]=o.units;
      r[H('Address')]=o.address; r[H('Region')]=dp_region_(o.address); r[H('Status')]=o.status; r[H('Ready Since')]=today; r[H('Payment')]='Pending';
      out.push(r); });
    if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,W).clearContent();
    if(out.length) sh.getRange(2,1,out.length,W).setValues(out);
    PropertiesService.getScriptProperties().setProperty('dpSync', String(Date.now()));
  } finally { lock.releaseLock(); }
}
/* Ready orders from Aluminum + UPVC (same filter as the Dispatch Tracker) */
function dp_readReady_(){
  var ss = SpreadsheetApp.openById(DT_PLAN_SS), map = {};
  [['Aluminum','Alu'],['UPVC','UPVC']].forEach(function(t){
    var sh = ss.getSheetByName(t[0]); if(!sh) return;
    var v = sh.getDataRange().getValues(); if(v.length<2) return;
    var h = v[0].map(function(x){ return String(x).replace(/\s+/g,' ').trim().toLowerCase(); });
    var cOffer=dt_col_(h,['offer']), cClient=dt_col_(h,['client']), cSales=dt_col_(h,['salesperson']), cCode=dt_col_(h,['code']),
        cStatus=dt_col_(h,['status & remarks','status/remarks','status']), cUnit=dt_col_(h,['unit']), cAddr=dt_col_(h,['address']);
    if(cAddr<0 && t[0]==='UPVC') cAddr = h.indexOf('');
    for(var i=1;i<v.length;i++){
      var r=v[i], offer=String(r[cOffer]||'').trim().replace(/\/+/g,'/'); if(!offer) continue;
      var st=String(r[cStatus]||'').trim(); if(!/^ready/i.test(st)) continue;
      var code = cCode>=0 ? String(r[cCode]||'').trim().toUpperCase() : '';
      if(dt_codeSkipped_(code)) continue;
      if(DT_SKIP_CLIENT.test(String(r[cClient]||''))) continue;
      map[offer] = { client:String(r[cClient]||'').trim(), sales:String(r[cSales]||'').trim(), dept:t[1], code:code,
                     units: cUnit>=0 ? String(r[cUnit]||'') : '', address: cAddr>=0 ? String(r[cAddr]||'').trim() : '', status:st };
    }
  });
  return map;
}
/* region from a free-text address: known city names first, then the PIN code, else "Address missing" */
var DP_CITY = [
  ['Delhi NCR', /\b(new\s*delhi|delhi|dallu?pura|dallipura|gurugr?a?m|gurgaon|gurugaon|faridabad|faridbad|noida|greater\s*noida|ghaziabad|gaziabad|gaziyabad|manesar|sohna)\b/i],
  ['Chandigarh & Punjab', /\b(chandigarh|mohali|panchkula|zirakpur|kharar|ludhiana|amritsar|jalandhar|patiala|bathinda|punjab(?!\s*bagh|\s*bhagh))\b/i],
  ['Haryana', /\b(kapriwas|dharuhera|rewari|rohtak|sonipat|sonepat|panipat|karnal|hisar|bahadurgarh|jhajjar|narnaul|narnoul|bhiwani|sirsa|ambala|kurukshetra|yamunanagar|pataudi|haryana)\b/i],
  ['Uttarakhand', /\b(dehradun|haridwar|rishikesh|haldwani|rudrapur|kashipur|nainital|ranikhet|roorkee|uttarakhand|uk)\b/i],
  ['Uttar Pradesh', /\b(agra|mathura|hathras|hatras|aligarh|meerut|moradabad|muradabad|bareilly|bareili|lucknow|kanpur|shamli|muzaffarnagar|saharanpur|chandausi|chadusi|varanasi|prayagraj|allahabad|up|u\.p\.?)\b/i],
  ['Rajasthan', /\b(jaipur|udaipur|jodhpur|ajmer|alwar|kota|bhiwadi|neemrana|behror|kaharani|rajasthan|rj)\b/i],
  ['Himachal', /\b(shimla|solan|kasauli|manali|dharamshala|baddi|himachal)\b/i],
  ['Jammu & Kashmir', /\b(jammu|srinagar|shrinagar|kashmir)\b/i],
  ['Other states', /\b(odisha|orissa|tirupati|hyderabad|bangalore|bengaluru|chennai|mumbai|pune|goa|kolkata|raipur|patna|nepal|kathmandu|bihar|gujarat|ahmedabad)\b/i]
];
var DP_PIN = [[110,110,'Delhi NCR'],[121,122,'Delhi NCR'],[201,201,'Delhi NCR'],[123,136,'Haryana'],[160,160,'Chandigarh & Punjab'],[140,159,'Chandigarh & Punjab'],
  [244,249,'Uttarakhand'],[262,263,'Uttarakhand'],[202,285,'Uttar Pradesh'],[301,345,'Rajasthan'],[171,177,'Himachal'],[180,194,'Jammu & Kashmir']];
function dp_region_(addr){
  var a = String(addr||'').trim(); if(!a) return 'Address missing';
  for(var i=0;i<DP_CITY.length;i++) if(DP_CITY[i][1].test(a)) return DP_CITY[i][0];
  var m = a.match(/\b(\d{3})\s?\d{3}\b/);
  if(m){ var n=+m[1]; for(var j=0;j<DP_PIN.length;j++) if(n>=DP_PIN[j][0] && n<=DP_PIN[j][1]) return DP_PIN[j][2]; }
  return 'Address missing';
}
