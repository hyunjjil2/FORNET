/**
 * 포워딩 매니페스트 — Apps Script Web App 백엔드
 * 이 스크립트는 반드시 "업무원본" 시트가 있는 구글시트에 바인딩된 상태로 (그 시트에서
 * 확장 프로그램 > Apps Script 로 열어서) 사용해야 합니다.
 *
 * 배포: 배포 > 새 배포 > 유형: 웹 앱 > 실행 계정: 나 > 액세스 권한: 전체 공개(익명 포함)
 * 배포 후 나오는 웹 앱 URL을 index.html의 WEBAPP_URL 에 붙여넣으세요.
 *
 * 성능 참고: 대량 업로드(수백~수천 행)를 빠르게 처리하기 위해 셀 단위 읽기/쓰기 대신
 * 한 번에 전체 데이터를 읽고(getValues) 한 번에 통째로 씁니다(setValues).
 */

var SHEET_NAME = '업무원본';
var SETTINGS_SHEET_NAME = '설정';
var BOOKING_HOLD_SHEET_NAME = '부킹확인중';
var BOOKING_HOLD_HEADERS = ['ID','제목','내용','상태','생성일시','수정일시']; // 자유형식 메모 — 항목별 입력칸 없이 제목+본문만
var BOOKING_HOLD_DATE_HEADERS = [];
var INFO_HEADERS = ['관리번호','루틴','청구전용','캔슬','서비스','수출입','거래처코드','거래처명','담당자','POL','POD',
  'DOC CLS','CFS CLS','ETD','ETA(PORT)','ETA(DOOR)','Door배송','배차필요','프리알러트필요','비엘홀딩','Line Booking No','House B/L No','Master B/L No','PO No','선사','운송사','VESSEL','터미널','특이사항',
  '생성일시','수정일시'];
var DATE_HEADERS = ['DOC CLS','CFS CLS','ETD','ETA(PORT)','ETA(DOOR)'];
var BOOL_HEADERS = ['청구전용','캔슬','Door배송','배차필요','프리알러트필요','비엘홀딩'];
var ARCHIVE_SHEET_NAME = '보관';
/* index.html의 ROUTES 설정(항목 key + billing 여부만)을 그대로 옮겨온 것 — "완료" 여부를
   서버에서도 판단하려면(오래된 완료건 자동 정리용) 프론트와 동일한 기준이 필요함.
   주의: index.html의 ROUTES.items를 바꾸면(항목 추가/삭제) 여기도 함께 맞춰줘야 함 —
   INFO_HEADERS와 마찬가지로 두 파일 간 수동 동기화가 필요한 설정. */
var ROUTE_ITEMS_ = {
  SEA_EXP_FCL: [
    {key:'myeonjang'},{key:'docCls'},{key:'jigyeol'},
    {key:'ichejeung'},{key:'chulhang'},{key:'blIssue'},{key:'invoice',billing:true}
  ],
  SEA_EXP_LCL: [
    {key:'myeonjang'},{key:'docCls'},{key:'jigyeol'},{key:'chulhang'},
    {key:'blIssue'},{key:'invoice',billing:true}
  ],
  SEA_IMP_FCL: [
    {key:'checkBill'},{key:'closeBill'},{key:'an'},{key:'jigyeol'},{key:'doIssue'},{key:'arrival'},
    {key:'customsInvoice'},{key:'customs'},{key:'carrierJigyeol'},
    {key:'finalInvoice',billing:true}
  ],
  SEA_IMP_LCL: [
    {key:'checkBill'},{key:'closeBill'},{key:'consolAn'},{key:'ownerReport'},{key:'consolJigyeol'},
    {key:'ichejeung'},{key:'doIssue'},{key:'customsInvoice'},{key:'customs'},{key:'carrierJigyeol'},
    {key:'finalInvoice',billing:true}
  ],
  AIR_EXP: [
    {key:'myeonjang'},{key:'chkBl'},{key:'preAlert'},{key:'airDocCls'},
    {key:'settlementDocs',billing:true},{key:'invoice',billing:true},{key:'carrierJigyeol',billing:true}
  ],
  AIR_IMP: [
    {key:'impDocCls'},{key:'impConsolJigyeol'},{key:'impDoIssue'},{key:'impCustomsFund'},
    {key:'impMyeonjang'},{key:'impBaecha'},{key:'impInvoice',billing:true}
  ]
};
function relevantItemKeys_(routeType, billingOnly){
  var items = ROUTE_ITEMS_[routeType] || [];
  return billingOnly ? items.filter(function(it){return it.billing;}) : items;
}
var PROGRESS_HEADERS = {'루틴':1,'청구전용':1,'캔슬':1};

function getSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) throw new Error('시트 "' + SHEET_NAME + '"를 찾을 수 없습니다.');
  return sh;
}
function headerMap_(sh){
  var last = sh.getLastColumn();
  var headers = sh.getRange(1,1,1,last).getValues()[0];
  var map = {};
  for(var i=0;i<headers.length;i++){ if(headers[i]) map[String(headers[i]).trim()] = i+1; }
  return map;
}
function findRowByManagementNo_(sh, hmap, managementNo){
  var col = hmap['관리번호'];
  if(!col) return -1;
  var last = sh.getLastRow();
  if(last < 2) return -1;
  var vals = sh.getRange(2, col, last-1, 1).getValues();
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][0]).trim() === String(managementNo).trim()) return i+2;
  }
  return -1;
}
/* 체크리스트 항목은 3단계 상태를 가짐: false(미완료) / true(완료) / 'na'(해당없음 — 이 건은
   해당 항목을 진행하지 않음, "완료"와 명확히 구분되어 오늘 할 일에서 완전히 제외됨).
   시트 셀에는 그대로 TRUE/FALSE/"NA" 문자열로 저장됨. */
function normalizeChkValue_(v){
  if(v===true || v==='TRUE' || v===1) return true;
  if(v==='na' || v==='NA') return 'na';
  return false;
}
/* 시트 셀에 실제로 쓸 값 — 프론트에서 오는 값(true/false/'na')을 그대로 시트에 반영 (NA는 "NA" 문자열로) */
function chkWriteValue_(v){
  if(v==='na' || v==='NA') return 'NA';
  return !!v;
}
function rowArrayToObject_(hmap, values){
  var obj = { checklist: {} };
  Object.keys(hmap).forEach(function(h){
    var v = values[hmap[h]-1];
    if(h.indexOf('CHK_')===0){
      obj.checklist[h.slice(4)] = normalizeChkValue_(v);
    } else if(DATE_HEADERS.indexOf(h)>=0){
      obj[h] = (v instanceof Date) ? v.toISOString() : (v || '');
    } else if(BOOL_HEADERS.indexOf(h)>=0){
      obj[h] = (v===true || v==='TRUE' || v===1);
    } else {
      obj[h] = (v===undefined||v===null) ? '' : v;
    }
  });
  return obj;
}

/* ================= 대한민국 공휴일(대체공휴일 포함) 자동 조회 =================
   서류마감·출항안내가 공휴일에 걸리는 건을 미리 알려주는 기능(index.html의 feedDueSoon·
   feedDocCls 등)을 위해, 구글이 공식 유지·관리하는 "대한민국 공휴일" 공개 캘린더에서 실제
   공휴일(설날·추석 연휴, 대체공휴일 포함)을 가져옴. 직접 날짜를 하드코딩하지 않는 이유는
   음력 기반 공휴일(설날·추석·부처님오신날)과 대체공휴일 적용 여부가 해마다 달라서 수작업으로
   관리하면 매년 틀리기 쉽기 때문 — 구글이 관리하는 캘린더를 쓰면 매년 자동으로 최신 상태 유지.
   혹시 캘린더 ID가 바뀌었거나 일시적으로 접근이 안 되는 경우에도 전체 조회(doGet) 자체가
   실패하면 안 되므로, 반드시 이 함수 내부에서 오류를 잡아 빈 배열을 반환하도록 함(대시보드
   핵심 기능에는 영향 없고, 공휴일 선반영 기능만 조용히 비활성화됨). 6시간 동안 캐시해서
   20초마다 도는 폴링마다 매번 캘린더를 조회하지 않도록 함. */
var KOREA_HOLIDAY_CALENDAR_ID_ = 'en.south_korea.official#holiday@group.v.calendar.google.com';
function getKoreaHolidays_(){
  try{
    var cache = CacheService.getScriptCache();
    var cached = cache.get('kr_holidays_v1');
    if(cached) return JSON.parse(cached);
    var cal = CalendarApp.getCalendarById(KOREA_HOLIDAY_CALENDAR_ID_);
    if(!cal) return [];
    var today = new Date();
    var start = new Date(today); start.setDate(start.getDate()-7);
    var end = new Date(today); end.setDate(end.getDate()+120); // 향후 120일치면 D-7 사전 안내에 충분히 여유 있음
    var events = cal.getEvents(start, end);
    var seen = {};
    var holidays = [];
    events.forEach(function(ev){
      var d = ev.isAllDayEvent() ? ev.getAllDayStartDate() : ev.getStartTime();
      var ymd = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
      if(seen[ymd]) return; // 같은 날짜에 캘린더 이벤트가 중복으로 잡혀도 한 번만
      seen[ymd] = 1;
      holidays.push({date: ymd, title: ev.getTitle()});
    });
    cache.put('kr_holidays_v1', JSON.stringify(holidays), 21600); // 6시간
    return holidays;
  }catch(err){
    return []; // 공휴일 조회 실패는 전체 조회를 막으면 안 되므로 조용히 빈 배열 반환
  }
}
function doGet(e){
  // 읽기 전용 호출이라 잠금(Lock)을 걸지 않음 — getValues()는 항상 일관된 스냅샷을 반환하므로
  // 안전하고, 여러 명이 동시에(또는 자동 20초 갱신이 겹쳐서) 조회해도 서로 기다리며 막히지 않음.
  try{
    var sh = getSheet_();
    var hmap = headerMap_(sh);
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    var out = [];
    if(lastRow >= 2){
      var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues(); // 전체를 한 번에 읽음 (행 단위 반복 X)
      var mgmtCol = hmap['관리번호'] - 1;
      for(var i=0;i<grid.length;i++){
        if(!grid[i][mgmtCol]) continue;
        out.push(rowArrayToObject_(hmap, grid[i]));
      }
    }
    var myClients = getMyClientCodes_();
    var labelOverrides = getItemLabelOverrides_();
    var bookingHolds = listBookingHolds_();
    var holidays = getKoreaHolidays_();
    var dispatchClients = [];
    try{ dispatchClients = getDispatchClientCodes_(); }catch(dcErr){ dispatchClients = []; }
    var preAlertClients = [];
    try{ preAlertClients = getPreAlertClientCodes_(); }catch(paErr){ preAlertClients = []; }
    var poNoteClients = [];
    try{ poNoteClients = getPoNoteClientCodes_(); }catch(pnErr){ poNoteClients = []; }
    return ContentService.createTextOutput(JSON.stringify({ok:true, rows: out, myClients: myClients, dispatchClients: dispatchClients, preAlertClients: preAlertClients, poNoteClients: poNoteClients, labelOverrides: labelOverrides, bookingHolds: bookingHolds, holidays: holidays}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e){
  // lock.waitLock()이 try 밖에 있으면, 락 대기 중 예외(LockService 타임아웃 등)가 발생했을 때
  // Apps Script가 우리 JSON 응답이 아니라 자체 HTML 오류 페이지를 반환해버려서
  // 프론트엔드에서 "Unexpected token '<', <!DOCTYPE...' is not valid JSON" 오류가 남.
  // (동시에 여러 체크리스트 항목을 빠르게 연속 클릭할 때 락 경합으로 발생하기 쉬움)
  // → try 안으로 옮기고, 락을 실제로 획득했을 때만 releaseLock 하도록 수정.
  var lock = LockService.getScriptLock();
  var haveLock = false;
  try{
    try{
      lock.waitLock(25000);
      haveLock = true;
    }catch(lockErr){
      throw new Error('다른 저장 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
    }
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var result;
    // 부킹확인중 메모 액션 — 업무원본 시트와 무관하게 별도 시트에서 처리
    if(action === 'upsertBookingHold'){ result = upsertBookingHold_(body.data||{}); }
    else if(action === 'deleteBookingHold'){ result = deleteBookingHold_(body.id); }
    else if(action === 'archiveOldCompleted'){ result = cleanupOldCompleted_(body.cutoffDays||180); }
    else {
      var sh = getSheet_();
      var hmap = headerMap_(sh);
      if(action === 'upsertMany'){
        result = upsertMany_(sh, hmap, body.rows||[]);
        // 엑셀 업로드(대량 반영)할 때마다 "내 거래처" 아닌 건을 자동으로 "보관"으로 옮겨 시트가
        // 계속 불어나지 않게 함 — upsertMany_가 헤더를 늘렸을 수 있으므로 sh를 그대로 넘겨 재조회.
        result.purge = purgeNonMyClients_(sh);
      }
      else if(action === 'upsertOne'){ result = upsertOne_(sh, hmap, body.id||null, body.data||{}); }
      else if(action === 'toggleItem'){ result = toggleItem_(sh, hmap, body.managementNo, body.itemKey, body.value); }
      else if(action === 'toggleMany'){ result = toggleMany_(sh, hmap, body.items||[], body.value); }
      else if(action === 'deleteOne'){ result = deleteOne_(sh, hmap, body.managementNo); }
      else if(action === 'setMyClients'){
        result = setMyClientCodes_(body.codes||[]);
        // 배차 거래처 목록은 같은 화면(내 거래처 관리)에서 함께 저장 — 보내오지 않았으면 기존 값 유지
        if(body.dispatchCodes){
          result.dispatchCodes = setDispatchClientCodes_(body.dispatchCodes).codes;
          // 이미 들어와 있는 해당 거래처 건(배차필요 미지정)에도 바로 적용
          result.dispatchApplied = applyDispatchToExisting_(result.dispatchCodes);
        }
        // PRE-ALERT 거래처 목록도 동일 — 보내오지 않았으면 기존 값 유지
        if(body.preAlertCodes){
          result.preAlertCodes = setPreAlertClientCodes_(body.preAlertCodes).codes;
          result.preAlertApplied = applyClientFlagToExisting_('프리알러트필요', result.preAlertCodes);
        }
        // PO No → 비고 자동 입력 거래처 목록도 동일 — 보내오지 않았으면 기존 값 유지
        if(body.poNoteCodes){
          result.poNoteCodes = setPoNoteClientCodes_(body.poNoteCodes).codes;
          result.poNoteApplied = applyPoNoteToExisting_(result.poNoteCodes);
        }
      }
      else if(action === 'setItemLabels'){ result = setItemLabelOverrides_(body.overrides||[]); }
      else { throw new Error('알 수 없는 action: ' + action); }
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true, result: result}))
      .setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }finally{
    if(haveLock) lock.releaseLock();
  }
}

function ensureCols_(sh, hmap, headers){
  var last = sh.getLastColumn();
  headers.forEach(function(h){
    if(!hmap[h]){
      last += 1;
      sh.getRange(1, last).setValue(h);
      hmap[h] = last;
    }
  });
}

function collectMissingHeaders_(hmap, rows){
  var missing = {};
  rows.forEach(function(r){
    Object.keys(r).forEach(function(h){
      if(h === 'checklist') return;
      if(!hmap[h]) missing[h] = 1;
    });
    if(r.checklist){
      Object.keys(r.checklist).forEach(function(k){
        var h = 'CHK_' + k;
        if(!hmap[h]) missing[h] = 1;
      });
    }
  });
  return Object.keys(missing);
}

function writeRow_(sh, hmap, row, data, opts){
  // opts.preserveProgress: 존재 시 진행 관련 필드(루틴/청구전용/캔슬/체크리스트)는 기존 값 유지 (재업로드 병합 규칙)
  opts = opts || {};
  var preserveProgress = !!opts.preserveProgress;
  var missing = collectMissingHeaders_(hmap, [data]);
  if(missing.length) ensureCols_(sh, hmap, missing);

  var lastCol = sh.getLastColumn();
  var rowArr = sh.getRange(row,1,1,lastCol).getValues()[0];
  Object.keys(data).forEach(function(h){
    if(h === 'checklist') return;
    if(preserveProgress && PROGRESS_HEADERS[h]) return; // 기존 값 유지
    var col = hmap[h]; if(!col) return;
    rowArr[col-1] = data[h];
  });
  if(data.checklist && !preserveProgress){
    Object.keys(data.checklist).forEach(function(k){
      var h = 'CHK_' + k;
      var col = hmap[h]; if(col) rowArr[col-1] = chkWriteValue_(data.checklist[k]);
    });
  }
  sh.getRange(row,1,1,lastCol).setValues([rowArr]);
}

function isBillingOnlyByNo_(no){ return /-X\d+$/i.test(String(no||'').trim()); }

/* 대량 업로드: 전체를 한 번에 읽고(getValues) 메모리에서 병합한 뒤 한 번에 씀(setValues).
   행이 수백~수천 개여도 실행 시간이 거의 늘지 않음. */
function upsertMany_(sh, hmap, rows){
  var dispatchSet = {};
  try{ getDispatchClientCodes_().forEach(function(c){ dispatchSet[c] = true; }); }catch(dcErr){}
  var preAlertSet = {};
  try{ getPreAlertClientCodes_().forEach(function(c){ preAlertSet[c] = true; }); }catch(paErr){}
  var missing = collectMissingHeaders_(hmap, rows);
  if(Object.keys(dispatchSet).length && !hmap['배차필요']) missing.push('배차필요');
  if(Object.keys(preAlertSet).length && !hmap['프리알러트필요']) missing.push('프리알러트필요');
  var poNoteSet = {};
  try{ getPoNoteClientCodes_().forEach(function(c){ poNoteSet[c] = true; }); }catch(pnErr){}
  if(Object.keys(poNoteSet).length){ if(!hmap['PO No']) missing.push('PO No'); if(!hmap['특이사항']) missing.push('특이사항'); }
  if(missing.length) ensureCols_(sh, hmap, missing);

  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  var grid = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,lastCol).getValues() : [];
  var mgmtCol = hmap['관리번호'] - 1;
  var indexByNo = {};
  for(var i=0;i<grid.length;i++){
    var v = String(grid[i][mgmtCol]||'').trim();
    if(v) indexByNo[v] = i;
  }

  var created=0, updated=0, dispatchFlagged=0, preAlertFlagged=0;
  var nowIso = new Date().toISOString();

  rows.forEach(function(r){
    var no = String(r['관리번호']||'').trim();
    if(!no) return;
    r['수정일시'] = nowIso;
    var idx = indexByNo.hasOwnProperty(no) ? indexByNo[no] : -1;
    if(idx >= 0){
      // 기존 건: 루틴·청구전용·캔슬·체크리스트 진행 상태는 절대 건드리지 않음
      Object.keys(r).forEach(function(h){
        if(h === 'checklist') return;
        if(PROGRESS_HEADERS[h]) return;
        var col = hmap[h]; if(!col) return;
        // 엑셀 칸이 비어 있으면 기존 값(대시보드에서 직접 입력한 서류마감일 등)을 그대로 둠 — 값이 있을 때만 덮어씀
        if(r[h]===''||r[h]===null||r[h]===undefined) return;
        grid[idx][col-1] = r[h];
      });
      // 배차 거래처의 기존 건 중 배차필요가 한 번도 정해지지 않은(빈칸) 건도 YES로 — 명시적 NO는 그대로 둠
      if(applyDispatchFlagToRow_(grid[idx], hmap, dispatchSet)) dispatchFlagged++;
      if(applyClientFlagToRow_(grid[idx], hmap, '프리알러트필요', preAlertSet)) preAlertFlagged++;
      applyPoNoteToRow_(grid[idx], hmap, poNoteSet);
      updated++;
    } else {
      var newArr = new Array(lastCol).fill('');
      r['생성일시'] = nowIso;
      if(!r['루틴']) r['루틴'] = '';
      r['청구전용'] = isBillingOnlyByNo_(no);
      r['캔슬'] = false;
      // 배차 거래처로 체크된 화주의 신규 건은 배차필요를 자동으로 YES — 이월·분할 청구 전용 건(-X01 등)은 배차 대상이 아니라 제외
      if(dispatchSet[String(r['거래처코드']||'').trim()] && !r['청구전용']) r['배차필요'] = true;
      if(preAlertSet[String(r['거래처코드']||'').trim()] && !r['청구전용']) r['프리알러트필요'] = true;
      Object.keys(r).forEach(function(h){
        if(h === 'checklist') return;
        var col = hmap[h]; if(!col) return;
        newArr[col-1] = r[h];
      });
      applyPoNoteToRow_(newArr, hmap, poNoteSet);
      if(r.checklist){
        Object.keys(r.checklist).forEach(function(k){
          var h = 'CHK_' + k;
          var col = hmap[h]; if(col) newArr[col-1] = chkWriteValue_(r.checklist[k]);
        });
      }
      grid.push(newArr);
      indexByNo[no] = grid.length - 1;
      created++;
    }
  });

  if(grid.length) sh.getRange(2,1,grid.length,lastCol).setValues(grid);
  return {created:created, updated:updated, dispatchFlagged:dispatchFlagged, preAlertFlagged:preAlertFlagged};
}

function upsertOne_(sh, hmap, id, data){
  var targetNo = id || data['관리번호'];
  var row = findRowByManagementNo_(sh, hmap, targetNo);
  data['수정일시'] = new Date().toISOString();
  if(row > 0){
    writeRow_(sh, hmap, row, data, {preserveProgress:false});
  } else {
    row = sh.getLastRow()+1;
    data['생성일시'] = data['수정일시'];
    writeRow_(sh, hmap, row, data, {preserveProgress:false});
  }
  return {row: row};
}

function toggleItem_(sh, hmap, managementNo, itemKey, value){
  var row = findRowByManagementNo_(sh, hmap, managementNo);
  if(row < 0) throw new Error('건을 찾을 수 없습니다: ' + managementNo);
  var h = 'CHK_' + itemKey;
  ensureCols_(sh, hmap, [h]);
  sh.getRange(row, hmap[h]).setValue(chkWriteValue_(value));
  sh.getRange(row, hmap['수정일시']).setValue(new Date().toISOString());
  return {row: row};
}

/* 일괄 체크 처리: 여러 건의 체크리스트 항목을 한 번에 완료 처리 (대시보드 "일괄 완료 처리" 기능).
   업로드와 동일한 방식으로 전체를 한 번에 읽고 메모리에서 처리한 뒤 한 번에 씀. */
function toggleMany_(sh, hmap, items, value){
  var lastRow = sh.getLastRow();
  if(lastRow < 2 || !items || !items.length) return {updated:0};
  var neededHeaders = {};
  items.forEach(function(it){ if(it && it.itemKey) neededHeaders['CHK_' + it.itemKey] = 1; });
  var missing = Object.keys(neededHeaders).filter(function(h){ return !hmap[h]; });
  if(missing.length) ensureCols_(sh, hmap, missing);

  var lastCol = sh.getLastColumn();
  var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var mgmtCol = hmap['관리번호'] - 1;
  var indexByNo = {};
  for(var i=0;i<grid.length;i++){
    var v = String(grid[i][mgmtCol]||'').trim();
    if(v) indexByNo[v] = i;
  }

  var nowIso = new Date().toISOString();
  var touchedRows = {};
  var updated = 0;
  items.forEach(function(it){
    if(!it || !it.itemKey) return;
    var no = String(it.managementNo||'').trim(); if(!no) return;
    var idx = indexByNo.hasOwnProperty(no) ? indexByNo[no] : -1;
    if(idx < 0) return;
    var col = hmap['CHK_' + it.itemKey]; if(!col) return;
    grid[idx][col-1] = chkWriteValue_(value);
    touchedRows[idx] = 1;
    updated++;
  });
  var updatedAtCol = hmap['수정일시'];
  if(updatedAtCol){ Object.keys(touchedRows).forEach(function(idx){ grid[idx][updatedAtCol-1] = nowIso; }); }
  if(updated > 0) sh.getRange(2,1,grid.length,lastCol).setValues(grid);
  return {updated: updated};
}

function deleteOne_(sh, hmap, managementNo){
  var row = findRowByManagementNo_(sh, hmap, managementNo);
  if(row < 0) return {deleted:false};
  sh.deleteRow(row);
  return {deleted:true};
}

/* ================= 내 거래처 목록 (설정 시트) ================= */
function getSettingsSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SETTINGS_SHEET_NAME);
    sh.getRange(1,1).setValue('거래처코드');
  }
  return sh;
}
function getMyClientCodes_(){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2,1,last-1,1).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var v = String(vals[i][0]||'').trim();
    if(v) out.push(v);
  }
  return out;
}
function setMyClientCodes_(codes){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if(last > 1) sh.getRange(2,1,last-1,1).clearContent();
  var clean = (codes||[]).map(function(c){ return String(c||'').trim(); }).filter(function(c){ return c; });
  if(clean.length) sh.getRange(2,1,clean.length,1).setValues(clean.map(function(c){ return [c]; }));
  return {codes: clean};
}

/* 배차 거래처의 "기존 건" 보정: 배차필요 칸이 빈칸(한 번도 정해진 적 없음)이고, 청구전용·캔슬이 아니며,
   ETD/ETA가 이미 지난 건(이미 출항·도착한 건)이 아니면 배차필요=TRUE로 바꿈. 명시적 FALSE(사용자가 NO로 둔 건)는 건드리지 않음.
   rowArr(시트 한 행 배열)를 직접 고치고, 바꿨으면 true 반환. */
function applyDispatchFlagToRow_(rowArr, hmap, dispatchSet){ return applyClientFlagToRow_(rowArr, hmap, '배차필요', dispatchSet); }
function applyClientFlagToRow_(rowArr, hmap, flagHeader, dispatchSet){
  var cCol = hmap['거래처코드'], dCol = hmap[flagHeader];
  if(!cCol || !dCol) return false;
  var code = String(rowArr[cCol-1]||'').trim();
  if(!code || !dispatchSet[code]) return false;
  var cur = rowArr[dCol-1];
  if(cur !== '' && cur !== null && cur !== undefined) return false; // 이미 TRUE/FALSE로 정해진 건
  if(hmap['청구전용'] && isTruthyCell_(rowArr[hmap['청구전용']-1])) return false;
  if(hmap['캔슬'] && isTruthyCell_(rowArr[hmap['캔슬']-1])) return false;
  var no = hmap['관리번호'] ? rowArr[hmap['관리번호']-1] : '';
  if(isBillingOnlyByNo_(no)) return false;
  var today = new Date(); today.setHours(0,0,0,0);
  var latest = null;
  ['ETD','ETA(PORT)'].forEach(function(h){
    if(!hmap[h]) return;
    var v = rowArr[hmap[h]-1]; if(v === '' || v === null || v === undefined) return;
    var d = new Date(v); if(isNaN(d.getTime())) return;
    if(latest === null || d.getTime() > latest) latest = d.getTime();
  });
  if(latest !== null && latest < today.getTime()) return false;
  rowArr[dCol-1] = true;
  return true;
}
function isTruthyCell_(v){ return v === true || String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'YES' || v === 1; }

/* 내 거래처 관리에서 배차 체크를 저장할 때, 이미 시트에 있는 해당 거래처 건에도 즉시 반영 */
function applyDispatchToExisting_(codes){ return applyClientFlagToExisting_('배차필요', codes); }
function applyClientFlagToExisting_(flagHeader, codes){
  var dispatchSet = {};
  (codes||[]).forEach(function(c){ c = String(c||'').trim(); if(c) dispatchSet[c] = true; });
  if(!Object.keys(dispatchSet).length) return 0;
  var sh = getSheet_();
  var hmap = headerMap_(sh);
  if(!hmap['거래처코드']) return 0;
  if(!hmap[flagHeader]) ensureCols_(sh, hmap, [flagHeader]);
  var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if(lastRow < 2) return 0;
  var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var n = 0;
  for(var i=0;i<grid.length;i++){ if(applyClientFlagToRow_(grid[i], hmap, flagHeader, dispatchSet)) n++; }
  if(n) sh.getRange(2,1,grid.length,lastCol).setValues(grid);
  return n;
}

/* ================= PO No → 비고 자동 입력 (설정 시트 G열) =================
   "내 거래처 관리"에서 PO→비고에 체크한 화주(예: 풍산 — 자기들 INV 번호를 PO No 칸에 적어 건을 구분)의 건은
   엑셀의 PO No를 '특이사항'(비고)에 자동으로 넣음. 비고가 비어 있으면 PO No로 채우고, 이미 다른 메모가 있으면
   지우지 않고 뒤에 " / PO No"를 덧붙임(이미 PO No가 들어 있으면 그대로 둠 — 재업로드해도 중복되지 않음). */
var PONOTE_CLIENT_COL_ = 7;
function getPoNoteClientCodes_(){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if(last < 2 || sh.getMaxColumns() < PONOTE_CLIENT_COL_) return [];
  var vals = sh.getRange(2, PONOTE_CLIENT_COL_, last-1, 1).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var v = String(vals[i][0]||'').trim();
    if(v) out.push(v);
  }
  return out;
}
function setPoNoteClientCodes_(codes){
  var sh = getSettingsSheet_();
  if(sh.getMaxColumns() < PONOTE_CLIENT_COL_) sh.insertColumnsAfter(sh.getMaxColumns(), PONOTE_CLIENT_COL_ - sh.getMaxColumns());
  if(sh.getRange(1, PONOTE_CLIENT_COL_).getValue() !== 'PO비고거래처코드') sh.getRange(1, PONOTE_CLIENT_COL_).setValue('PO비고거래처코드');
  var last = sh.getLastRow();
  if(last > 1) sh.getRange(2, PONOTE_CLIENT_COL_, last-1, 1).clearContent();
  var clean = (codes||[]).map(function(c){ return String(c||'').trim(); }).filter(function(c){ return c; });
  if(clean.length) sh.getRange(2, PONOTE_CLIENT_COL_, clean.length, 1).setValues(clean.map(function(c){ return [c]; }));
  return {codes: clean};
}
function applyPoNoteToRow_(rowArr, hmap, codeSet){
  if(!codeSet || !hmap['거래처코드'] || !hmap['PO No'] || !hmap['특이사항']) return false;
  var code = String(rowArr[hmap['거래처코드']-1]||'').trim();
  if(!code || !codeSet[code]) return false;
  var po = String(rowArr[hmap['PO No']-1]===undefined||rowArr[hmap['PO No']-1]===null ? '' : rowArr[hmap['PO No']-1]).trim();
  if(!po) return false;
  var cur = String(rowArr[hmap['특이사항']-1]===undefined||rowArr[hmap['특이사항']-1]===null ? '' : rowArr[hmap['특이사항']-1]).trim();
  if(cur.indexOf(po) >= 0) return false;
  rowArr[hmap['특이사항']-1] = cur ? (cur + ' / ' + po) : po;
  return true;
}
function applyPoNoteToExisting_(codes){
  var set = {};
  (codes||[]).forEach(function(c){ c = String(c||'').trim(); if(c) set[c] = true; });
  if(!Object.keys(set).length) return 0;
  var sh = getSheet_();
  var hmap = headerMap_(sh);
  if(!hmap['거래처코드'] || !hmap['PO No']) return 0; // PO No 열이 아직 없으면 적용할 값도 없음
  if(!hmap['특이사항']) ensureCols_(sh, hmap, ['특이사항']);
  var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
  if(lastRow < 2) return 0;
  var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var n = 0;
  for(var i=0;i<grid.length;i++){ if(applyPoNoteToRow_(grid[i], hmap, set)) n++; }
  if(n) sh.getRange(2,1,grid.length,lastCol).setValues(grid);
  return n;
}

/* ================= 배차 거래처 목록 (설정 시트 E열) =================
   "내 거래처 관리"에서 배차에 체크한 화주 — 엑셀 업로드로 새로 들어오는 건(upsertMany_의 신규 행)에
   '배차필요'를 자동으로 YES로 걸어 기존 배차 추적(오늘 할 일 > 배차)에 올림. 기존 건은 배차필요가 빈칸인 것만(명시적 NO·이미 출항/도착한 건 제외) 같이 YES로 바꿈.
   A열(거래처코드)·B~D열(라벨 재정의)과 별개 영역. */
var DISPATCH_CLIENT_COL_ = 5;
function getDispatchClientCodes_(){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  // 설정 시트에 E열이 아직 없으면(열이 5개 미만) 아직 저장된 배차 거래처가 없다는 뜻 — 오류 없이 빈 목록.
  // 조회(doGet) 전체가 이 부가 기능 때문에 실패하면 안 되므로 어떤 경우에도 예외를 밖으로 던지지 않음.
  if(last < 2 || sh.getMaxColumns() < DISPATCH_CLIENT_COL_) return [];
  var vals = sh.getRange(2, DISPATCH_CLIENT_COL_, last-1, 1).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var v = String(vals[i][0]||'').trim();
    if(v) out.push(v);
  }
  return out;
}
function setDispatchClientCodes_(codes){
  var sh = getSettingsSheet_();
  if(sh.getMaxColumns() < DISPATCH_CLIENT_COL_) sh.insertColumnsAfter(sh.getMaxColumns(), DISPATCH_CLIENT_COL_ - sh.getMaxColumns());
  if(sh.getRange(1, DISPATCH_CLIENT_COL_).getValue() !== '배차거래처코드') sh.getRange(1, DISPATCH_CLIENT_COL_).setValue('배차거래처코드');
  var last = sh.getLastRow();
  if(last > 1) sh.getRange(2, DISPATCH_CLIENT_COL_, last-1, 1).clearContent();
  var clean = (codes||[]).map(function(c){ return String(c||'').trim(); }).filter(function(c){ return c; });
  if(clean.length) sh.getRange(2, DISPATCH_CLIENT_COL_, clean.length, 1).setValues(clean.map(function(c){ return [c]; }));
  return {codes: clean};
}

/* ================= PRE-ALERT 거래처 목록 (설정 시트 F열) =================
   배차 거래처(E열)와 같은 방식 — "내 거래처 관리"에서 PRE-ALERT에 체크한 화주의 신규·기존(미지정) 건에
   '프리알러트필요'를 YES로 걸어 "오늘 할 일 > PRE-ALERT"에서 추적. */
var PREALERT_CLIENT_COL_ = 6;
function getPreAlertClientCodes_(){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if(last < 2 || sh.getMaxColumns() < PREALERT_CLIENT_COL_) return [];
  var vals = sh.getRange(2, PREALERT_CLIENT_COL_, last-1, 1).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var v = String(vals[i][0]||'').trim();
    if(v) out.push(v);
  }
  return out;
}
function setPreAlertClientCodes_(codes){
  var sh = getSettingsSheet_();
  if(sh.getMaxColumns() < PREALERT_CLIENT_COL_) sh.insertColumnsAfter(sh.getMaxColumns(), PREALERT_CLIENT_COL_ - sh.getMaxColumns());
  if(sh.getRange(1, PREALERT_CLIENT_COL_).getValue() !== '프리알러트거래처코드') sh.getRange(1, PREALERT_CLIENT_COL_).setValue('프리알러트거래처코드');
  var last = sh.getLastRow();
  if(last > 1) sh.getRange(2, PREALERT_CLIENT_COL_, last-1, 1).clearContent();
  var clean = (codes||[]).map(function(c){ return String(c||'').trim(); }).filter(function(c){ return c; });
  if(clean.length) sh.getRange(2, PREALERT_CLIENT_COL_, clean.length, 1).setValues(clean.map(function(c){ return [c]; }));
  return {codes: clean};
}

/* ================= 업무구분 항목 라벨 재정의 (설정 시트 B~D열) =================
   "거래처코드"(A열)와 별개 영역: B=루틴코드, C=항목키, D=변경된 라벨.
   프론트엔드의 ROUTES 기본 라벨을 덮어씀 — routeType+itemKey 조합별로 하나씩만 존재. */
function getItemLabelOverrides_(){
  var sh = getSettingsSheet_();
  var last = sh.getLastRow();
  if(last < 2) return [];
  var vals = sh.getRange(2,2,last-1,3).getValues(); // B,C,D
  var out = [];
  for(var i=0;i<vals.length;i++){
    var rt = String(vals[i][0]||'').trim();
    var key = String(vals[i][1]||'').trim();
    var label = String(vals[i][2]||'').trim();
    if(rt && key && label) out.push({routeType:rt, itemKey:key, label:label});
  }
  return out;
}
function setItemLabelOverrides_(overrides){
  var sh = getSettingsSheet_();
  if(sh.getRange(1,2).getValue() !== '루틴코드'){
    sh.getRange(1,2).setValue('루틴코드');
    sh.getRange(1,3).setValue('항목키');
    sh.getRange(1,4).setValue('변경된라벨');
  }
  var last = sh.getLastRow();
  if(last > 1) sh.getRange(2,2,last-1,3).clearContent();
  var clean = (overrides||[]).map(function(o){
    return {routeType:String(o.routeType||'').trim(), itemKey:String(o.itemKey||'').trim(), label:String(o.label||'').trim()};
  }).filter(function(o){ return o.routeType && o.itemKey && o.label; });
  if(clean.length){
    sh.getRange(2,2,clean.length,3).setValues(clean.map(function(o){ return [o.routeType, o.itemKey, o.label]; }));
  }
  return {overrides: clean};
}

/* ================= 부킹확인중 (임시 메모) =================
   정식 관리번호 체계(업무원본 시트)와 완전히 분리된 가벼운 메모 목록.
   부킹 컨펌 전, 대략의 정보로 이미 스케줄을 잡고 기다리는 건들을 적어두는 용도 —
   ALVIS 입력 후 정식 건(엑셀 업로드)으로 넘어가면 상태를 "확정전환됨"으로 바꾸거나 삭제. */
function getBookingHoldSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BOOKING_HOLD_SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(BOOKING_HOLD_SHEET_NAME);
    sh.getRange(1,1,1,BOOKING_HOLD_HEADERS.length).setValues([BOOKING_HOLD_HEADERS]);
  } else {
    // 예전 배포(POL/POD 등 항목별 입력칸)로 이미 만들어진 시트일 수 있으므로,
    // 지금 스키마(제목/내용 등)에 없는 헤더는 뒤에 추가해줌 — 기존 열은 그대로 두어 데이터 보존.
    var hmap = bookingHoldHeaderMap_(sh);
    var last = sh.getLastColumn();
    BOOKING_HOLD_HEADERS.forEach(function(h){
      if(!hmap[h]){ last += 1; sh.getRange(1,last).setValue(h); hmap[h] = last; }
    });
  }
  return sh;
}
function bookingHoldHeaderMap_(sh){
  var last = sh.getLastColumn();
  var headers = sh.getRange(1,1,1,last).getValues()[0];
  var map = {};
  for(var i=0;i<headers.length;i++){ if(headers[i]) map[String(headers[i]).trim()] = i+1; }
  return map;
}
function bookingHoldRowToObject_(hmap, values){
  var obj = {};
  Object.keys(hmap).forEach(function(h){
    var v = values[hmap[h]-1];
    if(BOOKING_HOLD_DATE_HEADERS.indexOf(h)>=0){
      obj[h] = (v instanceof Date) ? v.toISOString() : (v || '');
    } else {
      obj[h] = (v===undefined||v===null) ? '' : v;
    }
  });
  return obj;
}
function listBookingHolds_(){
  var sh = getBookingHoldSheet_();
  var hmap = bookingHoldHeaderMap_(sh);
  var last = sh.getLastRow();
  if(last < 2) return [];
  var lastCol = sh.getLastColumn();
  var grid = sh.getRange(2,1,last-1,lastCol).getValues();
  var idCol = hmap['ID'] - 1;
  var out = [];
  for(var i=0;i<grid.length;i++){
    if(!grid[i][idCol]) continue;
    out.push(bookingHoldRowToObject_(hmap, grid[i]));
  }
  return out;
}
function findBookingHoldRow_(sh, hmap, id){
  var col = hmap['ID'];
  if(!col) return -1;
  var last = sh.getLastRow();
  if(last < 2) return -1;
  var vals = sh.getRange(2, col, last-1, 1).getValues();
  for(var i=0;i<vals.length;i++){
    if(String(vals[i][0]).trim() === String(id).trim()) return i+2;
  }
  return -1;
}
function upsertBookingHold_(data){
  var sh = getBookingHoldSheet_();
  var hmap = bookingHoldHeaderMap_(sh);
  var nowIso = new Date().toISOString();
  var id = String(data.ID||'').trim();
  var isNew = !id;
  if(isNew) id = 'BH' + new Date().getTime() + Math.floor(Math.random()*1000);
  var rowIdx = isNew ? -1 : findBookingHoldRow_(sh, hmap, id);
  var rowArr = new Array(sh.getLastColumn()).fill('');
  if(rowIdx > 0){
    rowArr = sh.getRange(rowIdx,1,1,sh.getLastColumn()).getValues()[0];
  }
  BOOKING_HOLD_HEADERS.forEach(function(h){
    var col = hmap[h]; if(!col) return;
    if(h==='ID'){ rowArr[col-1] = id; return; }
    if(h==='생성일시'){ if(isNew) rowArr[col-1] = nowIso; return; }
    if(h==='수정일시'){ rowArr[col-1] = nowIso; return; }
    if(data.hasOwnProperty(h)) rowArr[col-1] = data[h];
  });
  if(!rowArr[hmap['상태']-1]) rowArr[hmap['상태']-1] = '대기중';
  if(rowIdx > 0){
    sh.getRange(rowIdx,1,1,rowArr.length).setValues([rowArr]);
  } else {
    sh.appendRow(rowArr);
  }
  return {id:id};
}
function deleteBookingHold_(id){
  var sh = getBookingHoldSheet_();
  var hmap = bookingHoldHeaderMap_(sh);
  var rowIdx = findBookingHoldRow_(sh, hmap, id);
  if(rowIdx > 0) sh.deleteRow(rowIdx);
  return {deleted: rowIdx>0};
}

/* ================= 오래된 완료건 자동 정리 =================
   "완료"는 저장된 날짜가 아니라 체크리스트가 전부 체크됐는지로 그때그때 판단되는 상태라서,
   정확한 "완료된 지 6개월"은 알 수 없음. 대신 "수정일시"(이미 있는 값 — 마지막으로 그 건이
   수정된 날짜)를 근사치로 사용: 완료된 건은 보통 그 이후로 다시 손댈 일이 없으므로 실질적으로
   거의 같은 의미. 대시보드(doGet)가 매번 "업무원본" 시트 전체를 읽어서 응답하기 때문에,
   오래된 완료건을 지우지 않고 "화면에서만 숨기는" 방식은 속도에 전혀 도움이 안 됨 — 반드시
   "업무원본" 시트 자체에서 빼내야 함. 완전 삭제 대신 "보관" 시트(별도 탭)로 옮기는 방식을
   쓰면, doGet이 "보관" 탭은 읽지 않으므로 완전 삭제와 속도 면에서 사실상 동일한 효과를
   내면서도 나중에 필요하면 "보관" 탭에서 다시 찾아볼 수 있음. */
function getArchiveSheet_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  var srcHeaders = headerMap_(getSheet_());
  var headerNames = Object.keys(srcHeaders).sort(function(a,b){return srcHeaders[a]-srcHeaders[b];});
  if(!sh){
    sh = ss.insertSheet(ARCHIVE_SHEET_NAME);
    sh.getRange(1,1,1,headerNames.length+1).setValues([headerNames.concat(['보관일시'])]);
  }
  return sh;
}
function dispatchItemKeyFor_(routeType){ return routeType==='AIR_IMP' ? 'impBaecha' : 'baecha'; }
/* PRE-ALERT는 2단계: (1) 아래 키 → (2) 공통 키 preAlertFinal(최종 발송). 항공 루틴의 (1)단계는 기존 PRE-ALERT 항목 키를 그대로 쓰고,
   그 외 루틴은 공통 키 preAlertDone을 씀. index.html의 preAlertItemKeyFor()와 동일해야 함. */
function preAlertItemKeyFor_(routeType){ return routeType==='AIR_EXP' ? 'preAlert' : (routeType==='AIR_IMP' ? 'impPreAlert' : 'preAlertDone'); }
function isChkTrue_(hmap, row, itemKey){
  var col = hmap['CHK_' + itemKey];
  if(!col) return false;
  var v = row[col-1];
  return v===true || v==='TRUE' || v===1;
}
var BL_HOLDING_FINAL_ITEM_KEY_ = 'blFinalIssue';
/* "미주"(ISF 대상) 여부는 DOOR/배차/비엘홀딩과 달리 사람이 직접 YES/NO로 지정하는 값이 아니라,
   루틴(해상수출·항공수출)과 POD(도착지) 값만으로 자동 판정함 — POD의 마지막 콤마 구분 항목이
   미국을 가리키면 자동으로 ISF 추적 대상이 됨. 단, 관리번호가 "-X01" 등으로 끝나는 이월·분할
   청구 전용 건(billingOnly)은 배차와 마찬가지로 제외함(호출부인 cleanupOldCompleted_에서 반영).
   프론트의 isUsPod()/isUsBoundRoute()/isUsBoundShipment()와 반드시 동일한 판정 기준을 유지해야 함. */
var US_BOUND_ROUTES_ = {SEA_EXP_FCL:1, SEA_EXP_LCL:1, AIR_EXP:1};
function isUsPod_(pod){
  var s = String(pod||'').trim();
  if(!s) return false;
  var parts = s.split(',');
  var country = parts[parts.length-1].trim().toUpperCase().replace(/\./g,'');
  return country==='USA' || country==='US' || country==='UNITED STATES' || country==='UNITED STATES OF AMERICA';
}
var US_ISF_ITEM_KEY_ = 'isfFiling';
/* 프론트의 isDone()과 반드시 동일한 기준이어야 함 — 여기서만 다르게 판단하면
   cleanupOldCompleted_(180일 경과 자동 보관)이 DOOR/배차/비엘홀딩/ISF 후속 확인이 아직 안
   끝난 건을 기본 체크리스트만 다 끝났다는 이유로 완료 취급해 조용히 "보관"으로 옮겨버릴 수 있음. */
function isRowDone_(hmap, row, routeType, billingOnly, doorEligible, dispatchNeeded, blHolding, usBound, preAlertNeeded){
  var items = relevantItemKeys_(routeType, billingOnly);
  if(items.length===0) return false; // 프론트의 isDone()과 동일: 해당 항목이 없으면 완료로 치지 않음
  var baseDone = items.every(function(it){ return isChkTrue_(hmap, row, it.key); });
  if(!baseDone) return false;
  if(doorEligible && !isChkTrue_(hmap, row, 'doorArrival')) return false;
  if(dispatchNeeded && !billingOnly && !isChkTrue_(hmap, row, dispatchItemKeyFor_(routeType))) return false;
  if(blHolding && !isChkTrue_(hmap, row, BL_HOLDING_FINAL_ITEM_KEY_)) return false;
  if(usBound && !isChkTrue_(hmap, row, US_ISF_ITEM_KEY_)) return false;
  if(preAlertNeeded && !billingOnly && !(isChkTrue_(hmap, row, preAlertItemKeyFor_(routeType)) && isChkTrue_(hmap, row, 'preAlertFinal'))) return false;
  return true;
}
/* cutoffDays: 이 일수보다 "수정일시"가 오래된 완료건을 정리 대상으로 봄 (기본 180일 ≈ 6개월).
   대량 처리 성능을 위해 전체를 한 번에 읽고(getValues), 정리 대상이 아닌 행만 모아 통째로
   다시 씀(setValues) — 옮겨지는 행은 "보관" 시트에 한 번에 append. */
function cleanupOldCompleted_(cutoffDays){
  cutoffDays = cutoffDays || 180;
  var sh = getSheet_();
  var hmap = headerMap_(sh);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if(lastRow < 2) return {archived:0, kept:0};
  var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var mgmtCol = hmap['관리번호']-1, routeCol = hmap['루틴']-1, billingCol = hmap['청구전용']-1, modCol = hmap['수정일시']-1;
  var doorCol = hmap['Door배송'] ? hmap['Door배송']-1 : -1;
  var dispatchCol = hmap['배차필요'] ? hmap['배차필요']-1 : -1;
  var preAlertCol = hmap['프리알러트필요'] ? hmap['프리알러트필요']-1 : -1;
  var blHoldingCol = hmap['비엘홀딩'] ? hmap['비엘홀딩']-1 : -1;
  var podCol = hmap['POD'] ? hmap['POD']-1 : -1;
  var cutoff = new Date(); cutoff.setDate(cutoff.getDate() - cutoffDays);
  var kept = [], toArchive = [];
  for(var i=0;i<grid.length;i++){
    var row = grid[i];
    if(!row[mgmtCol]){ continue; } // 빈 행은 건너뜀(결과에도 남기지 않음 — 어차피 실질 데이터 없음)
    var routeType = String(row[routeCol]||'').trim();
    var billingOnly = (row[billingCol]===true || row[billingCol]==='TRUE');
    var doorEligible = doorCol>=0 && (row[doorCol]===true || row[doorCol]==='TRUE');
    var dispatchNeeded = dispatchCol>=0 && (row[dispatchCol]===true || row[dispatchCol]==='TRUE');
    var preAlertNeeded = preAlertCol>=0 && (row[preAlertCol]===true || row[preAlertCol]==='TRUE');
    var blHolding = blHoldingCol>=0 && (row[blHoldingCol]===true || row[blHoldingCol]==='TRUE');
    // 이월·분할 청구 전용 건(-X01 등, billingOnly)은 배차와 마찬가지로 ISF 추적 대상에서 제외
    var usBound = !!US_BOUND_ROUTES_[routeType] && podCol>=0 && isUsPod_(row[podCol]) && !billingOnly;
    var modDate = (row[modCol] instanceof Date) ? row[modCol] : new Date(row[modCol]);
    var isOld = !isNaN(modDate.getTime()) && modDate < cutoff;
    if(isOld && isRowDone_(hmap, row, routeType, billingOnly, doorEligible, dispatchNeeded, blHolding, usBound, preAlertNeeded)){
      toArchive.push(row);
    } else {
      kept.push(row);
    }
  }
  if(toArchive.length===0) return {archived:0, kept:kept.length};
  var archiveSh = getArchiveSheet_();
  var nowIso = new Date().toISOString();
  var archiveRows = toArchive.map(function(row){ return row.concat([nowIso]); });
  archiveSh.getRange(archiveSh.getLastRow()+1, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
  // 남길 행만 다시 씀 — 지워진 만큼 시트 크기를 줄여 "업무원본"이 계속 가벼운 상태를 유지하게 함
  sh.getRange(2,1,lastRow-1,lastCol).clearContent();
  if(kept.length>0) sh.getRange(2,1,kept.length,lastCol).setValues(kept);
  return {archived:toArchive.length, kept:kept.length};
}
/* ================= 엑셀 업로드 시 "내 거래처 아닌 건" 자동 정리 =================
   여러 거래처 정보가 섞인 엑셀을 업로드해도, "설정" 시트에 등록된 "내 거래처" 코드에 해당하는
   건만 "업무원본"에 남기고 나머지는 매번 자동으로 "보관" 시트로 옮김 — 완전 삭제가 아니라
   보관 탭으로 옮기는 이유는 cleanupOldCompleted_와 동일: 거래처코드가 잘못 입력된 경우
   실수로 자기 건까지 날아가는 사고를 막기 위한 안전장치이며, doGet은 "보관" 탭을 읽지 않으므로
   속도 면에서는 완전 삭제와 사실상 동일한 효과를 냄.
   안전장치 두 가지:
   1) "내 거래처" 목록이 비어 있으면(아직 설정 전) 아무것도 지우지 않고 그대로 반환 — 목록
      미설정 상태에서 전체 데이터가 통째로 날아가는 사고를 방지.
   2) 거래처코드가 비어 있는 행은 "내 거래처가 아님"이 확실하지 않으므로(단순 입력 누락일 수 있음)
      보수적으로 남겨둠 — 확실히 다른 거래처코드가 찍혀 있는 행만 옮김. */
function purgeNonMyClients_(sh){
  var myCodes = getMyClientCodes_();
  if(!myCodes.length) return {purged:0, kept:0, skipped:'내 거래처 목록이 비어 있어 정리를 건너뜀'};
  var myCodeSet = {};
  myCodes.forEach(function(c){ myCodeSet[String(c).trim()] = 1; });

  var hmap = headerMap_(sh);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if(lastRow < 2) return {purged:0, kept:0};
  var mgmtCol = hmap['관리번호'] - 1;
  var codeCol = hmap['거래처코드'] ? hmap['거래처코드'] - 1 : -1;
  if(codeCol < 0) return {purged:0, kept:0, skipped:'거래처코드 열을 찾을 수 없어 정리를 건너뜀'};

  var grid = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  var kept = [], toPurge = [];
  for(var i=0;i<grid.length;i++){
    var row = grid[i];
    if(!row[mgmtCol]){ continue; } // 빈 행은 건너뜀(결과에도 남기지 않음)
    var code = String(row[codeCol]||'').trim();
    if(!code || myCodeSet[code]){ kept.push(row); } // 코드 미입력 건은 보수적으로 유지
    else { toPurge.push(row); }
  }
  if(toPurge.length===0) return {purged:0, kept:kept.length};

  var archiveSh = getArchiveSheet_();
  var nowIso = new Date().toISOString();
  var archiveRows = toPurge.map(function(row){ return row.concat([nowIso]); });
  archiveSh.getRange(archiveSh.getLastRow()+1, 1, archiveRows.length, archiveRows[0].length).setValues(archiveRows);
  sh.getRange(2,1,lastRow-1,lastCol).clearContent();
  if(kept.length>0) sh.getRange(2,1,kept.length,lastCol).setValues(kept);
  return {purged:toPurge.length, kept:kept.length};
}
/* 매일 자동 실행용 — Apps Script 편집기에서 setupDailyArchiveTrigger()를 딱 한 번 실행해두면,
   그 뒤로는 대시보드를 아무도 열지 않아도 매일 자동으로 이 함수가 실행됨. */
function dailyArchiveTrigger(){
  cleanupOldCompleted_(180);
}
/* Apps Script 편집기 상단 함수 선택 드롭다운에서 setupDailyArchiveTrigger를 고르고 ▶ 실행 버튼을
   한 번만 눌러주면 됨(최초 실행 시 권한 승인 필요) — 그러면 매일 자동으로 정리가 실행되도록
   트리거가 등록됨. 이미 등록돼 있으면 중복 등록하지 않도록 기존 것을 먼저 지움(재실행해도 안전). */
function setupDailyArchiveTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==='dailyArchiveTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyArchiveTrigger').timeBased().everyDays(1).atHour(3).create();
  return 'ok';
}
