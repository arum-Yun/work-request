/**
 * ═══════════════════════════════════════════════════════
 *  업무 요청서 시스템 — Google Apps Script (Backend)
 *  도메인: @thebuilders.co.kr
 * ═══════════════════════════════════════════════════════
 *
 *  [설정 방법]
 *  1. Google Drive에서 새 Google Spreadsheet 생성
 *  2. 메뉴 → 확장 프로그램 → Apps Script
 *  3. 이 코드 전체를 붙여넣기
 *  4. CONFIG 섹션의 값들을 실제 값으로 교체
 *  5. 메뉴 → 배포 → 새 배포 → 웹 앱으로 배포
 *     - 실행 계정: 나(관리자 계정)
 *     - 액세스: thebuilders.co.kr 조직 내 사용자만
 *  6. 배포 URL을 index.html / dashboard.html의 API_URL에 입력
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// CONFIG — 실제 값으로 교체 필요
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const CONFIG = {
  DOMAIN:           'buildersncompany.com',        // 허용 도메인 (구글 + 슬랙 공통)
  SHEET_REQUESTS:   'requests',                    // 요청서 데이터 시트명
  SHEET_DEPT:       'dept_codes',                  // 부서 코드 시트명
  SHEET_SEQ:        'seq_counter',                 // 채번 시트명
  SLACK_BOT_TOKEN:  'xoxb-여기에-슬랙-봇-토큰-입력', // Slack Bot Token
};

// 부서 코드 테이블 (dept_codes 시트와 동기화)
const DEPARTMENTS = [
  { name: '영업팀',     code: 'SL' },
  { name: '부동산팀',   code: 'RE' },
  { name: '운영기획팀', code: 'OP' },
  { name: '고객경험팀', code: 'CX' },
  { name: '사업전략팀', code: 'BS' },
  { name: '디자인팀',   code: 'DS' },
  { name: '마케팅팀',   code: 'MK' },
  { name: '백본팀',     code: 'BB' },
];


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 보안: 도메인 검증
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function checkAuth() {
  const email = Session.getActiveUser().getEmail();
  if (!email.endsWith('@' + CONFIG.DOMAIN)) {
    throw new Error('Unauthorized: ' + email);
  }
  return email;
}

function corsHeaders() {
  return ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// GET 라우터 — 목록 조회 / 채번
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function doGet(e) {
  // CORS preflight 대응
  const output = handleRequest(e, 'GET');
  return output;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// POST 라우터 — 제출 / 업데이트
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function doPost(e) {
  return handleRequest(e, 'POST');
}

function handleRequest(e, method) {
  try {
    // 도메인 검증 — @buildersncompany.com 계정만 허용
    const userEmail = Session.getActiveUser().getEmail();
    if (userEmail && !userEmail.endsWith('@' + CONFIG.DOMAIN)) {
      return jsonResponse({ error: 'Unauthorized: ' + userEmail });
    }

    if (method === 'GET') {
      const action = e.parameter.action || 'list';
      if (action === 'list')   return handleList(e.parameter, userEmail);
      if (action === 'seq')    return handleSeq(e.parameter);
      if (action === 'detail') return handleDetail(e.parameter.id);
      return jsonResponse({ error: 'Unknown action' });
    }

    if (method === 'POST') {
      const body   = JSON.parse(e.postData.contents);
      const action = body.action || 'submit';
      if (action === 'submit') return handleSubmit(body, userEmail);
      if (action === 'update') return handleUpdate(body, userEmail);
      return jsonResponse({ error: 'Unknown action' });
    }

  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 1) 요청서 제출
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function handleSubmit(data, userEmail) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);

  // Request ID 채번
  const requestId = generateRequestId(data.to_dept);

  const now = new Date();
  const row = [
    requestId,                        // A: request_id
    formatDate(now),                  // B: created_at
    data.due_date     || '',          // C: due_date
    data.from_dept    || '',          // D: from_dept
    data.from_name    || '',          // E: from_name
    userEmail,                        // F: from_email (자동 — 로그인 계정)
    data.to_dept      || '',          // G: to_dept
    data.to_name      || '',          // H: to_name
    data.task_name    || '',          // I: task_name
    data.priority     || 'normal',    // J: priority
    'new',                            // K: status (접수 상태로 시작)
    data.purpose      || '',          // L: purpose
    data.condition    || '',          // M: condition
    data.constraint   || '',          // N: constraint
    data.promise_1    || '',          // O: promise_1
    data.promise_2    || '',          // P: promise_2
    data.promise_3    || '',          // Q: promise_3
    (data.cc || []).join(','),        // R: cc
    '',                               // S: handler_comment (빈칸)
    formatDate(now),                  // T: updated_at
  ];

  sheet.appendRow(row);

  // 슬랙 DM 알림 — 수신 담당자에게
  if (data.to_name && data.to_dept) {
    const toEmail = resolveEmail(data.to_name, data.to_dept);
    if (toEmail) {
      sendSlackDM(toEmail, buildNewRequestMessage(requestId, data));
    }
  }

  // 슬랙 DM 알림 — 참조자에게
  if (data.cc && data.cc.length > 0) {
    data.cc.forEach(ccName => {
      // CC는 이름만 있으므로 이메일 매핑 필요 (아래 resolveEmail 참고)
      const ccEmail = resolveEmailByName(ccName);
      if (ccEmail) {
        sendSlackDM(ccEmail, buildCCMessage(requestId, data));
      }
    });
  }

  return jsonResponse({ success: true, request_id: requestId });
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 2) 상태/코멘트 업데이트
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function handleUpdate(data, userEmail) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  const rows  = sheet.getDataRange().getValues();

  // request_id로 행 찾기 (1행 = 헤더)
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.request_id) {
      const rowNum = i + 1;

      if (data.status !== undefined) {
        sheet.getRange(rowNum, 11).setValue(data.status);  // K열: status

        // 상태 변경 시 요청자에게 DM 알림
        const fromEmail = rows[i][5]; // F열: from_email
        if (fromEmail) {
          sendSlackDM(fromEmail, buildStatusUpdateMessage(data.request_id, data.status, rows[i][8]));
        }
      }
      if (data.handler_comment !== undefined) {
        sheet.getRange(rowNum, 19).setValue(data.handler_comment); // S열
      }
      if (data.to_name !== undefined) {
        sheet.getRange(rowNum, 9).setValue(data.to_name);  // H열: to_name
      }

      // updated_at 갱신
      sheet.getRange(rowNum, 20).setValue(formatDate(new Date())); // T열

      return jsonResponse({ success: true });
    }
  }
  return jsonResponse({ error: 'Request not found: ' + data.request_id });
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 3) 요청서 목록 조회
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function handleList(params, userEmail) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  const rows  = sheet.getDataRange().getValues();

  // 헤더 제외
  let data = rows.slice(1).map(r => ({
    request_id:      r[0],
    created_at:      r[1],
    due_date:        r[2],
    from_dept:       r[3],
    from_name:       r[4],
    from_email:      r[5],
    to_dept:         r[6],
    to_name:         r[7],
    task_name:       r[8],
    priority:        r[9],
    status:          r[10],
    purpose:         r[11],
    condition:       r[12],
    constraint:      r[13],
    promise_1:       r[14],
    promise_2:       r[15],
    promise_3:       r[16],
    cc:              r[17] ? r[17].split(',') : [],
    handler_comment: r[18],
    updated_at:      r[19],
  }));

  // 필터 적용
  if (params.status)   data = data.filter(r => r.status   === params.status);
  if (params.to_dept)  data = data.filter(r => r.to_dept  === params.to_dept);
  if (params.priority) data = data.filter(r => r.priority === params.priority);

  // 내가 보낸 요청
  if (params.view === 'mine') {
    data = data.filter(r => r.from_email === userEmail);
  }
  // 내가 받은 요청 (담당자 이메일 기준)
  if (params.view === 'received') {
    data = data.filter(r => resolveEmail(r.to_name, r.to_dept) === userEmail);
  }

  return jsonResponse({ success: true, data: data, total: data.length });
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 4) 요청서 상세 조회
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function handleDetail(requestId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === requestId) {
      const r = rows[i];
      return jsonResponse({
        success: true,
        data: {
          request_id: r[0], created_at: r[1], due_date: r[2],
          from_dept: r[3], from_name: r[4], from_email: r[5],
          to_dept: r[6], to_name: r[7], task_name: r[8],
          priority: r[9], status: r[10],
          purpose: r[11], condition: r[12], constraint: r[13],
          promise_1: r[14], promise_2: r[15], promise_3: r[16],
          cc: r[17] ? r[17].split(',') : [],
          handler_comment: r[18], updated_at: r[19],
        }
      });
    }
  }
  return jsonResponse({ error: 'Not found' });
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Request ID 채번
// 포맷: [수신부서코드]-[YYMM]-[순번3자리]
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function handleSeq(params) {
  const nextId = generateRequestId(params.dept);
  return jsonResponse({ success: true, request_id: nextId });
}

function generateRequestId(deptCode) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(CONFIG.SHEET_SEQ);
  const now      = new Date();
  const yy       = String(now.getFullYear()).slice(2);
  const mm       = String(now.getMonth() + 1).padStart(2, '0');
  const seqKey   = `${deptCode}_${yy}${mm}`;
  const rows     = sheet.getDataRange().getValues();

  // 기존 키 찾기
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === seqKey) {
      const next = rows[i][1] + 1;
      sheet.getRange(i + 1, 2).setValue(next);
      return `${deptCode}-${yy}${mm}-${String(next).padStart(3, '0')}`;
    }
  }

  // 없으면 새로 추가 (월 첫 요청)
  sheet.appendRow([seqKey, 1]);
  return `${deptCode}-${yy}${mm}-001`;
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 슬랙 DM 발송
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 이메일 주소로 슬랙 유저 ID를 조회한 뒤 DM을 발송
 * @param {string} email  - @thebuilders.co.kr 이메일
 * @param {Object} blocks - Slack Block Kit 메시지
 */
function sendSlackDM(email, blocks) {
  try {
    // 1) 이메일 → Slack User ID 조회
    const lookupRes = UrlFetchApp.fetch(
      'https://slack.com/api/users.lookupByEmail?email=' + encodeURIComponent(email),
      {
        method: 'get',
        headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
        muteHttpExceptions: true,
      }
    );
    const lookupData = JSON.parse(lookupRes.getContentText());
    if (!lookupData.ok) {
      Logger.log('Slack lookup failed for ' + email + ': ' + lookupData.error);
      return;
    }
    const userId = lookupData.user.id;

    // 2) DM 채널 열기 (또는 기존 채널 재사용)
    const openRes = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
      payload: JSON.stringify({ users: userId }),
      muteHttpExceptions: true,
    });
    const openData = JSON.parse(openRes.getContentText());
    if (!openData.ok) {
      Logger.log('Slack open DM failed: ' + openData.error);
      return;
    }
    const channelId = openData.channel.id;

    // 3) 메시지 발송
    UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + CONFIG.SLACK_BOT_TOKEN },
      payload: JSON.stringify({
        channel: channelId,
        blocks:  blocks,
        text:    '새로운 업무 요청이 도착했습니다.',  // 알림 fallback 텍스트
      }),
      muteHttpExceptions: true,
    });
  } catch (err) {
    Logger.log('sendSlackDM error: ' + err.message);
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 슬랙 메시지 템플릿 (Block Kit)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/** 새 요청서 수신 알림 (담당자 → DM) */
function buildNewRequestMessage(requestId, data) {
  const priorityEmoji = { urgent:'🔴', high:'🟠', normal:'🔵', low:'🟢' };
  const deptName = d => DEPARTMENTS.find(x => x.code === d)?.name || d;
  return [
    {
      type: 'header',
      text: { type: 'plain_text', text: '📬 새로운 업무 요청이 도착했습니다' }
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: `*Request ID*\n\`${requestId}\`` },
        { type: 'mrkdwn', text: `*우선순위*\n${priorityEmoji[data.priority] || '🔵'} ${data.priority}` },
        { type: 'mrkdwn', text: `*요청 팀*\n${deptName(data.from_dept)} · ${data.from_name}` },
        { type: 'mrkdwn', text: `*희망 완료일*\n${data.due_date || '미정'}` },
      ]
    },
    {
      type: 'section',
      text: { type: 'mrkdwn', text: `*작업명*\n${data.task_name}` }
    },
    {
      type: 'section',
      text: { type: 'mrkdwn', text: `*요청 목적*\n${data.purpose || '—'}` }
    },
    { type: 'divider' },
    {
      type: 'context',
      elements: [{ type: 'mrkdwn', text: `요청서를 확인하고 진행 상태를 업데이트해 주세요.` }]
    }
  ];
}

/** 상태 변경 알림 (요청자 → DM) */
function buildStatusUpdateMessage(requestId, newStatus, taskName) {
  const statusLabel = { new:'접수', review:'검토 중', wip:'진행 중', done:'✅ 완료', rejected:'❌ 반려' };
  return [
    {
      type: 'header',
      text: { type: 'plain_text', text: '🔄 업무 요청 상태가 변경되었습니다' }
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: `*Request ID*\n\`${requestId}\`` },
        { type: 'mrkdwn', text: `*변경된 상태*\n${statusLabel[newStatus] || newStatus}` },
      ]
    },
    {
      type: 'section',
      text: { type: 'mrkdwn', text: `*작업명*\n${taskName}` }
    },
  ];
}

/** 참조자 알림 */
function buildCCMessage(requestId, data) {
  return [
    {
      type: 'header',
      text: { type: 'plain_text', text: '👀 업무 요청서에 참조되었습니다' }
    },
    {
      type: 'section',
      fields: [
        { type: 'mrkdwn', text: `*Request ID*\n\`${requestId}\`` },
        { type: 'mrkdwn', text: `*작업명*\n${data.task_name}` },
      ]
    },
  ];
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 유틸리티
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

/**
 * 이름 → 이메일 변환
 * Google Workspace Directory API로 @buildersncompany.com 도메인에서 조회
 */
function resolveEmail(name, deptCode) {
  try {
    const results = AdminDirectory.Users.list({
      domain: CONFIG.DOMAIN,  // buildersncompany.com 단일 도메인
      query:  `name="${name}"`,
      maxResults: 1,
    });
    if (results.users && results.users.length > 0) {
      return results.users[0].primaryEmail;
    }
    Logger.log('유저를 찾을 수 없음: ' + name);
  } catch (e) {
    Logger.log('resolveEmail error: ' + e.message);
  }
  return null;
}

function resolveEmailByName(name) {
  return resolveEmail(name, null);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 초기 시트 세팅 (최초 1회 실행)
// Apps Script 편집기에서 직접 실행하세요
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) requests 시트
  let req = ss.getSheetByName(CONFIG.SHEET_REQUESTS);
  if (!req) req = ss.insertSheet(CONFIG.SHEET_REQUESTS);
  req.clearContents();
  req.getRange(1, 1, 1, 20).setValues([[
    'request_id','created_at','due_date',
    'from_dept','from_name','from_email',
    'to_dept','to_name','task_name',
    'priority','status',
    'purpose','condition','constraint',
    'promise_1','promise_2','promise_3',
    'cc','handler_comment','updated_at'
  ]]);
  req.getRange(1,1,1,20).setFontWeight('bold').setBackground('#4F46E5').setFontColor('#ffffff');
  req.setFrozenRows(1);

  // 2) dept_codes 시트
  let dept = ss.getSheetByName(CONFIG.SHEET_DEPT);
  if (!dept) dept = ss.insertSheet(CONFIG.SHEET_DEPT);
  dept.clearContents();
  dept.getRange(1,1,1,2).setValues([['code','name']]);
  dept.getRange(1,1,1,2).setFontWeight('bold').setBackground('#4F46E5').setFontColor('#ffffff');
  DEPARTMENTS.forEach((d, i) => {
    dept.getRange(i+2, 1, 1, 2).setValues([[d.code, d.name]]);
  });

  // 3) seq_counter 시트
  let seq = ss.getSheetByName(CONFIG.SHEET_SEQ);
  if (!seq) seq = ss.insertSheet(CONFIG.SHEET_SEQ);
  seq.clearContents();
  seq.getRange(1,1,1,2).setValues([['seq_key','current_count']]);
  seq.getRange(1,1,1,2).setFontWeight('bold').setBackground('#4F46E5').setFontColor('#ffffff');

  SpreadsheetApp.getUi().alert('✅ 시트 설정 완료! requests / dept_codes / seq_counter 시트가 생성되었습니다.');
}
