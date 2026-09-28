/**
 * ═══════════════════════════════════════════════════════
 *  api.js — Apps Script Web App 연동 모듈
 *  index.html / dashboard.html / requests.html 에서 공통 사용
 * ═══════════════════════════════════════════════════════
 *
 *  [설정] Apps Script 배포 후 아래 URL을 교체하세요
 */

const API_URL = 'https://script.google.com/a/macros/buildersncompany.com/s/AKfycbzQ4Ba7iyTmA0soaJxP-nczm51Cay7fSfp7vTAXLsU5gctILsxu8fdjIw2Cgnu_frru/exec';

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 공통 fetch 래퍼
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function apiGet(params = {}) {
  const qs = new URLSearchParams(params).toString();
  const res = await fetch(`${API_URL}?${qs}`, { credentials: 'include' });
  const data = await res.json();
  if (!data.success) throw new Error(data.error || 'API 오류');
  return data;
}

async function apiPost(body = {}) {
  const res = await fetch(API_URL, {
    method: 'POST',
    credentials: 'include',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const data = await res.json();
  if (!data.success) throw new Error(data.error || 'API 오류');
  return data;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// API 함수들
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/** 요청서 제출 */
async function submitRequest(formData) {
  return apiPost({ action: 'submit', ...formData });
}

/** 요청서 목록 조회 */
async function fetchRequests(filters = {}) {
  return apiGet({ action: 'list', ...filters });
}

/** 요청서 상세 조회 */
async function fetchDetail(requestId) {
  return apiGet({ action: 'detail', id: requestId });
}

/** 상태/코멘트 업데이트 */
async function updateRequest(requestId, changes) {
  return apiPost({ action: 'update', request_id: requestId, ...changes });
}

/** 다음 Request ID 미리 채번 (수신 부서 선택 시 호출) */
async function fetchNextId(deptCode) {
  return apiGet({ action: 'seq', dept: deptCode });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 개발 모드 (API_URL 미설정 시 샘플 데이터 반환)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const IS_DEV = API_URL.includes('여기에');

if (IS_DEV) {
  console.warn('[api.js] 개발 모드: 샘플 데이터를 사용합니다. Apps Script 배포 후 API_URL을 교체하세요.');

  // 개발 중에는 목업 데이터 반환하도록 override
  window._mockRequests = [
    { request_id:'DS-2609-001', task_name:'10월 신제품 런칭 키비주얼 제작', from_dept:'MK', from_name:'홍길동', to_dept:'DS', to_name:'김디자인', priority:'urgent', due_date:'2026-09-30', status:'wip',    created_at:'2026-09-01' },
    { request_id:'DS-2609-002', task_name:'회사 소개 브로셔 리뉴얼',        from_dept:'BS', from_name:'이전략', to_dept:'DS', to_name:'김디자인', priority:'high',   due_date:'2026-10-05', status:'review', created_at:'2026-09-05' },
    { request_id:'DS-2609-003', task_name:'SNS 카드뉴스 템플릿 3종',        from_dept:'MK', from_name:'박마케팅', to_dept:'DS', to_name:'',         priority:'normal', due_date:'2026-10-10', status:'new',    created_at:'2026-09-10' },
    { request_id:'MK-2609-001', task_name:'Q4 이메일 뉴스레터 기획',        from_dept:'SL', from_name:'최영업', to_dept:'MK', to_name:'이마케팅', priority:'normal', due_date:'2026-10-01', status:'wip',    created_at:'2026-09-03' },
    { request_id:'MK-2609-002', task_name:'신규 고객 온보딩 콘텐츠 제작',   from_dept:'CX', from_name:'정경험', to_dept:'MK', to_name:'이마케팅', priority:'high',   due_date:'2026-09-28', status:'wip',    created_at:'2026-09-07' },
  ];
}
