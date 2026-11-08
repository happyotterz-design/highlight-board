/* 사건파일: 수상한 보고서 · 공통 */
const API_URL = 'https://script.google.com/macros/s/AKfycbyyk05Wu6D3DUrO4F8wzG2DGzP9UUEnU8Bmi8qERaXupv2Cv9jutUMPe6v3zq-4XLIZag/exec';

function call(fn, ...args) {
  return fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ fn, args }) })
    .then(r => r.json(), () => { throw '연결이 잠깐 끊겼어요. 다시 해 주세요.'; })
    .then(j => { if (!j.ok) throw j.error; return j.data; });
}

function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
const $ = id => document.getElementById(id);
const ROLE_COLORS = ['#2a5db0', '#7b3fa0', '#b8651a', '#2c8a4b', '#b0405e', '#4f7f8a'];
function roleColor(roles, r) { return ROLE_COLORS[Math.max(0, roles.indexOf(r)) % ROLE_COLORS.length]; }
function store(k, v) { try { if (v === undefined) return JSON.parse(localStorage.getItem(k) || 'null'); localStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
function sortKeys(arr) { return arr.slice().sort((x, y) => { const [a, b] = x.split('-').map(Number), [c, d] = y.split('-').map(Number); return a - c || b - d; }); }

const STORY = {
  code: 'CASE 2026-OSAEK',
  title: '사건파일: 수상한 보고서',
  lead: [
    '내일 오전, 국립공원위원회에 「오색 케이블카 검토 보고서」가 제출된다. 작성자는 환경공학자다.',
    '그런데 어젯밤 편집국에 익명 제보가 들어왔다. "이 보고서에는 결함이 8개 숨어 있다. 이대로 제출되면 위원회는 잘못된 판단을 내릴 것이다."',
    '까칠한 기자단 수사팀, 출동이다. 자료 상자를 단서 삼아 결함을 찾고, 찾을 때마다 열리는 「수정 단서」를 모아 보고서를 다시 써서 복원하라.'
  ],
  missions: [
    ['결함 수사', '자료 상자와 보고서를 비교하며 숨은 결함 8개를 찾는다. 문장을 고르고 결함 이름을 맞힌다. 고치는 건 다음 라운드!'],
    ['보고서 복원', '고칠 문장을 하나씩 고치고 바꾼 이유를 적는다. 역할마다 맡은 문장이 있다.'],
    ['상호 심사', '다른 수사팀이 바꾼 문장과 이유를 보고 타당한지 판단한다.'],
    ['베스트 투표', '구간마다 가장 잘 고친 글에 투표해 우리 반 보고서를 완성한다.']
  ],
  points: [
    ['결함 발견 (문장 + 결함 이름)', '+3'],
    ['문장만 맞힘 (결함마다 기회 2번)', '+1'],
    ['내 역할 담당 결함을 내가 발견', '+1'],
    ['상호 심사에서 받은 \'타당해요\' 1개', '+1'],
    ['구간 베스트 문장 (투표 1등, 공동 1등도)', '+3'],
    ['1라운드 힌트', '−1'],
    ['2라운드 문장 도움 ① 방향 보기', '−1'],
    ['2라운드 문장 도움 ② 예시 문장까지', '총 −2'],
    ['예시 문장을 열고 80% 이상 겹치게 쓴 글', '베스트 +3 없음']
  ],
  note: '명탐정 기자 점수는 우리 모둠이 명탐정 기자단이 되기 위한 게임 점수예요. 수행평가 성적과는 관계없어요. 점수와 순위는 마지막에 공개해요.'
};

/* 투표 후보 순서: 반·구간마다 무작위로 섞되, 같은 반 모든 기기와 칠판에서 같은 순서 */
function shuffleFor(arr, seed) {
  let h = 2166136261; for (const ch of seed) { h ^= ch.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; }
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) { h = (Math.imul(h, 1103515245) + 12345) >>> 0; const j = h % (i + 1); [a[i], a[j]] = [a[j], a[i]]; }
  return a;
}

/* 결말: 브리핑의 익명 제보와 이어지는 사건 종결 기사 */
function caseClosed(R, cls) {
  const total = R.caught.length, found = R.caught.filter(f => f.groups).length;
  const top = R.rank[0];
  const end = R.restored.find(s => s.name === '결론');
  const changed = R.restored.filter(s => s.by).length;
  return {
    kicker: '까칠한 기자단 · 사건 종결 보도',
    headline: `[단독] ${cls ? cls + ' ' : ''}기자단, '수상한 보고서' 결함 ${found === total ? total + '개 모두' : found + '개'} 밝혀내`,
    body: [
      `익명 제보로 시작된 수사가 끝났다. 국립공원위원회 제출을 하루 앞둔 「오색 케이블카 검토 보고서」에서 기자단은 ${found === total ? `숨은 결함 ${total}개를 모두` : `숨은 결함 ${total}개 가운데 ${found}개를`} 찾아냈다.`,
      `수사팀은 결함을 하나씩 고쳐 보고서를 복원했고, 서로의 수정을 심사한 뒤 구간마다 가장 설득력 있는 글을 골라 '우리 반 보고서'를 완성했다. ${changed}개 구간이 새로 쓰였고, 모범 답안과 비교한 완성도는 ${R.overallBefore}%에서 ${R.overall}%로 올랐다.`,
      end && end.by ? `복원된 보고서는 이렇게 끝난다. "${end.text}"` : '',
      '수정된 보고서는 내일 오전 국립공원위원회에 다시 제출된다. 위원회가 잘못된 판단을 내릴 뻔한 위기는 기자단의 손에서 막혔다.'
    ].filter(Boolean),
    credit: top ? `최우수 수사팀 ${top.group}모둠 (${top.grade})` : ''
  };
}
