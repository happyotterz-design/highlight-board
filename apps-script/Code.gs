/**
 * 사건파일: 수상한 보고서 (까칠한 기자단 게임) · 서버
 *
 * 학생 화면:  https://happyotterz-design.github.io/highlight-board/
 * 칠판 화면:  https://happyotterz-design.github.io/highlight-board/board.html
 *
 * 처음 한 번
 *  1. 이 코드를 Code.gs에 통째로 붙여 넣고 저장
 *  2. 함수 목록에서 setupGame 을 골라 [실행] → 시트 탭이 만들어짐
 *  3. 「설정」 탭의 교사 비밀번호를 바꿔 두기
 *  4. 배포 → 배포 관리 → ✎ 수정 → 버전: 새 버전 → 배포  (주소 유지)
 *
 * 다른 반 수업 전: 시트 메뉴 [사건파일 → 새 게임 시작] (기록 탭이 보관 탭으로 복사된 뒤 비워짐)
 *
 * 시트
 *  본문  : 구간 / 내용 (예시 글)
 *  설정  : 제목, 모둠 수, 역할, 교사 비밀번호
 *  결함  : 숨은 결함 정답표 (학생 화면으로는 보내지 않음)
 *  자료  : 자료 상자 (증거물)
 *  단서  : 역할별 비밀 단서
 *  복원구역 : 모둠별로 고쳐 쓸 구간
 *  기록  : 학생 제출이 쌓이는 곳 (숨김 칸에 아무 글자나 적으면 점수·화면에서 빠짐)
 */

const SH = { PHRASE: '표현카드', TEXT: '본문', CONF: '설정', FLAW: '결함', DATA: '자료', CLUE: '단서', ZONE: '복원구역', LOG: '기록' };
const LOG_HEAD = ['id', '시각', '종류', '모둠', '역할', '문장번호', '고른 문장', '결함 카드', '문제점', '이렇게 바꾸자', '판정', '대상', '숨김', '반', '이름'];
const PHASES = ['브리핑', '1라운드 결함 수사', '2라운드 보고서 복원', '3라운드 상호 심사', '4라운드 내 초안 수사', '결과 공개'];
const CARDS = [
  ['찬반형 질문', '탐구 질문이 예/아니오로 끝난다'],
  ['위치 없는 출처', '자료 번호는 있지만 몇 번째 문단인지 없다'],
  ['사실만 반복하는 해석', '해석 문장이 자료 내용을 다시 말할 뿐이다'],
  ['믿을 수 없는 자료', '블로그, "~라고 한다"처럼 인정되지 않는 자료다'],
  ['전문가 관점 없음', '맡은 전문가라면 따질 내용이 빠져 있다'],
  ['한계 회피', '내 주장의 진짜 한계를 피하고 다른 이야기를 한다'],
  ['추상적 대안', '누가, 얼마나, 어기면 어떻게 되는지가 없다'],
  ['다짐형 결론', '조건 없이 "노력해야 한다"로 끝난다']
];

/* ================= 처음 설정 ================= */
function setupGame() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const make = (name, rows, widths) => {
    let s = ss.getSheetByName(name);
    if (s && s.getLastRow() > 0) return s;
    s = s || ss.insertSheet(name);
    s.getRange(1, 1, Math.max(rows.length, 300), rows[0].length).setNumberFormat('@');   // "1-1"이 날짜로 바뀌지 않게 글자 형식
    s.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    s.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
    s.setFrozenRows(1);
    (widths || []).forEach((w, i) => s.setColumnWidth(i + 1, w));
    s.getDataRange().setWrap(true).setVerticalAlignment('top');
    return s;
  };

  make(SH.TEXT, [
    ['구간', '내용'],
    ['탐구 질문', '환경공학자의 시선에서, 설악산 오색 케이블카는 설치되어야 하는가?'],
    ['서론', '설악산 오색 케이블카는 40년 넘게 논란이 되어 온 사업이다. 양양군에 따르면 케이블카가 설치되면 연간 방문객이 지금보다 크게 늘어나고, 지역 경제 효과가 1,000억 원이 넘는다고 한다(자료 1). 이처럼 케이블카는 오랫동안 인구가 줄어 온 양양 지역의 경제를 살릴 수 있는 좋은 기회이다. 이 자료는 내 주장에서 [케이블카가 지역 경제에 도움이 된다]라는 의미를 갖는다. 따라서 나는 환경공학자로서 설악산 케이블카는 설치되어야 한다고 주장한다.'],
    ['본론 1', '케이블카가 필요한 가장 큰 이유는 접근성이다. 현재 오색에서 대청봉까지는 걸어서 4~5시간이 걸리기 때문에 노약자나 장애인은 정상에 오르기 어렵다. 케이블카가 생기면 누구나 설악산의 풍경을 즐길 수 있게 된다. 어떤 여행 블로그에 따르면 케이블카 설치를 원하는 사람이 70%나 된다고 한다. 많은 사람이 원하는 만큼 케이블카 설치는 꼭 필요하다.'],
    ['본론 2', '물론 케이블카에도 문제점은 있다. 공사를 하는 동안 소음과 먼지가 생기고, 등산로 일부가 막혀 등산객이 불편할 수 있다(자료 3 · 2번째 문단). 이 자료는 내 주장에서 [공사 기간에 불편이 생긴다]라는 의미를 갖는다. 하지만 공사는 몇 년이면 끝나기 때문에 큰 문제가 아니다.'],
    ['본론 3', '이런 문제를 해결하기 위해 세 가지 대안을 제안한다. 첫째, 환경부는 케이블카 주변에 사는 산양을 보호하기 위해 적극적으로 노력해야 한다. 둘째, 양양군은 케이블카 이용 요금의 일부를 환경 보호에 쓴다(자료 4 · 3번째 문단). 이 자료는 내 주장에서 [요금을 환경 보호에 쓸 수 있다]라는 의미를 갖는다. 셋째, 등산객들이 자연을 아끼는 마음을 갖도록 캠페인을 펼친다.'],
    ['결론', '결론적으로 설악산 케이블카는 지역 경제와 접근성을 위해 꼭 필요하다. 앞으로 우리 모두가 자연을 보호하면서 개발할 수 있도록 노력해야 할 것이다.']
  ], [90, 700]);

  make(SH.CONF, [
    ['항목', '값'],
    ['글 제목', '오색 케이블카 검토 보고서'],
    ['모둠 수', 4],
    ['역할 (쉼표로 구분)', '팩트체크,윤리·공정,현실·실행,환경·미래'],
    ['교사 비밀번호', '1234'],
    ['힌트 감점', 1],
    ['반 목록 (쉼표로 구분)', '2-1반,2-2반,2-3반,2-4반,2-5반,2-6반']
  ], [160, 420]);

  // 문장번호: "구간순서-문장순서" (0부터). 본문을 고치면 여기 번호도 맞춰야 함.
  make(SH.FLAW, [
    ['코드', '구간', '문장번호', '결함 카드', '담당 역할', '증거물', '해설', '예시 단어'],
    ['A1', '탐구 질문', '0-0', '찬반형 질문', '윤리·공정', '자료 5', '"설치되어야 하는가?"는 예/아니오로 끝난다. 어떤 기준과 조건으로 운영해야 하는지 묻는 질문으로 바꾼다.', '어떤 기준으로, 어떤 조건에서, 함께 지키려면, 운영되어야 하는가, 접근권과 생태 보전'],
    ['B-b', '서론', '1-1', '위치 없는 출처', '팩트체크', '자료 1-2', '"(자료 1)"만 있고 몇 번째 문단인지 없다. "자료 1 · 2번째 문단"처럼 위치까지 쓴다.', '(자료 1 · 2번째 문단), 약 50만 명, 1,000억 원 이상'],
    ['B-c1', '서론', '1-3', '사실만 반복하는 해석', '팩트체크', '자료 1-5', '[케이블카가 지역 경제에 도움이 된다]는 앞 문장을 다시 말할 뿐이다. 이 수치가 무엇을 뜻하는지, 환경공학자에게 왜 중요한지를 쓴다.', '이 수치는, 을/를 보여 준다, 반면, 산양 서식지, 환경공학자에게, 라는 의미가 있다'],
    ['B-a', '본론 1', '2-3', '믿을 수 없는 자료', '팩트체크', '자료 6', '"어떤 여행 블로그"는 인정 자료가 아니다. 위치가 분명한 조사 자료로 바꾼다.', '자료 6 · 2번째 문단에 따르면, 고령층 응답자의 62%, 정상 경관을 포기'],
    ['C1', '본론 1', '2-4', '전문가 관점 없음', '환경·미래', '자료 2', '많은 사람이 원한다는 이유만으로 "꼭 필요하다"고 한다. 환경공학자라면 생태 보전 기준을 먼저 따진다.', '환경공학자로서, 생태 보전 기준, 다만, 조건이 먼저 필요하다, 아고산대 식생'],
    ['B-c2', '본론 2', '3-2', '사실만 반복하는 해석', '현실·실행', '', '[공사 기간에 불편이 생긴다]는 자료 내용을 그대로 되풀이한다. 이 자료가 내 주장에 어떤 의미인지 쓴다.', '이 자료는, 의 근거가 된다, 왜냐하면, 때문이다, 라는 의미가 있다'],
    ['C2', '본론 2', '3-3', '한계 회피', '윤리·공정', '자료 5', '"몇 년이면 끝난다"며 공사 소음만 다룬다. 정상 탐방 제한, 아고산대 회복처럼 운영 뒤의 진짜 한계를 인정한다.', '물론, 한계가 있다, 정상 방향 탐방로, 갈 수 없다, 회복에 수십 년, 그럼에도, 보완할 수 있다'],
    ['C3', '본론 3', '4-1,4-2', '추상적 대안', '현실·실행', '자료 4', '"노력해야 한다", "요금의 일부"는 실행 기준이 없다. 누가, 몇 %를, 어기면 어떻게 되는지 쓴다.', '의무로, 10%, 생태 복원 기금, 의무 적립, 관람 구역, 과태료, 조례로 정한다'],
    ['A3', '결론', '5-1', '다짐형 결론', '환경·미래', '', '조건 없이 "노력해야 할 것이다"로 끝난다. "어떤 조건이 지켜질 때에만"이라는 조건부 통찰과 전문가 가치관을 쓴다.', '조건이 지켜질 때에만, 함께 이룰 수 있다, 환경공학자로서 나는, 가장 중요하게 생각한다']
  ], [50, 80, 110, 130, 90, 70, 420, 300]);

  make(SH.DATA, [
    ['번호', '자료', '위치', '내용'],
    ['자료 1-2', '자료 1', '2번째 문단', '양양군은 케이블카를 설치하면 연간 방문객이 약 50만 명 늘고, 지역 경제 효과가 1,000억 원 이상일 것으로 예상했다.'],
    ['자료 1-5', '자료 1', '5번째 문단', '환경단체는 상부 정류장 예정지가 산양 서식지와 겹치며, 이 일대에서 산양 흔적이 해마다 100건 이상 확인된다고 밝혔다.'],
    ['자료 2', '자료 2', '3번째 문단', '해발 1,400m 이상 아고산대의 식생은 한 번 훼손되면 회복에 수십 년 이상 걸린다.'],
    ['자료 3', '자료 3', '2번째 문단', '공사 기간에는 등산로 일부가 통제되고 소음과 먼지가 발생한다.'],
    ['자료 4', '자료 4', '3번째 문단', '해외의 한 국립공원은 케이블카 요금의 10%를 생태 복원 기금으로 의무 적립하고, 관람 구역을 벗어난 이용객에게 과태료를 부과한다.'],
    ['자료 5', '자료 5', '4번째 문단', '생태 보호를 위해 상부 정류장 이용객은 정상 방향 탐방로로 갈 수 없도록 하는 조건이 붙었다.'],
    ['자료 6', '자료 6', '2번째 문단', '고령층 응답자의 62%가 체력 때문에 국립공원 정상 경관을 포기한 적이 있다고 답했다.']
  ], [80, 60, 90, 520]);

  make(SH.CLUE, [
    ['역할', '비밀 단서'],
    ['팩트체크', '출처는 "자료 번호 · 몇 번째 문단"까지 있어야 증거로 인정된다. "~라고 한다", 블로그는 의심하라. 해석이라면서 앞 문장을 그대로 되풀이하는 곳도 있다.'],
    ['윤리·공정', '질문이 예/아니오로 끝나면 생각도 거기서 멈춘다. 그리고 글쓴이가 진짜 아픈 곳을 피해 다른 이야기를 하는 구간이 있다.'],
    ['현실·실행', '"노력해야 한다"는 대안이 아니다. 누가, 얼마나, 안 지키면 어떻게 되는지 물어라. 그리고 "이 자료는 내 주장에서 [ ]" 문장이 정말 해석인지 의심하라.'],
    ['환경·미래', '글쓴이는 환경공학자다. 환경공학자라면 무엇을 먼저 따질까? 끝맺음이 조건 없는 다짐뿐인지도 확인하라.']
  ], [90, 600]);

  make(SH.ZONE, [
    ['모둠', '고쳐 쓸 구간 (쉼표로 구분, 본문 탭의 구간 이름 그대로)'],
    [1, '탐구 질문,서론'],
    [2, '본론 1'],
    [3, '본론 2'],
    [4, '본론 3,결론']
  ], [60, 360]);

  make(SH.PHRASE, [
    ['구간', '종류', '카드'],
    ['공통', '자료 인용', '( 자료 번호 ) · ( 몇 ) 번째 문단에 따르면, '],
    ['공통', '해석', '이 수치는 ( )을/를 보여 준다.'],
    ['공통', '해석', '환경공학자의 관점에서 이는 ( )라는 의미가 있다.'],
    ['공통', '해석', '이 자료는 내 주장에서 ( )의 근거가 된다. 왜냐하면 ( ) 때문이다.'],
    ['탐구 질문', '질문 틀', '환경공학자의 시선에서, 설악산 오색 케이블카는 ( )와/과 ( )을/를 함께 지키려면 어떤 기준으로 운영되어야 하는가?'],
    ['탐구 질문', '질문 틀', '환경공학자의 시선에서, 설악산 오색 케이블카는 어떤 조건에서만 운영될 수 있는가?'],
    ['서론', '연결', '반면, '],
    ['서론', '딜레마', '이처럼 케이블카는 ( )와/과 ( ) 사이의 딜레마를 안고 있다.'],
    ['서론', '주장', '따라서 나는 환경공학자로서 ( )이/가 지켜질 때에만 케이블카를 운영해야 한다고 주장한다.'],
    ['본론 1', '근거', '케이블카가 필요한 이유는 ( )의 접근권 때문이다.'],
    ['본론 1', '전문가 관점', '다만 환경공학자로서 ( )이/가 지켜진다는 조건이 먼저 필요하다.'],
    ['본론 2', '한계 인정', '물론 내 주장에도 한계가 있다.'],
    ['본론 2', '한계 내용', '( )에서는 ( )할 수 없다는 점이다.'],
    ['본론 2', '재반박', '그럼에도 ( )(으)로 이 한계를 보완할 수 있다.'],
    ['본론 3', '대안 틀', '첫째, ( )은/는 ( )을/를 의무로 해야 한다.'],
    ['본론 3', '강제 장치', '이를 어기면 ( )을/를 부과한다.'],
    ['본론 3', '재원', '이용 요금의 ( )%를 ( )에 쓰도록 정한다.'],
    ['결론', '조건부 통찰', '( )(이)라는 조건이 지켜질 때에만 ( )와/과 ( )을/를 함께 이룰 수 있다.'],
    ['결론', '가치관', '환경공학자로서 나는 ( )을/를 가장 중요하게 생각한다.'],
    ['공통', '핵심 단어', '접근권'], ['공통', '핵심 단어', '고령층'], ['공통', '핵심 단어', '산양 서식지'], ['공통', '핵심 단어', '아고산대 식생'],
    ['공통', '핵심 단어', '회복에 수십 년'], ['공통', '핵심 단어', '생태 복원 기금'], ['공통', '핵심 단어', '의무 적립'], ['공통', '핵심 단어', '관람 구역 제한'],
    ['공통', '핵심 단어', '과태료'], ['공통', '핵심 단어', '조건부 운영'], ['공통', '핵심 단어', '생태 보전 기준'], ['공통', '핵심 단어', '정상 탐방로 제한']
  ], [80, 90, 560]);

  const lg = make(SH.LOG, [LOG_HEAD], [70, 130, 60, 50, 80, 90, 220, 110, 220, 220, 70, 60, 50, 60, 90]);
  if (String(lg.getRange(1, 14).getValue()) !== '반') lg.getRange(1, 14).setValue('반').setFontWeight('bold');   // 예전 기록 탭에 '반' 칸 추가
  if (String(lg.getRange(1, 15).getValue()) !== '이름') lg.getRange(1, 15).setValue('이름').setFontWeight('bold');
  const cf = sheet_(SH.CONF);
  if (!cf.getDataRange().getValues().some(r => String(r[0]).trim() === '반 목록 (쉼표로 구분)')) cf.appendRow(['반 목록 (쉼표로 구분)', '2-1반,2-2반,2-3반,2-4반,2-5반,2-6반']);
  bustAll_();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('사건파일')
    .addItem('처음 설정 (탭 만들기)', 'setupGame')
    .addItem('전체 기록 보관 후 비우기', 'newGame')
    .addToUi();
}

function newGame() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('새 게임을 시작할까요?', '지금 기록 탭은 「기록_보관_날짜」 탭으로 복사해 두고 비웁니다.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SH.LOG);
  if (log.getLastRow() > 1) {
    const copy = log.copyTo(ss);
    copy.setName('기록_보관_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd_HHmm'));
    log.getRange(2, 1, log.getLastRow() - 1, LOG_HEAD.length).clearContent();
  }
  saveAll_({ cur: '', by: {} });
  bustAll_();
}

/* ================= 웹 창구 ================= */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '사건파일 서버가 켜져 있어요.' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents || '{}');
    const api = {
      boot: boot, state: studentState, submitFind: submitFind, useHint: useHint,
      saveLine: saveLine, submitRline: submitRline, submitReviewAll: submitReviewAll, submitSelf: submitSelf,
      board: boardState, setPhase: setPhase, toggleHide: toggleHide, checkPin: checkPin,
      admin: admin, openClass: openClass, addClass: addClass, resetClass: resetClass, setHidden: setHidden
    };
    if (!api[req.fn]) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, data: api[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= 읽기 도우미 ================= */
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function rows_(n) { const s = sheet_(n); if (!s) return []; return s.getLastRow() < 2 ? [] : s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).getValues(); }

function cached_(k, fn) {
  const c = CacheService.getScriptCache(), h = c.get(k);
  if (h) return JSON.parse(h);
  const v = fn();
  try { c.put(k, JSON.stringify(v), 30); } catch (e) {}
  return v;
}

function conf_() { return cached_('conf', conf0_); }
function conf0_() {
  const c = {};
  rows_(SH.CONF).forEach(r => c[String(r[0]).trim()] = r[1]);
  return {
    title: String(c['글 제목'] || ''),
    groups: Number(c['모둠 수']) || 4,
    roles: String(c['역할 (쉼표로 구분)'] || '').split(',').map(s => s.trim()).filter(Boolean),
    pin: String(c['교사 비밀번호'] || ''),
    hintCost: Number(c['힌트 감점'] || 1),
    classes: keyStr_(c['반 목록 (쉼표로 구분)'] || '2-1반,2-2반,2-3반,2-4반,2-5반,2-6반').split(',').map(s => s.trim()).filter(Boolean)
  };
}

function splitSentences(text) {
  const out = [], re = /[^.?!。]+[.?!。]+["')\]]*\s*|[^.?!。]+$/g;
  let m;
  while ((m = re.exec(text)) !== null) { const s = m[0].trim(); if (s) out.push(s); }
  return out;
}

function sections_() { return cached_('secs', sections0_); }
function sections0_() {
  return rows_(SH.TEXT).filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({ name: String(r[0]).trim(), text: String(r[1]).trim(), sentences: splitSentences(String(r[1])) }));
}

function flaws_() { return cached_('flaws', flaws0_); }
/* 2라운드 고쳐 쓰기용 예시 단어 (결함 탭 8번째 칸 '예시 단어'가 있으면 그걸 씀) */
const FLAW_WORDS = {
  'A1': '어떤 기준으로, 어떤 조건에서, 함께 지키려면, 운영되어야 하는가, 접근권과 생태 보전',
  'B-b': '(자료 1 · 2번째 문단), 약 50만 명, 1,000억 원 이상',
  'B-c1': '이 수치는, 을/를 보여 준다, 반면, 산양 서식지, 환경공학자에게, 라는 의미가 있다',
  'B-a': '자료 6 · 2번째 문단에 따르면, 고령층 응답자의 62%, 정상 경관을 포기',
  'C1': '환경공학자로서, 생태 보전 기준, 다만, 조건이 먼저 필요하다, 아고산대 식생',
  'B-c2': '이 자료는, 의 근거가 된다, 왜냐하면, 때문이다, 라는 의미가 있다',
  'C2': '물론, 한계가 있다, 정상 방향 탐방로, 갈 수 없다, 회복에 수십 년, 그럼에도, 보완할 수 있다',
  'C3': '의무로, 10%, 생태 복원 기금, 의무 적립, 관람 구역, 과태료, 조례로 정한다',
  'A3': '조건이 지켜질 때에만, 함께 이룰 수 있다, 환경공학자로서 나는, 가장 중요하게 생각한다'
};
function flaws0_() {
  return rows_(SH.FLAW).filter(r => r[0]).map(r => ({
    code: String(r[0]).trim(), section: String(r[1]), keys: keyStr_(r[2]).split(',').map(s => s.trim()).filter(Boolean),
    card: String(r[3]).trim(), role: String(r[4]).trim(), evidence: String(r[5]).trim(), note: String(r[6]),
    words: String((r[7] != null && String(r[7]).trim()) ? r[7] : (FLAW_WORDS[String(r[0]).trim()] || '')).split(',').map(x => x.trim()).filter(Boolean)
  }));
}

function data_() { return cached_('data', () => rows_(SH.DATA).filter(r => r[0]).map(r => ({ id: String(r[0]).trim(), name: String(r[1]), loc: String(r[2]), text: String(r[3]) }))); }
function clues_() { return cached_('clues', () => { const o = {}; rows_(SH.CLUE).forEach(r => o[String(r[0]).trim()] = String(r[1])); return o; }); }
function phrases_() { return cached_('phrases', () => rows_(SH.PHRASE).filter(r => String(r[2]).trim()).map(r => ({ section: String(r[0]).trim(), kind: String(r[1]).trim(), text: String(r[2]) }))); }
function zones_() { return cached_('zones', () => { const o = {}; rows_(SH.ZONE).forEach(r => o[Number(r[0])] = String(r[1]).split(',').map(s => s.trim()).filter(Boolean)); return o; }); }

/* 상태(라운드)는 반별로 스크립트 속성에 저장: { cur: '2-2반', by: { '2-2반': { phase, reveal, at } } } */
// 상태는 캐시하지 않고 매번 스크립트 속성에서 읽음 (여러 기기가 동시에 읽을 때 옛 단계가 남는 문제 방지)
function allState_() {
  const p = PropertiesService.getScriptProperties().getProperty('state');
  let s = p ? JSON.parse(p) : {};
  if (!s.by) s = { cur: '', by: {} };            // 예전 형식이면 새로 시작
  return s;
}
function saveAll_(s) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { PropertiesService.getScriptProperties().setProperty('state', JSON.stringify(s)); }
  finally { lock.releaseLock(); }
}
function curCls_() { return allState_().cur || ''; }
function getState_() {
  const s = allState_(), one = s.by[s.cur] || { phase: 0, reveal: false };
  return { cls: s.cur || '', phase: one.phase || 0, reveal: !!one.reveal, at: one.at || 0 };
}
function setState_(one) {
  const s = allState_();
  s.by[s.cur || ''] = { phase: one.phase || 0, reveal: !!one.reveal, at: one.at || Date.now() };
  saveAll_(s);
}

/* 기록 읽기 (3초 캐시) */
function log_() { const c = curCls_(); return logAll_().filter(e => e.cls === c); }
function logAll_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('log');
  if (hit) return JSON.parse(hit);
  const list = rows_(SH.LOG).filter(r => r[0] && String(r[12]).trim() === '').map(r => ({
    id: String(r[0]), time: r[1] instanceof Date ? r[1].getTime() : 0, type: String(r[2]), group: Number(r[3]), role: String(r[4]),
    keys: keyStr_(r[5]).split(',').filter(Boolean), quote: String(r[6]), card: String(r[7]), problem: String(r[8]), fix: String(r[9]),
    result: String(r[10]), target: String(r[11]), cls: keyStr_(r[13] == null ? '' : r[13]).trim(), who: String(r[14] == null ? '' : r[14])
  }));
  try { cache.put('log', JSON.stringify(list), 3); } catch (e) {}
  return list;
}
function bust_() { CacheService.getScriptCache().remove('log'); }
function bustAll_() { CacheService.getScriptCache().removeAll(['log', 'conf', 'secs', 'flaws', 'data', 'clues', 'zones', 'phrases']); }

function append_(row, who) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const id = Utilities.getUuid().slice(0, 8);
    const log = sheet_(SH.LOG);
    log.appendRow([id, new Date()].concat(row.slice(0, 11), [curCls_(), clean_(who, 60)]).map(v => typeof v === 'string' && /^\d+-\d+$/.test(v) ? "'" + v : v));
    bust_();
    return id;
  } finally { lock.releaseLock(); }
}

/* 시트가 "1-1" 같은 문장번호를 날짜(1월 1일)로 바꿔 버린 경우 되돌림 */
function keyStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M-d');
  return String(v).replace(/^'/, '');
}

function need_(cond, msg) { if (!cond) throw new Error(msg); }
function needClass_() { need_(curCls_(), '선생님이 아직 반을 열지 않았어요.'); }
function checkGroupRole_(g, role, conf) {
  need_(g >= 1 && g <= conf.groups, '모둠을 다시 골라 주세요.');
  need_(conf.roles.indexOf(role) >= 0, '역할을 다시 골라 주세요.');
}
function clean_(s, n) { return String(s || '').trim().slice(0, n || 500); }

/* ================= 점수 계산 ================= */
function score_(log, flaws, conf) {
  const groups = {};
  for (let g = 1; g <= conf.groups; g++) groups[g] = { group: g, flaws: {}, find: 0, bonus: 0, hint: 0, review: 0, total: 0 };
  const byCode = {}; flaws.forEach(f => byCode[f.code] = f);

  log.forEach(e => {
    const G = groups[e.group]; if (!G) return;
    if (e.type === 'find' && e.result) {
      const [kind, code] = e.result.split(':');           // full:A1 / part:A1
      const f = byCode[code]; if (!f) return;
      const cur = G.flaws[code] || { pts: 0, roleHit: false, tries: 0 };
      if (kind === 'full') { cur.pts = 3; if (e.role === f.role) cur.roleHit = true; }
      else if (kind === 'part') { cur.tries++; cur.pts = Math.max(cur.pts, 1); }
      G.flaws[code] = cur;
    }
    if (e.type === 'hint') G.hint += conf.hintCost;
  });

  // 상호 심사: 심사 모둠·대상 문장별 마지막 판정만, '타당해요' 1개당 +1
  Object.values(lastRline_(log)).forEach(e => {
    const T = groups[Number(e.target)]; if (!T || Number(e.target) === e.group) return;
    if (e.card === 'ok') T.review += 1;
  });

  Object.values(groups).forEach(G => {
    Object.values(G.flaws).forEach(f => { G.find += f.pts; if (f.roleHit) G.bonus += 1; });
    G.total = G.find + G.bonus + G.review - G.hint;
  });
  return groups;
}

function grade_(total, max) {
  const r = total / max;
  return r >= 0.7 ? '명탐정' : r >= 0.5 ? '유능한 탐정' : r >= 0.3 ? '평범한 탐정' : '견습 탐정';
}

/* ================= 학생 API ================= */
function boot() {
  const conf = conf_();
  return {
    title: conf.title, groups: conf.groups, roles: conf.roles, phases: PHASES,
    cards: CARDS.map(c => ({ name: c[0], desc: c[1] })),
    sections: sections_().map(s => ({ name: s.name, sentences: s.sentences })),
    clues: clues_(), flawCount: flaws_().length, hintCost: conf.hintCost, phrases: phrases_(),
    state: getState_()
  };
}

/* 우리 모둠 화면에 필요한 것만 */
function studentState(group) {
  const st = getState_();
  const conf = conf_();
  group = Number(group);
  const log = log_();
  const flaws = flaws_();
  const mine = log.filter(e => e.group === group);
  const found = {};
  mine.filter(e => e.type === 'find' && e.result).forEach(e => {
    const [kind, code] = e.result.split(':');
    const f = flaws.find(x => x.code === code); if (!f) return;
    const cur = found[code] || { code, section: f.section, keys: f.keys, kind: 'part', card: '', evidence: '', note: '', tries: 0, hitKeys: [] };
    e.keys.forEach(k => { if (f.keys.indexOf(k) >= 0 && cur.hitKeys.indexOf(k) < 0) cur.hitKeys.push(k); });   // 학생이 실제로 누른 문장만
    if (kind === 'full') { cur.kind = 'full'; cur.card = f.card; cur.evidence = f.evidence; cur.note = f.note; }
    else cur.tries++;
    found[code] = cur;
  });
  const data = data_();
  const opened = {};
  Object.values(found).forEach(f => { if (f.kind === 'full' && f.evidence) opened[f.evidence] = true; });
  const allOpen = st.phase >= 1;   // 자료 상자는 1라운드부터 모두 공개 (결함 판단에 필요)
  const hints = mine.filter(e => e.type === 'hint').map(e => ({ text: e.problem }));
  const target = reviewTarget_(group, conf.groups);
  const myLines = st.phase >= 2 ? lines_(log, group) : {};
  const myR = {}; Object.values(lastRline_(log)).filter(e => e.group === group).forEach(e => myR[e.keys[0]] = { verdict: e.card, comment: e.problem, who: e.who });
  const gotR = {}; Object.values(lastRline_(log)).filter(e => Number(e.target) === group).forEach(e => (gotR[e.keys[0]] = gotR[e.keys[0]] || []).push({ verdict: e.card, comment: e.problem }));
  const allC = mine.filter(e => e.type === 'review').pop();
  return {
    state: st,
    finds: mine.filter(e => e.type === 'find').map(e => ({ id: e.id, role: e.role, keys: e.keys, quote: e.quote, card: e.card, problem: e.problem, fix: e.fix, result: e.result.split(':')[0] || 'none' })).reverse(),
    found: Object.values(found).map(f => ({ code: f.code, section: f.section, keys: f.hitKeys, kind: f.kind, card: f.card, evidence: f.evidence, note: f.note, locked: f.kind !== 'full' && f.tries >= 2 })),
    data: data.map(d => (allOpen || opened[d.id]) ? d : { id: d.id, name: d.name, loc: d.loc, locked: true }),
    hints: hints,
    lines: myLines,
    review: st.phase >= 3 ? { target, changes: changes_(lines_(log, target)), mine: myR, comment: allC ? allC.problem : '', got: gotR } : null,
    selfs: mine.filter(e => e.type === 'self').map(e => ({ role: e.role, card: e.card, problem: e.problem, fix: e.fix })),
    allFlaws: st.phase >= 2 ? flaws.map(f => ({ code: f.code, card: f.card, note: f.note, evidence: f.evidence, keys: f.keys, role: f.role, words: f.words || [], found: !!(found[f.code] && found[f.code].kind === 'full') })) : [],
    foundCards: st.phase >= 4 ? classCards_(log, flaws) : [],
    result: st.phase >= 5 && st.reveal ? resultPack_(log, flaws, conf) : null
  };
}

function classCards_(log, flaws) {
  const set = {};
  log.filter(e => e.type === 'find' && e.result.indexOf('full:') === 0).forEach(e => {
    const f = flaws.find(x => x.code === e.result.split(':')[1]); if (f) set[f.card] = true;
  });
  return CARDS.map(c => ({ name: c[0], desc: c[1], found: !!set[c[0]] }));
}

function submitFind(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 1, '지금은 결함 수사 시간이 아니에요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  const keys = (p.keys || []).map(String).filter(k => /^\d+-\d+$/.test(k)).slice(0, 8);
  need_(keys.length, '문장을 하나 이상 골라 주세요.');
  need_(CARDS.some(c => c[0] === p.card), '결함 카드를 하나 골라 주세요.');
  const problem = clean_(p.problem, 300), fix = '';   // 1라운드는 문제 표시만. 고치기는 2라운드에서

  // 판정
  const flaws = flaws_();
  const hitFlaws = flaws.filter(f => f.keys.some(k => keys.indexOf(k) >= 0));
  const log = log_().filter(e => e.group === g && e.type === 'find');
  const status = code => {
    let full = false, tries = 0;
    log.forEach(e => { if (e.result === 'full:' + code) full = true; if (e.result === 'part:' + code) tries++; });
    return { full, tries };
  };
  let result = '', msg, evidence = '';
  const exact = hitFlaws.find(f => f.card === p.card);
  if (exact) {
    const s = status(exact.code);
    result = 'full:' + exact.code;
    evidence = exact.evidence;
    msg = s.full ? 'already' : 'full';
    if (s.full) result = 'dup:' + exact.code;
    else if (s.tries >= 2) { result = 'late:' + exact.code; msg = 'late'; }
  } else if (hitFlaws.length) {
    // 아직 못 찾은 결함 중 하나로 '아깝다' 판정
    const open = hitFlaws.find(f => !status(f.code).full && status(f.code).tries < 2);
    if (open) { result = 'part:' + open.code; msg = 'part'; }
    else { msg = hitFlaws.some(f => !status(f.code).full) ? 'locked' : 'already'; result = 'dup:' + hitFlaws[0].code; }
  } else { msg = 'none'; }

  const quote = keys.map(k => sentenceOf_(k)).join(' ');
  append_(['find', g, p.role, keys.join(','), quote.slice(0, 1000), p.card, problem, fix, result, '', ''], p.who);
  const out = { msg };
  if (msg === 'full') {
    const f = flaws.find(x => 'full:' + x.code === result);
    out.card = f.card; out.section = f.section; out.roleHit = f.role === p.role; out.note = f.note;
    const d = data_().find(x => x.id === evidence); if (d) out.evidence = d;
  }
  return out;
}

function sentenceOf_(k) {
  const [a, b] = k.split('-').map(Number); const s = sections_()[a];
  return s && s.sentences[b] ? s.sentences[b] : '';
}

function useHint(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 1, '힌트는 결함 수사 시간에만 쓸 수 있어요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  const flaws = flaws_();
  const log = log_().filter(e => e.group === g);
  const done = {}; log.filter(e => e.type === 'find' && e.result.indexOf('full:') === 0).forEach(e => done[e.result.split(':')[1]] = true);
  const hinted = {}; log.filter(e => e.type === 'hint').forEach(e => hinted[e.result] = true);
  const left = flaws.filter(f => !done[f.code] && !hinted[f.code]);
  need_(left.length, '더 줄 힌트가 없어요. 이미 다 찾았거나 힌트를 다 썼어요.');
  // 역할 담당 결함을 먼저
  const pick = left.find(f => f.role === p.role) || left[0];
  const text = `${pick.section}에 「${pick.card}」 결함이 숨어 있다.`;
  append_(['hint', g, p.role, '', '', '', text, '', pick.code, '', ''], p.who);
  return { text };
}

/* ---------- 2라운드: 문장 단위 고쳐 쓰기 ---------- */
function lines_(log, group) {
  const m = {};
  log.filter(e => e.type === 'line' && e.group === group).forEach(e => {
    let meta = {}; try { meta = JSON.parse(e.quote || '{}'); } catch (x) {}
    const k = e.keys[0] || '';
    m[k] = { key: k, sec: e.target, text: e.fix, reason: e.problem, del: e.card === 'del', who: e.who, role: e.role, time: e.time, after: meta.after || '' };
  });
  return m;
}
function assemble_(si, sentences, lines) {
  const out = [];
  const adds = Object.values(lines).filter(l => l.key.indexOf('n:') === 0 && !l.del && l.text.trim());
  sentences.forEach((orig, j) => {
    const k = si + '-' + j, L = lines[k];
    if (!(L && L.del)) out.push(L ? L.text : orig);
    adds.filter(a => a.after === k).sort((a, b) => a.time - b.time).forEach(a => out.push(a.text));
  });
  return out.join(' ');
}
function changes_(lines) {
  const secs = sections_();
  return Object.values(lines).filter(l => {
    if (l.key.indexOf('n:') === 0) return !l.del && l.text.trim();
    return l.del || l.text !== sentenceOf_(l.key);
  }).map(l => ({ key: l.key, sec: l.sec, orig: l.key.indexOf('n:') === 0 ? '' : sentenceOf_(l.key), text: l.text, reason: l.reason, del: l.del, who: l.who, role: l.role, added: l.key.indexOf('n:') === 0, after: l.after }))
    .sort((a, b) => order_(a) - order_(b));
}
function order_(c) {
  const k = c.added ? c.after : c.key;
  const [a, b] = String(k).split('-').map(Number); return a * 1000 + b * 10 + (c.added ? 5 : 0);
}
function groupEssay_(log, group) {
  const L = lines_(log, group);
  return sections_().map((s, si) => ({ name: s.name, text: assemble_(si, s.sentences, L) }));
}
function reviewTarget_(g, n) { return g % n + 1; }   // 1→2, 2→3, …, 마지막→1
function lastRline_(log) {
  const m = {};
  log.filter(e => e.type === 'rline').forEach(e => { m[e.group + '>' + e.target + '>' + (e.keys[0] || '')] = e; });
  return m;
}

function saveLine(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 2, '지금은 보고서 복원 시간이 아니에요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  const secs = sections_();
  const key = String(p.key || '');
  let sec = '', after = '';
  if (/^\d+-\d+$/.test(key)) {
    const [a, b] = key.split('-').map(Number);
    need_(secs[a] && secs[a].sentences[b] != null, '없는 문장이에요.');
    sec = secs[a].name;
  } else {
    need_(/^n:[a-z0-9]{4,12}$/.test(key), '문장 번호가 이상해요.');
    after = String(p.after || '');
    need_(/^\d+-\d+$/.test(after) && secs[Number(after.split('-')[0])], '어디 뒤에 넣을지 알 수 없어요.');
    sec = secs[Number(after.split('-')[0])].name;
  }
  const text = clean_(p.text, 800), reason = clean_(p.reason, 300), del = !!p.del;
  const isOrig = /^\d+-\d+$/.test(key) && text === sentenceOf_(key) && !del;
  if (!isOrig && !(del && key.indexOf('n:') === 0)) need_(reason.replace(/\s/g, '').length >= 4, '왜 이렇게 바꿨는지 짧게라도 적어 주세요. (상호 심사 때 다른 모둠이 봐요)');
  if (!del && !isOrig) need_(text.replace(/\s/g, '').length >= 4, '고친 문장이 너무 짧아요.');
  append_(['line', g, p.role, key, JSON.stringify({ after }), del ? 'del' : '', reason, text, '', sec, ''], p.who);
  return { ok: true };
}

/* ---------- 3라운드: 문장별 타당성 심사 ---------- */
function submitRline(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 3, '지금은 상호 심사 시간이 아니에요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  const t = reviewTarget_(g, conf.groups);
  need_(p.verdict === 'ok' || p.verdict === 'no', '타당해요 / 아쉬워요 중에 골라 주세요.');
  const comment = clean_(p.comment, 300);
  if (p.verdict === 'no') need_(comment.replace(/\s/g, '').length >= 4, '아쉬운 점을 짧게 적어 주세요.');
  const key = String(p.key || '');
  need_(changes_(lines_(log_(), t)).some(c => c.key === key), '심사할 수 없는 문장이에요.');
  append_(['rline', g, p.role, key, '', p.verdict, comment, '', '', String(t), ''], p.who);
  return { ok: true };
}
function submitReviewAll(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 3, '지금은 상호 심사 시간이 아니에요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  const comment = clean_(p.comment, 300);
  need_(comment.replace(/\s/g, '').length >= 5, '한 줄 평을 적어 주세요.');
  append_(['review', g, p.role, '', '', '', comment, '', '', String(reviewTarget_(g, conf.groups)), ''], p.who);
  return { ok: true };
}

function submitSelf(p) {
  needClass_();
  const st = getState_(); need_(st.phase === 4, '지금은 내 초안 수사 시간이 아니에요.');
  const conf = conf_(); const g = Number(p.group); checkGroupRole_(g, p.role, conf);
  need_(CARDS.some(c => c[0] === p.card), '결함 카드를 하나 골라 주세요.');
  const where = clean_(p.where, 300), plan = clean_(p.plan, 400);
  need_(where.replace(/\s/g, '').length >= 4, '내 초안의 어느 부분인지 적어 주세요.');
  need_(plan.replace(/\s/g, '').length >= 8, '어떻게 고칠지 조금 더 자세히 적어 주세요.');
  append_(['self', g, p.role, '', '', p.card, where, plan, '', '', ''], p.who);
  return { ok: true };
}

/* ================= 교사 API ================= */
function pinOk_(pin) { const c = conf_(); return c.pin && String(pin) === c.pin; }
function checkPin(pin) { return pinOk_(pin); }

function setPhase(pin, phase, reveal) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  needClass_();
  phase = Math.max(0, Math.min(PHASES.length - 1, Number(phase)));
  setState_({ phase, reveal: phase === PHASES.length - 1 ? !!reveal : false, at: Date.now() });
  return getState_();
}

function toggleHide(pin, id) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  const s = sheet_(SH.LOG);
  const ids = s.getRange(2, 1, Math.max(1, s.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) {
    s.getRange(i + 2, 13).setValue('교사 숨김'); bust_(); return true;
  }
  return false;
}

/* ---------- 교사 관리 ---------- */
function admin(pin) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  const conf = conf_(), cur = curCls_(), all = allState_();
  const raw = rows_(SH.LOG).filter(r => r[0]).map(r => ({
    id: String(r[0]), time: r[1] instanceof Date ? r[1].getTime() : 0, type: String(r[2]), group: Number(r[3]), role: String(r[4]),
    card: String(r[7]), problem: String(r[8]), fix: String(r[9]), result: String(r[10]), target: String(r[11]),
    hidden: String(r[12]).trim() !== '', cls: keyStr_(r[13] == null ? '' : r[13]).trim()
  }));
  const names = conf.classes.slice();
  raw.forEach(e => { if (e.cls && names.indexOf(e.cls) < 0) names.push(e.cls); });
  return {
    cur,
    classes: names.map(n => ({ name: n, count: raw.filter(e => e.cls === n && !e.hidden).length, phase: (all.by[n] || {}).phase || 0 })),
    records: raw.filter(e => e.cls === cur).reverse().slice(0, 400)
  };
}

function openClass(pin, name) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  name = String(name || '').trim(); need_(name, '반 이름이 없어요.');
  const s = allState_(); s.cur = name; if (!s.by[name]) s.by[name] = { phase: 0, reveal: false, at: Date.now() };
  saveAll_(s); bust_();
  return getState_();
}

function addClass(pin, name) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  name = String(name || '').trim().slice(0, 20); need_(name && name.indexOf(',') < 0, '반 이름을 확인해 주세요.');
  const conf = conf0_();
  if (conf.classes.indexOf(name) < 0) {
    const sh = sheet_(SH.CONF), vals = sh.getDataRange().getValues();
    const list = conf.classes.concat([name]).join(',');
    const i = vals.findIndex(r => String(r[0]).trim() === '반 목록 (쉼표로 구분)');
    if (i >= 0) sh.getRange(i + 1, 2).setValue(list); else sh.appendRow(['반 목록 (쉼표로 구분)', list]);
    CacheService.getScriptCache().remove('conf');
  }
  return true;
}

function resetClass(pin) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  const cur = curCls_(); need_(cur, '열린 반이 없어요.');
  const sh = sheet_(SH.LOG), n = sh.getLastRow() - 1;
  if (n > 0) {
    const vals = sh.getRange(2, 1, n, LOG_HEAD.length).getValues();
    vals.forEach((r, i) => { if (r[0] && keyStr_(r[13] == null ? '' : r[13]).trim() === cur && String(r[12]).trim() === '') sh.getRange(i + 2, 13).setValue('초기화'); });
  }
  setState_({ phase: 0, reveal: false, at: Date.now() });
  bust_();
  return true;
}

function setHidden(pin, id, hide) {
  need_(pinOk_(pin), '비밀번호가 맞지 않아요.');
  const s = sheet_(SH.LOG);
  const ids = s.getRange(2, 1, Math.max(1, s.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) {
    s.getRange(i + 2, 13).setValue(hide ? '교사 숨김' : ''); bust_(); return true;
  }
  return false;
}

/* 칠판 화면: 순위는 결과 공개 때만 */
function boardState(pin) {
  const teacher = pinOk_(pin);
  const st = getState_();
  const conf = conf_();
  const log = log_();
  const flaws = flaws_();
  const groups = score_(log, flaws, conf);
  // 결함 지도: 반 전체에서 발견된 결함
  const foundBy = {};
  log.filter(e => e.type === 'find' && e.result.indexOf('full:') === 0).forEach(e => {
    const code = e.result.split(':')[1]; (foundBy[code] = foundBy[code] || {})[e.group] = true;
  });
  const flawMap = flaws.map(f => ({
    code: f.code, section: f.section, found: !!foundBy[f.code], groups: Object.keys(foundBy[f.code] || {}).length,
    // 1라운드 중에는 다른 모둠이 베끼지 않도록 결함 이름·문장을 숨김 (2라운드부터 공개)
    card: (foundBy[f.code] && st.phase >= 2) || (st.phase >= 5 && st.reveal) ? f.card : '',
    keys: (foundBy[f.code] && st.phase >= 2) || (st.phase >= 5 && st.reveal) ? f.keys : [],
    note: st.phase >= 5 && st.reveal ? f.note : ''
  }));
  const progress = {};
  for (let g = 1; g <= conf.groups; g++) progress[g] = {
    finds: log.filter(e => e.group === g && e.type === 'find').length,
    lines: changes_(lines_(log, g)).length,
    reviews: Object.values(lastRline_(log)).filter(e => e.group === g).length,
    toReview: changes_(lines_(log, reviewTarget_(g, conf.groups))).length,
    selfs: log.filter(e => e.group === g && e.type === 'self').length
  };
  const essays = {}, changes = {};
  for (let g = 1; g <= conf.groups; g++) { essays[g] = groupEssay_(log, g); changes[g] = changes_(lines_(log, g)); }
  const rlines = Object.values(lastRline_(log)).map(e => ({ group: e.group, target: Number(e.target), key: e.keys[0], verdict: e.card, comment: e.problem }));
  const reviews = log.filter(e => e.type === 'review').map(e => ({ group: e.group, target: Number(e.target), comment: e.problem }));
  const selfs = log.filter(e => e.type === 'self').map(e => ({ group: e.group, role: e.role, card: e.card, where: e.problem, plan: e.fix }));
  return {
    teacher, state: st,
    finds: log.filter(e => e.type === 'find').map(e => ({ id: e.id, group: e.group, role: e.role, keys: e.keys, card: e.card, problem: e.problem, fix: e.fix, hit: e.result.split(':')[0] || 'none' })),
    flawMap, progress, essays, changes, rlines, reviews, selfs,
    result: st.phase >= 5 && st.reveal ? resultPack_(log, flaws, conf) : null
  };
}

function resultPack_(log, flaws, conf) {
  const groups = score_(log, flaws, conf);
  const flawKeys = new Set(); flaws.forEach(f => f.keys.forEach(k => flawKeys.add(k)));
  const max = flaws.length * 4 + flawKeys.size;
  const rank = Object.values(groups).sort((a, b) => b.total - a.total).map(G => ({
    group: G.group, total: G.total, find: G.find, bonus: G.bonus, review: G.review, hint: G.hint,
    flawsFound: Object.values(G.flaws).filter(f => f.pts === 3).length, grade: grade_(G.total, max)
  }));
  // 복원된 보고서: 구간마다 '타당해요'를 가장 많이 받은 모둠의 글 (같으면 더 많이 고친 모둠)
  const secs = sections_();
  const okBy = {};
  Object.values(lastRline_(log)).forEach(e => {
    if (e.card !== 'ok') return;
    const t = Number(e.target), c = changes_(lines_(log, t)).find(x => x.key === e.keys[0]);
    if (c) okBy[t + '|' + c.sec] = (okBy[t + '|' + c.sec] || 0) + 1;
  });
  const restored = secs.map((s, si) => {
    let best = 0, bestScore = -1, bestText = s.text;
    for (let g = 1; g <= conf.groups; g++) {
      const L = lines_(log, g);
      const nChg = changes_(L).filter(c => c.sec === s.name).length;
      if (!nChg) continue;
      const sc = (okBy[g + '|' + s.name] || 0) * 100 + nChg;
      if (sc > bestScore) { bestScore = sc; best = g; bestText = assemble_(si, s.sentences, L); }
    }
    return { name: s.name, original: s.text, text: bestText, by: best, ok: best ? (okBy[best + '|' + s.name] || 0) : 0 };
  });
  return { rank, max, restored, flaws: flaws.map(f => ({ code: f.code, section: f.section, card: f.card, note: f.note })) };
}
