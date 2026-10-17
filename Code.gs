/**
 * @file Code.gs
 * @description 초등학교 인성체크리스트 구글 스프레드시트 연동용 Google Apps Script
 * 
 * [배포 방법]
 * 1. 구글 스프레드시트 생성 (또는 기존 시트 열기)
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 본 Code.gs 내용을 붙여넣기
 * 4. 상단 [배포] -> [새 배포] 클릭
 * 5. 유형 선택: [웹 앱]
 *    - 설명: 인성체크리스트 연동 API
 *    - 다음 사용자 권한으로 실행: 나 (소유자)
 *    - 액세스 권한이 있는 사용자: 모든 사용자 (Anyone)
 * 6. [배포] 클릭 후 생성된 웹 앱 URL(https://script.google.com/macros/s/.../exec)을
 *    인성체크리스트 앱의 [구글 시트 연동] 설정에 입력하세요.
 */

// 5대 시트 이름 정의 (요구사항 스키마와 100% 일치)
const SHEET_NAMES = {
  STUDENTS: '학생정보',
  CHECKLIST: '인성체크리스트',
  EVALUATIONS: '평가결과',
  FEEDBACK: 'AI피드백',
  GROWTH: '성장기록'
};

/**
 * 초기 시트 헤더 및 탭 자동 생성 함수
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 학생정보
  let s1 = ss.getSheetByName(SHEET_NAMES.STUDENTS);
  if (!s1) {
    s1 = ss.insertSheet(SHEET_NAMES.STUDENTS);
    s1.appendRow(['학생ID', '표시명', '학년', '반']);
    s1.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#F1F5F9');
  }

  // 2. 인성체크리스트
  let s2 = ss.getSheetByName(SHEET_NAMES.CHECKLIST);
  if (!s2) {
    s2 = ss.insertSheet(SHEET_NAMES.CHECKLIST);
    s2.appendRow(['문항ID', '인성덕목', '생활영역', '체크리스트문항', '평가척도']);
    s2.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#F1F5F9');
  }

  // 3. 평가결과
  let s3 = ss.getSheetByName(SHEET_NAMES.EVALUATIONS);
  if (!s3) {
    s3 = ss.insertSheet(SHEET_NAMES.EVALUATIONS);
    s3.appendRow(['평가ID', '학생ID', '평가날짜', '문항ID', '평가유형', '평가점수', '평가메모']);
    s3.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#F1F5F9');
  }

  // 4. AI피드백
  let s4 = ss.getSheetByName(SHEET_NAMES.FEEDBACK);
  if (!s4) {
    s4 = ss.insertSheet(SHEET_NAMES.FEEDBACK);
    s4.appendRow(['피드백ID', '학생ID', '평가날짜', '잘한점', '노력이필요한점', 'AI성장피드백', '추천실천목표']);
    s4.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#F1F5F9');
  }

  // 5. 성장기록
  let s5 = ss.getSheetByName(SHEET_NAMES.GROWTH);
  if (!s5) {
    s5 = ss.insertSheet(SHEET_NAMES.GROWTH);
    s5.appendRow(['기록ID', '학생ID', '평가날짜', '인성덕목', '학생평균점수', '교사평균점수', '최종평균점수', '실천목표', '실천여부', '성장변화']);
    s5.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#F1F5F9');
  }

  return '모든 시트 초기화가 완료되었습니다.';
}

/**
 * GET 요청 핸들러 (연결 상태 확인 및 데이터 조회)
 */
function doGet(e) {
  try {
    setupSheets();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
    
    if (action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Google Apps Script 인성체크리스트 API가 정상 작동 중입니다.',
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Ready'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POST 요청 핸들러 (평가결과, AI피드백, 성장기록 저장)
 */
function doPost(e) {
  try {
    setupSheets();
    const contents = JSON.parse(e.postData.contents);
    const action = contents.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'saveEvaluation') {
      // 1. 학생정보 동기화 (기존 미존재 시 추가)
      const student = contents.student;
      if (student) {
        const sSheet = ss.getSheetByName(SHEET_NAMES.STUDENTS);
        const sData = sSheet.getDataRange().getValues();
        let exists = false;
        for (let i = 1; i < sData.length; i++) {
          if (String(sData[i][0]) === String(student.id)) {
            exists = true;
            break;
          }
        }
        if (!exists) {
          sSheet.appendRow([student.id, student.name, student.grade, student.classNum]);
        }
      }

      // 2. 인성체크리스트 문항 동기화
      const questions = contents.questions || [];
      if (questions.length > 0) {
        const qSheet = ss.getSheetByName(SHEET_NAMES.CHECKLIST);
        const qData = qSheet.getDataRange().getValues();
        const existingQIds = new Set();
        for (let i = 1; i < qData.length; i++) {
          existingQIds.add(String(qData[i][0]));
        }
        questions.forEach(q => {
          if (!existingQIds.has(String(q.id))) {
            qSheet.appendRow([q.id, q.virtue, q.domain, q.question, q.scale || '1~4점']);
            existingQIds.add(String(q.id));
          }
        });
      }

      // 3. 평가결과 저장 (시트3: 평가ID, 학생ID, 평가날짜, 문항ID, 평가유형, 평가점수, 평가메모)
      const evalList = contents.evaluations || [];
      const eSheet = ss.getSheetByName(SHEET_NAMES.EVALUATIONS);
      evalList.forEach(item => {
        eSheet.appendRow([
          item.evalId || ('EV_' + Utilities.getUuid().slice(0, 8)),
          item.studentId,
          item.date,
          item.questionId,
          item.type, // '학생', '교사', '최종'
          item.score,
          item.memo || ''
        ]);
      });

      // 4. AI피드백 저장 (시트4: 피드백ID, 학생ID, 평가날짜, 잘한점, 노력이필요한점, AI성장피드백, 추천실천목표)
      const feedback = contents.feedback;
      if (feedback) {
        const fSheet = ss.getSheetByName(SHEET_NAMES.FEEDBACK);
        fSheet.appendRow([
          feedback.feedbackId || ('FB_' + Utilities.getUuid().slice(0, 8)),
          feedback.studentId,
          feedback.date,
          feedback.strength || '',
          feedback.needEffort || '',
          feedback.aiMessage || '',
          feedback.recommendedGoal || ''
        ]);
      }

      // 5. 성장기록 저장 (시트5: 기록ID, 학생ID, 평가날짜, 인성덕목, 학생평균점수, 교사평균점수, 최종평균점수, 실천목표, 실천여부, 성장변화)
      const growthList = contents.growthRecords || [];
      if (growthList.length > 0) {
        const gSheet = ss.getSheetByName(SHEET_NAMES.GROWTH);
        growthList.forEach(rec => {
          gSheet.appendRow([
            rec.recordId || ('GR_' + Utilities.getUuid().slice(0, 8)),
            rec.studentId,
            rec.date,
            rec.virtue,
            rec.studentAvg != null ? rec.studentAvg : '',
            rec.teacherAvg != null ? rec.teacherAvg : '',
            rec.finalAvg != null ? rec.finalAvg : '',
            rec.goal || '',
            rec.isPracticed || '진행중',
            rec.growthChange != null ? rec.growthChange : 0
          ]);
        });
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: '구글 스프레드시트에 평가 데이터가 성공적으로 저장되었습니다.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '알 수 없는 요청 action 입니다: ' + action
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: '스프레드시트 처리 중 오류: ' + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
