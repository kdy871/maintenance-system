// 전기부 회비 관리 통합 웹/앱 서비스 스크립트 
// ( Google Apps Script ) - doGet 함수

function doGet(e) {
    try {
        // 1. 스프레드시트 아이디 (링크의 /d/ 와 /edit 사이의 값)
        const SPREADSHEET_ID = '18oMbUxtvLGQCJ7feyER97xqjiTRmqPJXEBeP7v0dJRE';
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

        // 파라미터 확인 (존재하지 않으면 기본 데이터 반환 등)
        const action = e.parameter.action || 'getAllData';

        if (action === 'getAllData') {
            return getAllData(ss);
        }

        else {
            return ContentService.createTextOutput(JSON.stringify({
                error: "잘못된 action 파라미터입니다."
            })).setMimeType(ContentService.MimeType.JSON);
        }

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            error: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function getAllData(ss) {
    // 시트 이름들을 정의 (실제 사용하시는 시트 이름으로 변경 필요)
    // 예: '회비내역', '입출금상세', '미납자명단' 등
    const summarySheet = ss.getSheetByName('요약') || ss.getSheets()[0];
    const historySheet = ss.getSheetByName('입출금내역') || ss.getSheets()[0];
    const membersSheet = ss.getSheetByName('회원명단') || ss.getSheets()[0];

    // 데이터 가공 로직 (임시)
    // *** 주의 사항 ***
    // 현재 구글 시트의 정확한 '시트 이름'과 '열 번호', '행 구조'를 알 수 없기 때문에 
    // 표준적인 데이터 파싱 구조의 뼈대만 작성했습니다.

    // 1. 총회비 정보 가져오기 (예: D2 셀에 총액이 있다고 가정)
    const totalBalance = summarySheet.getRange("D2").getValue();

    // 2. 미납 정보 가져오기
    // membersSheet 에서 미납된 사람의 이름과 개월수, 금액을 가져옵니다.
    const unpaidMembers = []; // mock data format

    // 3. 입출금 내역
    // historySheet 에서 [날짜, 구분, 내용, 금액] 순으로 데이터를 가져옵니다.
    const transactions = []; // mock data format

    const responseData = {
        totalBalance: totalBalance,
        unpaidCount: unpaidMembers.length,
        unpaidAmount: 0, // 계산 로직 추가 필요
        unpaidMembers: unpaidMembers,
        transactions: transactions
    };

    return ContentService.createTextOutput(JSON.stringify(responseData))
        .setMimeType(ContentService.MimeType.JSON);
}

// 개발자가 초기 세팅 안내를 위해 임시로 포함한 템플릿 코드입니다.
// 실제 사용중인 시트의 형식(열 이름 등)을 알려주셔야 스크립트를 올바르게 수정할 수 있습니다!
