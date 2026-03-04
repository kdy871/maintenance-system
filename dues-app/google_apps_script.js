function doGet(e) {
    var output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);

    try {
        const SPREADSHEET_ID = '18oMbUxtvLGQCJ7feyER97xqjiTRmqPJXEBeP7v0dJRE';
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName('회비 현황');

        // 2행부터 시작 (김기용 부장이 2행에 위치함)
        const lastRow = Math.max(sheet.getLastRow(), 5);
        const data = sheet.getRange(2, 1, lastRow - 1, 25).getDisplayValues();

        let incomeTotal = 0;
        let expenseTotal = 0;

        const transactions = [];
        const members = [];

        for (let i = 0; i < data.length; i++) {
            const row = data[i];

            // -- 1. 입금 내역 (A~D열) --
            // A(0): 날짜, B(1): 입금자, C(2): 입금금액, D(3): 내용
            if (row[0] && row[1] && row[0] !== '날짜' && row[1] !== '입금자') {
                const dateStr = String(row[0]).trim();
                const amount = Number(String(row[2]).replace(/,/g, '')) || 0;

                if (amount > 0) {
                    incomeTotal += amount;

                    let m = 0, d = 0;
                    const dateMatch = dateStr.match(/(\d+)월\s*(\d+)일/);
                    if (dateMatch) {
                        m = parseInt(dateMatch[1]) || 0;
                        d = parseInt(dateMatch[2]) || 0;
                    }

                    transactions.push({
                        type: 'income',
                        date: dateStr,
                        name: String(row[1]).trim(),
                        note: row[3] ? String(row[3]).trim() : '',
                        amount: amount,
                        _month: m,
                        _day: d
                    });
                }
            }

            // -- 2. 지출 내역 (W, X, Y열) --
            // W(22): 날짜, X(23): 금액, Y(24): 사유
            if (i !== 35 && row[22] && row[22] !== '날짜') { // 37행(T37 총잔액)은 데이터 배열의 index 35가 됨
                const dateRaw = String(row[22]).trim(); // ex: "1. 22"
                let dateStr = dateRaw;
                let m = 0, d = 0;

                const dateMatch = dateRaw.match(/(\d+)\.\s*(\d+)/);
                if (dateMatch) {
                    m = parseInt(dateMatch[1]) || 0;
                    d = parseInt(dateMatch[2]) || 0;
                    dateStr = m + "월 " + d + "일";
                }

                const amount = Number(String(row[23]).replace(/,/g, '')) || 0;
                if (amount > 0) {
                    expenseTotal += amount;
                    transactions.push({
                        type: 'expense',
                        date: dateStr,
                        name: row[24] ? String(row[24]).trim() : '지출',
                        note: '',
                        amount: amount,
                        _month: m,
                        _day: d
                    });
                }
            }

            // -- 3. 회원 명단 및 납부 현황 (F~T열) --
            // F(5): 순번, G(6): 성명, H~S(7~18): 1~12월, T(19): 납입총합
            const memberName = String(row[6]).replace(/\s/g, ''); // 공백 제거 (ex: 김 동 용 -> 김동용)

            // F열에 순번(숫자)이 있고, G열에 이름이 있는 경우만 회원으로 간주
            const isMemberRow = row[5] && !isNaN(Number(row[5])) && memberName && memberName !== '성명';

            if (isMemberRow) {
                const totalPaid = Number(String(row[19]).replace(/,/g, '')) || 0; // T열

                const monthly = [];
                // H열(7)부터 12개월치 읽기
                for (let mIdx = 0; mIdx < 12; mIdx++) {
                    const monthlyVal = Number(String(row[7 + mIdx]).replace(/,/g, '')) || 0;
                    monthly.push(monthlyVal);
                }

                members.push({
                    name: String(row[6]).trim(), // 원래 이름 (공백 유지, 화면 표시용)
                    totalPaid: totalPaid,
                    monthly: monthly
                });
            }
        }

        // 거래 내역 최신순 정렬 (월, 일 내림차순)
        transactions.sort((a, b) => {
            if (b._month !== a._month) return b._month - a._month;
            return b._day - a._day;
        });

        // 잔액은 철저하게 '총 수입 - 총 지출' 로 수학적 계산만 반환
        let totalBalance = incomeTotal - expenseTotal;

        const responseData = {
            totalBalance: totalBalance,
            incomeTotal: incomeTotal,
            expenseTotal: expenseTotal,
            transactions: transactions,
            members: members
        };

        output.setContent(JSON.stringify(responseData));
        return output;

    } catch (error) {
        output.setContent(JSON.stringify({ error: error.toString(), stack: error.stack }));
        return output;
    }
}
