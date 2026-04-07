const fs = require('fs');

const codeToAppend = `

function 次年度マスター用データ出力() {
    console.log('Starting 次年度マスター用データ生成...');
    const ss = getOrCreateSpreadsheet();
    if (!ss) return;

    const sheetName = '次年度用_前年実績マスター';
    let outputSheet = ss.getSheetByName(sheetName);
    
    if (outputSheet) {
        outputSheet.clear();
    } else {
        outputSheet = ss.insertSheet(sheetName);
    }
    
    // マスター用ヘッダー
    const headers = [
        '年月',
        '宅配_全_金', '宅配_乳_金', '宅配_乳_本', '宅配_400_金', '宅配_400_本', '宅配_1000_金', '宅配_1000_本',
        '直販_全_金', '直販_乳_金', '直販_乳_本', '直販_400_金', '直販_400_本', '直販_1000_金', '直販_1000_本',
        'R_全社_金', 'S_全乳_金', 'T_全乳_本'
    ];
    outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#c9daf8').setFontWeight('bold');
    
    const fy = getDashboardFiscalYear(ss);
    const qtyIndices = [2, 4, 6, 9, 11, 13, 16];
    
    const rows = [];
    for (let i = 0; i < 12; i++) {
        // 例: FY2025なら 2025/04, 2025/05...
        const d = new Date(fy, 3 + i, 1);
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy_MM');
        
        // 当月の実績値（金額・本数問わず全て「合計」が入っている）
        const totals = getSheetTotals(ss, key);
        
        // 当月の日数
        const dim = new Date(fy, 3 + i + 1, 0).getDate();
        
        const rowData = [d];
        for (let j = 0; j < 17; j++) {
            let val = totals.sums[j];
            
            // 数量(本数)の場合は日別平均に変換する (Dimで割る)
            if (qtyIndices.includes(j)) {
                val = dim > 0 ? val / dim : 0;
            }
            rowData.push(val);
        }
        rows.push(rowData);
    }
    
    // データ書き込み
    outputSheet.getRange(2, 1, 12, 18).setValues(rows);
    
    // フォーマット調整
    outputSheet.getRange('A:A').setNumberFormat('yyyy/MM');
    // B〜R列はカンマ区切り。平均値を含めExcelと同様の一般的な表示にするため小数点含める
    outputSheet.getRange(2, 2, 12, 17).setNumberFormat('0.00'); // 小数点以下も正確に残すため
    
    // 行幅調整
    outputSheet.setColumnWidth(1, 100);
    outputSheet.setColumnWidths(2, 17, 80);
    
    const msg = '次年度用マスターデータのシート「' + sheetName + '」を生成しました。\\nこのシートのデータをコピーして、次年度ダッシュボードの「前年実績マスター」に貼り付けてください。';
    console.log(msg);
    SpreadsheetApp.getUi().alert('完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
`;

fs.appendFileSync('Code.js', codeToAppend, 'utf8');
console.log('Appended to Code.js successfully.');
