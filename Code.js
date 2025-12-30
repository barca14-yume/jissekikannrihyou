/**
 * Code.gs - 全社実績ダッシュボード管理
 * - Dashboard: Month/Quarter/Half/Year x 5 Columns
 * - Metrics: 17 items (Qty -> Average, Amt -> Sum)
 * - Y1000 Header Fix
 * - Target/Prev assumed to be Daily Averages in CSV
 */

const CONFIG = {
    FOLDER_NAME: '実績CSVアップロード',
    PROCESSED_FOLDER_NAME: 'processed',
    FOLDER_MONTHLY_FIX: '月次確定CSVアップロード',
    FOLDER_MASTER_IMPORT: 'マスタデータ取込',
    SS_NAME: '全社実績ダッシュボード',
    SHEET_SUMMARY: 'サマリーシート',
    SHEET_MONTHLY_REPORT: '月別レポート',
    SHEET_TARGET: '目標マスター',
    SHEET_PREV: '前年実績マスター',
    SHEET_DAYS: '稼働日マスター',
    NAME_S1: '営業計',
    NAME_S2: '直販他計'
};

function onOpen() {
    try {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('ダッシュボード管理')
            .addItem('CSVを取り込む（今すぐ実行）', 'importCSV')
            .addItem('マスタデータをCSVから取り込む', 'importMasterData')
            .addItem('サマリーのみ更新', 'updateSummarySheet')
            .addItem('月別レポートのみ更新', 'updateMonthlyReportSheet')
            .addItem('サマリーのみ更新', 'updateSummarySheet')
            .addItem('月別レポートのみ更新', 'updateMonthlyReportSheet')
            .addItem('月次確定データを取り込む', 'importMonthlyFixCSV')
            .addItem('手動入力用シート作成', 'createEmptyMonthSheetUI')
            .addToUi();
    } catch (e) {
        console.warn('onOpen UI check failed: ' + e.message);
    }
}

function setup() {
    console.log('Starting setup...');
    const ss = getOrCreateSpreadsheet();
    if (!ss) throw new Error("Spreadsheet error.");

    let sSheet = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
    if (!sSheet) {
        sSheet = ss.insertSheet(CONFIG.SHEET_SUMMARY);
        sSheet.getRange('A1').setValue('全社販売実績ダッシュボード').setFontSize(14).setFontWeight('bold');
        sSheet.getRange('A3').setValue('基準日:');
        sSheet.getRange('B3').setValue(new Date());
    }

    createMasterSheet(ss, CONFIG.SHEET_TARGET, 0);
    createMasterSheet(ss, CONFIG.SHEET_TARGET, 0);
    createMasterSheet(ss, CONFIG.SHEET_PREV, -1);
    createWorkingDaysSheet(ss);

    ensureFolder(CONFIG.FOLDER_NAME);
    ensureFolder(CONFIG.FOLDER_MASTER_IMPORT);
    ensureFolder(CONFIG.FOLDER_MONTHLY_FIX);

    console.log('Setup completed.');
}

function importMasterData() {
    console.log('Starting Master Import...');
    const ss = getOrCreateSpreadsheet();
    if (!ss) {
        logOrAlert('エラー: スプレッドシートが見つかりません。');
        return;
    }

    const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_MASTER_IMPORT);
    if (!folders.hasNext()) {
        logOrAlert('フォルダ「' + CONFIG.FOLDER_MASTER_IMPORT + '」が見つかりません。');
        return;
    }

    const folder = folders.next();
    const files = folder.getFiles();
    let importedCount = 0;

    while (files.hasNext()) {
        const file = files.next();
        const name = file.getName().toLowerCase();

        if (name === 'target.csv') {
            importMasterCSV(ss, file, CONFIG.SHEET_TARGET);
            importedCount++;
        } else if (name === 'prev.csv') {
            importMasterCSV(ss, file, CONFIG.SHEET_PREV);
            importedCount++;
        } else if (name === 'days.csv') {
            importDaysCSV(ss, file);
            importedCount++;
        }
    }

    if (importedCount > 0) {
        logOrAlert('マスタデータの取り込みが完了しました。件数: ' + importedCount);
        updateSummarySheet();
    } else {
        logOrAlert('target.csv または prev.csv が見つかりませんでした。');
    }
}

function logOrAlert(msg) {
    console.log(msg);
    try {
        SpreadsheetApp.getUi().alert(msg);
    } catch (e) { }
}

function importMasterCSV(ss, file, sheetName) {
    if (!ss) return;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    console.log('Importing ' + file.getName() + ' to ' + sheetName);
    const csvString = file.getBlob().getDataAsString('Shift_JIS');
    const data = Utilities.parseCsv(csvString);

    let startRow = 0;
    if (data.length > 0 && isNaN(Date.parse(data[0][0]))) {
        startRow = 1;
    }

    const rowsToWrite = [];
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        if (row.length < 18 || !row[0]) continue;

        const dDate = new Date(row[0]);
        if (isNaN(dDate.getTime()) || dDate.getFullYear() < 2000) continue;

        const values = row.slice(1, 18).map(v => parseNumber(v));
        rowsToWrite.push([dDate, ...values]);
    }

    if (rowsToWrite.length === 0) {
        logOrAlert('有効なデータが見つかりませんでした。CSVの日付形式を確認してください。');
        return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
        sheet.getRange(2, 1, lastRow - 1, 18).clearContent();
    }

    sheet.getRange(2, 1, rowsToWrite.length, 18).setValues(rowsToWrite);
    console.log('Imported ' + rowsToWrite.length + ' rows.');
}

function createMasterSheet(ss, sheetName, offsetYear) {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) return;

    sheet = ss.insertSheet(sheetName);
    const tHeaders = [
        '年月',
        '宅配_全_金', '宅配_乳_金', '宅配_乳_本', '宅配_400_金', '宅配_400_本', '宅配_1000_金', '宅配_1000_本',
        '直販_全_金', '直販_乳_金', '直販_乳_本', '直販_400_金', '直販_400_本', '直販_1000_金', '直販_1000_本',
        'R_全社_金', 'S_全乳_金', 'T_全乳_本'
    ];
    sheet.getRange(1, 1, 1, tHeaders.length).setValues([tHeaders]).setFontWeight('bold').setBackground('#c9daf8');

    const today = new Date();
    const currentFY = getFiscalYear(today);
    const targetFY = currentFY + offsetYear;

    const rows = [];
    for (let i = 0; i < 12; i++) {
        const d = new Date(targetFY, 3 + i, 1);
        const row = [d, ...new Array(17).fill(0)];
        rows.push(row);
    }
    sheet.getRange(2, 1, 12, 18).setValues(rows);
    sheet.getRange('A:A').setNumberFormat('yyyy/MM');
    sheet.getRange(2, 2, 12, 17).setNumberFormat('#,##0');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidths(2, 17, 80);
}

function importCSV() {
    const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
    if (!folders.hasNext()) return;
    const folder = folders.next();
    const files = folder.getFilesByType(MimeType.CSV);
    const ss = getOrCreateSpreadsheet();
    let processedAny = false;
    while (files.hasNext()) {
        try {
            processFile(files.next(), ss);
            processedAny = true;
        } catch (e) {
            console.error('Error: ' + e.message);
        }
    }
    if (processedAny) {
        updateSummarySheet();
        updateMonthlyReportSheet();
    }
}

// Extracted CSV parsing logic
function parseCsvToRecord(file) {
    const csvString = file.getBlob().getDataAsString('Shift_JIS');
    const csvData = Utilities.parseCsv(csvString);
    let header = csvData.shift();
    if (!header) return null;
    header = header.map(h => h.trim());

    const cols = {
        name: header.indexOf('名称'),
        item: header.indexOf('項目'),
        total: findHeaderIndex(header, ['合　計', '合計']),
        dairy: findHeaderIndex(header, ['乳製品計']),
        y400: findHeaderIndex(header, ['Y400類']),
        // 宅配用: yakult1000類 (小文字) 優先
        y1000_home: findHeaderIndex(header, ['yakult1000類', 'yakult1000', 'Yakult1000類', 'Y1000類', 'Y1000', 'Yakult1000', 'Y1000本', 'ｙａｋｕｌｔ１０００類', 'Ｙ１０００類', 'Ｙ１０００', 'Ｙａｋｕｌｔ１０００類']),
        // 直販用: 従来通り (Y1000類, Yakult1000類...) 優先
        y1000_direct: findHeaderIndex(header, ['Y1000類', 'Yakult1000類', 'Y1000', 'Yakult1000', 'Y1000本', 'yakult1000類'])
    };
    if (cols.name === -1 || cols.item === -1) return null;

    const record = {
        [CONFIG.NAME_S1]: { amt: {}, qty: {} },
        [CONFIG.NAME_S2]: { amt: {}, qty: {} }
    };

    csvData.forEach(row => {
        const name = row[cols.name] ? row[cols.name].trim() : '';
        const rawItem = row[cols.item] ? row[cols.item] : '';
        const item = rawItem.replace(/\s+/g, '');

        if (!record[name]) return;
        let typeKey = null;
        if (item === '金額') typeKey = 'amt';
        else if (item === '総本数') typeKey = 'qty';

        if (typeKey) {
            const t = record[name][typeKey];
            t.total = parseNumber(row[cols.total]);
            t.dairy = parseNumber(row[cols.dairy]);
            t.y400 = parseNumber(row[cols.y400]);

            // Switch Y1000 column based on name (Home vs Direct)
            if (name === CONFIG.NAME_S1) {
                t.y1000 = parseNumber(row[cols.y1000_home]);
            } else {
                t.y1000 = parseNumber(row[cols.y1000_direct]);
            }
        }
    });

    const v = (val) => val || 0;
    const s1 = record[CONFIG.NAME_S1];
    const s2 = record[CONFIG.NAME_S2];
    const allTotalAmt = v(s1.amt.total) + v(s2.amt.total);
    const allDairyAmt = v(s1.amt.dairy) + v(s2.amt.dairy);
    const allDairyQty = v(s1.qty.dairy) + v(s2.qty.dairy);

    const rowValues = [
        v(s1.amt.total),
        v(s1.amt.dairy), v(s1.qty.dairy),
        v(s1.amt.y400), v(s1.qty.y400),
        v(s1.amt.y1000), v(s1.qty.y1000),
        v(s2.amt.total),
        v(s2.amt.dairy), v(s2.qty.dairy),
        v(s2.amt.y400), v(s2.qty.y400),
        v(s2.amt.y1000), v(s2.qty.y1000),
        allTotalAmt, allDairyAmt, allDairyQty
    ];

    return rowValues;
}

function processFile(file, ss) {
    console.log('Processing: ' + file.getName());
    const rowValues = parseCsvToRecord(file);
    if (!rowValues) return;

    let date = extractDateFromFilename(file.getName());
    if (!date) date = new Date();

    writeToMonthlySheet(ss, date, rowValues);
    moveFileToProcessed(file);
}

// New: Import Monthly Fix Data
function importMonthlyFixCSV() {
    const folders = DriveApp.getFoldersByName(CONFIG.FOLDER_MONTHLY_FIX);
    if (!folders.hasNext()) return;
    const folder = folders.next();
    const files = folder.getFilesByType(MimeType.CSV);
    const ss = getOrCreateSpreadsheet();
    let processedAny = false;

    while (files.hasNext()) {
        try {
            const file = files.next();
            console.log('Processing Fix: ' + file.getName());
            const rowValues = parseCsvToRecord(file);
            if (!rowValues) continue;

            // Extract YYYYMM
            const m = file.getName().match(/(\d{4})(\d{2})/);
            if (!m) {
                console.log('Skipping file (no YYYYMM match): ' + file.getName());
                continue;
            }

            const year = parseInt(m[1], 10);
            const month = parseInt(m[2], 10);
            const date = new Date(year, month - 1, 1);

            if (isNaN(date.getTime())) {
                console.error('Invalid date parsed from: ' + file.getName());
                continue;
            }

            overwriteMonthlySheet(ss, date, rowValues);
            moveFileToProcessed(file);
            processedAny = true;
        } catch (e) {
            console.error('Error: ' + e.message);
        }
    }
    if (processedAny) {
        updateSummarySheet();
        updateMonthlyReportSheet();
        logOrAlert('月次確定取り込みが完了しました。');
    }
}

function overwriteMonthlySheet(ss, date, dataArray) {
    const sheetName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy_MM');
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = createMonthlySheet(ss, sheetName, date);

    // Clear B2:R32 (Day 1 to 31)
    sheet.getRange(2, 1, 31, 18).clearContent();
    // Re-fill dates
    const days = [];
    for (let i = 1; i <= 31; i++) days.push([i]);
    sheet.getRange(2, 1, 31, 1).setValues(days);

    // Write Fix to Day 1 (Row 2)
    sheet.getRange(2, 2, 1, 17).setValues([dataArray]);
    sheet.getRange(2, 1).setValue('確定値');

    // Update formulas just in case
    updateAchievementFormulas(ss, sheet, date);
}

function writeToMonthlySheet(ss, date, dataArray) {
    const sheetName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy_MM');
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = createMonthlySheet(ss, sheetName, date);

    // Ensure structure is up to date (add S-U if missing, update formulas)
    updateAchievementFormulas(ss, sheet, date);

    const day = date.getDate();
    const row = day + 1;
    if (row <= 32) sheet.getRange(row, 2, 1, dataArray.length).setValues([dataArray]);
}

function updateAchievementFormulas(ss, sheet, date) {
    // 1. Ensure Headers
    const headers = ['S_宅配率', 'T_直販率', 'U_全体率'];
    sheet.getRange(1, 19, 1, 3).setValues([headers]).setFontWeight('bold').setBackground('#ddd');

    // 2. Fetch Data
    const key = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy_MM');
    const targetMap = getMasterValues(ss, CONFIG.SHEET_TARGET);
    const daysMap = getWorkingDays(ss);

    // 3. Calc Daily Targets
    let dailyS1 = 0, dailyS2 = 0, dailyTotal = 0;

    if (targetMap[key] && daysMap[key]) {
        const tVals = targetMap[key];
        const wDays = daysMap[key];

        const tgtS1 = tVals[0] || 0;
        const daysS1 = wDays.home || 0;
        dailyS1 = daysS1 > 0 ? tgtS1 / daysS1 : 0;

        const tgtS2 = tVals[7] || 0;
        const daysS2 = wDays.direct || 0;
        const adjS2 = Math.max(0, tgtS2 - 6000000);
        dailyS2 = daysS2 > 0 ? adjS2 / daysS2 : 0;

        dailyTotal = dailyS1 + dailyS2;
    }

    // 4. Set Formulas (S2:U31 in general usually, but let's cover 2 to 32 just in case)
    const formulas = [];
    for (let i = 0; i < 31; i++) {
        const r = i + 2;
        const fS = dailyS1 > 0 ? `=B${r}/${dailyS1}` : '=0';
        const fT = dailyS2 > 0 ? `=I${r}/${dailyS2}` : '=0';
        const fU = dailyTotal > 0 ? `=P${r}/${dailyTotal}` : '=0';
        formulas.push([fS, fT, fU]);
    }
    sheet.getRange(2, 19, 31, 3).setFormulas(formulas).setNumberFormat('0.0%');
}

function createMonthlySheet(ss, sheetName, date) {
    const sheet = ss.insertSheet(sheetName);
    const headers = [
        'Date',
        '宅配_全_金', '宅配_乳_金', '宅配_乳_本', '宅配_400_金', '宅配_400_本', '宅配_1000_金', '宅配_1000_本',
        '直販_全_金', '直販_乳_金', '直販_乳_本', '直販_400_金', '直販_400_本', '直販_1000_金', '直販_1000_本',
        'R_全社_金', 'S_全乳_金', 'T_全乳_本',
        'S_宅配率', 'T_直販率', 'U_全体率'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#ddd');

    const days = [];
    for (let i = 1; i <= 31; i++) days.push([i]);
    sheet.getRange(2, 1, 31, 1).setValues(days);

    updateAchievementFormulas(ss, sheet, date);

    sheet.getRange(33, 1).setValue('月計・平均').setFontWeight('bold');

    const qtyIndices = [2, 4, 6, 9, 11, 13, 16];
    const footerFormulas = [];
    for (let i = 0; i < 17; i++) {
        if (qtyIndices.includes(i)) footerFormulas.push('=AVERAGE(R[-31]C:R[-1]C)');
        else footerFormulas.push('=SUM(R[-31]C:R[-1]C)');
    }
    sheet.getRange(33, 2, 1, 17).setFormulasR1C1([footerFormulas]).setFontWeight('bold').setBackground('#fff0f0');
    return sheet;
}

function createEmptyMonthSheetUI() {
    const ui = SpreadsheetApp.getUi();
    const result = ui.prompt('シート作成', '作成する年月を入力してください (例: 2024_04)', ui.ButtonSet.OK_CANCEL);

    if (result.getSelectedButton() == ui.Button.OK) {
        const sheetName = result.getResponseText().trim();
        if (!sheetName.match(/^\d{4}_\d{2}$/)) {
            ui.alert('形式が正しくありません。半角で "YYYY_MM" の形式で入力してください。\n例: 2024_04');
            return;
        }

        const ss = getOrCreateSpreadsheet();
        if (ss.getSheetByName(sheetName)) {
            ui.alert('シート "' + sheetName + '" は既に存在します。');
            return;
        }

        const parts = sheetName.split('_');
        const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);

        createMonthlySheet(ss, sheetName, date);
        ui.alert('シート "' + sheetName + '" を作成しました。\nデータを貼り付けた後、「サマリーのみ更新」を実行してください。');
    }
}

// UIが動かない場合の非常用（コード内の createTarget を書き換えて直接実行してください）
function forceCreateSheet() {
    const createTarget = '2024_04'; // ←ここを作りたい年月に書き換える

    const ss = getOrCreateSpreadsheet();
    if (ss.getSheetByName(createTarget)) {
        console.log('エラー: シート ' + createTarget + ' は既に存在します。');
        return;
    }
    const parts = createTarget.split('_');
    const date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, 1);
    createMonthlySheet(ss, createTarget, date);
    console.log('シート ' + createTarget + ' を作成しました。');
}

function updateSummarySheet() {
    const ss = getOrCreateSpreadsheet();
    const sSheet = ss.getSheetByName(CONFIG.SHEET_SUMMARY);
    if (!sSheet) return;

    sSheet.getRange('B3').setValue(new Date());
    const baseDate = new Date(sSheet.getRange('B3').getValue());
    const fy = getFiscalYear(baseDate);
    const targetMap = getMasterValues(ss, CONFIG.SHEET_TARGET);
    const prevMap = getMasterValues(ss, CONFIG.SHEET_PREV);

    const metrics = [
        '宅配 全社売上', '宅配 乳製品売上', '宅配 乳製品本数', '宅配 Y400売上', '宅配 Y400本数', '宅配 Y1000売上', '宅配 Y1000本数',
        '直販 全社売上', '直販 乳製品売上', '直販 乳製品本数', '直販 Y400売上', '直販 Y400本数', '直販 Y1000売上', '直販 Y1000本数',
        '全社売上計', '全社乳製品計', '全社乳製品本数'
    ];
    // Qty indices for Average calculation
    const qtyIndices = [2, 4, 6, 9, 11, 13, 16];

    // Stats now tracks SUM and COUNT (for averages)
    const createStatObj = () => ({ sums: createZeroStats(), counts: createZeroStats() });
    const initStats = () => ({ actual: createStatObj(), target: createStatObj(), prev: createStatObj() });

    const sumRange = (startMonthIndex, duration) => {
        let stats = initStats();
        for (let i = 0; i < duration; i++) {
            const mOffset = startMonthIndex + i;
            const d = new Date(fy, 3 + mOffset, 1);

            if (d <= baseDate) {
                const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy_MM');
                const totals = getSheetTotals(ss, key); // Returns {sums, counts}

                // Change: Use total days of month for Qty counts, not data counts
                const dim = new Date(fy, 3 + mOffset + 1, 0).getDate();
                qtyIndices.forEach(idx => { totals.counts[idx] = dim; });

                // Actuals: Accumulate Sum and Count of valid days (for calculating Daily Average)
                addToStatObj(stats.actual, totals.sums, totals.counts);

                // Target/Prev: Values are ALREADY Monthly Averages (for Qty) or Totals (for Amt).
                // For Qty Average over Quarter/Year: We sum the Monthly Averages and divide by Month Count.
                // For Amt Total: We just sum.
                const tgtVals = targetMap[key] || createZeroStats();
                const countArr = new Array(17).fill(1); // Count as 1 month
                addToStatObj(stats.target, tgtVals, countArr);

                const prevD = new Date(d.getFullYear() - 1, d.getMonth(), 1);
                const prevKey = Utilities.formatDate(prevD, Session.getScriptTimeZone(), 'yyyy_MM');
                const prevVals = prevMap[prevKey] || createZeroStats();
                addToStatObj(stats.prev, prevVals, countArr);
            }
        }
        return stats;
    };

    // Calc Stats
    const q1Stats = sumRange(0, 3);
    const q2Stats = sumRange(3, 3);
    const q3Stats = sumRange(6, 3);
    const q4Stats = sumRange(9, 3);

    const h1Stats = sumRange(0, 6);
    const h2Stats = sumRange(6, 6);

    const yStats = sumRange(0, 12);

    // Rendering Function
    const drawBlock = (title, startRow, stats) => {
        sSheet.getRange(startRow, 1).setValue(title).setFontWeight('bold');
        const headerRow = ['項目', '実績', '目標', '達成率', '前年実績', '前年比'];
        sSheet.getRange(startRow + 1, 1, 1, 6).setValues([headerRow]).setBackground('#dedede').setFontWeight('bold');

        const rows = [];
        for (let i = 0; i < 17; i++) {
            // Calculate Value (Average for Qty, Sum for Amt)
            const isQty = qtyIndices.includes(i);

            let act = stats.actual.sums[i];
            let tgt = stats.target.sums[i];
            let prv = stats.prev.sums[i];

            if (isQty) {
                // Actual Qty: Divide by Day Count (Sum / Days)
                const actCount = stats.actual.counts[i];
                act = actCount > 0 ? act / actCount : 0;

                // Target Qty: Divide by Month Count (Sum / Months). 
                // Because CSV values are ALREADY averages.
                const tgtCount = stats.target.counts[i];
                tgt = tgtCount > 0 ? tgt / tgtCount : 0;

                const prvCount = stats.prev.counts[i];
                prv = prvCount > 0 ? prv / prvCount : 0;
            }

            const rate = tgt ? act / tgt : 0;
            const yoy = prv ? act / prv : 0;
            rows.push([metrics[i], act, tgt, rate, prv, yoy]);
        }
        sSheet.getRange(startRow + 2, 1, 17, 6).setValues(rows);
        sSheet.getRange(startRow + 2, 2, 17, 1).setNumberFormat('#,##0');
        sSheet.getRange(startRow + 2, 3, 17, 1).setNumberFormat('#,##0');
        sSheet.getRange(startRow + 2, 5, 17, 1).setNumberFormat('#,##0');
        sSheet.getRange(startRow + 2, 4, 17, 1).setNumberFormat('0.0%');
        sSheet.getRange(startRow + 2, 6, 17, 1).setNumberFormat('0.0%');
        return startRow + 20;
    };

    let r = 5;
    r = drawBlock('【第1四半期実績】(4月-6月)', r, q1Stats);
    r = drawBlock('【第2四半期実績】(7月-9月)', r, q2Stats);
    r = drawBlock('【第3四半期実績】(10月-12月)', r, q3Stats);
    r = drawBlock('【第4四半期実績】(1月-3月)', r, q4Stats);
    r = drawBlock('【上期実績】(4月-9月)', r, h1Stats);
    r = drawBlock('【下期実績】(10月-3月)', r, h2Stats);
    r = drawBlock('【年間累計実績】', r, yStats);

    // Clear remaining rows if space exists
    const maxRows = sSheet.getMaxRows();
    if (r <= maxRows) {
        sSheet.getRange(r, 1, maxRows - r + 1, sSheet.getLastColumn()).clearContent();
    }

    sSheet.getRange('A1').setValue(fy + '年度 全社販売実績ダッシュボード');
    console.log('Summary updated.');
}

function updateMonthlyReportSheet() {
    const ss = getOrCreateSpreadsheet();
    let mSheet = ss.getSheetByName(CONFIG.SHEET_MONTHLY_REPORT);
    if (!mSheet) {
        mSheet = ss.insertSheet(CONFIG.SHEET_MONTHLY_REPORT);
        mSheet.getRange('A1').setValue('月別実績レポート').setFontSize(14).setFontWeight('bold');
        mSheet.getRange('A3').setValue('基準日:');
        mSheet.getRange('B3').setValue(new Date());
    }

    const baseDate = new Date(); // Always use today effectively, or check B3 if needed. 
    // Ideally we sync with summary logic, but here we just show ALL months for the fiscal year of today.
    const fy = getFiscalYear(baseDate);

    // Refresh B3
    mSheet.getRange('B3').setValue(baseDate);
    mSheet.getRange('A1').setValue(fy + '年度 月別実績レポート');

    const targetMap = getMasterValues(ss, CONFIG.SHEET_TARGET);
    const prevMap = getMasterValues(ss, CONFIG.SHEET_PREV);

    const metrics = [
        '宅配 全社売上', '宅配 乳製品売上', '宅配 乳製品本数', '宅配 Y400売上', '宅配 Y400本数', '宅配 Y1000売上', '宅配 Y1000本数',
        '直販 全社売上', '直販 乳製品売上', '直販 乳製品本数', '直販 Y400売上', '直販 Y400本数', '直販 Y1000売上', '直販 Y1000本数',
        '全社売上計', '全社乳製品計', '全社乳製品本数'
    ];
    const qtyIndices = [2, 4, 6, 9, 11, 13, 16];
    const createStatObj = () => ({ sums: createZeroStats(), counts: createZeroStats() });
    const initStats = () => ({ actual: createStatObj(), target: createStatObj(), prev: createStatObj() });

    // Reuse similar logic, simplified for single month
    const getMonthStats = (monthIndex) => {
        let stats = initStats();
        // monthIndex 0 = Apr, 1 = May...
        const d = new Date(fy, 3 + monthIndex, 1);

        // Always try to fetch actuals regardless of date? 
        // Or only if d <= baseDate?  User said "Show all months", implies showing empty/zeros for future is fine or expected.
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy_MM');

        // Actuals
        const totals = getSheetTotals(ss, key);

        // Change: Use total days of month for Qty counts
        const dim = new Date(fy, 3 + monthIndex + 1, 0).getDate();
        qtyIndices.forEach(idx => { totals.counts[idx] = dim; });

        addToStatObj(stats.actual, totals.sums, totals.counts);

        // Target/Prev
        const countArr = new Array(17).fill(1);
        const tgtVals = targetMap[key] || createZeroStats();
        addToStatObj(stats.target, tgtVals, countArr);

        const prevD = new Date(d.getFullYear() - 1, d.getMonth(), 1);
        const prevKey = Utilities.formatDate(prevD, Session.getScriptTimeZone(), 'yyyy_MM');
        const prevVals = prevMap[prevKey] || createZeroStats();
        addToStatObj(stats.prev, prevVals, countArr);

        return stats;
    };

    const drawBlock = (title, startRow, stats) => {
        mSheet.getRange(startRow, 1).setValue(title).setFontWeight('bold');
        const headerRow = ['項目', '実績', '目標', '達成率', '前年実績', '前年比'];
        mSheet.getRange(startRow + 1, 1, 1, 6).setValues([headerRow]).setBackground('#e0f0ff').setFontWeight('bold');

        const rows = [];
        for (let i = 0; i < 17; i++) {
            const isQty = qtyIndices.includes(i);
            let act = stats.actual.sums[i];
            let tgt = stats.target.sums[i];
            let prv = stats.prev.sums[i];

            if (isQty) {
                const actCount = stats.actual.counts[i];
                act = actCount > 0 ? act / actCount : 0;
                const tgtCount = stats.target.counts[i];
                tgt = tgtCount > 0 ? tgt / tgtCount : 0;
                const prvCount = stats.prev.counts[i];
                prv = prvCount > 0 ? prv / prvCount : 0;
            }

            const rate = tgt ? act / tgt : 0;
            const yoy = prv ? act / prv : 0;
            rows.push([metrics[i], act, tgt, rate, prv, yoy]);
        }
        mSheet.getRange(startRow + 2, 1, 17, 6).setValues(rows);
        mSheet.getRange(startRow + 2, 2, 17, 1).setNumberFormat('#,##0');
        mSheet.getRange(startRow + 2, 3, 17, 1).setNumberFormat('#,##0');
        mSheet.getRange(startRow + 2, 5, 17, 1).setNumberFormat('#,##0');
        mSheet.getRange(startRow + 2, 4, 17, 1).setNumberFormat('0.0%');
        mSheet.getRange(startRow + 2, 6, 17, 1).setNumberFormat('0.0%');
        return startRow + 20;
    };

    let r = 5;
    for (let i = 0; i < 12; i++) {
        const m = (i + 3) % 12 + 1; // 0->4, 8->12, 9->1
        const stats = getMonthStats(i);
        r = drawBlock('【' + m + '月実績】', r, stats);
    }

    // Clear rest
    const maxRows = mSheet.getMaxRows();
    if (r <= maxRows) {
        mSheet.getRange(r, 1, maxRows - r + 1, mSheet.getLastColumn()).clearContent();
    }
    console.log('Monthly Report updated.');
}

function getMasterValues(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const data = sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    const map = {};
    data.forEach(row => {
        if (!row[0]) return;
        const d = new Date(row[0]);
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy_MM');
        const values = row.slice(1, 18).map(v => parseNumber(v));
        map[key] = values;
    });
    return map;
}

function getFiscalYear(date) {
    const m = date.getMonth();
    return (m < 3) ? date.getFullYear() - 1 : date.getFullYear();
}

// Updated to return sums AND counts
function getSheetTotals(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { sums: createZeroStats(), counts: createZeroStats() }; // Empty stats

    const data = sheet.getRange(2, 2, 31, 17).getValues();
    const sums = createZeroStats();
    const counts = createZeroStats();

    for (let r = 0; r < 31; r++) {
        for (let c = 0; c < 17; c++) {
            const val = data[r][c];
            if (typeof val === 'number') { // Only count if number (even 0)
                sums[c] += val;
                // Treat 0 as valid data? Usually yes for daily logs.
                // Assuming empty cells are '' (string).
                if (val !== '' && val !== null) counts[c]++;
            }
        }
    }
    return { sums, counts };
}

// Working Days Map: { 'yyyy_MM': { home: 20, direct: 22 } }
function getWorkingDays(ss) {
    const sheet = ss.getSheetByName(CONFIG.SHEET_DAYS);
    if (!sheet) return {};
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    // Format: A:Month, B:Home, C:Direct
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const map = {};
    data.forEach(row => {
        if (!row[0]) return;
        const d = new Date(row[0]);
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy_MM');
        map[key] = {
            home: Number(row[1]) || 0,
            direct: Number(row[2]) || 0
        };
    });
    return map;
}

function createWorkingDaysSheet(ss) {
    let sheet = ss.getSheetByName(CONFIG.SHEET_DAYS);
    if (sheet) return;

    sheet = ss.insertSheet(CONFIG.SHEET_DAYS);
    sheet.getRange(1, 1, 1, 3).setValues([['年月', '稼働日_宅配', '稼働日_直販']]).setFontWeight('bold').setBackground('#c9daf8');

    const today = new Date();
    const fy = getFiscalYear(today);
    const rows = [];
    for (let i = 0; i < 12; i++) {
        const d = new Date(fy, 3 + i, 1);
        rows.push([d, 0, 0]);
    }
    sheet.getRange(2, 1, 12, 3).setValues(rows);
    sheet.getRange('A:A').setNumberFormat('yyyy/MM');
    sheet.setColumnWidth(1, 100);
}

function importDaysCSV(ss, file) {
    if (!ss) return;
    const sheet = ss.getSheetByName(CONFIG.SHEET_DAYS);
    if (!sheet) return;

    const csvString = file.getBlob().getDataAsString('Shift_JIS');
    const data = Utilities.parseCsv(csvString);
    let startRow = 0;
    if (data.length > 0 && isNaN(Date.parse(data[0][0]))) startRow = 1;

    const rows = [];
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        if (row.length < 3) continue;
        const d = new Date(row[0]);
        if (isNaN(d.getTime())) continue;
        rows.push([d, row[1], row[2]]); // Date, Home, Direct
    }

    if (rows.length === 0) return;

    // Simple overwrite logic for now, or match/replace?
    // Let's clear and overwrite for simplicity as it's a master import
    if (sheet.getLastRow() >= 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    console.log('Imported Working Days.');
}

function createZeroStats() { return new Array(17).fill(0); }
function addToStatObj(targetObj, sums, counts) {
    for (let i = 0; i < 17; i++) {
        targetObj.sums[i] += (Number(sums[i]) || 0);
        targetObj.counts[i] += (Number(counts[i]) || 0);
    }
}
function findHeaderIndex(h, c) {
    for (const x of c) { const i = h.indexOf(x); if (i !== -1) return i; }
    return -1;
}
function extractDateFromFilename(n) {
    const m = n.match(/(\d{4})(\d{2})(\d{2})/);
    if (m) return new Date(m[1], m[2] - 1, m[3]);
    return null;
}
function parseNumber(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'string') return Number(v.replace(/,/g, '').trim()) || 0;
    return 0;
}
function getOrCreateSpreadsheet() {
    const fs = DriveApp.getFilesByName(CONFIG.SS_NAME);
    while (fs.hasNext()) {
        const f = fs.next();
        if (!f.isTrashed() && f.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.openById(f.getId());
    }
    return SpreadsheetApp.create(CONFIG.SS_NAME);
}
function moveFileToProcessed(f) {
    const fs = DriveApp.getFoldersByName(CONFIG.PROCESSED_FOLDER_NAME);
    const folder = fs.hasNext() ? fs.next() : DriveApp.createFolder(CONFIG.PROCESSED_FOLDER_NAME);
    f.moveTo(folder);
}
function ensureFolder(name) {
    const fs = DriveApp.getFoldersByName(name);
    if (!fs.hasNext()) DriveApp.createFolder(name);
}
