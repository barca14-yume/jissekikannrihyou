const fs = require('fs');

let text = fs.readFileSync('Code.js', 'utf8');

// 1. parseCsvToRecord cols object
const colsTarget = `        dairy: findHeaderIndex(header, ['乳製品計']),
        y400: findHeaderIndex(header, ['Y400類']),
        // 宅配用: yakult1000類 (小文字) 優先`;
const colsRepl = `        dairy: findHeaderIndex(header, ['乳製品計']),
        y400: findHeaderIndex(header, ['Y400類']),
        new_yakult: findHeaderIndex(header, ['Newヤクルト類', 'newヤクルト類', 'Ｎｅｗヤクルト類', 'Newヤクルト', 'newヤクルト', 'Ｎｅｗヤクルト', 'ＮＥＷヤクルト類', 'Newヤクルト類']),
        // 宅配用: yakult1000類 (小文字) 優先`;
text = text.replace(colsTarget, colsRepl);
text = text.replace(colsTarget.replace(/\n/g, '\r\n'), colsRepl.replace(/\n/g, '\r\n'));

// 2. parseCsvToRecord logic
const logicTarget = `        if (typeKey) {
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
        }`;
const logicRepl = `        if (typeKey) {
            const t = record[name][typeKey];
            t.total = parseNumber(row[cols.total]);
            t.dairy = parseNumber(row[cols.dairy]);

            // Switch Y400/NewYakult and Y1000 based on name (Home vs Direct)
            if (name === CONFIG.NAME_S1) {
                t.y400 = parseNumber(row[cols.y400]);
                t.y1000 = parseNumber(row[cols.y1000_home]);
            } else {
                t.y400 = parseNumber(row[cols.new_yakult]); // 直販はY400枠にNewヤクルトを入れる
                t.y1000 = parseNumber(row[cols.y1000_direct]);
            }
        }`;
text = text.replace(logicTarget, logicRepl);
text = text.replace(logicTarget.replace(/\n/g, '\r\n'), logicRepl.replace(/\n/g, '\r\n'));

// 3. metrics label replacement (occurs in 2 places)
const metricsTarget = `'直販 乳製品本数', '直販 Y400売上', '直販 Y400本数', '直販 Y1000売上'`;
const metricsRepl = `'直販 乳製品本数', '直販 ヤクルト売上', '直販 ヤクルト本数', '直販 Y1000売上'`;
// using global replace
text = text.split(metricsTarget).join(metricsRepl);

// 4. header label replacement (occurs in 3 places)
const headerTarget = `'直販_全_金', '直販_乳_金', '直販_乳_本', '直販_400_金', '直販_400_本', '直販_1000_金', '直販_1000_本'`;
const headerRepl = `'直販_全_金', '直販_乳_金', '直販_乳_本', '直販_ヤクルト_金', '直販_ヤクルト_本', '直販_1000_金', '直販_1000_本'`;
text = text.split(headerTarget).join(headerRepl);

fs.writeFileSync('Code.js', text, 'utf8');
console.log('Replacements applied successfully');
