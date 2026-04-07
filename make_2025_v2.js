const fs = require('fs');

let text = fs.readFileSync('Code.js', 'utf8');

text = text.replace(/'全社実績ダッシュボード_2026年度'/g, "'全社実績ダッシュボード'");
text = text.replace(/'実績CSVアップロード_2026年度'/g, "'実績CSVアップロード_2025年度'");
text = text.replace(/'processed_2026年度'/g, "'processed_2025年度'");
text = text.replace(/'月次確定CSVアップロード_2026年度'/g, "'月次確定CSVアップロード_2025年度'");
text = text.replace(/'マスタデータ取込_2026年度'/g, "'マスタデータ取込_2025年度'");

const s_target = "    sSheet.getRange('B3').setValue(new Date());\n    const baseDate = new Date(sSheet.getRange('B3').getValue());\n    const fy = getDashboardFiscalYear(ss);";
const s_repl = "    // 2025年度完全固定（B3の日付も2026年3月末に固定）\n    const baseDate = new Date(2026, 2, 31);\n    sSheet.getRange('B3').setValue(baseDate);\n    const fy = 2025;";
text = text.replace(s_target, s_repl);

const s_target_rn = s_target.replace(/\n/g, '\r\n');
const s_repl_rn = s_repl.replace(/\n/g, '\r\n');
text = text.replace(s_target_rn, s_repl_rn);

const m_target = "    const baseDate = new Date(); // Always use today effectively, or check B3 if needed. \n    // Ideally we sync with summary logic, but here we just show ALL months for the fiscal year of today.\n    const fy = getDashboardFiscalYear(ss);";
const m_repl = "    // 2025年度完全固定\n    const baseDate = new Date(2026, 2, 31);\n    const fy = 2025;";
text = text.replace(m_target, m_repl);

const m_target_rn = m_target.replace(/\n/g, '\r\n');
const m_repl_rn = m_repl.replace(/\n/g, '\r\n');
text = text.replace(m_target_rn, m_repl_rn);

fs.writeFileSync('Code_2025_fixed_v2.js', text, 'utf8');
console.log('Fixed v2 generated');
