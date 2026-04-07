const fs = require('fs');
let text = fs.readFileSync('Code.js', 'utf8');

const regex = /new_yakult:\s*findHeaderIndex\(header,\s*\[(.*?)\]\)/;
const match = text.match(regex);
if (match) {
    const list = match[1];
    if (!list.includes("'Newﾔｸﾙﾄ類'")) {
        const newList = list + ", 'Newﾔｸﾙﾄ類', 'Newﾔｸﾙﾄ'";
        const replaceString = "new_yakult: findHeaderIndex(header, [" + newList + "])";
        text = text.replace(match[0], replaceString);
        fs.writeFileSync('Code.js', text, 'utf8');
        console.log('Successfully added half-width katakana Yakult to list.');
    } else {
        console.log('Already added.');
    }
} else {
    console.log('Could not find RegExp match.');
}
