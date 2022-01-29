function setup() { // 初期設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('しりとり');
    if (sheet !== null)
        ss.deleteSheet(sheet);
    sheet = ss.insertSheet();
    sheet.setName('しりとり');
    sheet.getRange(1, 1, 1, 1).setValue('しりとり');
    sheet.getRange(1, 3, 1, 1).setValue('ひらがな，カタカナ以外を入力しないでください．');
}

function judge() {
    function convert2katakana(string) {
        let newString = '';

        for (let i = 0; i < string.length; i++) {
            if (string[i] === '　' || string[i] === ' ') {
                newString += '　';
            }
            else if (string[i].charCodeAt() < 12353 || string[i].charCodeAt() > 12435) { // ひらがな以外の文字列の処理;
                newString += string[i];
            }
            else {
                newString += String.fromCharCode(string[i].charCodeAt() + 96);
            }
        }
        return newString;
    }

    function tail_head_judge(tail, head) {
        function isYoon(w) {
            if (w === 'ァ' || w === 'ィ' || w === 'ゥ' || w === 'ェ' || w === 'ォ' || w === 'ッ' || w === 'ャ' || w === 'ュ' || w === 'ョ') return true; // 拗音・促音（ぁ、ぃ、ぅ、ぇ、ぉ、ゃ、ゅ、ょ、っ）の場合
            return false;
        }

        console.log(tail + " " + head);
        if (tail.slice(-1) === 'ー') {
            if (tail.slice(tail.length - 2, tail.length - 1) === head.slice(0, 1)) return true; // 末尾が伸ばし棒のとき，1つ前の文字で
        }
        else if (isYoon(tail.slice(-1))) {
            if (tail.slice(tail.length - 2, tail.length) === head.slice(0, 2))
                return true; // 最後の文字が拗音・促音（ぁ、ぃ、ぅ、ぇ、ぉ、ゃ、ゅ、ょ、っ）の場合、そのままのルール（2文字判定）
        }
        else if (tail.slice(-1) === 'ヂ' && head.slice(0, 1) === 'ジ') return true; // 例外ルール（単語が少ないため）
        else if (tail.slice(-1) === 'ヅ' && head.slice(0, 1) === 'ズ') return true; // 例外ルール（単語が少ないため）

        if (tail.slice(-1) === head.slice(0, 1)) return true; // しりとりが成立する場合はTrue（標準的なルール）
        return false; // しりとりになっていない
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('しりとり');
    if (sheet === null) {
        setup();
        return;
    }

    let lastRow = sheet.getLastRow();
    console.log(lastRow);
    let data = sheet.getRange(1, 1, lastRow, 1).getValues();
    let wordIndex = [];
    for (let i = 0; i < data.length; i++)
        wordIndex.push(convert2katakana(data[i][0]));
    console.log(wordIndex);
    let w1, w2;

    for (let i = 1; i < lastRow; i++) {
        w1 = sheet.getRange(i, 1, 1, 1).getValue();
        w1 = convert2katakana(w1); // カタカナに変換
        w2 = sheet.getRange(i + 1, 1, 1, 1).getValue();
        w2 = convert2katakana(w2); // カタカナに変換
        //console.log(wordIndex.lastIndexOf(w2));
        if (w2.slice(-1) === 'ン' || wordIndex.lastIndexOf(w2) !== i) {
            sheet.getRange(i + 1, 2, 1, 1).setValue('GAME OVER').setHorizontalAlignment('center').setFontColor('black'); // Game Overを書き込み
            return;
        }
        else if (tail_head_judge(w1, w2)) sheet.getRange(i + 1, 2, 1, 1).setValue('True').setFontColor('red'); // Trueを書き込み
        else sheet.getRange(i + 1, 2, 1, 1).setValue('False').setFontColor('blue'); // Falseを書き込み
    }
}