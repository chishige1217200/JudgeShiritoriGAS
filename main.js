function setup() { // 初期設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('しりとり');
    if (sheet !== null)
        ss.deleteSheet(sheet);
    sheet = ss.insertSheet();
    sheet.setName('しりとり');
    sheet.getRange(1, 1, 1, 1).setValue('しりとり');
    sheet.getRange(1, 3, 1, 1).setValue('ひらがな，カタカナ以外を入力しないでください．');
    sheet.getRange(2, 3, 1, 1).setValue('成功').setFontColor('red');
    sheet.getRange(2, 4, 1, 1).setValue('失敗').setFontColor('blue');
}

function judge() {
    function convert_to_katakana(string) {
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
        console.log(tail + " " + head);
        if (tail.slice(-1) === head.slice(0, 1)) return true; // しりとりが成立する場合はTrue
        else return false;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('しりとり');
    if (sheet === null) {
        setup();
        return;
    }

    let lastRow = sheet.getLastRow();
    console.log(lastRow);
    let w1, w2;

    for (let i = 1; i < lastRow; i++) {
        w1 = sheet.getRange(i, 1, 1, 1).getValue();
        w1 = convert_to_katakana(w1);
        w2 = sheet.getRange(i + 1, 1, 1, 1).getValue();
        w2 = convert_to_katakana(w2);
        if (w2.slice(-1) === 'ん') sheet.getRange(i + 1, 2, 1, 1).setValue('GAME OVER').setHorizontalAlignment('center').setFontColor('black'); // Game Overを書き込み
        else if (tail_head_judge(w1, w2)) sheet.getRange(i + 1, 2, 1, 1).setValue('True').setFontColor('red'); // Trueを書き込み
        else sheet.getRange(i + 1, 2, 1, 1).setValue('False').setFontColor('blue'); // Falseを書き込み
    }

    // 集計機能がほしい
}