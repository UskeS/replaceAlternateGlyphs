/**
 * @fileoverview aalt/nalt字形を基底グリフに置換するInDesign用スクリプト
 * @copyright Yusuke SAEGUSA 2020 <https://twitter.com/Uske_S>
 * @version 0.1.0
 * @license Apache-2.0 
 * @description monokanoさん作成のAppleScript(https://gist.github.com/monokano/1f7ba52c50e71494486636ff9e3f7889)をjsx化したもの  
 */

if (app.documents.length === 0 || app.selection.length === 0) {
    alert("テキストを選択してから実行してください");
    exit();
}

var doc = app.activeDocument;
var sels = doc.selection;
app.doScript(main, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT);

function main() {
    for (var i = 0; i < sels.length; i++) {
        if (!sels[i].hasOwnProperty("findGrep")) { continue; }
        var tgtChrs = sels[i].characters;
        for (var j = 0, lenJ = tgtChrs.length; j < lenJ; j++) {
            var tgtOtf = tgtChrs[j].opentypeFeatures[0]; // openTypeFeatureが入るときは配列になる
            if (tgtOtf && (tgtOtf[0] === "aalt" || tgtOtf[0] === "nalt")) { reWrite(tgtChrs[j]); }
        }
    }
}

function reWrite(chr) {
    var baseContent = chr.contents;
    // ルビがない場合
    if (!chr.rubyFlag || chr.rubyString === "") {
        chr.contents = baseContent;
        return;
    }
    // ルビがある場合
    var curParagraph = chr.paragraphs[0]; // 段落全体の情報を取得しておく
    var chrsRubyFlagInParagraph = {
        before: curParagraph.characters.everyItem().rubyFlag
    };
    var curRubyString = chr.rubyString;
    chr.insertionPoints[1].contents = baseContent;
    chr.characters[0].remove();
    chrsRubyFlagInParagraph.after = curParagraph.characters.everyItem().rubyFlag;
    var tgtChrIndexArray = [];
    for (var i = 0, len = chrsRubyFlagInParagraph.before.length; i < len; i++) {
        if (chrsRubyFlagInParagraph.before[i] !== chrsRubyFlagInParagraph.after[i]) {
            tgtChrIndexArray.push(i);
        }
    }
    try {
        curParagraph.characters.itemByRange(tgtChrIndexArray[0], tgtChrIndexArray[tgtChrIndexArray.length - 1]).properties = {
            rubyFlag: true,
            rubyString: curRubyString
        };
    } catch (e) {
        alert(e);
    }
}
