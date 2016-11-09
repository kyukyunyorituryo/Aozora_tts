// -----------------------------------------------------------------------------
// 青空文庫朗読
// 説明 選択した文字列の青空文庫のルビを開いて読み上げる
//
// -----------------------------------------------------------------------------
var str = document.selection.Text;
str = AozoraRuby(str);
var spkr = new ActiveXObject('SAPI.SpVoice');
spkr.rate = 2; //読み上げ速度：大きいと速い
//alert(str);
spkr.Speak(str);

function AozoraRuby(str) {
str = str.replace(/｜(.+?)《(.+?)》/mg, "$2");
    //カタカナにふりがな
    str = str.replace(/([ァ-ヺー]+?)《(.+?)》/mg, "$2");
    //ひらがなにふりがな
    str = str.replace(/([ぁ-ゖ]+?)《(.+?)》/mg, "$2");
    //漢字にふりがな
    str = str.replace(/([々仝〆〇ヶ\u3400-\u4DBF\u4E00-\u9FFF\uD840-\uD87F\uDC00-\uDFFF\uF900-\uFAFF]+?)《(.+?)》/mg, "$2");
    //アルファベットにふりがな
    str = str.replace(/([A-Za-z]+?)《(.+?)》/mgi, "$2");
    return str;
}
/*
Unicodeプロパティが使えない場合の正規表現
http://so-zou.jp/software/tech/programming/tech/regular-expression/meta-character/variable-width-encoding.htm
http://d.hatena.ne.jp/kazuhooku/20090723/1248309720
http://tools.m-bsys.com/data/charlist_kana.php
ひらがな：[ぁ-ゖ]
カタカナ：[ァ-ヺー]
使用している省略漢字：[\u3400-\u4DBF\u4E00-\u9FFF\uD840-\uD87F\uDC00-\uDFFF\uF900-\uFAFF]+
欠けている漢字：々仝〆〇ヶ
*/
