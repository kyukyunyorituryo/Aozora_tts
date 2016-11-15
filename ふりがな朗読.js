// -----------------------------------------------------------------------------
// 音声読み上げ＋ルビ振り
// 説明 選択した文字をYahooAPIで漢字に読み仮名をつけて音声を読み上げる。
//
// 参考
// http://qiita.com/tnakagawa/items/3bce99d49b1aa3fc9a72
// http://qiita.com/tnakagawa/items/4b501c21abcd39f30fbe
// 使用API
// http://developer.yahoo.co.jp/webapi/jlp/furigana/v1/furigana.html
// -----------------------------------------------------------------------------

//読み上げ２.js
var Grade =1;
var text = document.selection.Text;
var str ="";//出力される文字

var API_URL = "http://jlp.yahooapis.jp/FuriganaService/V1/furigana";
var Appid = "dj0zaiZpPVJqcVRmNUk2S0p1SSZzPWNvbnN1bWVyc2VjcmV0Jng9MTc-";//変更する場合はappidを取得してください



// -----------------------------------------------------------------------------
try {
    // 「ServerXMLHTTP」オブジェクト生成
    var http = new ActiveXObject("Msxml2.ServerXMLHTTP");
    // 要求初期化
    http.open("POST", API_URL, false);
     // 要求ヘッダ設定
    http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    // 要求
    var params = { appid: Appid, grade: Grade, sentence: text};
    http.send(escapeParams(params));
    // 応答結果表示
    
var x = http.responseText;
//document.selection.Text = x;

} catch (e) {
    // エラーの場合
    Alert("リクエスト失敗");
}

function escapeParams(params) {
    var param = "";
    // パラメータ数分ループ
    for (var key in params) {
        // 連結チェック
        if (param.length > 0) {
            param += "&";
        }
        // パラメータ設定
        param += encodeURIComponent(key).split("%20").join("+")
            + "=" + encodeURIComponent(params[key]).split("%20").join("+");
    }
    return param;
}
// DOMオブジェクト生成
var dom = new ActiveXObject("Msxml2.DOMDocument");
// 同期化
dom.async = false;
// パース
dom.loadXML(x);
if (dom.parseError.errorCode == 0) {
    // XML出力
    //Alert(dom.xml);
   }
    var root = dom.documentElement;
    // タグ名がWordのエレメント取得
    var elements = root.getElementsByTagName("Word");
//        Alert(elements.length);
        //
    for (var i = 0; i < elements.length; i++) {
        // エレメント取得
        var element = elements[i];
        // 子取得
        var child = element.childNodes;
        switch (element.childNodes.length) {
            case 1: str += child[0].text; break;
            case 3:
                if (child[0].text == child[1].text) {
                    str += child[1].text;
                } else {
                    str += child[1].text;
                } break;
            case 4:
                if (child[0].text == child[1].text) {
                    str += child[1].text;
                } else {
                    str += child[1].text;
                } break;
        }
    }

//document.selection.Text = str;
var spkr = new ActiveXObject('SAPI.SpVoice');
  var selword = str;
spkr.rate = 2; //読み上げ速度：大きいと速い
spkr.Speak(selword);
