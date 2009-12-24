// dispkeyfunc3.mac.js
// dispkeyfunc3.macのためのJScript。SendKeyするだけ。

// キーコード発行モード。
//   2文字以上
//      Numから始まらない場合、大文字に変換して、{}でくくる。
//      右の1文字
//   配列から読み取る
//   nullの場合、key
// キーコード発行

var tblMod = new Array();
tblMod[0] = "";						tblMod[1] = "+";
tblMod[2] = "^";					tblMod[3] = "(+^)";
tblMod[4] = "%";					tblMod[5] = "(%+)";
tblMod[6] = "(^%)";				tblMod[7] = "(+^%)";

var tblKey = new Array();
tblKey["Bksp"] = "{BKSP}";
tblKey["Tab"] = "{TAB}";
tblKey["Enter"] = "{ENTER}";
tblKey["Esc"] = "{ESC}";
tblKey["Space"] = "{SPACE}";
tblKey["PgUp"] = "{PGUP}";
tblKey["PgDn"] = "{PGDN}";
tblKey["End"] = "{END}";
tblKey["Home"] = "{HOME}";
tblKey["Left"] = "{LEFT}";
tblKey["Up"] = "{UP}";
tblKey["Right"] = "{RIGHT}";
tblKey["Down"] = "{DOWN}";
tblKey["Ins"] = "{INS}";
tblKey["Del"] = "{DEL}";
tblKey["F1"] = "{F1}";
tblKey["F2"] = "{F2}";
tblKey["F3"] = "{F3}";
tblKey["F4"] = "{F4}";
tblKey["F5"] = "{F5}";
tblKey["F6"] = "{F6}";
tblKey["F7"] = "{F7}";
tblKey["F8"] = "{F8}";
tblKey["F9"] = "{F9}";
tblKey["F10"] = "{F10}";
tblKey["F11"] = "{F11}";
tblKey["F12"] = "{F12}";
tblKey["+"] = "{+}";
tblKey["^"] = "{^}";
tblKey["%"] = "{%}";
tblKey["~"] = "{~}";
tblKey["{"] = "{{}";
tblKey["}"] = "{}}";
tblKey["["] = "{[}";
tblKey["]"] = "{]}";

var strMod = tblMod[ parseInt(WScript.Arguments.Named.Item("mod")) ];
var strKey = WScript.Arguments.Named.Item("key");

if( tblKey[strKey] == null )
	strKey = strKey.substr(strKey.length - 1);		//右端の1文字を取得
else
	strKey = tblKey[strKey];

var objShell = WScript.CreateObject("WScript.Shell");
objShell.SendKeys( strMod + strKey );

WScript.Quit();		//正常終了
