// CAUTION
// 
// このファイルは文字コードを SJIS として保存すること。
// (SJIS 形式で保存しないと、`WScript.Echo` などで文字化けする)

// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

var str_crlf = "\r\n";

//ActiveXObject生成
var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject生成
try{
	var itobj = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>
}catch(e){
	WScript.Echo("Cannot create object `iTunes.Application`");
	WScript.Quit(); // 終了
}

var track = itobj.CurrentTrack;    
var info;

//ファイル・フォルダ
var mydocu = wshobj.SpecialFolders("MyDocuments");//マイドキュメント場所
var nowpfol = "iTunesNowPlaying";//専用フォルダ名
var nowpfil = "NowPlaying.txt";//専用ファイル名

//フォルダ存在確認
if(!(axobj.FolderExists(mydocu + "\\" + nowpfol))){
	axobj.CreateFolder(mydocu + "\\" + nowpfol);//フォルダ作成
}

//曲情報収集
try{
	if(track.Kind == 1){ //ローカル再生中
		info=track.Artist + "\t" + track.Name;
	}else if(track.Kind == 3) {//ストリーム再生中
		var titles = itobj.currentStreamTitle.split(",");
		info=titles[0];
	}else{
		WScript.Echo("不明な動作中");
		WScript.Quit(); // 終了
	}
}catch(e){
	if(e == "[object Error]"){
		var str_errMsg =
			"`" + e + "` detected." + str_crlf + str_crlf +
			"iTunesが動作していません"
		;
		WScript.Echo(str_errMsg);

	}else{
		var str_errMsg =
			"`" + e + "` detected." + str_crlf + str_crlf +
			WScript.Echo("Unkown Error.");
		;
	}
	WScript.Quit(); // 終了
}

//ファイルオープン
try{
	//https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
	var txfl = axobj.OpenTextFile(mydocu + "\\" + nowpfol + "\\" + nowpfil, 8, true);
}catch(e){
	WScript.Echo(nowpfil + "が開けません");
	WScript.Quit(); // 終了
}

//ファイルへ書き込み
txfl.Write(info + "\n");

//ファイルクローズ
txfl.Close();
