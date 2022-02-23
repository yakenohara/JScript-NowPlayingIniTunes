// 再生中のトラックの曲名、アーティストを TSV ファイルに保存する

// <CAUTION>
// このファイルは文字コードを ※BOM付き※> UTF8 として保存すること。
// (`ü` などの文字化け回避)
// </CAUTION>

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

var track = itobj.CurrentTrack; // <SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
var info;

//ファイル・フォルダ
var mydocu = wshobj.SpecialFolders("MyDocuments");//マイドキュメント場所
var nowpfol = "iTunesNowPlaying";//専用フォルダ名
var nowpfil = "NowPlaying.tsv";//専用ファイル名
var str_toSaveFilePath = mydocu + "\\" + nowpfol + "\\" + nowpfil;

//フォルダ存在確認
if(!(axobj.FolderExists(mydocu + "\\" + nowpfol))){
	axobj.CreateFolder(mydocu + "\\" + nowpfol);//フォルダ作成
}

//曲情報収集
try{
	// <SDKREF>iTunesCOM.chm::/iTunesTrackCOM_8idl.html#a12</SDKREF>
	if(track.Kind == 1){ //ローカル再生中
		info=track.Name + "\t" + track.Artist;
	
	}else if(track.Kind == 3) {//ストリーム再生中
		// ↓現在は使用不可↓
		// var titles = itobj.currentStreamTitle.split(",");
		// info=titles[0];

		info=track.Name + "\t" + track.Artist;

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

// ファイル関連の操作を提供する（ストリーム）オブジェクトを取得
var fh = new ActiveXObject( "ADODB.Stream" );
	
// 読み込むファイルのタイプを指定
fh.Type    = 2;         // 1:Binary, 2:Text

// 読み込むファイルの文字コードを指定
fh.charset = "UTF-8";   // Shift_JIS, EUC-JP, UTF-8、等々

// 読み込むファイルの改行コードを指定
fh.LineSeparator = -1;  //  -1 CrLf , 10 Lf , 13 Cr

// ストリームを開く
fh.Open();

//  ファイル関連の操作を提供するオブジェクトを取得
var fs = new ActiveXObject( "Scripting.FileSystemObject" );

if( !fs.FileExists(str_toSaveFilePath) ){ // ファイルが存在しない場合
	// 空ファイルを作成
	fh.WriteText( "", 0);  // 第2引数が 0:改行なし, 1:改行あり
    fh.SaveToFile( str_toSaveFilePath , 1 ); // 第2引数が 1:新規作成, 2:上書き
	
	// 一旦ストリームをクローズ＆オブジェクトを破棄
	fh.Close();
	fh = null;
	
	// 新たにストリームオブジェクトを作り直して
	fh = new ActiveXObject( "ADODB.Stream" );

	// 読み込むファイルのタイプを指定
	fh.Type    = 2;         // 1:Binary, 2:Text

	// 読み込むファイルの文字コードを指定
	fh.charset = "UTF-8";   // Shift_JIS, EUC-JP, UTF-8、等々

	// 読み込むファイルの改行コードを指定
	fh.LineSeparator = -1;  //  -1 CrLf , 10 Lf , 13 Cr

	// ストリームを開く
	fh.Open();	
}

// ファイルオープン
fh.LoadFromFile(str_toSaveFilePath);

// ポインタを終端へ
fh.Position = fh.Size; 

// ファイルに格納したいテキストをストリームに登録
fh.WriteText( info, 1);  // 第2引数が 0:改行なし, 1:改行あり

// トラック情報の保存
fh.SaveToFile( str_toSaveFilePath , 2 ); // 第2引数が 1:新規作成, 2:上書き

//ファイルクローズ
fh.Close();
