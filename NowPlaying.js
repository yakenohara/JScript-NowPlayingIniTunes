//ActiveXObject生成
	var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
	var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject生成
	var itobj = WScript.CreateObject("iTunes.Application");
	var track = itobj.CurrentTrack;    
	var info;

//ファイル・フォルダ
	var mydocu = wshobj.SpecialFolders("MyDocuments");//マイドキュメント場所
	var nowpfol = "NowPlaying";//専用フォルダ名
	var nowpfil = "NowPlaying.txt";//専用ファイル名
	
//フォルダ存在確認
	if(!(axobj.FolderExists(mydocu + "\\" + nowpfol))){
		axobj.CreateFolder(mydocu + "\\" + nowpfol);//フォルダ作成
	}

//曲情報収集
	try{
		if(track.Kind == 1){ //ローカル再生中
			info=track.Artist + " - " + track.Name;
		}else if(track.Kind == 3) {//ストリーム再生中
			var titles = itobj.currentStreamTitle.split(",");
			info=titles[0];
		}else{
			WScript.Echo("不明な動作中");
			WScript.Quit(); // 終了
		}
	}catch(e){
		WScript.Echo("iTunesが動作していません。");
	}
	
//ファイルオープン
	try{
		var txfl = axobj.OpenTextFile(mydocu + "\\" + nowpfol + "\\" + nowpfil, 8, true); //https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
	}catch(e){
		WScript.Echo(nowpfil + "が開けません。");
	}

//ファイルへ書き込み
	txfl.Write(info + "\n");
	
//ファイルクローズ
	txfl.Close();