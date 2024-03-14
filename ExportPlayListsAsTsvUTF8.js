// iTunes のプレイリストを TSV ファイルに吐き出す

// <CAUTION>
// このファイルは文字コードを ※BOM付き※ UTF8 として保存すること。
// (`ü` などの文字化け回避)
// </CAUTION>

// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

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

//ファイル・フォルダ
var mydocu = wshobj.SpecialFolders("MyDocuments");//マイドキュメント場所
var str_fol = "iTunesPlayLists";//専用フォルダ名

//プレイリストの取得
var objPlaylists = itobj
	.LibrarySource //<SDKREF>iTunesCOM.chm::/interfaceIITSource.html</SDKREF>
	.Playlists //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylistCollection.html</SDKREF>
;

//フォルダ存在確認
if(!(axobj.FolderExists(mydocu + "\\" + str_fol))){
    axobj.CreateFolder(mydocu + "\\" + str_fol);//フォルダ作成
}

//プレイリスト毎ループ
for( var int_idxOfPlayelists = 1 ; int_idxOfPlayelists <= objPlaylists.Count; int_idxOfPlayelists++ ){
    
    var objPlaylist = objPlaylists.Item(int_idxOfPlayelists); //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html</SDKREF>
    var objTracks = objPlaylist.Tracks; //<SDKREF>iTunesCOM.chm::/interfaceIITTrackCollection.html</SDKREF>

    if(0 < objTracks.Count){
        
        var str_fileName = objPlaylist.Name + ".tsv"
    
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
    
        //Track 毎ループ
        for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
            
            var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
            var str_trackInfo = objTrack.Name + "\t" + objTrack.Artist + "\t" + objTrack.Album + "\t" + objTrack.Location
    
            // ファイルに格納したいテキストをストリームに登録
            fh.WriteText( str_trackInfo, 1);  // 第2引数が 0:改行なし, 1:改行あり
            
        }
    
        fh.SaveToFile( mydocu + "\\" + str_fol + "\\" + str_fileName , 2 ); // 第2引数が 1:新規作成, 2:上書き
        fh.Close();
        fh = null;
    }
}

WScript.Echo("Done!");
