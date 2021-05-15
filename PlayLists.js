
// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

var int_errCountObjectError = 0;
var int_errCountUnkown = 0;
var int_errTotal = 0;

//ActiveXObject生成
var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject生成
var itobj = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>

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
    var str_fileName = objPlaylist.Name + ".txt"

    //ファイルオープン
    try{
        var txfl = axobj.OpenTextFile(mydocu + "\\" + str_fol + "\\" + str_fileName, 2, true); //https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
    }catch(e){
        WScript.Echo(str_fileName + " が開けません。");
        WScript.Quit(); // 終了
    }

    //Track 毎ループ
    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
        
        var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
        
        try{
            var str_trackInfo = objTrack.Artist + "\t" + objTrack.Name;
            txfl.Write(str_trackInfo + "\n");
        
        }catch(e){
            if (e =="[object Error]"){
                //NOTE
                // .Name プロパティにアクセスした時にエラーになる場合がある。原因不明。 Message -> `プロシージャの呼び出し、または引数が不正です`
                int_errCountObjectError++;
                
            }else{
                int_errCountUnkown++;
            }
            txfl.Write(e + "\n");
        }
        
    }

}

int_errTotal = int_errCountObjectError + int_errCountUnkown;

if(0 < int_errTotal){
    var str_errMsg =
        "Error Detected.\n\n" + 
        "  [object Error] : " + int_errCountObjectError + "\n" +
        "  Unkown : " + int_errCountUnkown
    ;
    WScript.Echo(str_errMsg);

}else{
    WScript.Echo("Done!");
}
