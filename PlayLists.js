// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

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
	.Playlists; //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylistCollection.html</SDKREF>

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
        // WScript.Echo("開けません。");
        WScript.Quit(); // 終了
    }

    //曲毎ループ
    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
        var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
        // txfl.Write((typeof objTrack) + "\n");
        // txfl.Write(objTrack.Name + "\n");
        txfl.Write(objTrack.Kind + "\n");
    }

    //ファイルへ書き込み
	// txfl.Write("3298472" + "\n");
    // WScript.Echo(int_idxOfPlayelists);
    // WScript.Echo(objPlaylist.Name);

}

WScript.Echo("Done!");
