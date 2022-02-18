// <CAUTION>
// このファイルは文字コードを ※BOM付き※ UTF8 として保存すること。
// (`ü` などの文字化け回避)
// </CAUTION>

// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

// <Settings>--------------------------------------------
var str_playlistName = "Run";
var strarr_tsvPaths = new Array(
    "C:\\Users\\username\\Documents\\iTunesPlayLists\\Oldies.tsv",
    "C:\\Users\\username\\Documents\\iTunesPlayLists\\Run.tsv",
    "C:\\Users\\username\\Documents\\iTunesPlayLists\\Anime.tsv",
    "C:\\Users\\username\\Documents\\iTunesPlayLists\\Drive.tsv",
    "C:\\Users\\username\\Documents\\iTunesPlayLists\\Loves.tsv"
);

// --------------------------------------------</Settings>

// 終了
try{
    var	iTunesApp = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>
}catch(e){
    WScript.Echo("Cannot create object `iTunes.Application`");
	WScript.Quit(); // 終了
}

var	mainLibrarySource = iTunesApp.LibrarySource; //<SDKREF>iTunesCOM.chm::/interfaceIITSource.html</SDKREF>
var	mainLibrary = iTunesApp.LibraryPlaylist; //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html</SDKREF>
var	tracks = mainLibrary.Tracks; //<SDKREF>iTunesCOM.chm::/interfaceIITTrackCollection.html</SDKREF>

//  ファイル関連の操作を提供するオブジェクトを取得
var filesystem = new ActiveXObject( "Scripting.FileSystemObject" );

// ファイル関連の操作を提供する（ストリーム）オブジェクトを取得
var filestream = new ActiveXObject ("ADODB.Stream");
// 読み込むファイルのタイプを指定
filestream.Type = 2             // 1:Binary, 2:Text

// 読み込むファイルの文字コードを指定
filestream.charset = "UTF-8";   // Shift_JIS, EUC-JP, UTF-8、等々

// 読み込むファイルの改行コードを指定
filestream.LineSeparator = -1;  //  -1 CrLf , 10 Lf , 13 Cr

var strarr_ = new Array();

// ファイル毎ループ
for (var int_idxOfPaths in strarr_tsvPaths){
    
    var str_tsvPath = strarr_tsvPaths[int_idxOfPaths];

    try{
        // ファイルオープン
        filestream.Open();
        filestream.loadFromFile(str_tsvPath);
        
        var int_notFoundTracks = 0;

        //プレイリストの作成
        var albumPlaylist = iTunesApp.CreatePlaylist(filesystem.GetBaseName(str_tsvPath));

        // 行単位読み込みループ
        while (!filestream.EOS){
            
            str_line = filestream.ReadText(-2) // 1 行読み取り
            var strarr_trackInfo = str_line.split("\t")

            var obj_track = func_selectTrack(tracks, strarr_trackInfo[0], strarr_trackInfo[1], strarr_trackInfo[2]);
        
            if(obj_track !== undefined){ // トラックが見つかった場合
                albumPlaylist.AddTrack(obj_track);
            
            }else{ // トラックが見つからなかった場合
                if(int_notFoundTracks == 0){
                    strarr_.push("While processing `" + str_tsvPath + "` following songs not found.");
                }
                int_notFoundTracks++;
                strarr_.push(strarr_trackInfo.toString()); // 配列に格納されている要素を「,(コンマ)」区切りで文字列化
            }
        }

        filestream.Close();

    }catch(e){ // ファイルオープン失敗の場合
        strarr_.push("Cannot open file `" + str_tsvPath + "`");
        filestream.Close();
    }

}

// 終了メッセージ
if(0 < strarr_.length){
    WScript.Echo(strarr_.join("\n"));
}else{
    WScript.Echo("Done!");
}

//
// トラックコレクションから指定曲名、アーティスト名、アルバム名に該当するトラックを返す
//
function func_selectTrack(trackClctn, nm, artst, albm){

    var obj_ret;

    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= trackClctn.Count; int_idxOfTracks++ ){
        
        var obj_track = trackClctn.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>

        if(obj_track.Name === nm){

            var bl_found = true;

            // アーティスト名が指定された場合はアーティスト名をチェック
            if(typeof artst == "string"){
                if(obj_track.Artist != artst){ // アーティスト名が異なる場合
                    bl_found = false;
                }
            }

            // アルバム名が指定された場合はアーティスト名をチェック
            if(typeof albm == "string"){
                if(obj_track.Album != albm){ // アルバム名が異なる場合
                    bl_found = false;
                }
            }

            if(bl_found){ // ヒット判定
                obj_ret = obj_track;
                break;
            }
        }
    }

    return obj_ret;
}
