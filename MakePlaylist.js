// CAUTION
// 
// このファイルは文字コードを SJIS として保存すること。
// (SJIS 形式で保存しないと、`WScript.Echo` などで文字化けする)

// NOTE
//
// `<SDKREF>~~</SDKREF>` には、
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" 内の SDK Document の場所を記載
// 

// <Settings>--------------------------------------------
var str_playlistName = "Run";
var strarr_songNames = new Array(
    
    // Usage:
    // Song Name | Artist Name | Album Name
    // 
    //   NOTE Artist Name, and Album Name can be specify `undefined`

    new Array("Mr. Saxo Beat", "Alexandra Stan", "Original"),
    new Array("Turn Up the Music", "Chris Brown", "Turn Up the Music"),
    new Array("Viva La Vida", "Coldplay", "Viva La Vida Or Death And All His Friends"),
    new Array("Without You", "Harry Nilsson", "NISSAN We Love Drive"),
    new Array("Turn Me On", "David Guetta Feat. Nicki Minaj", "Nothing But The Beat 2.0"),
    new Array("Cities of The Future", "Infected Mushroom", "Original"),
    new Array("The Messenger 2012", "Infected Mushroom", "Original"),
    new Array("Love Foolosophy", "Jamiroquai", "A Funk Odyssey [Japan]"),
    new Array("The Other Side", "Jason Derülo", "Platinum Hits"),
    new Array("Party Rock Anthem", "LMFAO Feat. Lauren Bennet & GoonRock", "Sorry For Party Rocking"),
    new Array("Can't Hold Us (feat. Ray Dalton)", "Macklemore & Ryan Lewis", "Original"),
    new Array("Wrapped Up (feat. Travie McCoy)", "Olly Murs", "Original"),
    new Array("Dam Dariram", "Joga", "Dancemania DELUX4 [Disc 2]"),
    new Array("Good Time", "Owl City & Carly Rae Jepsen", "Kiss"),
    new Array("Breakin' A Sweat (Zedd Remix) - http://soundvor.ru/", "Skrillex ft. The Doors", "Original"),
    new Array("She's Got Me Dancing", "Tommy Sparks", "Original"),
    new Array("Don't Stop Movin'", "Livin' Joy", "Floorfillers Anthems (Floor 2)"),
    new Array("All Night (Europa XL Remix)", "Tezla", "Floorfillers 40 Massive Hits From The Clubs"),
    new Array("I Really Like You", "Carly Rae Jepsen", "Emotion")

);

// --------------------------------------------</Settings>

var int_notFoundTracks = 0;
var strarr_ = new Array();

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

//プレイリストの作成
var albumPlaylist = iTunesApp.CreatePlaylist(str_playlistName);

// プレイリストにトラックを追加
for (var int_idxOfTracks in strarr_songNames){
    
    var strarr_trackInfo = strarr_songNames[int_idxOfTracks];

    var obj_track = func_selectTrack(tracks, strarr_trackInfo[0], strarr_trackInfo[1], strarr_trackInfo[2]);

    if(obj_track !== undefined){ // トラックが見つかった場合
        albumPlaylist.AddTrack(obj_track);
    
    }else{ // トラックが見つからなかった場合
        int_notFoundTracks++;
        strarr_.push(strarr_trackInfo.toString());
    }
}

if(0 < int_notFoundTracks){
    WScript.Echo(int_notFoundTracks + " track(s) not found.");
    WScript.Echo(strarr_.toString());
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

            // アーティスト名が指定されたいた場合はアーティスト名をチェック
            if(typeof artst == "string"){
                if(obj_track.Artist != artst){ // アーティスト名が異なる場合
                    bl_found = false;
                }
            }

            // アーティスト名が指定されたいた場合はアーティスト名をチェック
            if(typeof albm == "string"){
                if(obj_track.Album != albm){ // アーティスト名が異なる場合
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
