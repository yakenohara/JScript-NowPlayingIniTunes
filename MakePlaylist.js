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
    "Mr. Saxo Beat",
    "Turn Up the Music",
    "Viva La Vida",
    "Without You",
    "Turn Me On",
    "Cities of The Future",
    "The Messenger 2012",
    "Love Foolosophy",
    "The Other Side",
    "Party Rock Anthem",
    "Can't Hold Us (feat. Ray Dalton)",
    "Wrapped Up (feat. Travie McCoy)",
    "Dam Dariram",
    "Good Time",
    "Breakin' A Sweat (Zedd Remix) - http://soundvor.ru/",
    "She's Got Me Dancing",
    "Don't Stop Movin'",
    "All Night (Europa XL Remix)",
    "I Really Like You"
);

// --------------------------------------------</Settings>

var int_notFoundTracks = 0;

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
    
    var str_trackName = strarr_songNames[int_idxOfTracks];

    try{
        var obj_track = tracks.ItemByName(str_trackName); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
        albumPlaylist.AddTrack(obj_track);
    
    }catch(e){ // トラックが見つからない場合
        
        if(e == "[object Error]"){
            int_notFoundTracks++;
        }else{
            var str_errMsg =
                "`" + e + "` detected." + str_crlf + str_crlf +
                WScript.Echo("Unkown Error.");
            ;
            WScript.Echo(str_errMsg);
            WScript.Quit(); // 終了
        }
    }
}


if(0 < int_notFoundTracks){
    WScript.Echo(int_notFoundTracks + " track(s) not found.");
}else{
    WScript.Echo("Done!");
}
