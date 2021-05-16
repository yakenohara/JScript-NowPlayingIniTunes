// CAUTION
// 
// ���̃t�@�C���͕����R�[�h�� SJIS �Ƃ��ĕۑ����邱�ƁB
// (SJIS �`���ŕۑ����Ȃ��ƁA`WScript.Echo` �Ȃǂŕ�����������)

// NOTE
//
// `<SDKREF>~~</SDKREF>` �ɂ́A
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" ���� SDK Document �̏ꏊ���L��
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
    new Array("The Other Side", undefined/*"Jason Der?lo"*/, "Platinum Hits"), //todo �@��ˑ�����? ����舵���Ȃ�
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

// �I��
try{
    var	iTunesApp = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>
}catch(e){
    WScript.Echo("Cannot create object `iTunes.Application`");
	WScript.Quit(); // �I��
}

var	mainLibrarySource = iTunesApp.LibrarySource; //<SDKREF>iTunesCOM.chm::/interfaceIITSource.html</SDKREF>
var	mainLibrary = iTunesApp.LibraryPlaylist; //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html</SDKREF>
var	tracks = mainLibrary.Tracks; //<SDKREF>iTunesCOM.chm::/interfaceIITTrackCollection.html</SDKREF>

//�v���C���X�g�̍쐬
var albumPlaylist = iTunesApp.CreatePlaylist(str_playlistName);

// �v���C���X�g�Ƀg���b�N��ǉ�
for (var int_idxOfTracks in strarr_songNames){
    
    var strarr_trackInfo = strarr_songNames[int_idxOfTracks];

    var obj_track = func_selectTrack(tracks, strarr_trackInfo[0], strarr_trackInfo[1], strarr_trackInfo[2]);

    if(obj_track !== undefined){ // �g���b�N�����������ꍇ
        albumPlaylist.AddTrack(obj_track);
    
    }else{ // �g���b�N��������Ȃ������ꍇ
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
// �g���b�N�R���N�V��������w��Ȗ��A�A�[�e�B�X�g���A�A���o�����ɊY������g���b�N��Ԃ�
//
function func_selectTrack(trackClctn, nm, artst, albm){

    var obj_ret;

    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= trackClctn.Count; int_idxOfTracks++ ){
        
        var obj_track = trackClctn.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>

        if(obj_track.Name === nm){

            var bl_found = true;

            // �A�[�e�B�X�g�����w�肳�ꂽ�����ꍇ�̓A�[�e�B�X�g�����`�F�b�N
            if(typeof artst == "string"){
                if(obj_track.Artist != artst){ // �A�[�e�B�X�g�����قȂ�ꍇ
                    bl_found = false;
                }
            }

            // �A�[�e�B�X�g�����w�肳�ꂽ�����ꍇ�̓A�[�e�B�X�g�����`�F�b�N
            if(typeof albm == "string"){
                if(obj_track.Album != albm){ // �A�[�e�B�X�g�����قȂ�ꍇ
                    bl_found = false;
                }
            }

            if(bl_found){ // �q�b�g����
                obj_ret = obj_track;
                break;
            }
        }
    }

    return obj_ret;
}
