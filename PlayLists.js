
// NOTE
//
// `<SDKREF>~~</SDKREF>` �ɂ́A
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" ���� SDK Document �̏ꏊ���L��
// 

var int_errCountObjectError = 0;
var int_errCountUnkown = 0;
var int_errTotal = 0;

//ActiveXObject����
var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject����
var itobj = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>

//�t�@�C���E�t�H���_
var mydocu = wshobj.SpecialFolders("MyDocuments");//�}�C�h�L�������g�ꏊ
var str_fol = "iTunesPlayLists";//��p�t�H���_��

//�v���C���X�g�̎擾
var objPlaylists = itobj
	.LibrarySource //<SDKREF>iTunesCOM.chm::/interfaceIITSource.html</SDKREF>
	.Playlists //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylistCollection.html</SDKREF>
;

//�t�H���_���݊m�F
if(!(axobj.FolderExists(mydocu + "\\" + str_fol))){
    axobj.CreateFolder(mydocu + "\\" + str_fol);//�t�H���_�쐬
}

//�v���C���X�g�����[�v
for( var int_idxOfPlayelists = 1 ; int_idxOfPlayelists <= objPlaylists.Count; int_idxOfPlayelists++ ){
    
    var objPlaylist = objPlaylists.Item(int_idxOfPlayelists); //<SDKREF>iTunesCOM.chm::/interfaceIITPlaylist.html</SDKREF>
    var objTracks = objPlaylist.Tracks; //<SDKREF>iTunesCOM.chm::/interfaceIITTrackCollection.html</SDKREF>
    var str_fileName = objPlaylist.Name + ".txt"

    //�t�@�C���I�[�v��
    try{
        var txfl = axobj.OpenTextFile(mydocu + "\\" + str_fol + "\\" + str_fileName, 2, true); //https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
    }catch(e){
        WScript.Echo(str_fileName + " ���J���܂���B");
        WScript.Quit(); // �I��
    }

    //Track �����[�v
    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
        
        var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
        
        try{
            var str_trackInfo = objTrack.Artist + "\t" + objTrack.Name;
            txfl.Write(str_trackInfo + "\n");
        
        }catch(e){
            if (e =="[object Error]"){
                //NOTE
                // .Name �v���p�e�B�ɃA�N�Z�X�������ɃG���[�ɂȂ�ꍇ������B�����s���B Message -> `�v���V�[�W���̌Ăяo���A�܂��͈������s���ł�`
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
