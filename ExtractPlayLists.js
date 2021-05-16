// CAUTION
// 
// ���̃t�@�C���͕����R�[�h�� SJIS �Ƃ��ĕۑ����邱�ƁB
// (SJIS �`���ŕۑ����Ȃ��ƁA`WScript.Echo` �Ȃǂŕ�����������)

// NOTE
//
// `<SDKREF>~~</SDKREF>` �ɂ́A
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" ���� SDK Document �̏ꏊ���L��
// 

var str_crlf = "\r\n";

var int_errCountObjectError = 0;
var int_errCountUnkown = 0;
var int_errTotal = 0;

//ActiveXObject����
var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject����
try{
	var itobj = WScript.CreateObject("iTunes.Application"); //<SDKREF>iTunesCOM.chm::/interfaceIiTunes.html</SDKREF>
}catch(e){
	WScript.Echo("Cannot create object `iTunes.Application`");
	WScript.Quit(); // �I��
}

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
        //https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
        var txfl = axobj.OpenTextFile(mydocu + "\\" + str_fol + "\\" + str_fileName, 2, true);
    }catch(e){
        WScript.Echo(str_fileName + " ���J���܂���");
        WScript.Quit(); // �I��
    }

    //Track �����[�v
    for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
        
        var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
        
        try{
            var str_trackInfo = objTrack.Artist + "\t" + objTrack.Name;
            txfl.Write(str_trackInfo + str_crlf);
        
        }catch(e){
            if (e =="[object Error]"){
                //NOTE
                // .Name �v���p�e�B�ɃA�N�Z�X�������ɃG���[�ɂȂ�ꍇ������B�����s���B Message -> `�v���V�[�W���̌Ăяo���A�܂��͈������s���ł�`
                int_errCountObjectError++;
                
            }else{
                int_errCountUnkown++;
            }
            txfl.Write(e + str_crlf);
        }
        
    }

    //�t�@�C���N���[�Y
    txfl.Close();

}

int_errTotal = int_errCountObjectError + int_errCountUnkown;

if(0 < int_errTotal){
    var str_errMsg =
        "Done!" + str_crlf + str_crlf +
        "But some error detected." + str_crlf + str_crlf +
        "  [object Error] : " + int_errCountObjectError + str_crlf +
        "  Unkown : " + int_errCountUnkown
    ;
    WScript.Echo(str_errMsg);

}else{
    WScript.Echo("Done!");
}
