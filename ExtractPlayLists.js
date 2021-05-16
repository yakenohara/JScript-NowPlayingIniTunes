// CAUTION
// 
// ���̃t�@�C���͕����R�[�h�� SJIS �Ƃ��ĕۑ����邱�ƁB
// (SJIS �`���ŕۑ����Ȃ��ƁA`WScript.Echo` �Ȃǂŕ�����������)

// NOTE
//
// `<SDKREF>~~</SDKREF>` �ɂ́A
// "\SDK Reference\iTunes_COM_9.1.0.80\iTunes COM 9.1.0.80\iTunesCOM.chm" ���� SDK Document �̏ꏊ���L��
// 

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

    if(0 < objTracks.Count){
        
        var str_fileName = objPlaylist.Name + ".txt"
    
        // �t�@�C���֘A�̑����񋟂���i�X�g���[���j�I�u�W�F�N�g���擾
        var fh = new ActiveXObject( "ADODB.Stream" );
            
        // �ǂݍ��ރt�@�C���̃^�C�v���w��
        fh.Type    = 2;         // -1:Binary, 2:Text
        
        // �ǂݍ��ރt�@�C���̕����R�[�h���w��
        fh.charset = "UTF-8";   // Shift_JIS, EUC-JP, UTF-8�A���X
        
        // �ǂݍ��ރt�@�C���̉��s�R�[�h���w��
        fh.LineSeparator = -1;  // ' -1 CrLf , 10 Lf , 13 Cr
        
        // �X�g���[�����J��
        fh.Open();
    
        //Track �����[�v
        for( var int_idxOfTracks = 1 ; int_idxOfTracks <= objTracks.Count; int_idxOfTracks++ ){
            
            var objTrack = objTracks.Item(int_idxOfTracks); //<SDKREF>iTunesCOM.chm::/interfaceIITTrack.html</SDKREF>
            var str_trackInfo = objTrack.Name + "\t" + objTrack.Artist + "\t" + objTrack.Album
    
            // �t�@�C���Ɋi�[�������e�L�X�g���X�g���[���ɓo�^
            fh.WriteText( str_trackInfo, 1);  // ��2������ 0:���s�Ȃ�, 1:���s����
            
        }
    
        //<Save as UTF-8>------------------------------------------------
    
        //�t�@�C���N���[�Y
        // �|�C���^���f�[�^�̐擪�Ɉړ�������
        fh.Position = 0;
            
        // �o�C�i�����[�h�ɕύX����
        fh.Type = 1; 
        
        // �|�C���^��BOM�̕��i3�o�C�g�j�������ɂ��炵��
        fh.Position = 3;
        
        // �K���ȕϐ��Ƀo�C�i���f�[�^�Ƃ��ăf�[�^��ޔ�
        var bin = fh.Read();
        
        // ��U�X�g���[�����N���[�Y���I�u�W�F�N�g��j��
        fh.Close();
        fh = null;
        
        // �V���ɃX�g���[���I�u�W�F�N�g����蒼����
        fh = new ActiveXObject( "ADODB.Stream" );
        fh.Type    = 1; // �o�C�i�����[�h�ɐݒ肵��
        fh.Open();
        fh.Write(bin);  // �ޔ����Ă������f�[�^��ǂݍ��ݒ�����
        
        // �������珑�����߂�BOM�Ȃ�UTF-8�t�@�C���̏o���オ��
        fh.SaveToFile( mydocu + "\\" + str_fol + "\\" + str_fileName , 2 ); // ��2������ 1:�V�K�쐬, 2:�㏑��
        fh.Close();
        fh = null;
    
        //------------------------------------------------<Save as UTF-8>
    }
}

WScript.Echo("Done!");
