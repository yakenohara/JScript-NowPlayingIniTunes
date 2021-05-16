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

var track = itobj.CurrentTrack;    
var info;

//�t�@�C���E�t�H���_
var mydocu = wshobj.SpecialFolders("MyDocuments");//�}�C�h�L�������g�ꏊ
var nowpfol = "iTunesNowPlaying";//��p�t�H���_��
var nowpfil = "NowPlaying.txt";//��p�t�@�C����

//�t�H���_���݊m�F
if(!(axobj.FolderExists(mydocu + "\\" + nowpfol))){
	axobj.CreateFolder(mydocu + "\\" + nowpfol);//�t�H���_�쐬
}

//�ȏ����W
try{
	if(track.Kind == 1){ //���[�J���Đ���
		info=track.Artist + "\t" + track.Name;
	}else if(track.Kind == 3) {//�X�g���[���Đ���
		var titles = itobj.currentStreamTitle.split(",");
		info=titles[0];
	}else{
		WScript.Echo("�s���ȓ��쒆");
		WScript.Quit(); // �I��
	}
}catch(e){
	if(e == "[object Error]"){
		var str_errMsg =
			"`" + e + "` detected." + str_crlf + str_crlf +
			"iTunes�����삵�Ă��܂���"
		;
		WScript.Echo(str_errMsg);

	}else{
		var str_errMsg =
			"`" + e + "` detected." + str_crlf + str_crlf +
			WScript.Echo("Unkown Error.");
		;
	}
	WScript.Quit(); // �I��
}

//�t�@�C���I�[�v��
try{
	//https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
	var txfl = axobj.OpenTextFile(mydocu + "\\" + nowpfol + "\\" + nowpfil, 8, true);
}catch(e){
	WScript.Echo(nowpfil + "���J���܂���");
	WScript.Quit(); // �I��
}

//�t�@�C���֏�������
txfl.Write(info + "\n");

//�t�@�C���N���[�Y
txfl.Close();
