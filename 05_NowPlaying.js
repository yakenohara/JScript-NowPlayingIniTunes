//ActiveXObject����
	var axobj = new ActiveXObject("Scripting.FileSystemObject"); //FileSystem
	var wshobj = new ActiveXObject("WScript.Shell");//WScript

//iTunesObject����
	var itobj = WScript.CreateObject("iTunes.Application");
	var track = itobj.CurrentTrack;    
	var info;

//�t�@�C���E�t�H���_
	var mydocu = wshobj.SpecialFolders("MyDocuments");//�}�C�h�L�������g�ꏊ
	var nowpfol = "NowPlaying";//��p�t�H���_��
	var nowpfil = "NowPlaying.txt";//��p�t�@�C����
	
//�t�H���_���݊m�F
	if(!(axobj.FolderExists(mydocu + "\\" + nowpfol))){
		axobj.CreateFolder(mydocu + "\\" + nowpfol);//�t�H���_�쐬
	}

//�ȏ����W
	try{
		if(track.Kind == 1){ //���[�J���Đ���
			info=track.Artist + " - " + track.Name;
		}else if(track.Kind == 3) {//�X�g���[���Đ���
			var titles = itobj.currentStreamTitle.split(",");
			info=titles[0];
		}else{
			WScript.Echo("�s���ȓ��쒆");
			WScript.Quit(); // �I��
		}
	}catch(e){
		WScript.Echo("iTunes�����삵�Ă��܂���B");
	}
	
//�t�@�C���I�[�v��
	try{
		var txfl = axobj.OpenTextFile(mydocu + "\\" + nowpfol + "\\" + nowpfil, 8, true); //https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
	}catch(e){
		WScript.Echo(nowpfil + "���J���܂���B");
	}

//�t�@�C���֏�������
	txfl.Write(info + "\n");
	
//�t�@�C���N���[�Y
	txfl.Close();