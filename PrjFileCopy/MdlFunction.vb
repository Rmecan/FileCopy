Option Strict Off
Option Explicit On
Module MdlFunction
	'----------------------------------------
	'�@�ėp�֐��Q
	'
	'�@�@�@   �@AUTH Last Jedai
	'
	'       �@Create 2017/12/25 Sunshine Ikebukuro
	'----------------------------------------
    'ini�t�@�C��
    Public ReadOnly GsIniFile As String = My.Application.Info.DirectoryPath & "\" & "config.ini"

    '[�A�v�����O]
    Public GbLogEnable As Boolean
    Public GsLogFile As String

    '[�R�s�[]
    Public GbCopyLogEnable As Boolean
    Public GsCopyLogFile As String
    Public GsCopySrcDir As String
    Public GsCopyDstDir As String
    Public GdCopyDateFrom As Date
    Public GdCopyDateTo As Date
    Public GsCopyPatternOk As String
    Public GsCopyPatternNg As String
    	
	'INI�擾�`�o�h�錾
    Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
		
	'Window�̈ʒu�⻲�ށA�\����ݒ肷��API
	Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    
	'=====================================================
	'INI�X�V�֐�
	'=====================================================
    Public Sub Initial_IniUpload()
        '----------------------
        '�O���[�o���ϐ��̏�����
        '----------------------
        '���O�֌W
        GbLogEnable = False
        GsLogFile = ""

        '�R�s�[�֌W
        GbCopyLogEnable = False
        GsCopyLogFile = ""
        GsCopySrcDir = "C:\daifuku"
        GsCopyDstDir = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Desktop\")
        GdCopyDateFrom = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0)
        GdCopyDateTo = New Date(Now.Year, Now.Month, Now.Day, 23, 59, 59)
        GsCopyPatternOk = "*"
        GsCopyPatternNg = "*.scc"

        'INI�t�@�C������������ǂݍ���
        If My.Computer.FileSystem.FileExists(GsIniFile) Then
            '----------------------
            'ini�t�@�C���̓Ǎ�
            '----------------------
            Dim temp As String = ""
            '���O�֌W
            GfsGetINI("APP_LOG", "ENABLE", temp)
            GbLogEnable = IIf(temp = "1", True, False)
            GfsGetINI("APP_LOG", "FILE", temp)
            GsLogFile = Environment.ExpandEnvironmentVariables(temp)

            '�R�s�[�֌W
            GfsGetINI("COPY", "LOG_ENABLE ", temp)
            GbCopyLogEnable = IIf(temp = "1", True, False)
            GfsGetINI("COPY", "LOG_FILE", temp)
            GsCopyLogFile = Environment.ExpandEnvironmentVariables(temp)
            GfsGetINI("COPY", "SRC_DIR", temp)
            GsCopySrcDir = Environment.ExpandEnvironmentVariables(temp)
            GfsGetINI("COPY", "DST_DIR", temp)
            GsCopyDstDir = Environment.ExpandEnvironmentVariables(temp)
            GfsGetINI("COPY", "DATE_FROM", temp)
            If temp <> "" And IsDate(temp) Then
                GdCopyDateFrom = CDate(temp)
            End If
            GfsGetINI("COPY", "DATE_TO", temp)
            If temp <> "" And IsDate(temp) Then
                GdCopyDateTo = CDate(temp)
            End If
            GfsGetINI("COPY", "PATTERN_OK", temp)
            GsCopyPatternOk = temp
            GfsGetINI("COPY", "PATTERN_NG", temp)
            GsCopyPatternNg = temp
        End If

    End Sub
	
	'=====================================================
	'INI�擾�֐�
	'����(INI�O���[�v�L�[
	'     INI�L�[���[�h
	'     INI�擾�l�i�߂�l�j
	'Func�߂�l(�I��=TRUE,FLASE)
	'=====================================================
	Public Function GfsGetINI(ByVal sGroup As String, ByVal sKeyWord As String, ByRef sBackword As String) As Boolean
		Dim lRet As Integer 'API�֐��߂�l
		
		'�����l�Z�b�g
		sBackword = New String(" ", 256)
		GfsGetINI = True
		
		'API�֐��Ăяo��
		lRet = GetPrivateProfileString(sGroup, sKeyWord, "", sBackword, 256, GsIniFile)
		
		'SPACE�؂�Ƃ�
        sBackword = Trim(sBackword)

		'���s�R�[�h�؂���
		sBackword = Replace(sBackword, Chr(0), "")
		
    End Function

    '********************************************************
    ' �֐����FGetLineNumText
    ' �@�\�@�F�e�L�X�g�t�@�C��"FilePath"�̍s�����擾����
    ' �����@�FFilePath...�s�����擾�������t�@�C���̐�΃p�X
    ' �߂�l�F�s��
    '********************************************************
    Function GetLineNumText(ByVal filePath As String) As Int32
        GetLineNumText = 0

        If Not My.Computer.FileSystem.FileExists(filePath) Then
            Exit Function
        End If

        Dim FSO As Object = CreateObject("Scripting.FileSystemObject")
        With FSO.OpenTextFile(filePath, 8)
            GetLineNumText = .Line
            .Close()
        End With
        FSO = Nothing
    End Function

End Module