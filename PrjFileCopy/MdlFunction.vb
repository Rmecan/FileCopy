Option Strict Off
Option Explicit On
Module MdlFunction
	'----------------------------------------
	'　汎用関数群
	'
	'　　　   　AUTH Last Jedai
	'
	'       　Create 2017/12/25 Sunshine Ikebukuro
	'----------------------------------------
    'iniファイル
    Public ReadOnly GsIniFile As String = My.Application.Info.DirectoryPath & "\" & "config.ini"

    '[アプリログ]
    Public GbLogEnable As Boolean
    Public GsLogFile As String

    '[コピー]
    Public GbCopyLogEnable As Boolean
    Public GsCopyLogFile As String
    Public GsCopySrcDir As String
    Public GsCopyDstDir As String
    Public GdCopyDateFrom As Date
    Public GdCopyDateTo As Date
    Public GsCopyPatternOk As String
    Public GsCopyPatternNg As String
    	
	'INI取得ＡＰＩ宣言
    Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
		
	'Windowの位置やｻｲｽﾞ、表示を設定するAPI
	Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    
	'=====================================================
	'INI更新関数
	'=====================================================
    Public Sub Initial_IniUpload()
        '----------------------
        'グローバル変数の初期化
        '----------------------
        'ログ関係
        GbLogEnable = False
        GsLogFile = ""

        'コピー関係
        GbCopyLogEnable = False
        GsCopyLogFile = ""
        GsCopySrcDir = "C:\daifuku"
        GsCopyDstDir = Environment.ExpandEnvironmentVariables("%USERPROFILE%\Desktop\")
        GdCopyDateFrom = New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0)
        GdCopyDateTo = New Date(Now.Year, Now.Month, Now.Day, 23, 59, 59)
        GsCopyPatternOk = "*"
        GsCopyPatternNg = "*.scc"

        'INIファイルがあったら読み込む
        If My.Computer.FileSystem.FileExists(GsIniFile) Then
            '----------------------
            'iniファイルの読込
            '----------------------
            Dim temp As String = ""
            'ログ関係
            GfsGetINI("APP_LOG", "ENABLE", temp)
            GbLogEnable = IIf(temp = "1", True, False)
            GfsGetINI("APP_LOG", "FILE", temp)
            GsLogFile = Environment.ExpandEnvironmentVariables(temp)

            'コピー関係
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
	'INI取得関数
	'引数(INIグループキー
	'     INIキーワード
	'     INI取得値（戻り値）
	'Func戻り値(終了=TRUE,FLASE)
	'=====================================================
	Public Function GfsGetINI(ByVal sGroup As String, ByVal sKeyWord As String, ByRef sBackword As String) As Boolean
		Dim lRet As Integer 'API関数戻り値
		
		'初期値セット
		sBackword = New String(" ", 256)
		GfsGetINI = True
		
		'API関数呼び出し
		lRet = GetPrivateProfileString(sGroup, sKeyWord, "", sBackword, 256, GsIniFile)
		
		'SPACE切りとり
        sBackword = Trim(sBackword)

		'改行コード切り取り
		sBackword = Replace(sBackword, Chr(0), "")
		
    End Function

    '********************************************************
    ' 関数名：GetLineNumText
    ' 機能　：テキストファイル"FilePath"の行数を取得する
    ' 引数　：FilePath...行数を取得したいファイルの絶対パス
    ' 戻り値：行数
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