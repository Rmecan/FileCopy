Option Strict Off
Option Explicit On
Module MdlLog
    Public Const gLogInformation As Short = 1
    Public Const gLogWarnig As Short = 2
    Public Const gLogError As Short = 3
	
    Public Sub LogOutput(ByVal iLogMode As Short, ByVal sLogMessage As String)
        '無効になっていたらバイバイ
        If False = GbLogEnable Or sLogMessage = "" Then
            Exit Sub
        End If

        Dim sOutputString As String = ""

        'メッセージ編集
        Select Case iLogMode
            '// Informationメッセージ
            Case gLogInformation
                sOutputString = Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & "[Information]" & vbTab & sLogMessage
                '// Warnigメッセージ
            Case gLogWarnig
                sOutputString = Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & "[Warning]" & vbTab & sLogMessage
                '// Errorメッセージ
            Case gLogError
                sOutputString = Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & "[Error]" & vbTab & sLogMessage
        End Select

        Dim fiLog As New IO.FileInfo(GsLogFile)

        'ログ書き込み
        If Not fiLog.Directory.Exists Then
            fiLog.Directory.Create()
        End If

        Using sw As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fiLog.FullName, True)
            sw.WriteLine(sOutputString)
        End Using
        System.Threading.Thread.Sleep(100)

    End Sub
	
	Public Sub KillLogFile()
        If My.Computer.FileSystem.DirectoryExists(GsLogFile) Then
            My.Computer.FileSystem.DeleteDirectory(GsLogFile, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

	End Sub

End Module