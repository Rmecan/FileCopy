Public Class FrmMain

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'ログ出力
        LogOutput(gLogInformation, "アプリ起動")
        'ini 再読み込み
        FrmMain_Initialize()
    End Sub

    Private Sub FrmMain_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        'ログ出力
        LogOutput(gLogInformation, "アプリ終了")
    End Sub

    Private Sub FrmMain_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Select Case KeyCode
            Case Keys.F5
                'ini 再読み込み
                FrmMain_Initialize()
        End Select
    End Sub

    Private Sub FrmMain_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter
        'コントロール内にドラッグされたとき実行される
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            'ドラッグされたデータ形式を調べ、ファイルのときはコピーとする
            e.Effect = DragDropEffects.Copy
        Else
            'ファイル以外は受け付けない
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub FrmMain_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop
        'ドロップされたすべてのファイル名を取得する
        Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        Dim file As String = Nothing
        'ドロップされたすべてのファイルから先頭1個だけを対象とする
        For Each wk As String In files
            file = wk
            Exit For
        Next wk
        'ゴミは無視する
        If IsNothing(file) Or (Not My.Computer.FileSystem.FileExists(file)) Then
            Return
        End If
        ' チェック処理
        If False = IsDragAndDropStartOk() Then
            Return
        End If
        '実行処理
        DoFileCopy(file)
    End Sub

    Private Sub BtnDstDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDstDir.Click
        'ダイアログを出す
        DlgDstDir.Description = LblHozon.Text + "のフォルダを指定してください。"
        DlgDstDir.RootFolder = Environment.SpecialFolder.Desktop
        DlgDstDir.SelectedPath = TxtDstDir.Text
        '選択したフォルダを画面に表示
        If DlgDstDir.ShowDialog(Me) = DialogResult.OK Then
            TxtDstDir.Text = DlgDstDir.SelectedPath
        End If
    End Sub

    Private Sub BtnSrcDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSrcDir.Click
        'ダイアログを出す
        DlgSrcDir.Description = LblKensaku.Text + "のフォルダを指定してください。"
        DlgSrcDir.RootFolder = Environment.SpecialFolder.Desktop
        DlgSrcDir.SelectedPath = TxtSrcDir.Text
        '選択したフォルダを画面に表示
        If DlgSrcDir.ShowDialog(Me) = DialogResult.OK Then
            TxtSrcDir.Text = DlgSrcDir.SelectedPath
        End If
    End Sub

    Private Sub FrmMain_Initialize()
        'iniを再読み込み
        Initial_IniUpload()
        'iniの内容を画面にセット
        TxtDstDir.Text = GsCopyDstDir
        TxtSrcDir.Text = GsCopySrcDir
        DateKoshinhiFrom.Value = GdCopyDateFrom
        DateKoshinhiTo.Value = GdCopyDateTo
        TxtOkPattern.Text = GsCopyPatternOk
        TxtNgPattern.Text = GsCopyPatternNg
        ChkOutput.Checked = GbCopyLogEnable
    End Sub

    Private Sub Enable_Control(ByVal IsEnable As Boolean)
        If True = IsEnable Then
            BtnDstDir.Enabled = True
            BtnSrcDir.Enabled = True
            BtnFileCopy.Enabled = True
            BtnListSakusei.Enabled = True
            DateKoshinhiFrom.Enabled = True
            DateKoshinhiTo.Enabled = True
            ChkSubKensaku.Enabled = True
            ChkOutput.Enabled = True
            TxtOkPattern.Enabled = True
            TxtNgPattern.Enabled = True
        Else
            BtnDstDir.Enabled = False
            BtnSrcDir.Enabled = False
            BtnFileCopy.Enabled = False
            BtnListSakusei.Enabled = False
            DateKoshinhiFrom.Enabled = False
            DateKoshinhiTo.Enabled = False
            ChkSubKensaku.Enabled = False
            ChkOutput.Enabled = False
            TxtOkPattern.Enabled = False
            TxtNgPattern.Enabled = False
        End If
    End Sub

    Private Function IsDragAndDropStartOk() As Boolean
        IsDragAndDropStartOk = False

        '前提チェック
        If String.IsNullOrEmpty(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "は必須です。", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "が存在しません。", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        IsDragAndDropStartOk = True
    End Function

    Private Function IsJikkoStartOk() As Boolean
        IsJikkoStartOk = False

        '前提チェック
        If String.IsNullOrEmpty(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "は必須です。", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "が存在しません。", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If String.IsNullOrEmpty(TxtSrcDir.Text) Then
            MsgBox(LblKensaku.Text + "は必須です。", MsgBoxStyle.OkOnly, Me.Text)
            BtnSrcDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtSrcDir.Text) Then
            MsgBox(LblKensaku.Text + "が存在しません。", MsgBoxStyle.OkOnly, Me.Text)
            BtnSrcDir.Focus()
            Exit Function
        End If

        If TxtDstDir.Text = TxtSrcDir.Text Then
            MsgBox(LblHozon.Text + "と" + LblKensaku.Text + "は別々のフォルダを指定してください。", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not IsDate(DateKoshinhiFrom.Value) Then
            MsgBox("更新日時[開始]が不正です。", MsgBoxStyle.OkOnly, Me.Text)
            DateKoshinhiFrom.Focus()
            Exit Function
        End If

        If Not IsDate(DateKoshinhiTo.Value) Then
            MsgBox("更新日時[終了]が不正です。", MsgBoxStyle.OkOnly, Me.Text)
            DateKoshinhiTo.Focus()
            Exit Function
        End If

        If DateTime.Compare(DateKoshinhiFrom.Value, DateKoshinhiTo.Value) > 0 Then
            MsgBox("更新日時の開始、終了が逆転しています。", MsgBoxStyle.OkOnly, Me.Text)
            DateKoshinhiTo.Focus()
            Exit Function
        End If

        If String.IsNullOrEmpty(TxtOkPattern.Text) Then
            TxtOkPattern.Text = "*.*"
        Else
            Dim okPatterns As String() = TxtOkPattern.Text.Split(",")
            Dim hasAstr As Boolean = False
            Dim hasExtention As Boolean = False
            For Each okPattern As String In okPatterns
                If "*" = okPattern Or "*.*" = okPattern Then
                    hasAstr = True
                Else
                    hasExtention = True
                End If
            Next
            If hasAstr And hasExtention Then
                MsgBox("*または*.*を入力した場合、他の条件は指定できません。", MsgBoxStyle.OkOnly, Me.Text)
                TxtOkPattern.Focus()
                Exit Function
            End If
        End If

        If String.IsNullOrEmpty(TxtNgPattern.Text) Then
            TxtNgPattern.Text = ""
        Else
            Dim ngPatterns As String() = TxtNgPattern.Text.Split(",")
            Dim hasAstr As Boolean = False
            Dim hasExtention As Boolean = False
            For Each ngPattern As String In ngPatterns
                If "*" = ngPattern Or "*.*" = ngPattern Then
                    MsgBox("*または*.*は指定できません。", MsgBoxStyle.OkOnly, Me.Text)
                    TxtNgPattern.Focus()
                    Exit Function
                End If
            Next
        End If

        IsJikkoStartOk = True
    End Function

    Private Sub BtnJikko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFileCopy.Click
        'チェック処理
        If False = IsJikkoStartOk() Then
            Return
        End If
        '実行処理
        DoFileCopy()
    End Sub

    Private Sub BtnListSakusei_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnListSakusei.Click
        'チェック処理
        If False = IsJikkoStartOk() Then
            Return
        End If
        '実行処理
        DoListSakusei()
    End Sub

    Private Sub DoFileCopy(Optional ByVal filePath As String = "")
        '処理中にボタンを押されたくないのでボタンを無効にする
        Enable_Control(False)
        'コピー処理
        BgFileCopy.RunWorkerAsync(filePath)
    End Sub

    Private Sub DoListSakusei(Optional ByVal filePath As String = "")
        '処理中にボタンを押されたくないのでボタンを無効にする
        Enable_Control(False)
        '作成処理
        BgListSakusei.RunWorkerAsync(filePath)
    End Sub

    Private Sub BgFileCopy_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgFileCopy.DoWork
        Dim pathCopyList As String = CStr(e.Argument)
        Dim pathCopyLog As String = ""
        Dim wkCopyList As List(Of String) = New List(Of String)

        '生成するコピーログの名称決定
        If False = ChkOutput.Checked Then
            pathCopyLog = ""
        ElseIf String.IsNullOrEmpty(GsCopyLogFile) Then
            pathCopyLog = TxtDstDir.Text + "\" + Now.ToString("yyyyMMddHHmmss") + "COPY.log"
        Else
            pathCopyLog = TxtDstDir.Text + "\" + GsCopyLogFile
        End If

        'コピー一覧の指定がない時は検索条件からコピー一覧を作成する
        If String.IsNullOrEmpty(pathCopyList) Then
            Dim diKensaku As New System.IO.DirectoryInfo(TxtSrcDir.Text)
            Dim opKensaku As System.IO.SearchOption = System.IO.SearchOption.TopDirectoryOnly
            If ChkSubKensaku.Checked Then
                opKensaku = System.IO.SearchOption.AllDirectories
            End If

            '名前が被ってるファイルが居たら無条件削除
            If My.Computer.FileSystem.FileExists(pathCopyList) Then
                My.Computer.FileSystem.DeleteFile(pathCopyList)
            End If

            '全検索
            For Each tempFile As System.IO.FileInfo In diKensaku.GetFiles("*", opKensaku)
                '更新日時[From]より小さい
                If tempFile.LastWriteTime < DateKoshinhiFrom.Value Then
                    Continue For
                End If
                '更新日時[To]より大きい
                If DateKoshinhiTo.Value < tempFile.LastWriteTime Then
                    Continue For
                End If
                'OKワードになかったらダメ
                Dim hasOk As Boolean = False
                For Each pattern As String In TxtOkPattern.Text.Split(",")
                    If tempFile.FullName Like pattern Then
                        hasOk = True
                        Exit For
                    End If
                Next
                If False = hasOk Then
                    Continue For
                End If
                'NGワードに載っていたらダメ
                Dim hasNg As Boolean = False
                For Each pattern As String In TxtNgPattern.Text.Split(",")
                    If tempFile.FullName Like pattern Then
                        hasNg = True
                        Exit For
                    End If
                Next
                If True = hasNg Then
                    Continue For
                End If
                'ここまでこれたらコピーリストに追加
                wkCopyList.Add(tempFile.FullName)
            Next

        Else
            'ファイルからコピーリストを作成
            Using sr As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(pathCopyList)
                While (sr.Peek() >= 0)
                    Dim filePath As String = sr.ReadLine().Trim
                    '最初と最後がダブルクォーテーションの場合は取り除く
                    If filePath.StartsWith("""") And filePath.EndsWith("""") Then
                        filePath = filePath.Substring(1, filePath.Length - 2)
                    End If
                    If My.Computer.FileSystem.FileExists(filePath) Then
                        wkCopyList.Add(filePath)
                    Else
                        LogOutput(gLogError, "ファイルが存在しないためコピーできません。[" + filePath + "]")
                    End If
                End While
            End Using
        End If

        'コピーするのがなかったら即終了
        If wkCopyList.Count < 1 Then
            Exit Sub
        End If

        'コピーログを作成するので事前準備
        If Not String.IsNullOrEmpty(pathCopyLog) Then
            If My.Computer.FileSystem.FileExists(pathCopyLog) Then
                My.Computer.FileSystem.DeleteFile(pathCopyLog)
            End If
        End If

        'コピーリストを元にファイルをコピーをする
        For Each srcFileName As String In wkCopyList
            Dim srcFileInfo As IO.FileInfo = New IO.FileInfo(srcFileName)
            Dim dstFileInfo As IO.FileInfo = New IO.FileInfo(TxtDstDir.Text + srcFileName.Substring(srcFileName.IndexOf(":") + 1))
            'フォルダ作成
            If Not My.Computer.FileSystem.DirectoryExists(dstFileInfo.DirectoryName) Then
                My.Computer.FileSystem.CreateDirectory(dstFileInfo.DirectoryName)
            End If
            'コピー
            srcFileInfo.CopyTo(dstFileInfo.FullName, True)
            'コピーログを作成する
            If Not String.IsNullOrEmpty(pathCopyLog) Then
                Using sw As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(pathCopyLog, True)
                    sw.WriteLine("""" + srcFileInfo.FullName + """")
                End Using
                System.Threading.Thread.Sleep(10)
            End If
        Next
    End Sub

    Private Sub BgFileCopy_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BgFileCopy.ProgressChanged

    End Sub

    Private Sub BgFileCopy_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgFileCopy.RunWorkerCompleted
        LogOutput(gLogInformation, "コピー処理終了")
        'ボタンを有効にする
        Enable_Control(True)
    End Sub

    Private Sub BgListSakusei_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgListSakusei.DoWork
        Dim pathCopyLog As String = TxtDstDir.Text + "\" + Now.ToString("yyyyMMddHHmmss") + "_抽出結果.txt"
        Dim wkCopyList As List(Of String) = New List(Of String)

        Dim diKensaku As New System.IO.DirectoryInfo(TxtSrcDir.Text)
        Dim opKensaku As System.IO.SearchOption = System.IO.SearchOption.TopDirectoryOnly
        If ChkSubKensaku.Checked Then
            opKensaku = System.IO.SearchOption.AllDirectories
        End If

        '全検索
        For Each tempFile As System.IO.FileInfo In diKensaku.GetFiles("*", opKensaku)
            '更新日時[From]より小さい
            If tempFile.LastWriteTime < DateKoshinhiFrom.Value Then
                Continue For
            End If
            '更新日時[To]より大きい
            If DateKoshinhiTo.Value < tempFile.LastWriteTime Then
                Continue For
            End If
            'OKワードになかったらダメ
            Dim hasOk As Boolean = False
            For Each pattern As String In TxtOkPattern.Text.Split(",")
                If tempFile.FullName Like pattern Then
                    hasOk = True
                    Exit For
                End If
            Next
            If False = hasOk Then
                Continue For
            End If
            'NGワードに載っていたらダメ
            Dim hasNg As Boolean = False
            For Each pattern As String In TxtNgPattern.Text.Split(",")
                If tempFile.FullName Like pattern Then
                    hasNg = True
                    Exit For
                End If
            Next
            If True = hasNg Then
                Continue For
            End If
            'ここまでこれたらコピーリストに追加
            wkCopyList.Add(tempFile.FullName)
        Next

        'コピーするのがなかったら即終了
        If wkCopyList.Count < 1 Then
            Exit Sub
        End If

        'コピーログを作成するので事前準備
        If Not String.IsNullOrEmpty(pathCopyLog) Then
            If My.Computer.FileSystem.FileExists(pathCopyLog) Then
                My.Computer.FileSystem.DeleteFile(pathCopyLog)
            End If
        End If

        'コピーリストを元にファイルをコピーをする
        For Each srcFileName As String In wkCopyList
            Dim srcFileInfo As IO.FileInfo = New IO.FileInfo(srcFileName)
            'コピーログを作成する
            Using sw As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(pathCopyLog, True)
                sw.WriteLine("""" + srcFileInfo.FullName + """")
            End Using
            System.Threading.Thread.Sleep(10)
        Next
    End Sub

    Private Sub BgListSakusei_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BgListSakusei.ProgressChanged

    End Sub

    Private Sub BgListSakusei_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgListSakusei.RunWorkerCompleted
        LogOutput(gLogInformation, "リスト作成終了")
        'ボタンを有効にする
        Enable_Control(True)
    End Sub
End Class
