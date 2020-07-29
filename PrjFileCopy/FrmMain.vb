Public Class FrmMain

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '���O�o��
        LogOutput(gLogInformation, "�A�v���N��")
        'ini �ēǂݍ���
        FrmMain_Initialize()
    End Sub

    Private Sub FrmMain_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        '���O�o��
        LogOutput(gLogInformation, "�A�v���I��")
    End Sub

    Private Sub FrmMain_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Select Case KeyCode
            Case Keys.F5
                'ini �ēǂݍ���
                FrmMain_Initialize()
        End Select
    End Sub

    Private Sub FrmMain_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter
        '�R���g���[�����Ƀh���b�O���ꂽ�Ƃ����s�����
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            '�h���b�O���ꂽ�f�[�^�`���𒲂ׁA�t�@�C���̂Ƃ��̓R�s�[�Ƃ���
            e.Effect = DragDropEffects.Copy
        Else
            '�t�@�C���ȊO�͎󂯕t���Ȃ�
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub FrmMain_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop
        '�h���b�v���ꂽ���ׂẴt�@�C�������擾����
        Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        Dim file As String = Nothing
        '�h���b�v���ꂽ���ׂẴt�@�C������擪1������ΏۂƂ���
        For Each wk As String In files
            file = wk
            Exit For
        Next wk
        '�S�~�͖�������
        If IsNothing(file) Or (Not My.Computer.FileSystem.FileExists(file)) Then
            Return
        End If
        ' �`�F�b�N����
        If False = IsDragAndDropStartOk() Then
            Return
        End If
        '���s����
        DoFileCopy(file)
    End Sub

    Private Sub BtnDstDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDstDir.Click
        '�_�C�A���O���o��
        DlgDstDir.Description = LblHozon.Text + "�̃t�H���_���w�肵�Ă��������B"
        DlgDstDir.RootFolder = Environment.SpecialFolder.Desktop
        DlgDstDir.SelectedPath = TxtDstDir.Text
        '�I�������t�H���_����ʂɕ\��
        If DlgDstDir.ShowDialog(Me) = DialogResult.OK Then
            TxtDstDir.Text = DlgDstDir.SelectedPath
        End If
    End Sub

    Private Sub BtnSrcDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSrcDir.Click
        '�_�C�A���O���o��
        DlgSrcDir.Description = LblKensaku.Text + "�̃t�H���_���w�肵�Ă��������B"
        DlgSrcDir.RootFolder = Environment.SpecialFolder.Desktop
        DlgSrcDir.SelectedPath = TxtSrcDir.Text
        '�I�������t�H���_����ʂɕ\��
        If DlgSrcDir.ShowDialog(Me) = DialogResult.OK Then
            TxtSrcDir.Text = DlgSrcDir.SelectedPath
        End If
    End Sub

    Private Sub FrmMain_Initialize()
        'ini���ēǂݍ���
        Initial_IniUpload()
        'ini�̓��e����ʂɃZ�b�g
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

        '�O��`�F�b�N
        If String.IsNullOrEmpty(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "�͕K�{�ł��B", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "�����݂��܂���B", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        IsDragAndDropStartOk = True
    End Function

    Private Function IsJikkoStartOk() As Boolean
        IsJikkoStartOk = False

        '�O��`�F�b�N
        If String.IsNullOrEmpty(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "�͕K�{�ł��B", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtDstDir.Text) Then
            MsgBox(LblHozon.Text + "�����݂��܂���B", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If String.IsNullOrEmpty(TxtSrcDir.Text) Then
            MsgBox(LblKensaku.Text + "�͕K�{�ł��B", MsgBoxStyle.OkOnly, Me.Text)
            BtnSrcDir.Focus()
            Exit Function
        End If

        If Not My.Computer.FileSystem.DirectoryExists(TxtSrcDir.Text) Then
            MsgBox(LblKensaku.Text + "�����݂��܂���B", MsgBoxStyle.OkOnly, Me.Text)
            BtnSrcDir.Focus()
            Exit Function
        End If

        If TxtDstDir.Text = TxtSrcDir.Text Then
            MsgBox(LblHozon.Text + "��" + LblKensaku.Text + "�͕ʁX�̃t�H���_���w�肵�Ă��������B", MsgBoxStyle.OkOnly, Me.Text)
            BtnDstDir.Focus()
            Exit Function
        End If

        If Not IsDate(DateKoshinhiFrom.Value) Then
            MsgBox("�X�V����[�J�n]���s���ł��B", MsgBoxStyle.OkOnly, Me.Text)
            DateKoshinhiFrom.Focus()
            Exit Function
        End If

        If Not IsDate(DateKoshinhiTo.Value) Then
            MsgBox("�X�V����[�I��]���s���ł��B", MsgBoxStyle.OkOnly, Me.Text)
            DateKoshinhiTo.Focus()
            Exit Function
        End If

        If DateTime.Compare(DateKoshinhiFrom.Value, DateKoshinhiTo.Value) > 0 Then
            MsgBox("�X�V�����̊J�n�A�I�����t�]���Ă��܂��B", MsgBoxStyle.OkOnly, Me.Text)
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
                MsgBox("*�܂���*.*����͂����ꍇ�A���̏����͎w��ł��܂���B", MsgBoxStyle.OkOnly, Me.Text)
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
                    MsgBox("*�܂���*.*�͎w��ł��܂���B", MsgBoxStyle.OkOnly, Me.Text)
                    TxtNgPattern.Focus()
                    Exit Function
                End If
            Next
        End If

        IsJikkoStartOk = True
    End Function

    Private Sub BtnJikko_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFileCopy.Click
        '�`�F�b�N����
        If False = IsJikkoStartOk() Then
            Return
        End If
        '���s����
        DoFileCopy()
    End Sub

    Private Sub BtnListSakusei_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnListSakusei.Click
        '�`�F�b�N����
        If False = IsJikkoStartOk() Then
            Return
        End If
        '���s����
        DoListSakusei()
    End Sub

    Private Sub DoFileCopy(Optional ByVal filePath As String = "")
        '�������Ƀ{�^���������ꂽ���Ȃ��̂Ń{�^���𖳌��ɂ���
        Enable_Control(False)
        '�R�s�[����
        BgFileCopy.RunWorkerAsync(filePath)
    End Sub

    Private Sub DoListSakusei(Optional ByVal filePath As String = "")
        '�������Ƀ{�^���������ꂽ���Ȃ��̂Ń{�^���𖳌��ɂ���
        Enable_Control(False)
        '�쐬����
        BgListSakusei.RunWorkerAsync(filePath)
    End Sub

    Private Sub BgFileCopy_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgFileCopy.DoWork
        Dim pathCopyList As String = CStr(e.Argument)
        Dim pathCopyLog As String = ""
        Dim wkCopyList As List(Of String) = New List(Of String)

        '��������R�s�[���O�̖��̌���
        If False = ChkOutput.Checked Then
            pathCopyLog = ""
        ElseIf String.IsNullOrEmpty(GsCopyLogFile) Then
            pathCopyLog = TxtDstDir.Text + "\" + Now.ToString("yyyyMMddHHmmss") + "COPY.log"
        Else
            pathCopyLog = TxtDstDir.Text + "\" + GsCopyLogFile
        End If

        '�R�s�[�ꗗ�̎w�肪�Ȃ����͌�����������R�s�[�ꗗ���쐬����
        If String.IsNullOrEmpty(pathCopyList) Then
            Dim diKensaku As New System.IO.DirectoryInfo(TxtSrcDir.Text)
            Dim opKensaku As System.IO.SearchOption = System.IO.SearchOption.TopDirectoryOnly
            If ChkSubKensaku.Checked Then
                opKensaku = System.IO.SearchOption.AllDirectories
            End If

            '���O������Ă�t�@�C���������疳�����폜
            If My.Computer.FileSystem.FileExists(pathCopyList) Then
                My.Computer.FileSystem.DeleteFile(pathCopyList)
            End If

            '�S����
            For Each tempFile As System.IO.FileInfo In diKensaku.GetFiles("*", opKensaku)
                '�X�V����[From]��菬����
                If tempFile.LastWriteTime < DateKoshinhiFrom.Value Then
                    Continue For
                End If
                '�X�V����[To]���傫��
                If DateKoshinhiTo.Value < tempFile.LastWriteTime Then
                    Continue For
                End If
                'OK���[�h�ɂȂ�������_��
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
                'NG���[�h�ɍڂ��Ă�����_��
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
                '�����܂ł��ꂽ��R�s�[���X�g�ɒǉ�
                wkCopyList.Add(tempFile.FullName)
            Next

        Else
            '�t�@�C������R�s�[���X�g���쐬
            Using sr As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(pathCopyList)
                While (sr.Peek() >= 0)
                    Dim filePath As String = sr.ReadLine().Trim
                    '�ŏ��ƍŌオ�_�u���N�H�[�e�[�V�����̏ꍇ�͎�菜��
                    If filePath.StartsWith("""") And filePath.EndsWith("""") Then
                        filePath = filePath.Substring(1, filePath.Length - 2)
                    End If
                    If My.Computer.FileSystem.FileExists(filePath) Then
                        wkCopyList.Add(filePath)
                    Else
                        LogOutput(gLogError, "�t�@�C�������݂��Ȃ����߃R�s�[�ł��܂���B[" + filePath + "]")
                    End If
                End While
            End Using
        End If

        '�R�s�[����̂��Ȃ������瑦�I��
        If wkCopyList.Count < 1 Then
            Exit Sub
        End If

        '�R�s�[���O���쐬����̂Ŏ��O����
        If Not String.IsNullOrEmpty(pathCopyLog) Then
            If My.Computer.FileSystem.FileExists(pathCopyLog) Then
                My.Computer.FileSystem.DeleteFile(pathCopyLog)
            End If
        End If

        '�R�s�[���X�g�����Ƀt�@�C�����R�s�[������
        For Each srcFileName As String In wkCopyList
            Dim srcFileInfo As IO.FileInfo = New IO.FileInfo(srcFileName)
            Dim dstFileInfo As IO.FileInfo = New IO.FileInfo(TxtDstDir.Text + srcFileName.Substring(srcFileName.IndexOf(":") + 1))
            '�t�H���_�쐬
            If Not My.Computer.FileSystem.DirectoryExists(dstFileInfo.DirectoryName) Then
                My.Computer.FileSystem.CreateDirectory(dstFileInfo.DirectoryName)
            End If
            '�R�s�[
            srcFileInfo.CopyTo(dstFileInfo.FullName, True)
            '�R�s�[���O���쐬����
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
        LogOutput(gLogInformation, "�R�s�[�����I��")
        '�{�^����L���ɂ���
        Enable_Control(True)
    End Sub

    Private Sub BgListSakusei_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BgListSakusei.DoWork
        Dim pathCopyLog As String = TxtDstDir.Text + "\" + Now.ToString("yyyyMMddHHmmss") + "_���o����.txt"
        Dim wkCopyList As List(Of String) = New List(Of String)

        Dim diKensaku As New System.IO.DirectoryInfo(TxtSrcDir.Text)
        Dim opKensaku As System.IO.SearchOption = System.IO.SearchOption.TopDirectoryOnly
        If ChkSubKensaku.Checked Then
            opKensaku = System.IO.SearchOption.AllDirectories
        End If

        '�S����
        For Each tempFile As System.IO.FileInfo In diKensaku.GetFiles("*", opKensaku)
            '�X�V����[From]��菬����
            If tempFile.LastWriteTime < DateKoshinhiFrom.Value Then
                Continue For
            End If
            '�X�V����[To]���傫��
            If DateKoshinhiTo.Value < tempFile.LastWriteTime Then
                Continue For
            End If
            'OK���[�h�ɂȂ�������_��
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
            'NG���[�h�ɍڂ��Ă�����_��
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
            '�����܂ł��ꂽ��R�s�[���X�g�ɒǉ�
            wkCopyList.Add(tempFile.FullName)
        Next

        '�R�s�[����̂��Ȃ������瑦�I��
        If wkCopyList.Count < 1 Then
            Exit Sub
        End If

        '�R�s�[���O���쐬����̂Ŏ��O����
        If Not String.IsNullOrEmpty(pathCopyLog) Then
            If My.Computer.FileSystem.FileExists(pathCopyLog) Then
                My.Computer.FileSystem.DeleteFile(pathCopyLog)
            End If
        End If

        '�R�s�[���X�g�����Ƀt�@�C�����R�s�[������
        For Each srcFileName As String In wkCopyList
            Dim srcFileInfo As IO.FileInfo = New IO.FileInfo(srcFileName)
            '�R�s�[���O���쐬����
            Using sw As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(pathCopyLog, True)
                sw.WriteLine("""" + srcFileInfo.FullName + """")
            End Using
            System.Threading.Thread.Sleep(10)
        Next
    End Sub

    Private Sub BgListSakusei_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BgListSakusei.ProgressChanged

    End Sub

    Private Sub BgListSakusei_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BgListSakusei.RunWorkerCompleted
        LogOutput(gLogInformation, "���X�g�쐬�I��")
        '�{�^����L���ɂ���
        Enable_Control(True)
    End Sub
End Class
