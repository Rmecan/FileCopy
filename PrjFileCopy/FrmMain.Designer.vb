<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMain
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.LblHozon = New System.Windows.Forms.Label
        Me.LblKensaku = New System.Windows.Forms.Label
        Me.LblKoshinhiFrom = New System.Windows.Forms.Label
        Me.LblKoshinhiTo = New System.Windows.Forms.Label
        Me.DateKoshinhiFrom = New System.Windows.Forms.DateTimePicker
        Me.DateKoshinhiTo = New System.Windows.Forms.DateTimePicker
        Me.BtnDstDir = New System.Windows.Forms.Button
        Me.BtnSrcDir = New System.Windows.Forms.Button
        Me.ChkSubKensaku = New System.Windows.Forms.CheckBox
        Me.TxtDstDir = New System.Windows.Forms.TextBox
        Me.TxtSrcDir = New System.Windows.Forms.TextBox
        Me.GrpKoshinhi = New System.Windows.Forms.GroupBox
        Me.LblOkPattern = New System.Windows.Forms.Label
        Me.LblNgPattern = New System.Windows.Forms.Label
        Me.TxtOkPattern = New System.Windows.Forms.TextBox
        Me.TxtNgPattern = New System.Windows.Forms.TextBox
        Me.BtnFileCopy = New System.Windows.Forms.Button
        Me.BgFileCopy = New System.ComponentModel.BackgroundWorker
        Me.ChkOutput = New System.Windows.Forms.CheckBox
        Me.DlgDstDir = New System.Windows.Forms.FolderBrowserDialog
        Me.DlgSrcDir = New System.Windows.Forms.FolderBrowserDialog
        Me.BgListSakusei = New System.ComponentModel.BackgroundWorker
        Me.BtnListSakusei = New System.Windows.Forms.Button
        Me.GrpKoshinhi.SuspendLayout()
        Me.SuspendLayout()
        '
        'LblHozon
        '
        Me.LblHozon.AutoSize = True
        Me.LblHozon.Location = New System.Drawing.Point(12, 11)
        Me.LblHozon.Name = "LblHozon"
        Me.LblHozon.Size = New System.Drawing.Size(41, 12)
        Me.LblHozon.TabIndex = 0
        Me.LblHozon.Text = "保存先"
        '
        'LblKensaku
        '
        Me.LblKensaku.AutoSize = True
        Me.LblKensaku.Location = New System.Drawing.Point(12, 39)
        Me.LblKensaku.Name = "LblKensaku"
        Me.LblKensaku.Size = New System.Drawing.Size(41, 12)
        Me.LblKensaku.TabIndex = 1
        Me.LblKensaku.Text = "参照元"
        '
        'LblKoshinhiFrom
        '
        Me.LblKoshinhiFrom.AutoSize = True
        Me.LblKoshinhiFrom.Location = New System.Drawing.Point(10, 20)
        Me.LblKoshinhiFrom.Name = "LblKoshinhiFrom"
        Me.LblKoshinhiFrom.Size = New System.Drawing.Size(41, 12)
        Me.LblKoshinhiFrom.TabIndex = 2
        Me.LblKoshinhiFrom.Text = "開始日"
        '
        'LblKoshinhiTo
        '
        Me.LblKoshinhiTo.AutoSize = True
        Me.LblKoshinhiTo.Location = New System.Drawing.Point(10, 44)
        Me.LblKoshinhiTo.Name = "LblKoshinhiTo"
        Me.LblKoshinhiTo.Size = New System.Drawing.Size(41, 12)
        Me.LblKoshinhiTo.TabIndex = 3
        Me.LblKoshinhiTo.Text = "終了日"
        '
        'DateKoshinhiFrom
        '
        Me.DateKoshinhiFrom.CustomFormat = "yyyy/MM/dd HH:mm:ss"
        Me.DateKoshinhiFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateKoshinhiFrom.Location = New System.Drawing.Point(57, 18)
        Me.DateKoshinhiFrom.Name = "DateKoshinhiFrom"
        Me.DateKoshinhiFrom.Size = New System.Drawing.Size(137, 19)
        Me.DateKoshinhiFrom.TabIndex = 4
        '
        'DateKoshinhiTo
        '
        Me.DateKoshinhiTo.CustomFormat = "yyyy/MM/dd HH:mm:ss"
        Me.DateKoshinhiTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateKoshinhiTo.Location = New System.Drawing.Point(57, 40)
        Me.DateKoshinhiTo.Name = "DateKoshinhiTo"
        Me.DateKoshinhiTo.Size = New System.Drawing.Size(137, 19)
        Me.DateKoshinhiTo.TabIndex = 5
        '
        'BtnDstDir
        '
        Me.BtnDstDir.Location = New System.Drawing.Point(333, 6)
        Me.BtnDstDir.Name = "BtnDstDir"
        Me.BtnDstDir.Size = New System.Drawing.Size(51, 23)
        Me.BtnDstDir.TabIndex = 6
        Me.BtnDstDir.Text = "選択"
        Me.BtnDstDir.UseVisualStyleBackColor = True
        '
        'BtnSrcDir
        '
        Me.BtnSrcDir.Location = New System.Drawing.Point(333, 34)
        Me.BtnSrcDir.Name = "BtnSrcDir"
        Me.BtnSrcDir.Size = New System.Drawing.Size(51, 23)
        Me.BtnSrcDir.TabIndex = 7
        Me.BtnSrcDir.Text = "選択"
        Me.BtnSrcDir.UseVisualStyleBackColor = True
        '
        'ChkSubKensaku
        '
        Me.ChkSubKensaku.AutoSize = True
        Me.ChkSubKensaku.Checked = True
        Me.ChkSubKensaku.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkSubKensaku.Location = New System.Drawing.Point(5, 222)
        Me.ChkSubKensaku.Name = "ChkSubKensaku"
        Me.ChkSubKensaku.Size = New System.Drawing.Size(168, 16)
        Me.ChkSubKensaku.TabIndex = 8
        Me.ChkSubKensaku.Text = "サブディレクトリ以下も検索する"
        Me.ChkSubKensaku.UseVisualStyleBackColor = True
        '
        'TxtDstDir
        '
        Me.TxtDstDir.Location = New System.Drawing.Point(64, 8)
        Me.TxtDstDir.Name = "TxtDstDir"
        Me.TxtDstDir.ReadOnly = True
        Me.TxtDstDir.Size = New System.Drawing.Size(259, 19)
        Me.TxtDstDir.TabIndex = 10
        '
        'TxtSrcDir
        '
        Me.TxtSrcDir.Location = New System.Drawing.Point(64, 36)
        Me.TxtSrcDir.Name = "TxtSrcDir"
        Me.TxtSrcDir.ReadOnly = True
        Me.TxtSrcDir.Size = New System.Drawing.Size(259, 19)
        Me.TxtSrcDir.TabIndex = 11
        '
        'GrpKoshinhi
        '
        Me.GrpKoshinhi.Controls.Add(Me.DateKoshinhiTo)
        Me.GrpKoshinhi.Controls.Add(Me.DateKoshinhiFrom)
        Me.GrpKoshinhi.Controls.Add(Me.LblKoshinhiTo)
        Me.GrpKoshinhi.Controls.Add(Me.LblKoshinhiFrom)
        Me.GrpKoshinhi.Location = New System.Drawing.Point(7, 139)
        Me.GrpKoshinhi.Name = "GrpKoshinhi"
        Me.GrpKoshinhi.Size = New System.Drawing.Size(214, 73)
        Me.GrpKoshinhi.TabIndex = 12
        Me.GrpKoshinhi.TabStop = False
        Me.GrpKoshinhi.Text = "更新日時"
        '
        'LblOkPattern
        '
        Me.LblOkPattern.AutoSize = True
        Me.LblOkPattern.Location = New System.Drawing.Point(5, 58)
        Me.LblOkPattern.Name = "LblOkPattern"
        Me.LblOkPattern.Size = New System.Drawing.Size(66, 12)
        Me.LblOkPattern.TabIndex = 13
        Me.LblOkPattern.Text = "検索パターン"
        '
        'LblNgPattern
        '
        Me.LblNgPattern.AutoSize = True
        Me.LblNgPattern.Location = New System.Drawing.Point(5, 99)
        Me.LblNgPattern.Name = "LblNgPattern"
        Me.LblNgPattern.Size = New System.Drawing.Size(66, 12)
        Me.LblNgPattern.TabIndex = 14
        Me.LblNgPattern.Text = "除外パターン"
        '
        'TxtOkPattern
        '
        Me.TxtOkPattern.Location = New System.Drawing.Point(5, 73)
        Me.TxtOkPattern.Name = "TxtOkPattern"
        Me.TxtOkPattern.Size = New System.Drawing.Size(379, 19)
        Me.TxtOkPattern.TabIndex = 15
        '
        'TxtNgPattern
        '
        Me.TxtNgPattern.Location = New System.Drawing.Point(5, 114)
        Me.TxtNgPattern.Name = "TxtNgPattern"
        Me.TxtNgPattern.Size = New System.Drawing.Size(379, 19)
        Me.TxtNgPattern.TabIndex = 16
        '
        'BtnFileCopy
        '
        Me.BtnFileCopy.Font = New System.Drawing.Font("MS UI Gothic", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnFileCopy.Location = New System.Drawing.Point(256, 222)
        Me.BtnFileCopy.Name = "BtnFileCopy"
        Me.BtnFileCopy.Size = New System.Drawing.Size(128, 60)
        Me.BtnFileCopy.TabIndex = 17
        Me.BtnFileCopy.Text = "コピー実行"
        Me.BtnFileCopy.UseVisualStyleBackColor = True
        '
        'BgFileCopy
        '
        '
        'ChkOutput
        '
        Me.ChkOutput.AutoSize = True
        Me.ChkOutput.Checked = True
        Me.ChkOutput.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkOutput.Location = New System.Drawing.Point(5, 244)
        Me.ChkOutput.Name = "ChkOutput"
        Me.ChkOutput.Size = New System.Drawing.Size(147, 16)
        Me.ChkOutput.TabIndex = 18
        Me.ChkOutput.Text = "保存先にコピーログを出力"
        Me.ChkOutput.UseVisualStyleBackColor = True
        '
        'DlgSrcDir
        '
        Me.DlgSrcDir.RootFolder = System.Environment.SpecialFolder.MyComputer
        '
        'BgListSakusei
        '
        '
        'BtnListSakusei
        '
        Me.BtnListSakusei.Font = New System.Drawing.Font("MS UI Gothic", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.BtnListSakusei.Location = New System.Drawing.Point(256, 152)
        Me.BtnListSakusei.Name = "BtnListSakusei"
        Me.BtnListSakusei.Size = New System.Drawing.Size(128, 60)
        Me.BtnListSakusei.TabIndex = 19
        Me.BtnListSakusei.Text = "抽出リスト作成"
        Me.BtnListSakusei.UseVisualStyleBackColor = True
        '
        'FrmMain
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(399, 295)
        Me.Controls.Add(Me.BtnListSakusei)
        Me.Controls.Add(Me.ChkOutput)
        Me.Controls.Add(Me.BtnFileCopy)
        Me.Controls.Add(Me.TxtNgPattern)
        Me.Controls.Add(Me.TxtOkPattern)
        Me.Controls.Add(Me.LblNgPattern)
        Me.Controls.Add(Me.LblOkPattern)
        Me.Controls.Add(Me.GrpKoshinhi)
        Me.Controls.Add(Me.TxtSrcDir)
        Me.Controls.Add(Me.TxtDstDir)
        Me.Controls.Add(Me.ChkSubKensaku)
        Me.Controls.Add(Me.BtnSrcDir)
        Me.Controls.Add(Me.BtnDstDir)
        Me.Controls.Add(Me.LblKensaku)
        Me.Controls.Add(Me.LblHozon)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmMain"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "ファイルコピー"
        Me.GrpKoshinhi.ResumeLayout(False)
        Me.GrpKoshinhi.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LblHozon As System.Windows.Forms.Label
    Friend WithEvents LblKensaku As System.Windows.Forms.Label
    Friend WithEvents LblKoshinhiFrom As System.Windows.Forms.Label
    Friend WithEvents LblKoshinhiTo As System.Windows.Forms.Label
    Friend WithEvents DateKoshinhiFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateKoshinhiTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents BtnDstDir As System.Windows.Forms.Button
    Friend WithEvents BtnSrcDir As System.Windows.Forms.Button
    Friend WithEvents ChkSubKensaku As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDstDir As System.Windows.Forms.TextBox
    Friend WithEvents TxtSrcDir As System.Windows.Forms.TextBox
    Friend WithEvents GrpKoshinhi As System.Windows.Forms.GroupBox
    Friend WithEvents LblOkPattern As System.Windows.Forms.Label
    Friend WithEvents LblNgPattern As System.Windows.Forms.Label
    Friend WithEvents TxtOkPattern As System.Windows.Forms.TextBox
    Friend WithEvents TxtNgPattern As System.Windows.Forms.TextBox
    Friend WithEvents BtnFileCopy As System.Windows.Forms.Button
    Friend WithEvents BgFileCopy As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkOutput As System.Windows.Forms.CheckBox
    Friend WithEvents DlgDstDir As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents DlgSrcDir As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BgListSakusei As System.ComponentModel.BackgroundWorker
    Friend WithEvents BtnListSakusei As System.Windows.Forms.Button

End Class
