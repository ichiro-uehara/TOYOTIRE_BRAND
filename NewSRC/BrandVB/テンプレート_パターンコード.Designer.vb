<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_TMP_PTNCODE
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents ImgThumbnail1 As System.Windows.Forms.PictureBox
	Public WithEvents w_hm_name As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents w_ptncode As System.Windows.Forms.TextBox
	Public WithEvents w_type As System.Windows.Forms.ComboBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label19 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command6 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImgThumbnail1 = New System.Windows.Forms.PictureBox()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.w_hm_name = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.w_ptncode = New System.Windows.Forms.TextBox()
        Me.w_type = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command6 = New System.Windows.Forms.Button()
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame4.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImgThumbnail1
        '
        Me.ImgThumbnail1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImgThumbnail1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ImgThumbnail1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ImgThumbnail1.Location = New System.Drawing.Point(64, 240)
        Me.ImgThumbnail1.Name = "ImgThumbnail1"
        Me.ImgThumbnail1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ImgThumbnail1.Size = New System.Drawing.Size(457, 193)
        Me.ImgThumbnail1.TabIndex = 13
        Me.ImgThumbnail1.TabStop = False
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.w_hm_name)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 120)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(561, 65)
        Me.Frame4.TabIndex = 11
        Me.Frame4.TabStop = False
        '
        'w_hm_name
        '
        Me.w_hm_name.AcceptsReturn = True
        Me.w_hm_name.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_hm_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_hm_name.Enabled = False
        Me.w_hm_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_hm_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_hm_name.Location = New System.Drawing.Point(136, 24)
        Me.w_hm_name.MaxLength = 0
        Me.w_hm_name.Name = "w_hm_name"
        Me.w_hm_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_hm_name.Size = New System.Drawing.Size(161, 21)
        Me.w_hm_name.TabIndex = 7
        Me.w_hm_name.TabStop = False
        Me.w_hm_name.Text = "w_hm_name"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(114, 25)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Editing characters"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.w_ptncode)
        Me.Frame7.Controls.Add(Me.w_type)
        Me.Frame7.Controls.Add(Me.Label19)
        Me.Frame7.Controls.Add(Me.Label3)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(0, 64)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(561, 65)
        Me.Frame7.TabIndex = 8
        Me.Frame7.TabStop = False
        '
        'w_ptncode
        '
        Me.w_ptncode.AcceptsReturn = True
        Me.w_ptncode.BackColor = System.Drawing.SystemColors.Window
        Me.w_ptncode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_ptncode.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_ptncode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_ptncode.Location = New System.Drawing.Point(408, 24)
        Me.w_ptncode.MaxLength = 6
        Me.w_ptncode.Name = "w_ptncode"
        Me.w_ptncode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_ptncode.Size = New System.Drawing.Size(97, 21)
        Me.w_ptncode.TabIndex = 2
        Me.w_ptncode.Text = "w_ptnco"
        '
        'w_type
        '
        Me.w_type.BackColor = System.Drawing.SystemColors.Window
        Me.w_type.Cursor = System.Windows.Forms.Cursors.Default
        Me.w_type.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_type.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_type.Location = New System.Drawing.Point(104, 24)
        Me.w_type.Name = "w_type"
        Me.w_type.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_type.Size = New System.Drawing.Size(137, 22)
        Me.w_type.TabIndex = 1
        Me.w_type.Text = "w_type"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(25, 27)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(73, 25)
        Me.Label19.TabIndex = 10
        Me.Label19.Text = "Font"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(289, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(113, 25)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Pattern code"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Command2)
        Me.Frame1.Controls.Add(Me.Command3)
        Me.Frame1.Controls.Add(Me.Command4)
        Me.Frame1.Controls.Add(Me.Command6)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(561, 73)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(296, 24)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(74, 37)
        Me.Command2.TabIndex = 4
        Me.Command2.Text = "Clear"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(384, 24)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(74, 37)
        Me.Command3.TabIndex = 5
        Me.Command3.Text = "End"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(464, 24)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(74, 37)
        Me.Command4.TabIndex = 6
        Me.Command4.Text = "Ｈｅｌｐ"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Command6
        '
        Me.Command6.BackColor = System.Drawing.SystemColors.Control
        Me.Command6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command6.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command6.Location = New System.Drawing.Point(24, 24)
        Me.Command6.Name = "Command6"
        Me.Command6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command6.Size = New System.Drawing.Size(148, 37)
        Me.Command6.TabIndex = 3
        Me.Command6.Text = "Substitution read"
        Me.Command6.UseVisualStyleBackColor = False
        '
        'F_TMP_PTNCODE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(575, 504)
        Me.ControlBox = False
        Me.Controls.Add(Me.ImgThumbnail1)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame7)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(229, 150)
        Me.Name = "F_TMP_PTNCODE"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Template (pattern code)"
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class