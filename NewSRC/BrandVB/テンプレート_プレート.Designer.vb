<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_TMP_PLATE
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
	Public WithEvents Command5 As System.Windows.Forms.Button
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_plate_n As System.Windows.Forms.TextBox
	Public WithEvents w_plate_r As System.Windows.Forms.TextBox
	Public WithEvents w_plate_h As System.Windows.Forms.TextBox
	Public WithEvents w_plate_w As System.Windows.Forms.TextBox
	Public WithEvents w_type As System.Windows.Forms.ComboBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label28 As System.Windows.Forms.Label
	Public WithEvents Label19 As System.Windows.Forms.Label
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
	Public WithEvents w_hm_name As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImgThumbnail1 = New System.Windows.Forms.PictureBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Command5 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.w_plate_n = New System.Windows.Forms.TextBox()
        Me.w_plate_r = New System.Windows.Forms.TextBox()
        Me.w_plate_h = New System.Windows.Forms.TextBox()
        Me.w_plate_w = New System.Windows.Forms.TextBox()
        Me.w_type = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.w_hm_name = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImgThumbnail1
        '
        Me.ImgThumbnail1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImgThumbnail1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ImgThumbnail1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ImgThumbnail1.Location = New System.Drawing.Point(48, 288)
        Me.ImgThumbnail1.Name = "ImgThumbnail1"
        Me.ImgThumbnail1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ImgThumbnail1.Size = New System.Drawing.Size(457, 193)
        Me.ImgThumbnail1.TabIndex = 11
        Me.ImgThumbnail1.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Command5)
        Me.Frame1.Controls.Add(Me.Command4)
        Me.Frame1.Controls.Add(Me.Command3)
        Me.Frame1.Controls.Add(Me.Command2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(553, 73)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        '
        'Command5
        '
        Me.Command5.BackColor = System.Drawing.SystemColors.Control
        Me.Command5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command5.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command5.Location = New System.Drawing.Point(24, 24)
        Me.Command5.Name = "Command5"
        Me.Command5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command5.Size = New System.Drawing.Size(120, 37)
        Me.Command5.TabIndex = 2
        Me.Command5.Text = "CAD reading"
        Me.Command5.UseVisualStyleBackColor = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(456, 24)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(74, 37)
        Me.Command4.TabIndex = 5
        Me.Command4.Text = "Ｈｅｌｐ"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(368, 24)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(74, 37)
        Me.Command3.TabIndex = 4
        Me.Command3.Text = "End"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(280, 24)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(74, 37)
        Me.Command2.TabIndex = 3
        Me.Command2.Text = "Clear"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.w_plate_n)
        Me.Frame7.Controls.Add(Me.w_plate_r)
        Me.Frame7.Controls.Add(Me.w_plate_h)
        Me.Frame7.Controls.Add(Me.w_plate_w)
        Me.Frame7.Controls.Add(Me.w_type)
        Me.Frame7.Controls.Add(Me.Label4)
        Me.Frame7.Controls.Add(Me.Label3)
        Me.Frame7.Controls.Add(Me.Label2)
        Me.Frame7.Controls.Add(Me.Label28)
        Me.Frame7.Controls.Add(Me.Label19)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(0, 64)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(553, 129)
        Me.Frame7.TabIndex = 0
        Me.Frame7.TabStop = False
        '
        'w_plate_n
        '
        Me.w_plate_n.AcceptsReturn = True
        Me.w_plate_n.BackColor = System.Drawing.SystemColors.Window
        Me.w_plate_n.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_plate_n.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_plate_n.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_plate_n.Location = New System.Drawing.Point(328, 88)
        Me.w_plate_n.MaxLength = 0
        Me.w_plate_n.Name = "w_plate_n"
        Me.w_plate_n.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_plate_n.Size = New System.Drawing.Size(73, 21)
        Me.w_plate_n.TabIndex = 18
        Me.w_plate_n.Text = "w_plate_n"
        '
        'w_plate_r
        '
        Me.w_plate_r.AcceptsReturn = True
        Me.w_plate_r.BackColor = System.Drawing.SystemColors.Window
        Me.w_plate_r.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_plate_r.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_plate_r.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_plate_r.Location = New System.Drawing.Point(232, 88)
        Me.w_plate_r.MaxLength = 0
        Me.w_plate_r.Name = "w_plate_r"
        Me.w_plate_r.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_plate_r.Size = New System.Drawing.Size(73, 21)
        Me.w_plate_r.TabIndex = 16
        Me.w_plate_r.Text = "w_plate_r"
        '
        'w_plate_h
        '
        Me.w_plate_h.AcceptsReturn = True
        Me.w_plate_h.BackColor = System.Drawing.SystemColors.Window
        Me.w_plate_h.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_plate_h.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_plate_h.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_plate_h.Location = New System.Drawing.Point(136, 88)
        Me.w_plate_h.MaxLength = 0
        Me.w_plate_h.Name = "w_plate_h"
        Me.w_plate_h.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_plate_h.Size = New System.Drawing.Size(73, 21)
        Me.w_plate_h.TabIndex = 14
        Me.w_plate_h.Text = "w_plate_h"
        '
        'w_plate_w
        '
        Me.w_plate_w.AcceptsReturn = True
        Me.w_plate_w.BackColor = System.Drawing.SystemColors.Window
        Me.w_plate_w.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_plate_w.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_plate_w.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_plate_w.Location = New System.Drawing.Point(40, 88)
        Me.w_plate_w.MaxLength = 0
        Me.w_plate_w.Name = "w_plate_w"
        Me.w_plate_w.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_plate_w.Size = New System.Drawing.Size(73, 21)
        Me.w_plate_w.TabIndex = 12
        Me.w_plate_w.Text = "w_plate_w"
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
        Me.w_type.Size = New System.Drawing.Size(161, 22)
        Me.w_type.TabIndex = 1
        Me.w_type.Text = "w_type"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(325, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(122, 25)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Screw position (X)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(229, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(72, 25)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Corner R"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(133, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(49, 25)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Height"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label28
        '
        Me.Label28.BackColor = System.Drawing.SystemColors.Control
        Me.Label28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label28.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label28.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label28.Location = New System.Drawing.Point(37, 64)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label28.Size = New System.Drawing.Size(49, 25)
        Me.Label28.TabIndex = 13
        Me.Label28.Text = "Width"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label19.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label19.Location = New System.Drawing.Point(49, 27)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label19.Size = New System.Drawing.Size(49, 25)
        Me.Label19.TabIndex = 9
        Me.Label19.Text = "Type"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.w_hm_name)
        Me.Frame4.Controls.Add(Me.Label1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 184)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(553, 65)
        Me.Frame4.TabIndex = 7
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
        Me.w_hm_name.Location = New System.Drawing.Point(137, 21)
        Me.w_hm_name.MaxLength = 0
        Me.w_hm_name.Name = "w_hm_name"
        Me.w_hm_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_hm_name.Size = New System.Drawing.Size(161, 21)
        Me.w_hm_name.TabIndex = 8
        Me.w_hm_name.Text = "w_hm_name"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(115, 25)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Editing characters"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'F_TMP_PLATE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(556, 529)
        Me.ControlBox = False
        Me.Controls.Add(Me.ImgThumbnail1)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame7)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(211, 150)
        Me.Name = "F_TMP_PLATE"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Template (plate)"
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class