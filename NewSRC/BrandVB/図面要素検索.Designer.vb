<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_ZSEARCH_YOUSO
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
	Public WithEvents w_show_no As System.Windows.Forms.TextBox
	Public WithEvents w_total As System.Windows.Forms.TextBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    'Public WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents cmd_Cancel As System.Windows.Forms.Button
	Public WithEvents cmd_ZumenRead As System.Windows.Forms.Button
	Public WithEvents cmd_Help As System.Windows.Forms.Button
	Public WithEvents cmd_End As System.Windows.Forms.Button
	Public WithEvents cmd_Clear As System.Windows.Forms.Button
	Public WithEvents cmd_Search As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_taisho As System.Windows.Forms.ComboBox
	Public WithEvents w_mojicd As System.Windows.Forms.TextBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_ZSEARCH_YOUSO))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImgThumbnail1 = New System.Windows.Forms.PictureBox()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.w_show_no = New System.Windows.Forms.TextBox()
        Me.w_total = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        'Me.MSFlexGrid1 = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cmd_Cancel = New System.Windows.Forms.Button()
        Me.cmd_ZumenRead = New System.Windows.Forms.Button()
        Me.cmd_Help = New System.Windows.Forms.Button()
        Me.cmd_End = New System.Windows.Forms.Button()
        Me.cmd_Clear = New System.Windows.Forms.Button()
        Me.cmd_Search = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.w_taisho = New System.Windows.Forms.ComboBox()
        Me.w_mojicd = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        Me.Frame4.SuspendLayout()
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImgThumbnail1
        '
        Me.ImgThumbnail1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImgThumbnail1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ImgThumbnail1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ImgThumbnail1.Location = New System.Drawing.Point(56, 144)
        Me.ImgThumbnail1.Name = "ImgThumbnail1"
        Me.ImgThumbnail1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ImgThumbnail1.Size = New System.Drawing.Size(457, 193)
        Me.ImgThumbnail1.TabIndex = 18
        Me.ImgThumbnail1.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.w_show_no)
        Me.Frame2.Controls.Add(Me.w_total)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(0, 352)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(561, 57)
        Me.Frame2.TabIndex = 11
        Me.Frame2.TabStop = False
        '
        'w_show_no
        '
        Me.w_show_no.AcceptsReturn = True
        Me.w_show_no.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_show_no.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_show_no.Enabled = False
        Me.w_show_no.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_show_no.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_show_no.Location = New System.Drawing.Point(472, 21)
        Me.w_show_no.MaxLength = 0
        Me.w_show_no.Name = "w_show_no"
        Me.w_show_no.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_show_no.Size = New System.Drawing.Size(65, 21)
        Me.w_show_no.TabIndex = 16
        Me.w_show_no.Text = "w_show_no"
        Me.w_show_no.Visible = False
        '
        'w_total
        '
        Me.w_total.AcceptsReturn = True
        Me.w_total.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_total.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_total.Enabled = False
        Me.w_total.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_total.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_total.Location = New System.Drawing.Point(238, 21)
        Me.w_total.MaxLength = 0
        Me.w_total.Name = "w_total"
        Me.w_total.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_total.Size = New System.Drawing.Size(73, 21)
        Me.w_total.TabIndex = 12
        Me.w_total.Text = "w_total"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(24, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(208, 18)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Conditions corresponding number"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        'Me.Frame4.Controls.Add(Me.MSFlexGrid1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 408)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(561, 281)
        Me.Frame4.TabIndex = 15
        Me.Frame4.TabStop = False
        '
        'MSFlexGrid1
        '
        'Me.MSFlexGrid1.Location = New System.Drawing.Point(8, 24)
        'Me.MSFlexGrid1.Name = "MSFlexGrid1"
        'Me.MSFlexGrid1.OcxState = CType(resources.GetObject("MSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        'Me.MSFlexGrid1.Size = New System.Drawing.Size(545, 249)
        'Me.MSFlexGrid1.TabIndex = 17
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmd_Cancel)
        Me.Frame1.Controls.Add(Me.cmd_ZumenRead)
        Me.Frame1.Controls.Add(Me.cmd_Help)
        Me.Frame1.Controls.Add(Me.cmd_End)
        Me.Frame1.Controls.Add(Me.cmd_Clear)
        Me.Frame1.Controls.Add(Me.cmd_Search)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(561, 57)
        Me.Frame1.TabIndex = 8
        Me.Frame1.TabStop = False
        '
        'cmd_Cancel
        '
        Me.cmd_Cancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Cancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Cancel.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Cancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Cancel.Location = New System.Drawing.Point(304, 15)
        Me.cmd_Cancel.Name = "cmd_Cancel"
        Me.cmd_Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Cancel.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Cancel.TabIndex = 5
        Me.cmd_Cancel.Text = "Cancel"
        Me.cmd_Cancel.UseVisualStyleBackColor = False
        '
        'cmd_ZumenRead
        '
        Me.cmd_ZumenRead.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_ZumenRead.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_ZumenRead.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_ZumenRead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_ZumenRead.Location = New System.Drawing.Point(96, 16)
        Me.cmd_ZumenRead.Name = "cmd_ZumenRead"
        Me.cmd_ZumenRead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_ZumenRead.Size = New System.Drawing.Size(122, 33)
        Me.cmd_ZumenRead.TabIndex = 3
        Me.cmd_ZumenRead.Text = "Drawing reading"
        Me.cmd_ZumenRead.UseVisualStyleBackColor = False
        '
        'cmd_Help
        '
        Me.cmd_Help.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Help.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Help.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Help.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Help.Location = New System.Drawing.Point(472, 16)
        Me.cmd_Help.Name = "cmd_Help"
        Me.cmd_Help.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Help.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Help.TabIndex = 7
        Me.cmd_Help.Text = "Ｈｅｌｐ"
        Me.cmd_Help.UseVisualStyleBackColor = False
        '
        'cmd_End
        '
        Me.cmd_End.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_End.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_End.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_End.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_End.Location = New System.Drawing.Point(384, 15)
        Me.cmd_End.Name = "cmd_End"
        Me.cmd_End.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_End.Size = New System.Drawing.Size(74, 33)
        Me.cmd_End.TabIndex = 6
        Me.cmd_End.Text = "End"
        Me.cmd_End.UseVisualStyleBackColor = False
        '
        'cmd_Clear
        '
        Me.cmd_Clear.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Clear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Clear.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Clear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Clear.Location = New System.Drawing.Point(224, 16)
        Me.cmd_Clear.Name = "cmd_Clear"
        Me.cmd_Clear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Clear.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Clear.TabIndex = 4
        Me.cmd_Clear.Text = "Clear"
        Me.cmd_Clear.UseVisualStyleBackColor = False
        '
        'cmd_Search
        '
        Me.cmd_Search.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Search.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Search.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Search.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Search.Location = New System.Drawing.Point(16, 16)
        Me.cmd_Search.Name = "cmd_Search"
        Me.cmd_Search.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Search.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Search.TabIndex = 2
        Me.cmd_Search.Text = "Search"
        Me.cmd_Search.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.w_taisho)
        Me.Frame3.Controls.Add(Me.w_mojicd)
        Me.Frame3.Controls.Add(Me.Label7)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(0, 48)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(561, 65)
        Me.Frame3.TabIndex = 9
        Me.Frame3.TabStop = False
        '
        'w_taisho
        '
        Me.w_taisho.BackColor = System.Drawing.SystemColors.Window
        Me.w_taisho.Cursor = System.Windows.Forms.Cursors.Default
        Me.w_taisho.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_taisho.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_taisho.Location = New System.Drawing.Point(407, 29)
        Me.w_taisho.Name = "w_taisho"
        Me.w_taisho.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_taisho.Size = New System.Drawing.Size(137, 22)
        Me.w_taisho.TabIndex = 1
        Me.w_taisho.Text = "w_taisho"
        '
        'w_mojicd
        '
        Me.w_mojicd.AcceptsReturn = True
        Me.w_mojicd.BackColor = System.Drawing.SystemColors.Window
        Me.w_mojicd.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_mojicd.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_mojicd.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_mojicd.Location = New System.Drawing.Point(137, 29)
        Me.w_mojicd.MaxLength = 0
        Me.w_mojicd.Name = "w_mojicd"
        Me.w_mojicd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_mojicd.Size = New System.Drawing.Size(161, 21)
        Me.w_mojicd.TabIndex = 0
        Me.w_mojicd.Text = "w_mojicd"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(288, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(113, 17)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Search drawing"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(115, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Character code"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'F_ZSEARCH_YOUSO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(561, 691)
        Me.ControlBox = False
        Me.Controls.Add(Me.ImgThumbnail1)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(438, 21)
        Me.Name = "F_ZSEARCH_YOUSO"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Drawing elements Search"
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class