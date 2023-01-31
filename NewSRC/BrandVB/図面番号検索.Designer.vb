<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_ZSEARCH_BANGO
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
	Public WithEvents cmd_Cancel As System.Windows.Forms.Button
	Public WithEvents cmd_ZumenRead As System.Windows.Forms.Button
	Public WithEvents cmd_Help As System.Windows.Forms.Button
	Public WithEvents cmd_End As System.Windows.Forms.Button
	Public WithEvents cmd_Clear As System.Windows.Forms.Button
	Public WithEvents cmd_Search As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_no2 As System.Windows.Forms.TextBox
	Public WithEvents w_no1 As System.Windows.Forms.TextBox
	Public WithEvents w_id As System.Windows.Forms.TextBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents w_show_no As System.Windows.Forms.TextBox
	Public WithEvents w_total As System.Windows.Forms.TextBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    'Public WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_ZSEARCH_BANGO))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cmd_Cancel = New System.Windows.Forms.Button()
        Me.cmd_ZumenRead = New System.Windows.Forms.Button()
        Me.cmd_Help = New System.Windows.Forms.Button()
        Me.cmd_End = New System.Windows.Forms.Button()
        Me.cmd_Clear = New System.Windows.Forms.Button()
        Me.cmd_Search = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.w_no2 = New System.Windows.Forms.TextBox()
        Me.w_no1 = New System.Windows.Forms.TextBox()
        Me.w_id = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.w_show_no = New System.Windows.Forms.TextBox()
        Me.w_total = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        'Me.MSFlexGrid1 = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame4.SuspendLayout()
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.cmd_Cancel.Location = New System.Drawing.Point(300, 16)
        Me.cmd_Cancel.Name = "cmd_Cancel"
        Me.cmd_Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Cancel.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Cancel.TabIndex = 6
        Me.cmd_Cancel.Text = "Cancel"
        Me.cmd_Cancel.UseVisualStyleBackColor = False
        '
        'cmd_ZumenRead
        '
        Me.cmd_ZumenRead.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_ZumenRead.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_ZumenRead.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_ZumenRead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_ZumenRead.Location = New System.Drawing.Point(88, 16)
        Me.cmd_ZumenRead.Name = "cmd_ZumenRead"
        Me.cmd_ZumenRead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_ZumenRead.Size = New System.Drawing.Size(126, 33)
        Me.cmd_ZumenRead.TabIndex = 4
        Me.cmd_ZumenRead.Text = "Drawing reading"
        Me.cmd_ZumenRead.UseVisualStyleBackColor = False
        '
        'cmd_Help
        '
        Me.cmd_Help.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Help.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Help.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Help.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Help.Location = New System.Drawing.Point(460, 16)
        Me.cmd_Help.Name = "cmd_Help"
        Me.cmd_Help.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Help.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Help.TabIndex = 9
        Me.cmd_Help.Text = "Ｈｅｌｐ"
        Me.cmd_Help.UseVisualStyleBackColor = False
        '
        'cmd_End
        '
        Me.cmd_End.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_End.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_End.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_End.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_End.Location = New System.Drawing.Point(380, 16)
        Me.cmd_End.Name = "cmd_End"
        Me.cmd_End.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_End.Size = New System.Drawing.Size(74, 33)
        Me.cmd_End.TabIndex = 7
        Me.cmd_End.Text = "End"
        Me.cmd_End.UseVisualStyleBackColor = False
        '
        'cmd_Clear
        '
        Me.cmd_Clear.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Clear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Clear.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Clear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Clear.Location = New System.Drawing.Point(220, 16)
        Me.cmd_Clear.Name = "cmd_Clear"
        Me.cmd_Clear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Clear.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Clear.TabIndex = 5
        Me.cmd_Clear.Text = "Clear"
        Me.cmd_Clear.UseVisualStyleBackColor = False
        '
        'cmd_Search
        '
        Me.cmd_Search.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Search.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Search.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Search.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Search.Location = New System.Drawing.Point(8, 16)
        Me.cmd_Search.Name = "cmd_Search"
        Me.cmd_Search.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Search.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Search.TabIndex = 3
        Me.cmd_Search.Text = "Search"
        Me.cmd_Search.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.w_no2)
        Me.Frame3.Controls.Add(Me.w_no1)
        Me.Frame3.Controls.Add(Me.w_id)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(0, 48)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(561, 105)
        Me.Frame3.TabIndex = 10
        Me.Frame3.TabStop = False
        '
        'w_no2
        '
        Me.w_no2.AcceptsReturn = True
        Me.w_no2.BackColor = System.Drawing.SystemColors.Window
        Me.w_no2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no2.Location = New System.Drawing.Point(264, 56)
        Me.w_no2.MaxLength = 2
        Me.w_no2.Name = "w_no2"
        Me.w_no2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no2.Size = New System.Drawing.Size(49, 21)
        Me.w_no2.TabIndex = 2
        Me.w_no2.Text = "w_no2"
        '
        'w_no1
        '
        Me.w_no1.AcceptsReturn = True
        Me.w_no1.BackColor = System.Drawing.SystemColors.Window
        Me.w_no1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no1.Location = New System.Drawing.Point(200, 56)
        Me.w_no1.MaxLength = 5
        Me.w_no1.Name = "w_no1"
        Me.w_no1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no1.Size = New System.Drawing.Size(49, 21)
        Me.w_no1.TabIndex = 1
        Me.w_no1.Text = "w_no1"
        '
        'w_id
        '
        Me.w_id.AcceptsReturn = True
        Me.w_id.BackColor = System.Drawing.SystemColors.Window
        Me.w_id.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_id.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_id.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_id.Location = New System.Drawing.Point(136, 56)
        Me.w_id.MaxLength = 0
        Me.w_id.Name = "w_id"
        Me.w_id.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_id.Size = New System.Drawing.Size(49, 21)
        Me.w_id.TabIndex = 0
        Me.w_id.Text = "w_id"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(264, 32)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(110, 25)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Revision number"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(200, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(58, 25)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Number"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(136, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 25)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Symbol"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(46, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(66, 20)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Drawing"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.w_show_no)
        Me.Frame2.Controls.Add(Me.w_total)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(0, 144)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(561, 57)
        Me.Frame2.TabIndex = 12
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
        Me.w_show_no.Location = New System.Drawing.Point(407, 21)
        Me.w_show_no.MaxLength = 0
        Me.w_show_no.Name = "w_show_no"
        Me.w_show_no.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_show_no.Size = New System.Drawing.Size(81, 21)
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
        Me.w_total.Location = New System.Drawing.Point(237, 21)
        Me.w_total.MaxLength = 0
        Me.w_total.Name = "w_total"
        Me.w_total.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_total.Size = New System.Drawing.Size(73, 21)
        Me.w_total.TabIndex = 13
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
        Me.Label6.Size = New System.Drawing.Size(207, 17)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Conditions corresponding number"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        'Me.Frame4.Controls.Add(Me.MSFlexGrid1)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 192)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(561, 305)
        Me.Frame4.TabIndex = 15
        Me.Frame4.TabStop = False
        '
        'MSFlexGrid1
        '
        'Me.MSFlexGrid1.Location = New System.Drawing.Point(8, 24)
        'Me.MSFlexGrid1.Name = "MSFlexGrid1"
        'Me.MSFlexGrid1.OcxState = CType(resources.GetObject("MSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        'Me.MSFlexGrid1.Size = New System.Drawing.Size(545, 273)
        'Me.MSFlexGrid1.TabIndex = 20
        '
        'F_ZSEARCH_BANGO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(562, 500)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(316, 250)
        Me.Name = "F_ZSEARCH_BANGO"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Drawing number search"
        Me.Frame1.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class