<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_TMP_PSI2
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
	Public WithEvents w_font As System.Windows.Forms.ComboBox
	Public WithEvents w_psi1 As System.Windows.Forms.TextBox
	Public WithEvents w_psi2 As System.Windows.Forms.TextBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.w_psi2c = New System.Windows.Forms.ComboBox()
        Me.w_font = New System.Windows.Forms.ComboBox()
        Me.w_psi1 = New System.Windows.Forms.TextBox()
        Me.w_psi2 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Frame2.SuspendLayout()
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.w_psi2c)
        Me.Frame2.Controls.Add(Me.w_font)
        Me.Frame2.Controls.Add(Me.w_psi1)
        Me.Frame2.Controls.Add(Me.w_psi2)
        Me.Frame2.Controls.Add(Me.Label7)
        Me.Frame2.Controls.Add(Me.Label1)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(0, 72)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(537, 81)
        Me.Frame2.TabIndex = 8
        Me.Frame2.TabStop = False
        '
        'w_psi2c
        '
        Me.w_psi2c.AutoCompleteCustomSource.AddRange(New String() {"kPa", "PSI"})
        Me.w_psi2c.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.w_psi2c.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_psi2c.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.w_psi2c.Items.AddRange(New Object() {"kPa", "PSI"})
        Me.w_psi2c.Location = New System.Drawing.Point(429, 32)
        Me.w_psi2c.Name = "w_psi2c"
        Me.w_psi2c.Size = New System.Drawing.Size(62, 21)
        Me.w_psi2c.TabIndex = 0
        '
        'w_font
        '
        Me.w_font.BackColor = System.Drawing.SystemColors.Window
        Me.w_font.Cursor = System.Windows.Forms.Cursors.Default
        Me.w_font.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_font.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_font.Location = New System.Drawing.Point(88, 32)
        Me.w_font.Name = "w_font"
        Me.w_font.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_font.Size = New System.Drawing.Size(129, 22)
        Me.w_font.TabIndex = 1
        Me.w_font.Text = "w_font"
        '
        'w_psi1
        '
        Me.w_psi1.AcceptsReturn = True
        Me.w_psi1.BackColor = System.Drawing.SystemColors.Window
        Me.w_psi1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_psi1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_psi1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_psi1.Location = New System.Drawing.Point(358, 32)
        Me.w_psi1.MaxLength = 3
        Me.w_psi1.Name = "w_psi1"
        Me.w_psi1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_psi1.Size = New System.Drawing.Size(65, 21)
        Me.w_psi1.TabIndex = 2
        Me.w_psi1.Text = "w_psi1"
        '
        'w_psi2
        '
        Me.w_psi2.AcceptsReturn = True
        Me.w_psi2.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_psi2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_psi2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_psi2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_psi2.Location = New System.Drawing.Point(427, 0)
        Me.w_psi2.MaxLength = 3
        Me.w_psi2.Name = "w_psi2"
        Me.w_psi2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_psi2.Size = New System.Drawing.Size(49, 21)
        Me.w_psi2.TabIndex = 7
        Me.w_psi2.TabStop = False
        Me.w_psi2.Text = "PSI"
        Me.w_psi2.Visible = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(17, 35)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(65, 25)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Font"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(223, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(129, 25)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Durable air pressure"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Command1)
        Me.Frame1.Controls.Add(Me.Command4)
        Me.Frame1.Controls.Add(Me.Command3)
        Me.Frame1.Controls.Add(Me.Command2)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(537, 81)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'Command1
        '
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command1.Location = New System.Drawing.Point(24, 24)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(146, 37)
        Me.Command1.TabIndex = 3
        Me.Command1.Text = "Substitution read"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(440, 24)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(74, 37)
        Me.Command4.TabIndex = 6
        Me.Command4.Text = "Ｈｅｌｐ"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(344, 24)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(74, 37)
        Me.Command3.TabIndex = 5
        Me.Command3.Text = "End"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(248, 24)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(74, 37)
        Me.Command2.TabIndex = 4
        Me.Command2.Text = "Clear"
        Me.Command2.UseVisualStyleBackColor = False
        '
        'F_TMP_PSI2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(540, 157)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(247, 310)
        Me.Name = "F_TMP_PSI2"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Template 2 (kPa or PSI)"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents w_psi2c As System.Windows.Forms.ComboBox
#End Region 
End Class