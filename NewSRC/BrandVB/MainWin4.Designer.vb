<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_MAIN4
#Region "Windows フォーム デザイナによって生成されたコード "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'この呼び出しは、Windows フォーム デザイナで必要です。
		InitializeComponent()
	End Sub
	'Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows フォーム デザイナで必要です。
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents SRflag As System.Windows.Forms.TextBox
	Public WithEvents POKE As System.Windows.Forms.Button
	Public WithEvents REQUEST As System.Windows.Forms.Button
	Public WithEvents LINK As System.Windows.Forms.Button
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Vbsql1 As System.Windows.Forms.PictureBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SRflag = New System.Windows.Forms.TextBox()
        Me.POKE = New System.Windows.Forms.Button()
        Me.REQUEST = New System.Windows.Forms.Button()
        Me.LINK = New System.Windows.Forms.Button()
        Me.Text2 = New System.Windows.Forms.TextBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Vbsql1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.Vbsql1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SRflag
        '
        Me.SRflag.AcceptsReturn = True
        Me.SRflag.BackColor = System.Drawing.SystemColors.Window
        Me.SRflag.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.SRflag.ForeColor = System.Drawing.SystemColors.WindowText
        Me.SRflag.Location = New System.Drawing.Point(96, 80)
        Me.SRflag.MaxLength = 0
        Me.SRflag.Name = "SRflag"
        Me.SRflag.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SRflag.Size = New System.Drawing.Size(41, 19)
        Me.SRflag.TabIndex = 6
        '
        'POKE
        '
        Me.POKE.BackColor = System.Drawing.SystemColors.Control
        Me.POKE.Cursor = System.Windows.Forms.Cursors.Default
        Me.POKE.ForeColor = System.Drawing.SystemColors.ControlText
        Me.POKE.Location = New System.Drawing.Point(256, 120)
        Me.POKE.Name = "POKE"
        Me.POKE.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.POKE.Size = New System.Drawing.Size(81, 25)
        Me.POKE.TabIndex = 4
        Me.POKE.Text = "POKE"
        Me.POKE.UseVisualStyleBackColor = False
        '
        'REQUEST
        '
        Me.REQUEST.BackColor = System.Drawing.SystemColors.Control
        Me.REQUEST.Cursor = System.Windows.Forms.Cursors.Default
        Me.REQUEST.ForeColor = System.Drawing.SystemColors.ControlText
        Me.REQUEST.Location = New System.Drawing.Point(136, 120)
        Me.REQUEST.Name = "REQUEST"
        Me.REQUEST.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.REQUEST.Size = New System.Drawing.Size(89, 25)
        Me.REQUEST.TabIndex = 3
        Me.REQUEST.Text = "REQUEST"
        Me.REQUEST.UseVisualStyleBackColor = False
        '
        'LINK
        '
        Me.LINK.BackColor = System.Drawing.SystemColors.Control
        Me.LINK.Cursor = System.Windows.Forms.Cursors.Default
        Me.LINK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.LINK.Location = New System.Drawing.Point(24, 120)
        Me.LINK.Name = "LINK"
        Me.LINK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.LINK.Size = New System.Drawing.Size(89, 25)
        Me.LINK.TabIndex = 2
        Me.LINK.Text = "LINK"
        Me.LINK.UseVisualStyleBackColor = False
        '
        'Text2
        '
        Me.Text2.AcceptsReturn = True
        Me.Text2.BackColor = System.Drawing.SystemColors.Window
        Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text2.Location = New System.Drawing.Point(96, 56)
        Me.Text2.MaxLength = 0
        Me.Text2.Name = "Text2"
        Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text2.Size = New System.Drawing.Size(217, 19)
        Me.Text2.TabIndex = 1
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Text1.ForeColor = System.Drawing.Color.Black
        Me.Text1.Location = New System.Drawing.Point(32, 24)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(281, 14)
        Me.Text1.TabIndex = 0
        Me.Text1.Text = "This is the main window hidden"
        Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Vbsql1
        '
        Me.Vbsql1.BackColor = System.Drawing.SystemColors.Control
        Me.Vbsql1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Vbsql1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Vbsql1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Vbsql1.Location = New System.Drawing.Point(312, 24)
        Me.Vbsql1.Name = "Vbsql1"
        Me.Vbsql1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Vbsql1.Size = New System.Drawing.Size(64, 26)
        Me.Vbsql1.TabIndex = 7
        Me.Vbsql1.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(81, 17)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "RecieveDATA"
        '
        'F_MAIN4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(378, 160)
        Me.Controls.Add(Me.SRflag)
        Me.Controls.Add(Me.POKE)
        Me.Controls.Add(Me.REQUEST)
        Me.Controls.Add(Me.LINK)
        Me.Controls.Add(Me.Text2)
        Me.Controls.Add(Me.Text1)
        Me.Controls.Add(Me.Vbsql1)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(443, 68)
        Me.Name = "F_MAIN4"
        Me.Opacity = 0.0R
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form8"
        CType(Me.Vbsql1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class