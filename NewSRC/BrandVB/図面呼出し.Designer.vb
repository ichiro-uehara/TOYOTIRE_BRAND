<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_ZMNCALL
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
	Public WithEvents w_taisho As System.Windows.Forms.ComboBox
	Public WithEvents w_id As System.Windows.Forms.TextBox
	Public WithEvents w_no1 As System.Windows.Forms.TextBox
	Public WithEvents w_no2 As System.Windows.Forms.TextBox
	Public WithEvents z_end As System.Windows.Forms.Button
	Public WithEvents z_read As System.Windows.Forms.Button
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.w_taisho = New System.Windows.Forms.ComboBox()
        Me.w_id = New System.Windows.Forms.TextBox()
        Me.w_no1 = New System.Windows.Forms.TextBox()
        Me.w_no2 = New System.Windows.Forms.TextBox()
        Me.z_end = New System.Windows.Forms.Button()
        Me.z_read = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'w_taisho
        '
        Me.w_taisho.BackColor = System.Drawing.SystemColors.Window
        Me.w_taisho.Cursor = System.Windows.Forms.Cursors.Default
        Me.w_taisho.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_taisho.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_taisho.Location = New System.Drawing.Point(24, 99)
        Me.w_taisho.Name = "w_taisho"
        Me.w_taisho.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_taisho.Size = New System.Drawing.Size(137, 22)
        Me.w_taisho.TabIndex = 1
        Me.w_taisho.Text = "w_taisho"
        '
        'w_id
        '
        Me.w_id.AcceptsReturn = True
        Me.w_id.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_id.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_id.Enabled = False
        Me.w_id.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_id.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_id.Location = New System.Drawing.Point(168, 99)
        Me.w_id.MaxLength = 0
        Me.w_id.Name = "w_id"
        Me.w_id.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_id.Size = New System.Drawing.Size(49, 21)
        Me.w_id.TabIndex = 0
        Me.w_id.Text = "w_id"
        '
        'w_no1
        '
        Me.w_no1.AcceptsReturn = True
        Me.w_no1.BackColor = System.Drawing.SystemColors.Window
        Me.w_no1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no1.Location = New System.Drawing.Point(232, 99)
        Me.w_no1.MaxLength = 5
        Me.w_no1.Name = "w_no1"
        Me.w_no1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no1.Size = New System.Drawing.Size(49, 21)
        Me.w_no1.TabIndex = 2
        Me.w_no1.Text = "w_no1"
        '
        'w_no2
        '
        Me.w_no2.AcceptsReturn = True
        Me.w_no2.BackColor = System.Drawing.SystemColors.Window
        Me.w_no2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no2.Location = New System.Drawing.Point(296, 99)
        Me.w_no2.MaxLength = 2
        Me.w_no2.Name = "w_no2"
        Me.w_no2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no2.Size = New System.Drawing.Size(49, 21)
        Me.w_no2.TabIndex = 3
        Me.w_no2.Text = "w_no2"
        '
        'z_end
        '
        Me.z_end.BackColor = System.Drawing.SystemColors.Control
        Me.z_end.Cursor = System.Windows.Forms.Cursors.Default
        Me.z_end.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.z_end.ForeColor = System.Drawing.SystemColors.ControlText
        Me.z_end.Location = New System.Drawing.Point(248, 152)
        Me.z_end.Name = "z_end"
        Me.z_end.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.z_end.Size = New System.Drawing.Size(121, 33)
        Me.z_end.TabIndex = 5
        Me.z_end.Text = "End"
        Me.z_end.UseVisualStyleBackColor = False
        '
        'z_read
        '
        Me.z_read.BackColor = System.Drawing.SystemColors.Control
        Me.z_read.Cursor = System.Windows.Forms.Cursors.Default
        Me.z_read.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.z_read.ForeColor = System.Drawing.SystemColors.ControlText
        Me.z_read.Location = New System.Drawing.Point(48, 152)
        Me.z_read.Name = "z_read"
        Me.z_read.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.z_read.Size = New System.Drawing.Size(137, 33)
        Me.z_read.TabIndex = 4
        Me.z_read.Text = "Drawing reading"
        Me.z_read.UseVisualStyleBackColor = False
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(24, 75)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(113, 17)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Search drawing"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(168, 75)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(49, 25)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Symbol"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(232, 75)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(58, 25)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Number"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(296, 75)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(106, 25)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Revision number"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 13.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(381, 25)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Please specify the name of the drawing to call"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'F_ZMNCALL
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(414, 220)
        Me.ControlBox = False
        Me.Controls.Add(Me.w_taisho)
        Me.Controls.Add(Me.w_id)
        Me.Controls.Add(Me.w_no1)
        Me.Controls.Add(Me.w_no2)
        Me.Controls.Add(Me.z_end)
        Me.Controls.Add(Me.z_read)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(260, 214)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "F_ZMNCALL"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Drawing reading form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region 
End Class