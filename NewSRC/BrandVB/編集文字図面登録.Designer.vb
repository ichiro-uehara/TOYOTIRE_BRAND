<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_HZSAVE
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
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_id As System.Windows.Forms.TextBox
	Public WithEvents w_no2 As System.Windows.Forms.TextBox
	Public WithEvents w_no1 As System.Windows.Forms.TextBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents w_entry_date As System.Windows.Forms.TextBox
	Public WithEvents w_entry_name As System.Windows.Forms.TextBox
	Public WithEvents w_dep_name As System.Windows.Forms.TextBox
	Public WithEvents w_comment As System.Windows.Forms.TextBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
    'Public WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents w_hm_num As System.Windows.Forms.TextBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_HZSAVE))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.w_id = New System.Windows.Forms.TextBox()
        Me.w_no2 = New System.Windows.Forms.TextBox()
        Me.w_no1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.w_entry_date = New System.Windows.Forms.TextBox()
        Me.w_entry_name = New System.Windows.Forms.TextBox()
        Me.w_dep_name = New System.Windows.Forms.TextBox()
        Me.w_comment = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        'Me.MSFlexGrid1 = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.w_hm_num = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame4.SuspendLayout()
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Text1)
        Me.Frame1.Controls.Add(Me.Command4)
        Me.Frame1.Controls.Add(Me.Command3)
        Me.Frame1.Controls.Add(Me.Command2)
        Me.Frame1.Controls.Add(Me.Command1)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(601, 73)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'Text1
        '
        Me.Text1.AcceptsReturn = True
        Me.Text1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Text1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Text1.Location = New System.Drawing.Point(424, 32)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(169, 19)
        Me.Text1.TabIndex = 12
        Me.Text1.TabStop = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(326, 24)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(74, 37)
        Me.Command4.TabIndex = 11
        Me.Command4.Text = "Ｈｅｌｐ"
        Me.Command4.UseVisualStyleBackColor = False
        '
        'Command3
        '
        Me.Command3.BackColor = System.Drawing.SystemColors.Control
        Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command3.Location = New System.Drawing.Point(246, 24)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(74, 37)
        Me.Command3.TabIndex = 10
        Me.Command3.Text = "End"
        Me.Command3.UseVisualStyleBackColor = False
        '
        'Command2
        '
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command2.Location = New System.Drawing.Point(134, 25)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(74, 37)
        Me.Command2.TabIndex = 9
        Me.Command2.Text = "Clear"
        Me.Command2.UseVisualStyleBackColor = False
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
        Me.Command1.Size = New System.Drawing.Size(104, 37)
        Me.Command1.TabIndex = 8
        Me.Command1.Text = "registration"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.w_id)
        Me.Frame3.Controls.Add(Me.w_no2)
        Me.Frame3.Controls.Add(Me.w_no1)
        Me.Frame3.Controls.Add(Me.Label4)
        Me.Frame3.Controls.Add(Me.Label2)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(0, 64)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(601, 97)
        Me.Frame3.TabIndex = 13
        Me.Frame3.TabStop = False
        '
        'w_id
        '
        Me.w_id.AcceptsReturn = True
        Me.w_id.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_id.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_id.Enabled = False
        Me.w_id.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_id.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_id.Location = New System.Drawing.Point(104, 48)
        Me.w_id.MaxLength = 0
        Me.w_id.Name = "w_id"
        Me.w_id.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_id.Size = New System.Drawing.Size(49, 21)
        Me.w_id.TabIndex = 25
        Me.w_id.Text = "w_id"
        '
        'w_no2
        '
        Me.w_no2.AcceptsReturn = True
        Me.w_no2.BackColor = System.Drawing.SystemColors.Window
        Me.w_no2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no2.Enabled = False
        Me.w_no2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no2.Location = New System.Drawing.Point(273, 48)
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
        Me.w_no1.Location = New System.Drawing.Point(176, 48)
        Me.w_no1.MaxLength = 4
        Me.w_no1.Name = "w_no1"
        Me.w_no1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no1.Size = New System.Drawing.Size(73, 21)
        Me.w_no1.TabIndex = 1
        Me.w_no1.Text = "w_no1"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(88, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(65, 25)
        Me.Label4.TabIndex = 24
        Me.Label4.Text = "Symbol"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(32, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(65, 25)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Drawing"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(270, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(115, 25)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Revision number"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(173, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 25)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Number"
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.w_entry_date)
        Me.Frame6.Controls.Add(Me.w_entry_name)
        Me.Frame6.Controls.Add(Me.w_dep_name)
        Me.Frame6.Controls.Add(Me.w_comment)
        Me.Frame6.Controls.Add(Me.Label9)
        Me.Frame6.Controls.Add(Me.Label8)
        Me.Frame6.Controls.Add(Me.Label7)
        Me.Frame6.Controls.Add(Me.Label6)
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(0, 152)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(601, 89)
        Me.Frame6.TabIndex = 16
        Me.Frame6.TabStop = False
        '
        'w_entry_date
        '
        Me.w_entry_date.AcceptsReturn = True
        Me.w_entry_date.BackColor = System.Drawing.SystemColors.ControlDark
        Me.w_entry_date.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_date.Enabled = False
        Me.w_entry_date.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_date.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_date.Location = New System.Drawing.Point(488, 48)
        Me.w_entry_date.MaxLength = 0
        Me.w_entry_date.Name = "w_entry_date"
        Me.w_entry_date.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_date.Size = New System.Drawing.Size(89, 21)
        Me.w_entry_date.TabIndex = 6
        Me.w_entry_date.Text = "w_entry_date"
        '
        'w_entry_name
        '
        Me.w_entry_name.AcceptsReturn = True
        Me.w_entry_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_entry_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_name.Location = New System.Drawing.Point(284, 48)
        Me.w_entry_name.MaxLength = 0
        Me.w_entry_name.Name = "w_entry_name"
        Me.w_entry_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_name.Size = New System.Drawing.Size(96, 21)
        Me.w_entry_name.TabIndex = 5
        Me.w_entry_name.Text = "w_entry_name"
        '
        'w_dep_name
        '
        Me.w_dep_name.AcceptsReturn = True
        Me.w_dep_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_dep_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_dep_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_dep_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_dep_name.Location = New System.Drawing.Point(104, 48)
        Me.w_dep_name.MaxLength = 0
        Me.w_dep_name.Name = "w_dep_name"
        Me.w_dep_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_dep_name.Size = New System.Drawing.Size(73, 21)
        Me.w_dep_name.TabIndex = 4
        Me.w_dep_name.Text = "w_dep_name"
        '
        'w_comment
        '
        Me.w_comment.AcceptsReturn = True
        Me.w_comment.BackColor = System.Drawing.SystemColors.Window
        Me.w_comment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_comment.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_comment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_comment.Location = New System.Drawing.Point(104, 16)
        Me.w_comment.MaxLength = 0
        Me.w_comment.Name = "w_comment"
        Me.w_comment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_comment.Size = New System.Drawing.Size(473, 21)
        Me.w_comment.TabIndex = 3
        Me.w_comment.Text = "w_comment"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(385, 51)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(97, 25)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "Record date"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(189, 51)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(89, 25)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Registrant"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 51)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(81, 25)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Unit"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(16, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(81, 25)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Comment"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        'Me.Frame4.Controls.Add(Me.MSFlexGrid1)
        Me.Frame4.Controls.Add(Me.w_hm_num)
        Me.Frame4.Controls.Add(Me.Label10)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 232)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(601, 497)
        Me.Frame4.TabIndex = 21
        Me.Frame4.TabStop = False
        '
        'MSFlexGrid1
        '
        'Me.MSFlexGrid1.Location = New System.Drawing.Point(8, 48)
        'Me.MSFlexGrid1.Name = "MSFlexGrid1"
        'Me.MSFlexGrid1.OcxState = CType(resources.GetObject("MSFlexGrid1.OcxState"), System.Windows.Forms.AxHost.State)
        'Me.MSFlexGrid1.Size = New System.Drawing.Size(577, 441)
        'Me.MSFlexGrid1.TabIndex = 26
        '
        'w_hm_num
        '
        Me.w_hm_num.AcceptsReturn = True
        Me.w_hm_num.BackColor = System.Drawing.SystemColors.ControlDark
        Me.w_hm_num.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_hm_num.Enabled = False
        Me.w_hm_num.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_hm_num.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_hm_num.Location = New System.Drawing.Point(203, 19)
        Me.w_hm_num.MaxLength = 0
        Me.w_hm_num.Name = "w_hm_num"
        Me.w_hm_num.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_hm_num.Size = New System.Drawing.Size(177, 21)
        Me.w_hm_num.TabIndex = 7
        Me.w_hm_num.Text = "w_hm_num"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(8, 22)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(189, 25)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Number of editing characters"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'F_HZSAVE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(601, 734)
        Me.ControlBox = False
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(220, 128)
        Me.Name = "F_HZSAVE"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Editing characters drawing registration"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        'CType(Me.MSFlexGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region 
End Class