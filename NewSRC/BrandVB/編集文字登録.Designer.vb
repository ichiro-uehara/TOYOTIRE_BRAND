<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_HMSAVE
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
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_haiti_pic As System.Windows.Forms.TextBox
	Public WithEvents w_no As System.Windows.Forms.TextBox
	Public WithEvents w_font_name As System.Windows.Forms.TextBox
	Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
	Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
	Public CommonDialog1Font As System.Windows.Forms.FontDialog
	Public CommonDialog1Color As System.Windows.Forms.ColorDialog
	Public CommonDialog1Print As System.Windows.Forms.PrintDialog
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents w_spell As System.Windows.Forms.TextBox
	Public WithEvents w_entry_date As System.Windows.Forms.TextBox
	Public WithEvents w_entry_name As System.Windows.Forms.TextBox
	Public WithEvents w_dep_name As System.Windows.Forms.TextBox
	Public WithEvents w_comment As System.Windows.Forms.TextBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Frame6 As System.Windows.Forms.GroupBox
	Public WithEvents w_ang As System.Windows.Forms.TextBox
	Public WithEvents w_high As System.Windows.Forms.TextBox
	Public WithEvents w_width As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label15 As System.Windows.Forms.Label
	Public WithEvents Frame7 As System.Windows.Forms.GroupBox
    'Public WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents w_gm_num As System.Windows.Forms.TextBox
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
	'Windows フォーム デザイナを使って変更できます。
	'コード エディタを使用して、変更しないでください。
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ImgThumbnail1 = New System.Windows.Forms.PictureBox()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.Text1 = New System.Windows.Forms.TextBox()
        Me.Command4 = New System.Windows.Forms.Button()
        Me.Command3 = New System.Windows.Forms.Button()
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.w_haiti_pic = New System.Windows.Forms.TextBox()
        Me.w_no = New System.Windows.Forms.TextBox()
        Me.w_font_name = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.Frame6 = New System.Windows.Forms.GroupBox()
        Me.w_spell = New System.Windows.Forms.TextBox()
        Me.w_entry_date = New System.Windows.Forms.TextBox()
        Me.w_entry_name = New System.Windows.Forms.TextBox()
        Me.w_dep_name = New System.Windows.Forms.TextBox()
        Me.w_comment = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Frame7 = New System.Windows.Forms.GroupBox()
        Me.w_ang = New System.Windows.Forms.TextBox()
        Me.w_high = New System.Windows.Forms.TextBox()
        Me.w_width = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.DataGridViewList = New System.Windows.Forms.DataGridView()
        Me.w_gm_num = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.Frame4.SuspendLayout()
        CType(Me.DataGridViewList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImgThumbnail1
        '
        Me.ImgThumbnail1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImgThumbnail1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ImgThumbnail1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ImgThumbnail1.Location = New System.Drawing.Point(88, 531)
        Me.ImgThumbnail1.Name = "ImgThumbnail1"
        Me.ImgThumbnail1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ImgThumbnail1.Size = New System.Drawing.Size(457, 185)
        Me.ImgThumbnail1.TabIndex = 35
        Me.ImgThumbnail1.TabStop = False
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
        Me.Frame1.Location = New System.Drawing.Point(0, -8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(633, 57)
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
        Me.Text1.Location = New System.Drawing.Point(504, 21)
        Me.Text1.MaxLength = 0
        Me.Text1.Name = "Text1"
        Me.Text1.ReadOnly = True
        Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text1.Size = New System.Drawing.Size(89, 19)
        Me.Text1.TabIndex = 15
        Me.Text1.TabStop = False
        '
        'Command4
        '
        Me.Command4.BackColor = System.Drawing.SystemColors.Control
        Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Command4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Command4.Location = New System.Drawing.Point(376, 16)
        Me.Command4.Name = "Command4"
        Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command4.Size = New System.Drawing.Size(74, 33)
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
        Me.Command3.Location = New System.Drawing.Point(288, 16)
        Me.Command3.Name = "Command3"
        Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command3.Size = New System.Drawing.Size(74, 33)
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
        Me.Command2.Location = New System.Drawing.Point(125, 16)
        Me.Command2.Name = "Command2"
        Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command2.Size = New System.Drawing.Size(74, 33)
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
        Me.Command1.Location = New System.Drawing.Point(16, 16)
        Me.Command1.Name = "Command1"
        Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Command1.Size = New System.Drawing.Size(103, 33)
        Me.Command1.TabIndex = 8
        Me.Command1.Text = "registration"
        Me.Command1.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.w_haiti_pic)
        Me.Frame3.Controls.Add(Me.w_no)
        Me.Frame3.Controls.Add(Me.w_font_name)
        Me.Frame3.Controls.Add(Me.Label11)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(0, 40)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(633, 73)
        Me.Frame3.TabIndex = 17
        Me.Frame3.TabStop = False
        '
        'w_haiti_pic
        '
        Me.w_haiti_pic.AcceptsReturn = True
        Me.w_haiti_pic.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_haiti_pic.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_haiti_pic.Enabled = False
        Me.w_haiti_pic.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_haiti_pic.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_haiti_pic.Location = New System.Drawing.Point(416, 40)
        Me.w_haiti_pic.MaxLength = 0
        Me.w_haiti_pic.Name = "w_haiti_pic"
        Me.w_haiti_pic.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_haiti_pic.Size = New System.Drawing.Size(153, 21)
        Me.w_haiti_pic.TabIndex = 13
        Me.w_haiti_pic.Text = "w_haiti_pic"
        '
        'w_no
        '
        Me.w_no.AcceptsReturn = True
        Me.w_no.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_no.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no.Enabled = False
        Me.w_no.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no.Location = New System.Drawing.Point(256, 40)
        Me.w_no.MaxLength = 0
        Me.w_no.Name = "w_no"
        Me.w_no.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no.Size = New System.Drawing.Size(49, 21)
        Me.w_no.TabIndex = 12
        Me.w_no.Text = "w_no"
        '
        'w_font_name
        '
        Me.w_font_name.AcceptsReturn = True
        Me.w_font_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_font_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_font_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_font_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_font_name.Location = New System.Drawing.Point(24, 40)
        Me.w_font_name.MaxLength = 0
        Me.w_font_name.Name = "w_font_name"
        Me.w_font_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_font_name.Size = New System.Drawing.Size(73, 21)
        Me.w_font_name.TabIndex = 1
        Me.w_font_name.Text = "w_font_name"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(416, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(121, 25)
        Me.Label11.TabIndex = 33
        Me.Label11.Text = "Picture number"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(240, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(122, 25)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Category number"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 25)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Font name"
        '
        'Frame6
        '
        Me.Frame6.BackColor = System.Drawing.SystemColors.Control
        Me.Frame6.Controls.Add(Me.w_spell)
        Me.Frame6.Controls.Add(Me.w_entry_date)
        Me.Frame6.Controls.Add(Me.w_entry_name)
        Me.Frame6.Controls.Add(Me.w_dep_name)
        Me.Frame6.Controls.Add(Me.w_comment)
        Me.Frame6.Controls.Add(Me.Label6)
        Me.Frame6.Controls.Add(Me.Label2)
        Me.Frame6.Controls.Add(Me.Label9)
        Me.Frame6.Controls.Add(Me.Label8)
        Me.Frame6.Controls.Add(Me.Label7)
        Me.Frame6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame6.Location = New System.Drawing.Point(0, 104)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame6.Size = New System.Drawing.Size(633, 113)
        Me.Frame6.TabIndex = 20
        Me.Frame6.TabStop = False
        '
        'w_spell
        '
        Me.w_spell.AcceptsReturn = True
        Me.w_spell.BackColor = System.Drawing.SystemColors.Window
        Me.w_spell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_spell.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_spell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_spell.Location = New System.Drawing.Point(104, 16)
        Me.w_spell.MaxLength = 0
        Me.w_spell.Name = "w_spell"
        Me.w_spell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_spell.Size = New System.Drawing.Size(489, 21)
        Me.w_spell.TabIndex = 2
        Me.w_spell.Text = "w_spell"
        '
        'w_entry_date
        '
        Me.w_entry_date.AcceptsReturn = True
        Me.w_entry_date.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_entry_date.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_date.Enabled = False
        Me.w_entry_date.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_date.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_date.Location = New System.Drawing.Point(504, 80)
        Me.w_entry_date.MaxLength = 0
        Me.w_entry_date.Name = "w_entry_date"
        Me.w_entry_date.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_date.Size = New System.Drawing.Size(89, 21)
        Me.w_entry_date.TabIndex = 16
        Me.w_entry_date.Text = "w_entry_date"
        '
        'w_entry_name
        '
        Me.w_entry_name.AcceptsReturn = True
        Me.w_entry_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_entry_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_name.Location = New System.Drawing.Point(312, 80)
        Me.w_entry_name.MaxLength = 0
        Me.w_entry_name.Name = "w_entry_name"
        Me.w_entry_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_name.Size = New System.Drawing.Size(91, 21)
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
        Me.w_dep_name.Location = New System.Drawing.Point(104, 80)
        Me.w_dep_name.MaxLength = 0
        Me.w_dep_name.Name = "w_dep_name"
        Me.w_dep_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_dep_name.Size = New System.Drawing.Size(82, 21)
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
        Me.w_comment.Location = New System.Drawing.Point(104, 48)
        Me.w_comment.MaxLength = 0
        Me.w_comment.Name = "w_comment"
        Me.w_comment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_comment.Size = New System.Drawing.Size(489, 21)
        Me.w_comment.TabIndex = 3
        Me.w_comment.Text = "w_comment"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(16, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(81, 25)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Comment"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(81, 25)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Spell"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(401, 83)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(97, 25)
        Me.Label9.TabIndex = 24
        Me.Label9.Text = "Record date"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(216, 83)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(89, 25)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Registrant"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 83)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(81, 25)
        Me.Label7.TabIndex = 22
        Me.Label7.Text = "Unit"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame7
        '
        Me.Frame7.BackColor = System.Drawing.SystemColors.Control
        Me.Frame7.Controls.Add(Me.w_ang)
        Me.Frame7.Controls.Add(Me.w_high)
        Me.Frame7.Controls.Add(Me.w_width)
        Me.Frame7.Controls.Add(Me.Label5)
        Me.Frame7.Controls.Add(Me.Label4)
        Me.Frame7.Controls.Add(Me.Label15)
        Me.Frame7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame7.Location = New System.Drawing.Point(0, 208)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame7.Size = New System.Drawing.Size(633, 65)
        Me.Frame7.TabIndex = 25
        Me.Frame7.TabStop = False
        '
        'w_ang
        '
        Me.w_ang.AcceptsReturn = True
        Me.w_ang.BackColor = System.Drawing.SystemColors.Window
        Me.w_ang.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_ang.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_ang.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_ang.Location = New System.Drawing.Point(416, 32)
        Me.w_ang.MaxLength = 0
        Me.w_ang.Name = "w_ang"
        Me.w_ang.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_ang.Size = New System.Drawing.Size(177, 21)
        Me.w_ang.TabIndex = 7
        Me.w_ang.Text = "w_ang"
        '
        'w_high
        '
        Me.w_high.AcceptsReturn = True
        Me.w_high.BackColor = System.Drawing.SystemColors.Window
        Me.w_high.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_high.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_high.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_high.Location = New System.Drawing.Point(216, 32)
        Me.w_high.MaxLength = 0
        Me.w_high.Name = "w_high"
        Me.w_high.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_high.Size = New System.Drawing.Size(177, 21)
        Me.w_high.TabIndex = 6
        Me.w_high.Text = "w_high"
        '
        'w_width
        '
        Me.w_width.AcceptsReturn = True
        Me.w_width.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_width.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_width.Enabled = False
        Me.w_width.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_width.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_width.Location = New System.Drawing.Point(16, 32)
        Me.w_width.MaxLength = 0
        Me.w_width.Name = "w_width"
        Me.w_width.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_width.Size = New System.Drawing.Size(177, 21)
        Me.w_width.TabIndex = 14
        Me.w_width.Text = "w_width"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(456, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(81, 25)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Base angle"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(240, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(81, 25)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Base height"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(48, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(81, 25)
        Me.Label15.TabIndex = 26
        Me.Label15.Text = "Width"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.DataGridViewList)
        Me.Frame4.Controls.Add(Me.w_gm_num)
        Me.Frame4.Controls.Add(Me.Label10)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 264)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(633, 258)
        Me.Frame4.TabIndex = 28
        Me.Frame4.TabStop = False
        '
        'DataGridViewList
        '
        Me.DataGridViewList.AllowUserToAddRows = False
        Me.DataGridViewList.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridViewList.ColumnHeadersHeight = 20
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewList.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewList.Location = New System.Drawing.Point(6, 47)
        Me.DataGridViewList.Name = "DataGridViewList"
        Me.DataGridViewList.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("ＭＳ ゴシック", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewList.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewList.RowHeadersWidth = 32
        Me.DataGridViewList.Size = New System.Drawing.Size(620, 205)
        Me.DataGridViewList.TabIndex = 33
        Me.DataGridViewList.TabStop = False
        '
        'w_gm_num
        '
        Me.w_gm_num.AcceptsReturn = True
        Me.w_gm_num.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_gm_num.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.w_gm_num.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_gm_num.Enabled = False
        Me.w_gm_num.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 13.5!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_gm_num.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_gm_num.Location = New System.Drawing.Point(199, 16)
        Me.w_gm_num.MaxLength = 0
        Me.w_gm_num.Name = "w_gm_num"
        Me.w_gm_num.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_gm_num.Size = New System.Drawing.Size(177, 18)
        Me.w_gm_num.TabIndex = 31
        Me.w_gm_num.Text = "w_gm_num"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(8, 19)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(185, 25)
        Me.Label10.TabIndex = 32
        Me.Label10.Text = "Number of primitive character"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'F_HMSAVE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(632, 741)
        Me.ControlBox = False
        Me.Controls.Add(Me.ImgThumbnail1)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame7)
        Me.Controls.Add(Me.Frame4)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(69, 88)
        Me.Name = "F_HMSAVE"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Editing characters registration"
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        Me.Frame6.ResumeLayout(False)
        Me.Frame6.PerformLayout()
        Me.Frame7.ResumeLayout(False)
        Me.Frame7.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        Me.Frame4.PerformLayout()
        CType(Me.DataGridViewList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridViewList As DataGridView
#End Region
End Class