<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class F_HMSEARCH
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
	Public WithEvents cmd_ReadClear As System.Windows.Forms.Button
	Public WithEvents cmd_AllRead As System.Windows.Forms.Button
	Public WithEvents w_show_no As System.Windows.Forms.TextBox
	Public WithEvents w_total As System.Windows.Forms.TextBox
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
    'Public WithEvents MSFlexGrid1 As AxMSFlexGridLib.AxMSFlexGrid
    Public WithEvents Frame4 As System.Windows.Forms.GroupBox
	Public WithEvents cmd_Cancel As System.Windows.Forms.Button
	Public WithEvents cmd_CadRead As System.Windows.Forms.Button
	Public WithEvents cmd_Help As System.Windows.Forms.Button
	Public WithEvents cmd_End As System.Windows.Forms.Button
	Public WithEvents cmd_Clear As System.Windows.Forms.Button
	Public WithEvents cmd_Search As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents w_hikaku As System.Windows.Forms.ComboBox
    Public WithEvents w_entry_date_1 As System.Windows.Forms.TextBox
    Public WithEvents w_high As System.Windows.Forms.TextBox
    Public WithEvents w_entry_name As System.Windows.Forms.TextBox
    Public WithEvents w_entry_date_0 As System.Windows.Forms.TextBox
    Public WithEvents w_spell As System.Windows.Forms.TextBox
    Public WithEvents w_no As System.Windows.Forms.TextBox
    Public WithEvents w_font_name As System.Windows.Forms.TextBox
    Public CommonDialog1Open As System.Windows.Forms.OpenFileDialog
    Public CommonDialog1Save As System.Windows.Forms.SaveFileDialog
    Public CommonDialog1Font As System.Windows.Forms.FontDialog
    Public CommonDialog1Color As System.Windows.Forms.ColorDialog
    Public CommonDialog1Print As System.Windows.Forms.PrintDialog
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents Label8 As System.Windows.Forms.Label
    Public WithEvents Label9 As System.Windows.Forms.Label
    Public WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    Public WithEvents w_entry_date As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
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
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.cmd_ReadClear = New System.Windows.Forms.Button()
        Me.cmd_AllRead = New System.Windows.Forms.Button()
        Me.w_show_no = New System.Windows.Forms.TextBox()
        Me.w_total = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Frame4 = New System.Windows.Forms.GroupBox()
        Me.DataGridViewList = New System.Windows.Forms.DataGridView()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.cmd_Cancel = New System.Windows.Forms.Button()
        Me.cmd_CadRead = New System.Windows.Forms.Button()
        Me.cmd_Help = New System.Windows.Forms.Button()
        Me.cmd_End = New System.Windows.Forms.Button()
        Me.cmd_Clear = New System.Windows.Forms.Button()
        Me.cmd_Search = New System.Windows.Forms.Button()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.w_hikaku = New System.Windows.Forms.ComboBox()
        Me.w_entry_date_1 = New System.Windows.Forms.TextBox()
        Me.w_high = New System.Windows.Forms.TextBox()
        Me.w_entry_name = New System.Windows.Forms.TextBox()
        Me.w_entry_date_0 = New System.Windows.Forms.TextBox()
        Me.w_spell = New System.Windows.Forms.TextBox()
        Me.w_no = New System.Windows.Forms.TextBox()
        Me.w_font_name = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CommonDialog1Open = New System.Windows.Forms.OpenFileDialog()
        Me.CommonDialog1Save = New System.Windows.Forms.SaveFileDialog()
        Me.CommonDialog1Font = New System.Windows.Forms.FontDialog()
        Me.CommonDialog1Color = New System.Windows.Forms.ColorDialog()
        Me.CommonDialog1Print = New System.Windows.Forms.PrintDialog()
        Me.w_entry_date = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(Me.components)
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame2.SuspendLayout()
        Me.Frame4.SuspendLayout()
        CType(Me.DataGridViewList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Frame1.SuspendLayout()
        Me.Frame3.SuspendLayout()
        CType(Me.w_entry_date, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImgThumbnail1
        '
        Me.ImgThumbnail1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ImgThumbnail1.Cursor = System.Windows.Forms.Cursors.Default
        Me.ImgThumbnail1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ImgThumbnail1.Location = New System.Drawing.Point(88, 192)
        Me.ImgThumbnail1.Name = "ImgThumbnail1"
        Me.ImgThumbnail1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ImgThumbnail1.Size = New System.Drawing.Size(457, 193)
        Me.ImgThumbnail1.TabIndex = 31
        Me.ImgThumbnail1.TabStop = False
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.SystemColors.Control
        Me.Frame2.Controls.Add(Me.cmd_ReadClear)
        Me.Frame2.Controls.Add(Me.cmd_AllRead)
        Me.Frame2.Controls.Add(Me.w_show_no)
        Me.Frame2.Controls.Add(Me.w_total)
        Me.Frame2.Controls.Add(Me.Label6)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(0, 408)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(625, 57)
        Me.Frame2.TabIndex = 23
        Me.Frame2.TabStop = False
        '
        'cmd_ReadClear
        '
        Me.cmd_ReadClear.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_ReadClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_ReadClear.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_ReadClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_ReadClear.Location = New System.Drawing.Point(417, 16)
        Me.cmd_ReadClear.Name = "cmd_ReadClear"
        Me.cmd_ReadClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_ReadClear.Size = New System.Drawing.Size(97, 33)
        Me.cmd_ReadClear.TabIndex = 11
        Me.cmd_ReadClear.Text = "Read clear"
        Me.cmd_ReadClear.UseVisualStyleBackColor = False
        '
        'cmd_AllRead
        '
        Me.cmd_AllRead.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_AllRead.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_AllRead.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_AllRead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_AllRead.Location = New System.Drawing.Point(314, 16)
        Me.cmd_AllRead.Name = "cmd_AllRead"
        Me.cmd_AllRead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_AllRead.Size = New System.Drawing.Size(97, 33)
        Me.cmd_AllRead.TabIndex = 10
        Me.cmd_AllRead.Text = "Read all"
        Me.cmd_AllRead.UseVisualStyleBackColor = False
        '
        'w_show_no
        '
        Me.w_show_no.AcceptsReturn = True
        Me.w_show_no.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.w_show_no.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_show_no.Enabled = False
        Me.w_show_no.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_show_no.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_show_no.Location = New System.Drawing.Point(551, 22)
        Me.w_show_no.MaxLength = 0
        Me.w_show_no.Name = "w_show_no"
        Me.w_show_no.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_show_no.Size = New System.Drawing.Size(65, 21)
        Me.w_show_no.TabIndex = 29
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
        Me.w_total.Location = New System.Drawing.Point(164, 21)
        Me.w_total.MaxLength = 0
        Me.w_total.Name = "w_total"
        Me.w_total.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_total.Size = New System.Drawing.Size(73, 21)
        Me.w_total.TabIndex = 24
        Me.w_total.Text = "w_total"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(3, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(155, 33)
        Me.Label6.TabIndex = 25
        Me.Label6.Text = "Conditions corresponding number"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Frame4
        '
        Me.Frame4.BackColor = System.Drawing.SystemColors.Control
        Me.Frame4.Controls.Add(Me.DataGridViewList)
        Me.Frame4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame4.Location = New System.Drawing.Point(0, 456)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame4.Size = New System.Drawing.Size(625, 281)
        Me.Frame4.TabIndex = 28
        Me.Frame4.TabStop = False
        '
        'DataGridViewList
        '
        Me.DataGridViewList.AllowUserToAddRows = False
        Me.DataGridViewList.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewList.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridViewList.ColumnHeadersHeight = 30
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridViewList.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewList.Location = New System.Drawing.Point(4, 12)
        Me.DataGridViewList.Name = "DataGridViewList"
        Me.DataGridViewList.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewList.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DataGridViewList.RowHeadersWidth = 32
        Me.DataGridViewList.Size = New System.Drawing.Size(620, 256)
        Me.DataGridViewList.TabIndex = 1
        Me.DataGridViewList.TabStop = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.cmd_Cancel)
        Me.Frame1.Controls.Add(Me.cmd_CadRead)
        Me.Frame1.Controls.Add(Me.cmd_Help)
        Me.Frame1.Controls.Add(Me.cmd_End)
        Me.Frame1.Controls.Add(Me.cmd_Clear)
        Me.Frame1.Controls.Add(Me.cmd_Search)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(0, 0)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(625, 57)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'cmd_Cancel
        '
        Me.cmd_Cancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Cancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Cancel.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Cancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Cancel.Location = New System.Drawing.Point(295, 16)
        Me.cmd_Cancel.Name = "cmd_Cancel"
        Me.cmd_Cancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Cancel.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Cancel.TabIndex = 14
        Me.cmd_Cancel.Text = "Cancel"
        Me.cmd_Cancel.UseVisualStyleBackColor = False
        '
        'cmd_CadRead
        '
        Me.cmd_CadRead.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_CadRead.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_CadRead.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_CadRead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_CadRead.Location = New System.Drawing.Point(96, 16)
        Me.cmd_CadRead.Name = "cmd_CadRead"
        Me.cmd_CadRead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_CadRead.Size = New System.Drawing.Size(113, 33)
        Me.cmd_CadRead.TabIndex = 12
        Me.cmd_CadRead.Text = "CAD reading"
        Me.cmd_CadRead.UseVisualStyleBackColor = False
        '
        'cmd_Help
        '
        Me.cmd_Help.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Help.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Help.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Help.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Help.Location = New System.Drawing.Point(528, 16)
        Me.cmd_Help.Name = "cmd_Help"
        Me.cmd_Help.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Help.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Help.TabIndex = 16
        Me.cmd_Help.Text = "Ｈｅｌｐ"
        Me.cmd_Help.UseVisualStyleBackColor = False
        '
        'cmd_End
        '
        Me.cmd_End.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_End.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_End.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_End.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_End.Location = New System.Drawing.Point(440, 16)
        Me.cmd_End.Name = "cmd_End"
        Me.cmd_End.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_End.Size = New System.Drawing.Size(74, 33)
        Me.cmd_End.TabIndex = 15
        Me.cmd_End.Text = "End"
        Me.cmd_End.UseVisualStyleBackColor = False
        '
        'cmd_Clear
        '
        Me.cmd_Clear.BackColor = System.Drawing.SystemColors.Control
        Me.cmd_Clear.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmd_Clear.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.cmd_Clear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmd_Clear.Location = New System.Drawing.Point(215, 16)
        Me.cmd_Clear.Name = "cmd_Clear"
        Me.cmd_Clear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmd_Clear.Size = New System.Drawing.Size(74, 33)
        Me.cmd_Clear.TabIndex = 13
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
        Me.cmd_Search.TabIndex = 9
        Me.cmd_Search.Text = "Search"
        Me.cmd_Search.UseVisualStyleBackColor = False
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.SystemColors.Control
        Me.Frame3.Controls.Add(Me.w_hikaku)
        Me.Frame3.Controls.Add(Me.w_entry_date_1)
        Me.Frame3.Controls.Add(Me.w_high)
        Me.Frame3.Controls.Add(Me.w_entry_name)
        Me.Frame3.Controls.Add(Me.w_entry_date_0)
        Me.Frame3.Controls.Add(Me.w_spell)
        Me.Frame3.Controls.Add(Me.w_no)
        Me.Frame3.Controls.Add(Me.w_font_name)
        Me.Frame3.Controls.Add(Me.Label10)
        Me.Frame3.Controls.Add(Me.Label7)
        Me.Frame3.Controls.Add(Me.Label8)
        Me.Frame3.Controls.Add(Me.Label9)
        Me.Frame3.Controls.Add(Me.Label5)
        Me.Frame3.Controls.Add(Me.Label3)
        Me.Frame3.Controls.Add(Me.Label1)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(0, 48)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(625, 113)
        Me.Frame3.TabIndex = 17
        Me.Frame3.TabStop = False
        '
        'w_hikaku
        '
        Me.w_hikaku.BackColor = System.Drawing.SystemColors.Window
        Me.w_hikaku.Cursor = System.Windows.Forms.Cursors.Default
        Me.w_hikaku.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_hikaku.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_hikaku.Location = New System.Drawing.Point(384, 16)
        Me.w_hikaku.Name = "w_hikaku"
        Me.w_hikaku.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_hikaku.Size = New System.Drawing.Size(73, 22)
        Me.w_hikaku.TabIndex = 4
        Me.w_hikaku.Text = "w_hikaku"
        '
        'w_entry_date_1
        '
        Me.w_entry_date_1.AcceptsReturn = True
        Me.w_entry_date_1.BackColor = System.Drawing.SystemColors.Window
        Me.w_entry_date_1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_date_1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_date_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_date.SetIndex(Me.w_entry_date_1, CType(1, Short))
        Me.w_entry_date_1.Location = New System.Drawing.Point(512, 80)
        Me.w_entry_date_1.MaxLength = 8
        Me.w_entry_date_1.Name = "w_entry_date_1"
        Me.w_entry_date_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_date_1.Size = New System.Drawing.Size(89, 21)
        Me.w_entry_date_1.TabIndex = 8
        Me.w_entry_date_1.Text = "w_entry_date"
        '
        'w_high
        '
        Me.w_high.AcceptsReturn = True
        Me.w_high.BackColor = System.Drawing.SystemColors.Window
        Me.w_high.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_high.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_high.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_high.Location = New System.Drawing.Point(463, 16)
        Me.w_high.MaxLength = 0
        Me.w_high.Name = "w_high"
        Me.w_high.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_high.Size = New System.Drawing.Size(97, 21)
        Me.w_high.TabIndex = 5
        Me.w_high.Text = "w_high"
        '
        'w_entry_name
        '
        Me.w_entry_name.AcceptsReturn = True
        Me.w_entry_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_entry_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_name.Location = New System.Drawing.Point(384, 48)
        Me.w_entry_name.MaxLength = 0
        Me.w_entry_name.Name = "w_entry_name"
        Me.w_entry_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_name.Size = New System.Drawing.Size(89, 21)
        Me.w_entry_name.TabIndex = 6
        Me.w_entry_name.Text = "w_entry_name"
        '
        'w_entry_date_0
        '
        Me.w_entry_date_0.AcceptsReturn = True
        Me.w_entry_date_0.BackColor = System.Drawing.SystemColors.Window
        Me.w_entry_date_0.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_entry_date_0.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_entry_date_0.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_entry_date.SetIndex(Me.w_entry_date_0, CType(0, Short))
        Me.w_entry_date_0.Location = New System.Drawing.Point(384, 80)
        Me.w_entry_date_0.MaxLength = 8
        Me.w_entry_date_0.Name = "w_entry_date_0"
        Me.w_entry_date_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_entry_date_0.Size = New System.Drawing.Size(89, 21)
        Me.w_entry_date_0.TabIndex = 7
        Me.w_entry_date_0.Text = "w_entry_date"
        '
        'w_spell
        '
        Me.w_spell.AcceptsReturn = True
        Me.w_spell.BackColor = System.Drawing.SystemColors.Window
        Me.w_spell.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_spell.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_spell.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_spell.Location = New System.Drawing.Point(123, 80)
        Me.w_spell.MaxLength = 0
        Me.w_spell.Name = "w_spell"
        Me.w_spell.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_spell.Size = New System.Drawing.Size(169, 21)
        Me.w_spell.TabIndex = 3
        Me.w_spell.Text = "w_spell"
        '
        'w_no
        '
        Me.w_no.AcceptsReturn = True
        Me.w_no.BackColor = System.Drawing.SystemColors.Window
        Me.w_no.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_no.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_no.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_no.Location = New System.Drawing.Point(123, 48)
        Me.w_no.MaxLength = 0
        Me.w_no.Name = "w_no"
        Me.w_no.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_no.Size = New System.Drawing.Size(49, 21)
        Me.w_no.TabIndex = 2
        Me.w_no.Text = "w_no"
        '
        'w_font_name
        '
        Me.w_font_name.AcceptsReturn = True
        Me.w_font_name.BackColor = System.Drawing.SystemColors.Window
        Me.w_font_name.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.w_font_name.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.w_font_name.ForeColor = System.Drawing.SystemColors.WindowText
        Me.w_font_name.Location = New System.Drawing.Point(123, 16)
        Me.w_font_name.MaxLength = 6
        Me.w_font_name.Name = "w_font_name"
        Me.w_font_name.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.w_font_name.Size = New System.Drawing.Size(87, 21)
        Me.w_font_name.TabIndex = 1
        Me.w_font_name.Text = "w_font_name"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(481, 83)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(25, 17)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "〜"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(328, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(49, 17)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Height"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(295, 51)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(82, 17)
        Me.Label8.TabIndex = 22
        Me.Label8.Text = "Registrant"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(298, 83)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(79, 17)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Record date"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(52, 83)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(65, 17)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Spell"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 51)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(111, 17)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Category number"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("ＭＳ Ｐゴシック", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(44, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(73, 17)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Font name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'F_HMSEARCH
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(628, 730)
        Me.ControlBox = False
        Me.Controls.Add(Me.ImgThumbnail1)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Location = New System.Drawing.Point(134, 65)
        Me.Name = "F_HMSEARCH"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Editing characters Search"
        CType(Me.ImgThumbnail1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.Frame4.ResumeLayout(False)
        CType(Me.DataGridViewList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Frame1.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame3.PerformLayout()
        CType(Me.w_entry_date, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridViewList As DataGridView
#End Region
End Class