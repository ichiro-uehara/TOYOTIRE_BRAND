Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_PLATE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_PLATE()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click

        InitFlag = False '20100628追加コード
		form_no.Close()
		End
		
		'form1.Show
		
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        On Error Resume Next
        Err.Clear()
        Dim oCommonDialog As Object
        oCommonDialog = CreateObject("MSComDlg.CommonDialog")

        If Err.Number = 0 Then
            With oCommonDialog
                .HelpCommand = cdlHelpContext
                .HelpFile = "c:\VBhelp\BRAND.HLP"
                .HelpContext = 808
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		Dim w_ret As Object
        Dim ZumenName As String
        Dim result As Integer
        Dim pic_no As Integer
		
		Dim w_mess As String

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


		'/* 入力チェック */
		If check_F_TMP_PLATE <> 0 Then
			Exit Sub
		End If
		
		'Brand Ver.3 変更
		'    DBTableName = DBName & "..hm_kanri"
		DBTableName = DBName & "..hm_kanri1"
		
		init_sql()
		form_no.Enabled = False
        F_MSG.Show()
		
		If FreePicNum < 1 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ピクチャ数が足りません" & Chr(13) & "空きピクチャ数 =" & FreePicNum)
            ErrMsg = "The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum
            ErrTtl = "Plate reading"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		
        pic_no = -1


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "SELECT haiti_pic")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE (")
        'result = SqlCmd(SqlConn, " font_name = '" & Mid(w_hm_name.Text, 6) & "' AND")
        'result = SqlCmd(SqlConn, " no = '" & Mid(w_hm_name.Text, 7, 2) & "')")
        '
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        pic_no = SqlData(SqlConn, 1)
        '    Loop
        'Else
        '    MsgBox("SQLｴﾗｰ", 64, "SQLｴﾗｰ")
        '    GoTo error_section
        'End If
		

        '検索コマンド作成
        sqlcmd = "SELECT haiti_pic"
        sqlcmd = sqlcmd & " FROM " & DBTableName
        sqlcmd = sqlcmd & " WHERE ("
        sqlcmd = sqlcmd & " font_name = '" & Mid(w_hm_name.Text, 6) & "' AND"
        sqlcmd = sqlcmd & " no = '" & Mid(w_hm_name.Text, 7, 2) & "')"

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        Do Until Rs.EOF
            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                pic_no = Val(Rs.rdoColumns(0).Value)
            Else
                pic_no = -1
            End If

            Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        pic_no = what_pic_from_hmcode(form_no.w_hm_name.Text)
        If pic_no < 1 Then GoTo error_section
        ZumenName = "HM-" & VB.Left(Trim(form_no.w_hm_name.Text), 6)

		'----- .NET 移行 -----
		'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
		w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

		w_ret = PokeACAD("ACADREAD", w_mess)


        ' -> watanabe add VerUP(2011)
        CommunicateMode = comTmpWait
        ' <- watanabe add VerUP(2011)

        w_ret = RequestACAD("TMPREAD")
		
		
		'   time_start = Now
		'   Do
		'     time_now = Now
		'     If Trim(form_main.text2.Text) = "" Then
		'      If time_now - time_start > timeOutSecond Then
		'        MsgBox "タイムアウトエラー", 64, "ERROR"
		'        w_ret = PokeACAD("ERROR", "TIMEOUT" & timeOutSecond & "秒が経過しました。")
		'        w_ret = RequestACAD("ERROR")
		'        GoTo LOOP_EXIT
		'      End If
		'     ElseIf Left$(Trim(form_main.text2.Text), 7) = "OK-DATA" Then
		'      form_main.text2.Text = ""
		'      GoTo LOOP_EXIT
		'     ElseIf Left(Trim(form_main.text2.Text), 5) = "ERROR" Then
		'      error_no = Mid$(Trim(form_main.text2.Text), 6, 3)
		'      MsgBox "エラーが返りました   ERROR_NO=" & error_no, 64, "ACAD戻り値ｴﾗｰ"
		'      form_main.text2.Text = ""
		'      GoTo LOOP_EXIT
		'     Else
		'      MsgBox "ﾘﾀｰﾝｺｰﾄﾞが不正です" & Chr(13) & Trim(form_main.text2.Text), 64, "ACAD戻り値ｴﾗｰ"
		'      form_main.text2.Text = ""
		'      GoTo LOOP_EXIT
		'     End If
		'   Loop
		
		' -> watanabe add 2007.03
		Dim hexdata As String
		Dim www As New VB6.FixedLengthString(16)
		
        hexdata = ""
        www.Value = "                "
        w_ret = DbltoHex(Val(Trim(form_no.w_plate_w.Text)), www.Value)
		hexdata = hexdata & www.Value
		
        www.Value = "                "
        w_ret = DbltoHex(Val(Trim(form_no.w_plate_h.Text)), www.Value)
		hexdata = hexdata & www.Value
		
        www.Value = "                "
        w_ret = DbltoHex(Val(Trim(form_no.w_plate_r.Text)), www.Value)
		hexdata = hexdata & www.Value
		
        www.Value = "                "
        w_ret = DbltoHex(Val(Trim(form_no.w_plate_n.Text)), www.Value)
		hexdata = hexdata & www.Value
		
        w_ret = PokeACAD("PLATEDATA", hexdata)


        ' -> watanabe add VerUP(2011)
        CommunicateMode = comNone
        ' <- watanabe add VerUP(2011)

        w_ret = RequestACAD("PLATEEDIT")
		' <- watanabe add 2007.03
		
LOOP_EXIT: 
		end_sql()
		Exit Sub
		
error_section: 
        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)
        ' <- watanabe add VerUP(2011)

        On Error Resume Next

        ' -> watanabe add VerUP(2011)
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        F_MSG.Close()
		form_no.Enabled = True
    End Sub
	
	Private Sub F_TMP_PLATE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_ret As Object
		Dim ret As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_w_str As String
		Dim i As Short
		
		form_no = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		'タイプ
		'(Brand Ver.3 変更)
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Plate_ini)
        ret = set_read5(w_w_str, "plate", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "PLATE1" Then
				'             form_no.w_type.AddItem "ネジ無し"
				'          ElseIf Tmp_hm_word(i) = "PLATE2" Then
				'             form_no.w_type.AddItem "前ネジ"
				'          ElseIf Tmp_hm_word(i) = "PLATE3" Then
				'             form_no.w_type.AddItem "後ネジ"
				'          ElseIf Tmp_hm_word(i) = "PLATE4" Then
				'             form_no.w_type.AddItem "前後ネジ"
				'          End If
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_PLATE()
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")

        InitFlag = True '20100628追加コード
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
        Dim w_file As String
        Dim TiffFile As String
        Dim w_text As String

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		On Error Resume Next
		Err.Clear()
        w_text = w_hm_name.Text
		
		If Trim(w_hm_name.Text) = "" Then Exit Sub
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'       TiffFile = TIFFDir & w_hm_name.Text & ".tif"
		'
		'       'Tiffﾌｧｲﾙ表示
		'       w_file = Dir(TiffFile)
		'       If w_file <> "" Then
		'           form_no.ImgThumbnail1.Image = TiffFile
		'           form_no.ImgThumbnail1.ThumbWidth = 500
		'           form_no.ImgThumbnail1.ThumbHeight = 200
		'       Else
		'           MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
		'           form_no.ImgThumbnail1.Image = ""
		'       End If
        TiffFile = TIFFDir & w_hm_name.Text & ".bmp"
		
		'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
		If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
            form_no.ImgThumbnail1.Height = 193 '200 '20100701コード変更
		Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
		End If
		'Brand Ver.5 TIFF->BMP 変更 end
		
		w_type.Focus()
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
		Dim i As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "ネジ無し" Then
		'       w_str = "PLATE1"
		'   ElseIf w_type.Text = "前ネジ" Then
		'       w_str = "PLATE2"
		'   ElseIf w_type.Text = "後ネジ" Then
		'       w_str = "PLATE3"
		'   ElseIf w_type.Text = "前後ネジ" Then
		'       w_str = "PLATE4"
		'   End If
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = w_type.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                Exit For
            End If
		Next i
		
	End Sub
	
	Private Sub w_type_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_type.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
            form_no.w_hm_name.Text = ""
			'Brand Ver.5 TIFF->BMP 変更 start
			'      form_no.ImgThumbnail1.Image = ""
            form_no.ImgThumbnail1.Image = Nothing
			'Brand Ver.5 TIFF->BMP 変更 end
		End If
		
	End Sub
	
	Private Sub w_type_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_type.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_type, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class