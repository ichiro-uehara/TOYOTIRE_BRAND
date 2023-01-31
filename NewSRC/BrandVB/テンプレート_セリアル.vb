Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_SERIARU
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
        Dim result As Integer
		Dim w_text As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim tmp_dbtable As Object
        'Dim select_item As Object
        ' <- watanabe del VerUP(2011)

        Dim select_where(5) As Object
		Dim select_num As Object
		Dim w_str As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


		'/* 入力チェック */
		If check_F_TMP_SERIARU <> 0 Then
			Exit Sub
		End If
		
		'// サイズコード検索
        w_text = Trim(form_no.w_size1.Text) & Trim(form_no.w_size2.Text) & Trim(form_no.w_size3.Text) & Trim(form_no.w_size5.Text) & Trim(form_no.w_size6.Text)
		w_str = ""
        w_str = "WHERE syurui = '" & Trim(form_no.w_syurui.Text) & "'"
        select_num = 0
        w_str = w_str & " AND size1 = '" & Trim(form_no.w_size1.Text) & "'"
        w_str = w_str & " AND size2 = '" & Trim(form_no.w_size2.Text) & "'"
        w_str = w_str & " AND size3 = '" & Trim(form_no.w_size3.Text) & "'"
        w_str = w_str & " AND size5 = '" & Trim(form_no.w_size5.Text) & "'" '
        w_str = w_str & " AND size6 = '" & Trim(form_no.w_size6.Text) & "'"
		
		init_sql()


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "SELECT size_code")
        'result = SqlCmd(SqlConn, " FROM " & STANDARD_DBName & "..size_code")
        'result = SqlCmd(SqlConn, " " & w_str)
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '    If SqlNextRow(SqlConn) = REGROW Then
        '        form_no.w_size_code.Text = SqlData(SqlConn, 1)
        '    End If
        'Else
        '    MsgBox("ﾃﾞｰﾀﾍﾞｰｽSELECTｴﾗｰ", MsgBoxStyle.Critical)
        'End If


        '検索コマンド作成
        sqlcmd = "SELECT size_code"
        sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..size_code"
        sqlcmd = sqlcmd & " " & w_str

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If GL_T_RDO.Con.RowsAffected() > 0 Then
            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                form_no.w_size_code.Text = Rs.rdoColumns(0).Value
            Else
                form_no.w_size_code.Text = ""
            End If
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")
		
		Command6.Focus()
		
        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "ｼｽﾃﾑｴﾗｰ")

        On Error Resume Next
        Err.Clear()
        Rs.Close()
        end_sql()
        ' <- watanabe add VerUP(2011)

    End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_SERIARU(0)
		
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
                .HelpContext = 802
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim gm_alph As Object
		Dim gm_no As Object
        Dim ZumenName As String
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		Dim i As Object
		
		Dim w_mess As String


        ' -> watanabe add VerUP(2011)
        w_mess = ""
        ' <- watanabe add VerUP(2011)


		'Brand Ver.5 TIFF->BMP 変更 start
		'    If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Target character is not selected.", 64)
            w_size1.Focus()
            Exit Sub
        End If

        init_sql()

        DBTableName = DBName & "..hm_kanri"

        form_no.Enabled = False
        F_MSG.Show()

        '(Brand Ver.3 追加)
        For i = 1 To Len(form_no.w_size_code.Text)
            w_str = Mid(form_no.w_size_code.Text, i, 1)
            If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("Replacement for the primitive character size code has not been set in the configuration file (" & Tmp_Serial1_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("Replacement for the primitive character size code has not been set in the configuration file (" & Tmp_Serial1_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
            End If
        Next i

        If FreePicNum < 1 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
            GoTo error_section
        End If

        '    pic_no = -1
        '    result = SqlCmd(SqlConn, "SELECT haiti_pic")
        '    result = SqlCmd(SqlConn, " FROM " & DBTableName)
        '    result = SqlCmd(SqlConn, " WHERE (")
        '    result = SqlCmd(SqlConn, " font_name = '" & Mid$(w_hm_name.Text, 6) & "' AND")
        '    result = SqlCmd(SqlConn, " no = '" & Mid$(w_hm_name.Text, 7, 2) & "')")
        '
        '    result = SqlExec(SqlConn)
        '    result = SqlResults(SqlConn)
        '    If result = SUCCEED Then
        '      Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        pic_no = SqlData$(SqlConn, 1)
        '      Loop
        '    Else
        '      MsgBox "SQLｴﾗｰ", 64, "SQLｴﾗｰ"
        '       GoTo error_section
        '    End If

        '** 画面情報を送信 **
        Call temp_bz_get(3)
        Call bz_spec_set(w_mess)
        w_ret = PokeACAD("SPECADD", w_mess)
        w_ret = RequestACAD("SPECADD")


        '// 置換モードの送信
        w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
        w_ret = RequestACAD("CHNGMODE")

        '（図面名）送信
        pic_no = what_pic_from_hmcode(form_no.w_hm_name.Text)


        If pic_no < 1 Then GoTo error_section

        ZumenName = "HM-" & VB.Left(Trim(form_no.w_hm_name.Text), 6)


		'----- .NET 移行 -----
		'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
		w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

		w_ret = PokeACAD("HMCODE", w_mess)

        For i = 1 To Len(form_no.w_size_code.Text)
            If Mid(form_no.w_size_code.Text, i, 1) >= "0" And Mid(form_no.w_size_code.Text, i, 1) <= "9" Then
                gm_no = Val(Mid(form_no.w_size_code.Text, i, 1))
                pic_no = what_pic_from_gmcode(GensiNUM(gm_no))

                If pic_no < 1 Then GoTo error_section

                ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE1", w_mess)
            Else
                gm_alph = Mid(form_no.w_size_code.Text, i, 1)

                pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))

                If pic_no < 1 Then GoTo error_section

                ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE1", w_mess)
            End If

        Next i

		'----- .NET 移行 -----
		'w_ret = PokeACAD("SERIWIDE", VB6.Format(TmpSerialWidth))
		'w_ret = PokeACAD("SERIMOVE", VB6.Format(TmpSerialMove))

		w_ret = PokeACAD("SERIWIDE", TmpSerialWidth.ToString())
		w_ret = PokeACAD("SERIMOVE", TmpSerialMove.ToString())

		'// 終了の送信
		w_ret = RequestACAD("TMPCHANG")

LOOP_EXIT:

        end_sql()
        Exit Sub

error_section:
        end_sql()
        On Error Resume Next
        F_MSG.Close()
        form_no.Enabled = True
	End Sub
	
	Private Sub F_TMP_SERIARU_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_w_str As String
		Dim i As Short
		
		form_no = Me

		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		
		'タイヤ種類
		w_syurui.Items.Clear()
		w_syurui.Items.Add("PC")
		w_syurui.Items.Add("LT")
		w_syurui.Items.Add("TB")
		
		'フォント
		'(Brand Ver.3 追加)
		w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
				w_font.Items.Add(Tmp_font_word(i))
			End If
		Next i
		
		
		'工場
		'(Brand Ver.3 変更)
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Serial1_ini)
        ret = set_read5(w_w_str, "serial1", 1)
		w_plant.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "TT" Then
				'             w_plant.AddItem "仙台(TT)"
				'          ElseIf Tmp_hm_word(i) = "KW" Then
				'             w_plant.AddItem "桑名(KW)"
				'          ElseIf Tmp_hm_word(i) = "CS" Then
				'             w_plant.AddItem "正新(CS)"
				'          ElseIf Tmp_hm_word(i) = "CH" Then
				'             w_plant.AddItem "上海(CH)"
				'          End If
				w_plant.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		w_tmp_seri_width.Text = CStr(TmpSerialWidth)
		w_tmp_seri_move.Text = CStr(TmpSerialMove)
		
		Call Clear_F_TMP_SERIARU(0)
		
        CommunicateMode = comSpecData
        RequestACAD("SPECDATA")
		
		If Trim(w_syurui.Text) = "" Then
			w_syurui.Text = "PC"
		End If
		
        InitFlag = True '20100628追加コード
	End Sub
	
    'UPGRADE_WARNING: イベント w_font.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_font_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font.SelectedIndexChanged
		Dim ret As Object
		
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		'(Brand Cad System Ver.3 UP)
		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_Serial1_ini)
                ret = set_read5(w_w_str, "serial1", i)
                If ret = False Then
                    MsgBox(Tmp_Serial1_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Serial1_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'工場
		'(Brand Ver.3 変更)
		w_plant.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "TT" Then
				'             w_plant.AddItem "仙台(TT)"
				'          ElseIf Tmp_hm_word(i) = "KW" Then
				'             w_plant.AddItem "桑名(KW)"
				'          ElseIf Tmp_hm_word(i) = "CS" Then
				'             w_plant.AddItem "正新(CS)"
				'          ElseIf Tmp_hm_word(i) = "CH" Then
				'             w_plant.AddItem "上海(CH)"
				'          End If
				w_plant.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
		Dim w_file As Object
        Dim TiffFile As String
        Dim w_text As Object

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		On Error Resume Next ' エラーのトラップを留保します。
		Err.Clear()
		
        w_text = w_hm_name.Text
		
        If Trim(w_text) <> "" Then
            'Brand Ver.5 TIFF->BMP 変更 end
            '       TiffFile = TIFFDir & w_hm_name.Text & ".tif"
            '
            '       'Tiffﾌｧｲﾙ表示
            '       w_file = Dir(TiffFile)
            '       If w_file <> "" Then
            '          form_no.ImgThumbnail1.Image = TiffFile
            '          form_no.ImgThumbnail1.ThumbWidth = 500
            '          form_no.ImgThumbnail1.ThumbHeight = 200
            '       Else
            '          MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
            '       End If
            TiffFile = TIFFDir & w_hm_name.Text & ".bmp"

            'bmpﾌｧｲﾙ表示
            w_file = Dir(TiffFile)
            If w_file <> "" Then
                form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
                form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
                form_no.ImgThumbnail1.Height = 193 '200 '20100701コード変更
            Else
                MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            End If
            'Brand Ver.5 TIFF->BMP 変更 end
        End If
		
		w_plant.Focus()
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_plant.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_plant_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_plant.SelectedIndexChanged
		Dim i As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

		'(Brand Ver.3 変更)
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
		'20000124 修正
        If w_plant.Text = "Sendai(TT)" Then
            w_plant_code.Text = "CX"
            '      w_str = "TT"
        ElseIf w_plant.Text = "Kuwana(KW)" Then
            w_plant_code.Text = "N3"
            '      w_str = "KW"
        ElseIf w_plant.Text = "Cheng shin(CS)" Then
            w_plant_code.Text = "UY"
            '      w_str = "CS"
        ElseIf w_plant.Text = "Shanghai(CH)" Then
            w_plant_code.Text = "9T"
            '      w_str = "CH"
        End If
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'      If Tmp_hm_word(i) = Trim(w_str) Then
            If Tmp_hm_word(i) = w_plant.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                Exit For
            End If
		Next i
		
	End Sub
	
	Private Sub w_plant_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_plant.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_plant_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_plant.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_plant, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size1.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.Leave
		
        form_no.w_size1.Text = UCase(Trim(form_no.w_size1.Text))
		
	End Sub
	
	Private Sub w_size2_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size2.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.Leave
		
        form_no.w_size2.Text = UCase(Trim(form_no.w_size2.Text))
		
	End Sub
	
	Private Sub w_size3_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size3.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size3_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size3.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size3_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.Leave
		
        form_no.w_size3.Text = UCase(Trim(form_no.w_size3.Text))
		
	End Sub
	
    Private Sub w_size4_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.Leave

        form_no.w_size4.Text = UCase(Trim(form_no.w_size4.Text))

    End Sub
	
	Private Sub w_size5_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size5.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size5_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size5.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii > 32 Then
			If (KeyAscii = CDbl("100")) Or (KeyAscii = CDbl("114")) Then
				KeyAscii = KeyAscii - 32
			ElseIf (KeyAscii <> CDbl("45")) And (KeyAscii <> CDbl("68")) And (KeyAscii <> CDbl("82")) And (KeyAscii <> CDbl("42")) Then 
				KeyAscii = 0
			End If
			GoTo EventExitSub
		End If
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size5_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.Leave
		
        form_no.w_size5.Text = UCase(Trim(form_no.w_size5.Text))
		
	End Sub
	
	Private Sub w_size6_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size6.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size6_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size6.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size6_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.Leave
		
        form_no.w_size6.Text = UCase(Trim(form_no.w_size6.Text))
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_syurui.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_syurui_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_syurui.SelectedIndexChanged
		
		If Trim(w_hm_name.Text) <> "" Then
			Call Clear_F_TMP_SERIARU(1)
		End If
		
	End Sub
	
	Private Sub w_syurui_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_syurui.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				Call Clear_F_TMP_SERIARU(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_syurui_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syurui.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_syurui, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class