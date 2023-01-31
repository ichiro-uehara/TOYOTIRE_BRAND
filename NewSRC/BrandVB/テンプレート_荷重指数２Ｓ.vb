Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_KAJUU2S
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
        Dim result As Integer
        Dim w_text As String
		
		Dim tmp_dbtable As Object
		Dim select_item As Object
		Dim select_where(5) As Object
		Dim select_num As Object
		Dim w_str As String
		Dim w_text1 As String
		Dim w_text2 As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

		On Error Resume Next
		Err.Clear()

        ' -> watanabe mov VerUP(2011)
        select_item = ""
        tmp_dbtable = ""
        ' <- watanabe mov VerUP(2011)


		'/* 入力チェック */
		If check_F_TMP_KAJUU <> 0 Then Exit Sub
		
		'''''   MsgBox "検索します。"
		
		'// 荷重指数検索
        w_text = Trim(form_no.w_size1.Text) & Trim(form_no.w_size2.Text) & Trim(form_no.w_size3.Text) & Trim(form_no.w_size5.Text) & Trim(form_no.w_size6.Text)
		w_str = ""
        w_str = "WHERE syurui = '" & Trim(form_no.w_syurui.Text) & "'"
		
		'  If w_text <> "" Then
        select_num = 0
		'     If Trim(form_no.w_size1.Text) <> "" Then
        w_str = w_str & " AND size1 = '" & Trim(form_no.w_size1.Text) & "'"
		'     End If
		'     If Trim(form_no.w_size2.Text) <> "" Then
        w_str = w_str & " AND size2 = '" & Trim(form_no.w_size2.Text) & "'"
		'     End If
		'     If Trim(form_no.w_size3.Text) <> "" Then
        w_str = w_str & " AND size3 = '" & Trim(form_no.w_size3.Text) & "'"
		'     End If
		'     If Trim(form_no.w_size5.Text) <> "" Then
        w_str = w_str & " AND size5 = '" & Trim(form_no.w_size5.Text) & "'"
		'     End If
		'     If Trim(form_no.w_size6.Text) <> "" Then
        w_str = w_str & " AND size6 = '" & Trim(form_no.w_size6.Text) & "'"
		'     End If
		'  End If
		
        Select Case form_no.w_kikaku.Text
            Case "JATMA"
                select_item = "standard_load_index"
                tmp_dbtable = "..jatma"
            Case "TRA (standard)"
                select_item = "standard_load_index"
                tmp_dbtable = "..tra"
            Case "TRA (special)"
                select_item = "extra_load_index"
                tmp_dbtable = "..tra"
            Case "TRA (light)"
                select_item = "light_load_index"
                tmp_dbtable = "..tra"
            Case "ETRTO(Standard)"
                select_item = "standard_load_index"
                tmp_dbtable = "..etrto"
            Case "ETRTO(Special)"
                select_item = "extra_load_index"
                tmp_dbtable = "..etrto"
        End Select

        init_sql()


        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT " & select_item)
        '      result = SqlCmd(SqlConn, " FROM " & STANDARD_DBName & tmp_dbtable)
        ''  If w_text <> "" Then
        '      result = SqlCmd(SqlConn, " " & w_str)
        ''  End If
        '      result = SqlExec(SqlConn)
        '      result = SqlResults(SqlConn)
        '      If result = SUCCEED Then
        '          If SqlNextRow(SqlConn) = REGROW Then
        '              w_load_index.Text = SqlData(SqlConn, 1)
        '          End If
        '      Else
        '          MsgBox("ﾃﾞｰﾀﾍﾞｰｽSELECTｴﾗｰ", MsgBoxStyle.Critical)
        '      End If


        '検索コマンド作成
        sqlcmd = "SELECT " & select_item
        sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & tmp_dbtable
        sqlcmd = sqlcmd & " " & w_str

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        w_load_index.Text = ""
        If GL_T_RDO.Con.RowsAffected() > 0 Then
            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                w_load_index.Text = Rs.rdoColumns(0).Value
            Else
                w_load_index.Text = ""
            End If
        Else
            MsgBox("Load index database select error.", MsgBoxStyle.Critical)
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()
		
		'// 該当データが無い場合
        If Trim(form_no.w_load_index.Text) = "" Then
            Exit Sub
        End If
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")
        form_no.w_hm_name.Focus() 'SetFocus()
		

        init_sql()


        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Dim gm_no As Object
        Dim ZumenName As String
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		
		Dim w_mess As String
		Dim w_w_str As String
		Dim i As Short
		Dim load_data As String


        ' -> watanabe add VerUP(2011)
        w_mess = ""
        ' <- watanabe add VerUP(2011)


		If check_F_TMP_KAJUU <> 0 Then Exit Sub
		
		If Len(w_load_index.Text) = 0 Then
            MsgBox("Target character is not selected.")
			Exit Sub
		End If
		
		load_data = w_load_index.Text & w_sokudo.Text
		
		form_no.Enabled = False
		F_MSG.Show()
		
		'(Brand Ver.3 追加)
		For i = 1 To Len(load_data)
            w_str = Mid(load_data, i, 1)
			If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for selected load index is not set to the configuration file (" & Tmp_Load2S_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for selected load index is not set to the configuration file (" & Tmp_Load2S_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
			End If
		Next i
		
		If FreePicNum < 1 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
			GoTo error_section
		End If
		
		'** 画面情報を送信 **
		Call temp_bz_get(2)
		Call bz_spec_set(w_mess)
        w_ret = PokeACAD("SPECADD", w_mess)
        w_ret = RequestACAD("SPECADD")
		
		'// 置換モードの送信
        w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
        w_ret = RequestACAD("CHNGMODE")
		
		'（図面名）送信
        pic_no = what_pic_from_hmcode(Tmp2_Dummy_HM)
		
        If pic_no < 1 Then
            MsgBox("From the database, you did not get to edit letter data.", 64, "SQL error")
            GoTo error_section
        End If
        ZumenName = "HM-" & VB.Left(Trim(Tmp2_Dummy_HM), 6)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
        w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("HMCODE", w_mess)
		
		'[荷重指数]
		For i = 1 To Len(load_data)
            w_str = Mid(load_data, i, 1)
			
			If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    gm_no = Val(w_str)
                    pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                    If pic_no < 1 Then
                        MsgBox("From the database, it was not possible to get data primitive letter (" & GensiNUM(gm_no) & ")", 64, "SQL error")
                        GoTo error_section
                    End If
                    ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                    '----- .NET 移行 -----
                    'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                    w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                    w_ret = PokeACAD("GMCODE1", w_mess)
                End If
				
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                gm_no = Asc(w_str) - Asc("A")
                pic_no = what_pic_from_gmcode(GensiALPH(gm_no))
                If pic_no < 1 Then
                    MsgBox("From the database, it was not possible to get data primitive letter (" & GensiALPH(gm_no) & ")", 64, "SQL error")
                    GoTo error_section
                End If
                ZumenName = "GM-" & Mid(GensiALPH(gm_no), 1, 6)

                '----- .NET 移行 -----
                'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                w_ret = PokeACAD("GMCODE1", w_mess)
			End If
			
		Next i
		
		'// 終了の送信
        w_mess = Tmp_Load2S_ini
        w_ret = PokeACAD("TMPNAME", w_mess)
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_mess = "TYPE" & i
                w_ret = PokeACAD("TMPDATANO", w_mess)
				Exit For
			End If
		Next i
		w_mess = Trim(load_data)
        w_ret = PokeACAD("TMPSPELL", w_mess)

        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG3")

        form_no.Command1.Enabled = False
        form_no.Command2.Enabled = False
        form_no.Command3.Enabled = False
        form_no.Command5.Enabled = False
        form_no.w_size1.Enabled = False
        form_no.w_size1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.w_size2.Enabled = False
        form_no.w_size2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_size3.Enabled = False
        form_no.w_size3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_size4.Enabled = False
        form_no.w_size4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_size5.Enabled = False
        form_no.w_size5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_size6.Enabled = False
        form_no.w_size6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_syurui.Enabled = False
        form_no.w_syurui.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_kikaku.Enabled = False
        form_no.w_kikaku.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_font.Enabled = False
        form_no.w_font.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_sokudo.Enabled = False
        form_no.w_sokudo.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		
        form_no.w_size1.Text = ""
        form_no.w_size2.Text = ""
        form_no.w_size3.Text = ""
        form_no.w_size4.Text = ""
        form_no.w_size5.Text = ""
        form_no.w_size6.Text = ""
        'form_no.w_syurui.ListIndex = 0
        form_no.w_syurui.Text = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(0)) '20100624変更コード
        'form_no.w_kikaku.ListIndex = 0
        form_no.w_kikaku.Text = form_no.w_kikaku.GetItemText(form_no.w_kikaku.Items(0))
        'form_no.w_font.ListIndex = 0
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0))
        form_no.w_sokudo.Text = ""
        form_no.w_load_index.Text = ""
		
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		
		form_no.Close()
		End
		
	End Sub
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
        On Error Resume Next
        Err.Clear()
        Dim oCommonDialog As Object
        oCommonDialog = CreateObject("MSComDlg.CommonDialog")

        If Err.Number = 0 Then
            With oCommonDialog
                .HelpCommand = cdlHelpContext
                .HelpFile = "c:\VBhelp\BRAND.HLP"
                .HelpContext = 801
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_TMP_KAJUU2S_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Object
		Dim i As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_w_str As String
		
		form_no = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		'タイヤ種類
		w_syurui.Items.Clear()
		w_syurui.Items.Add("PC")
		w_syurui.Items.Add("LT")
		w_syurui.Items.Add("TB")
		
		'規格
		w_kikaku.Items.Clear()
		w_kikaku.Items.Add("JATMA")
        w_kikaku.Items.Add("TRA (standard)")
        w_kikaku.Items.Add("TRA (special)")
        w_kikaku.Items.Add("TRA (light)")
        w_kikaku.Items.Add("ETRTO(Standard)")
        w_kikaku.Items.Add("ETRTO(Special)")
		
		'フォント
		'(Brand Ver.3 変更)
        form_no.w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt + 1
            If Trim(Tmp_font_word(i)) = "" Then
                Exit For
            Else
                form_no.w_font.Items.Add(Tmp_font_word(i))
            End If
		Next i
		
        form_no.w_size1.Text = ""
        form_no.w_size2.Text = ""
        form_no.w_size3.Text = ""
        form_no.w_size4.Text = ""
        form_no.w_size5.Text = ""
        form_no.w_size6.Text = ""
        'form_no.w_syurui.ListIndex = 0
        form_no.w_syurui.Text = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(0)) '20100624変更コード
        'form_no.w_kikaku.ListIndex = 0
        form_no.w_kikaku.Text = form_no.w_kikaku.GetItemText(form_no.w_kikaku.Items(0))
        'form_no.w_font.ListIndex = 0
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0))
        form_no.w_sokudo.Text = ""
        form_no.w_load_index.Text = ""
		
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Load2S_ini)
        ret = set_read4(w_w_str, "load2S", 1)

        CommunicateMode = comSpecData
		RequestACAD("SPECDATA")
		
		If Trim(w_syurui.Text) = "" Then
			w_syurui.Text = "PC"
		End If
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_font.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_font_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font.SelectedIndexChanged
		Dim ret As Object
		
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String
		
		'(Brand Cad System Ver.3 UP)
		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_Load2S_ini)
                ret = set_read4(w_w_str, "load2S", i)
                If ret = False Then
                    MsgBox(Tmp_Load2S_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Load2S_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
	End Sub
	
	Private Sub w_size1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.Leave
		
        form_no.w_size1.Text = UCase(Trim(form_no.w_size1.Text))
		
	End Sub
	
	Private Sub w_size2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.Leave
		
        form_no.w_size2.Text = UCase(Trim(form_no.w_size2.Text))
		
	End Sub
	
	Private Sub w_size3_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.Leave
		
        form_no.w_size3.Text = UCase(Trim(form_no.w_size3.Text))
		
	End Sub
	
	Private Sub w_size4_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.Leave
		
        form_no.w_size4.Text = UCase(Trim(form_no.w_size4.Text))
		
	End Sub
	
	Private Sub w_size5_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.Leave
		
        form_no.w_size5.Text = UCase(Trim(form_no.w_size5.Text))
		
	End Sub
	
	Private Sub w_size6_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.Leave
		
        form_no.w_size6.Text = UCase(Trim(form_no.w_size6.Text))
		
	End Sub
	
	Private Sub w_sokudo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_sokudo.Leave
		
        form_no.w_sokudo.Text = UCase(Trim(form_no.w_sokudo.Text))
		
	End Sub
End Class