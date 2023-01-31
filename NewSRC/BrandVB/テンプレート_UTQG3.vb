Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_UTQG3
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Call Clear_F_TMP_UTQG3()
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		form_no.Close()
		End
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
                .HelpContext = 805
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim w_mess As String
		Dim w_str As String
		Dim w_ret As Short
		Dim pic_no As Short
		Dim gm_no As Short
		Dim gm_alph As String
		
		Dim key_value As String
		Dim tmp_str As String
		
		Dim grp_num As Short
		Dim top_dumy_num As Short
		Dim top_hmcode As String
		
		Dim grp_datum_no() As Short
		Dim grp_dist_x() As Double
		Dim grp_dist_y() As Double
		Dim grp_dumy_num() As Short
		Dim grp_hmcode() As String
		
		Dim ZumenName As String
		Dim change_num As Short
		Dim sub_num As Short
		Dim poke_gmcode(2) As String
		Dim poke_holdgm(2) As String
		
		Dim hexdata As String
		Dim str_dbl As New VB6.FixedLengthString(16)
		Dim str_int As New VB6.FixedLengthString(8)
		
		Dim error_no As String
		Dim i As Short
		Dim j As Short
		

		'/* 入力チェック */
        form_no.w_treadwear.Text = Trim(form_no.w_treadwear.Text)
		form_no.w_traction.Text = Trim(form_no.w_traction.Text)
		form_no.w_temperature.Text = Trim(form_no.w_temperature.Text)
		
		If check_F_TMP_UTQG <> 0 Then Exit Sub
		
		form_no.Enabled = False
		F_MSG.Show()
		
		If chk_treadwear.CheckState = 0 Then
			For i = 1 To Len(form_no.w_treadwear.Text)
				w_str = Mid(form_no.w_treadwear.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input TREADWEAR is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input TREADWEAR  is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
		End If
		
		If chk_traction.CheckState = 0 Then
			For i = 1 To Len(form_no.w_traction.Text)
				w_str = Mid(form_no.w_traction.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＲＡＣＴＩＯＮ is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input ＴＲＡＣＴＩＯＮ  is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
		End If
		
		If chk_temperature.CheckState = 0 Then
			For i = 1 To Len(form_no.w_temperature.Text)
				w_str = Mid(form_no.w_temperature.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＥＭＰＥＲＡＴＵＲＥ is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input TEMPERATURE  is not set to the configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
		End If
		
		If FreePicNum < 2 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
			GoTo error_section
		End If
		
		
		' グループデータの分解取得
		key_value = Tmp_hm_group(w_type.SelectedIndex + 1)
		
        ' グループ数
		tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
		If IsNumeric(tmp_str) = False Then
            MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
			GoTo error_section
		End If
		grp_num = CShort(tmp_str)
		key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		
        ' 先頭ダミー部数
		tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
		If IsNumeric(tmp_str) = False Then
            MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
			GoTo error_section
		End If
		top_dumy_num = CShort(tmp_str)
		key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		
        ' 先頭編集文字コード取得
		If grp_num = 1 Then
			top_hmcode = Trim(key_value)
		Else
			If InStr(key_value, "|") = 0 Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			top_hmcode = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		End If
		
		'MsgBox "1" & Chr(13) & grp_num & Chr(13) & top_dumy_num & Chr(13) & top_hmcode
		
        ' 追加編集文字データ取得
		ReDim grp_datum_no(grp_num)
		ReDim grp_dist_x(grp_num)
		ReDim grp_dist_y(grp_num)
		ReDim grp_dumy_num(grp_num)
		ReDim grp_hmcode(grp_num)
		For i = 0 To grp_num - 2
            ' 基準行
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_datum_no(i) = CShort(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' 距離X
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dist_x(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' 距離Y
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dist_y(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' ダミー部数
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dumy_num(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' 先頭編集文字コード取得
			If i = (grp_num - 2) Then
				grp_hmcode(i) = Trim(key_value)
			Else
				If InStr(key_value, "|") = 0 Then
                    MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
					GoTo error_section
				End If
				grp_hmcode(i) = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
				key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			End If
			
			'MsgBox CStr(i + 2) & Chr(13) & grp_datum_no(i) & Chr(13) & grp_dist_x(i) & Chr(13) & grp_dist_y(i) & Chr(13) & grp_dumy_num(i) & Chr(13) & grp_hmcode(i)
			
		Next i
		
		
		' 先頭編集文字作成
		change_num = 0
		
		'// 置換モードの送信
		w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
		w_ret = RequestACAD("CHNGMODE")
		
		'（図面名）送信
		pic_no = what_pic_from_hmcode(top_hmcode)
		If pic_no < 1 Then GoTo error_section
		ZumenName = "HM-" & VB.Left(Trim(top_hmcode), 6)

		'----- .NET 移行 -----
		'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
		w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

		w_ret = PokeACAD("HMCODE", w_mess)
		
		'[[ TREADWEAR ]]
		If top_dumy_num > 0 Then
			If chk_treadwear.CheckState = 0 Then
				For i = 1 To Len(form_no.w_treadwear.Text)
					gm_no = Val(Mid(form_no.w_treadwear.Text, i, 1))
					pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
					If pic_no < 1 Then GoTo error_section
					ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

					'----- .NET 移行 -----
					'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
					w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

					w_ret = PokeACAD("GMCODE1", w_mess)
				Next i
			Else
				w_mess = ""
				w_ret = PokeACAD("HOLDGM1", w_mess)
			End If
			change_num = change_num + 1
		End If
		
		'[[ TRACTION ]]
		If top_dumy_num > 1 Then
			If chk_traction.CheckState = 0 Then
				For i = 1 To Len(form_no.w_traction.Text)
					gm_alph = Mid(form_no.w_traction.Text, i, 1)
					pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
					If pic_no < 1 Then GoTo error_section
					ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

					'----- .NET 移行 -----
					'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
					w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

					w_ret = PokeACAD("GMCODE2", w_mess)
				Next i
			Else
				w_mess = ""
				w_ret = PokeACAD("HOLDGM2", w_mess)
			End If
			change_num = change_num + 1
		End If
		
		'[[ TEMPERATURE ]]
		If top_dumy_num > 2 Then
			If chk_temperature.CheckState = 0 Then
				For i = 1 To Len(form_no.w_temperature.Text)
					gm_alph = Mid(form_no.w_temperature.Text, i, 1)
					pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
					If pic_no < 1 Then GoTo error_section
					ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

					'----- .NET 移行 -----
					'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
					w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

					w_ret = PokeACAD("GMCODE3", w_mess)
				Next i
			Else
				w_mess = ""
				w_ret = PokeACAD("HOLDGM3", w_mess)
			End If
			change_num = change_num + 1
		End If


        ' -> watanabe add VerUP(2011)
        CommunicateMode = comTmpWait
        ' <- watanabe add VerUP(2011)

		'// テンプレート変換の送信
		w_ret = RequestACAD("TMPCHANG")
		
        ' CAD処理終了チェック
		If check_cad_run = False Then
			GoTo error_section
		End If

		'// 作図実行ＰＩＣ保持の送信
		w_ret = RequestACAD("TMPTOPPIC")
		
		'' CAD処理終了チェック
		If check_cad_run = False Then
			GoTo error_section
		End If
		
        ' -> watanabe add VerUP(2011)
        CommunicateMode = comNone
        ' <- watanabe add VerUP(2011)

		
        ' 配列に送信コードをセット
		poke_gmcode(0) = "GMCODE1"
		poke_gmcode(1) = "GMCODE2"
		poke_gmcode(2) = "GMCODE3"
		
		poke_holdgm(0) = "HOLDGM1"
		poke_holdgm(1) = "HOLDGM2"
		poke_holdgm(2) = "HOLDGM3"
		
		
        ' グループ数分ループ
		For j = 0 To grp_num - 2

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

            ' 前回データクリア
			w_ret = RequestACAD("TMPDATCLR")
			
            ' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)

			
            ' グループ編集文字作成
			sub_num = 0
			
			'// 置換モードの送信
			w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
			w_ret = RequestACAD("CHNGMODE")
			
			'（図面名）送信
			pic_no = what_pic_from_hmcode(grp_hmcode(j))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "HM-" & VB.Left(Trim(grp_hmcode(j)), 6)

			'----- .NET 移行 -----
			'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
			w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

			w_ret = PokeACAD("HMCODE", w_mess)
			
			'[[ TREADWEAR ]]
			If change_num = 0 And grp_dumy_num(j) > sub_num Then
				If chk_treadwear.CheckState = 0 Then
					For i = 1 To Len(form_no.w_treadwear.Text)
						gm_no = Val(Mid(form_no.w_treadwear.Text, i, 1))
						pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
						If pic_no < 1 Then GoTo error_section
						ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_ret = PokeACAD(poke_gmcode(sub_num), w_mess)
					Next i
				Else
					w_mess = ""
					w_ret = PokeACAD(poke_holdgm(sub_num), w_mess)
				End If
				change_num = change_num + 1
				sub_num = sub_num + 1
			End If
			
			'[[ TRACTION ]]
			If change_num = 1 And grp_dumy_num(j) > sub_num Then
				If chk_traction.CheckState = 0 Then
                    For i = 1 To Len(form_no.w_traction.Text)
                        gm_alph = Mid(form_no.w_traction.Text, i, 1)
                        pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_ret = PokeACAD(poke_gmcode(sub_num), w_mess)
                    Next i
				Else
					w_mess = ""
					w_ret = PokeACAD(poke_holdgm(sub_num), w_mess)
				End If
				change_num = change_num + 1
				sub_num = sub_num + 1
			End If
			
			'[[ TEMPERATURE ]]
			If change_num = 2 And grp_dumy_num(j) > sub_num Then
				If chk_temperature.CheckState = 0 Then
                    For i = 1 To Len(form_no.w_temperature.Text)
                        gm_alph = Mid(form_no.w_temperature.Text, i, 1)
                        pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_ret = PokeACAD(poke_gmcode(sub_num), w_mess)
                    Next i
				Else
					w_mess = ""
					w_ret = PokeACAD(poke_holdgm(sub_num), w_mess)
				End If
				change_num = change_num + 1
				sub_num = sub_num + 1
			End If
			

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

            '// テンプレート変換の送信
			w_ret = RequestACAD("TMPCHANG")
			
            ' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If
			
            '// 作図実行ＰＩＣ保持の送信
			w_ret = RequestACAD("TMPADDPIC")
			
            ' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If
			
            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)


            ' グループ化
			hexdata = ""
			w_ret = InttoHex(grp_datum_no(j), str_int.Value)
			hexdata = hexdata & str_int.Value
			
			w_ret = DbltoHex(grp_dist_x(j), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value
			
			w_ret = DbltoHex(grp_dist_y(j), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value


            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

			w_ret = PokeACAD("TMPGRPDAT", hexdata)
			w_ret = RequestACAD("TMPGRPADD")
			
            ' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)

        Next j
		
		
		' VB終了
		End
		
		
		'エラー処理
error_section: 
		On Error Resume Next
		
		F_MSG.Close()
		form_no.Enabled = True
		
		Erase grp_datum_no
		Erase grp_dist_x
		Erase grp_dist_y
		Erase grp_dumy_num
		Erase grp_hmcode
	End Sub
	
	Private Sub F_TMP_UTQG3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_w_str As String
		Dim w_ret As Short
		Dim ret As Short
		Dim i As Short
		
		form_no = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		
		'フォント
        form_no.w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
                form_no.w_font.Items.Add(Tmp_font_word(i)) '20100624コード変更
			End If
		Next i
		
		'タイプ
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg3_ini)
		ret = set_read6(w_w_str, "utqg3", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_UTQG3()
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
		w_ret = RequestACAD("PICEMPTY")
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_font.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_font_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font.SelectedIndexChanged
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String
		Dim ret As Short
		
		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg3_ini)
				ret = set_read6(w_w_str, "utqg3", i)
				If ret = False Then
                    MsgBox(Tmp_Utqg3_ini & "File reading error.", 64, "BrandVB error")
					Exit Sub
				Else
					read_flg = 1
					Exit For
				End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Utqg3_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'タイプ
		w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i)) '20100624コード変更
			End If
		Next i
		
        form_no.w_type.Text = ""
        form_no.w_hm_name.Text = ""
        form_no.ImgThumbnail1.Image = Nothing
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
		Dim w_text As String
		Dim TiffFile As String
		Dim w_file As String
		
		On Error Resume Next ' エラーのトラップを留保します。
		
		Err.Clear()
		
        w_text = form_no.w_hm_name.Text
		
		If Trim(w_text) = "" Then Exit Sub
		
        TiffFile = TIFFDir & form_no.w_hm_name.Text & ".bmp"
		
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
		
	End Sub
	
	Private Sub w_temperature_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_temperature.Leave
		form_no.w_temperature.Text = UCase(Trim(form_no.w_temperature.Text))
	End Sub
	
	Private Sub w_traction_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_traction.Leave
		form_no.w_traction.Text = UCase(Trim(form_no.w_traction.Text))
	End Sub
	
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged

        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        Dim i As Short
		
		For i = 1 To MaxSelNum
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
            form_no.ImgThumbnail1.Image = Nothing
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