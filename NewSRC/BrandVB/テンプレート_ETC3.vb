Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_ETC3
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Call Clear_F_TMP_ETC3()
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        InitFlag = False '20100628追加コード
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
                .HelpContext = 811
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim w_mess As String
		Dim w_w_n As Short
		Dim w_ret As Short
		Dim w_w_gmcode As String
		Dim w_cmd As String
		Dim w_str As String
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
		
		Dim hexdata As String
		Dim str_dbl As New VB6.FixedLengthString(16)
		Dim str_int As New VB6.FixedLengthString(8)
		
		Dim error_no As String
		Dim i As Short
		Dim j As Short
		Dim k As Short
		
		
		'/* 入力チェック */
		w_w_n = 0
		
		For i = 1 To MaxSelNum
			If (w_type.Text = Tmp_hm_word(i)) Then
				If (Tmp_prcs_code(i) = "ETC1") Then
					If (w_etc(1).Text = "") Then
                        MsgBox("Input error")
						Exit Sub
					End If
					w_w_n = 1
				ElseIf (Tmp_prcs_code(i) = "ETC2") Then 
					For j = 1 To 2
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 2
				ElseIf (Tmp_prcs_code(i) = "ETC3") Then 
					For j = 1 To 3
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 3
				ElseIf (Tmp_prcs_code(i) = "ETC4") Then 
					For j = 1 To 4
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 4
				ElseIf (Tmp_prcs_code(i) = "ETC5") Then 
					For j = 1 To 5
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 5
				ElseIf (Tmp_prcs_code(i) = "ETC6") Then 
					For j = 1 To 6
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 6
				ElseIf (Tmp_prcs_code(i) = "ETC7") Then 
					For j = 1 To 7
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 7
				ElseIf (Tmp_prcs_code(i) = "ETC8") Then 
					For j = 1 To 8
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 8
				ElseIf (Tmp_prcs_code(i) = "ETC9") Then 
					For j = 1 To 9
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 9
				ElseIf (Tmp_prcs_code(i) = "ETC10") Then 
					For j = 1 To 10
						If w_etc(j).Text = "" Then
                            MsgBox("Input error")
							Exit Sub
						End If
					Next j
					w_w_n = 10
				End If
			End If
		Next i
		
		For i = 1 To 10
            form_no.w_etc(i).Text = Trim(form_no.w_etc(i).Text)
		Next i
		
		If check_F_TMP_ETC <> 0 Then
			Exit Sub
		End If
		
		form_no.Enabled = False
		F_MSG.Show()
		
		For j = 1 To w_w_n
            For i = 1 To Len(form_no.w_etc(j).Text)
                w_str = Mid(form_no.w_etc(j).Text, i, 1)
                If IsNumeric(w_str) Then
                    If Val(w_str) >= 0 And Val(w_str) < 10 Then
                        If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input " & w_str & "  is not set to the configuration file (" & Tmp_ETC3_ini & ")", 64, "Configuration file error")
                            GoTo error_section
                        End If
                    End If
                ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                    If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & "  is not set to the configuration file (" & Tmp_ETC3_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                ElseIf Asc("a") <= Asc(w_str) And Asc(w_str) <= Asc("z") Then
                    If GensiALPHS(Asc(w_str) - Asc("a")) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & "  is not set to the configuration file (" & Tmp_ETC3_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                ElseIf 33 <= Asc(w_str) And Asc(w_str) <= 126 Then
                    If GensiKIGO(Asc(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & "  is not set to the configuration file (" & Tmp_ETC3_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            Next i
		Next j
		
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
			
        Next i
		
		
		' 先頭編集文字作成
		change_num = 0
		
		'// 置換モードの送信
        w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
		w_ret = RequestACAD("CHNGMODE")
		
		
		' -> watanabe add 2007.06
		'（ブランド番号）送信
		w_ret = PokeACAD("TMPETC3BNO", CStr(Tmp_brd_no))
		' <- watanabe add 2007.06
		
		
		'（図面名）送信
		pic_no = what_pic_from_hmcode(top_hmcode)
		If pic_no < 1 Then GoTo error_section
		ZumenName = "HM-" & VB.Left(Trim(top_hmcode), 6)

		'----- .NET 移行 -----
		'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
		w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

		w_ret = PokeACAD("HMCODE", w_mess)
		
		'[[ TYPE(1～10) ]]
		For j = 1 To w_w_n
			If top_dumy_num > j - 1 Then
                For i = 1 To Len(form_no.w_etc(j).Text)
                    w_str = Mid(form_no.w_etc(j).Text, i, 1)
                    If IsNumeric(w_str) Then
                        gm_no = Val(w_str)
                        pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_cmd = "GMCODE" & j
                        w_ret = PokeACAD(w_cmd, w_mess)
                    ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                        gm_no = Asc(w_str) - Asc("A")
                        pic_no = what_pic_from_gmcode(GensiALPH(gm_no))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiALPH(gm_no), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_cmd = "GMCODE" & j
                        w_ret = PokeACAD(w_cmd, w_mess)
                    ElseIf Asc("a") <= Asc(w_str) And Asc(w_str) <= Asc("z") Then
                        gm_no = Asc(w_str) - Asc("a")
                        pic_no = what_pic_from_gmcode(GensiALPHS(gm_no))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiALPHS(gm_no), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_cmd = "GMCODE" & j
                        w_ret = PokeACAD(w_cmd, w_mess)
                    ElseIf 33 <= Asc(w_str) And Asc(w_str) <= 126 Then
                        gm_no = Asc(w_str)
                        pic_no = what_pic_from_gmcode(GensiKIGO(gm_no))
                        If pic_no < 1 Then GoTo error_section
                        ZumenName = "GM-" & Mid(GensiKIGO(gm_no), 1, 6)

						'----- .NET 移行 -----
						'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
						w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

						w_cmd = "GMCODE" & j
                        w_ret = PokeACAD(w_cmd, w_mess)
                    End If
                Next i
				change_num = change_num + 1
			End If
		Next j
		
		'// 終了の送信
        CommunicateMode = comNone


        ' -> watanabe add VerUP(2011)
        CommunicateMode = comTmpWait
        ' <- watanabe add VerUP(2011)

        w_ret = RequestACAD("TMPCHANG")

        ' CAD処理終了チェック
		If check_cad_run = False Then
			GoTo error_section
		End If
		
        '// 作図実行ＰＩＣ保持の送信
		w_ret = RequestACAD("TMPTOPPIC")
		
        ' CAD処理終了チェック
		If check_cad_run = False Then
			GoTo error_section
		End If
		
        ' -> watanabe add VerUP(2011)
        CommunicateMode = comNone
        ' <- watanabe add VerUP(2011)


        ' グループ数分ループ
		For k = 0 To grp_num - 2
			
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
			pic_no = what_pic_from_hmcode(grp_hmcode(k))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "HM-" & VB.Left(Trim(grp_hmcode(k)), 6)

			'----- .NET 移行 -----
			'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
			w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

			w_ret = PokeACAD("HMCODE", w_mess)
			
			'[[ TYPE(1～10) ]]
			For j = (change_num + 1) To w_w_n
				If grp_dumy_num(k) > sub_num Then
                    For i = 1 To Len(form_no.w_etc(j).Text)
                        w_str = Mid(form_no.w_etc(j).Text, i, 1)
                        If IsNumeric(w_str) Then
                            gm_no = Val(w_str)
                            pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                            If pic_no < 1 Then GoTo error_section
                            ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

							'----- .NET 移行 -----
							'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
							w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

							w_cmd = "GMCODE" & (sub_num + 1)
                            w_ret = PokeACAD(w_cmd, w_mess)
                        ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                            gm_no = Asc(w_str) - Asc("A")
                            pic_no = what_pic_from_gmcode(GensiALPH(gm_no))
                            If pic_no < 1 Then GoTo error_section
                            ZumenName = "GM-" & Mid(GensiALPH(gm_no), 1, 6)

							'----- .NET 移行 -----
							'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
							w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

							w_cmd = "GMCODE" & (sub_num + 1)
                            w_ret = PokeACAD(w_cmd, w_mess)
                        ElseIf Asc("a") <= Asc(w_str) And Asc(w_str) <= Asc("z") Then
                            gm_no = Asc(w_str) - Asc("a")
                            pic_no = what_pic_from_gmcode(GensiALPHS(gm_no))
                            If pic_no < 1 Then GoTo error_section
                            ZumenName = "GM-" & Mid(GensiALPHS(gm_no), 1, 6)

							'----- .NET 移行 -----
							'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
							w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

							w_cmd = "GMCODE" & (sub_num + 1)
                            w_ret = PokeACAD(w_cmd, w_mess)
                        ElseIf 33 <= Asc(w_str) And Asc(w_str) <= 126 Then
                            gm_no = Asc(w_str)
                            pic_no = what_pic_from_gmcode(GensiKIGO(gm_no))
                            If pic_no < 1 Then GoTo error_section
                            ZumenName = "GM-" & Mid(GensiKIGO(gm_no), 1, 6)

							'----- .NET 移行 -----
							'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
							w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

							w_cmd = "GMCODE" & (sub_num + 1)
                            w_ret = PokeACAD(w_cmd, w_mess)
                        End If
                    Next i
					change_num = change_num + 1
					sub_num = sub_num + 1
				End If
			Next j
			

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

            '// テンプレート変換の送信
			w_ret = RequestACAD("TMPCHANG")
			
			'' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If

			'// 作図実行ＰＩＣ保持の送信
			w_ret = RequestACAD("TMPADDPIC")
			
			'' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If
			
            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)


            '' グループ化
			hexdata = ""
			w_ret = InttoHex(grp_datum_no(k), str_int.Value)
			hexdata = hexdata & str_int.Value
			
			w_ret = DbltoHex(grp_dist_x(k), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value
			
			w_ret = DbltoHex(grp_dist_y(k), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value


            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

			w_ret = PokeACAD("TMPGRPDAT", hexdata)
			w_ret = RequestACAD("TMPGRPADD")
			
			'' CAD処理終了チェック
			If check_cad_run = False Then
				GoTo error_section
			End If

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)

        Next k
		
		
		' VB終了
		End
		
		
		'    '// 画面ロック
		'    form_no.Command2.Enabled = False
		'    form_no.Command4.Enabled = False
		'    form_no.Command6.Enabled = False
		'    form_no.w_type.Enabled = False
		'    form_no.w_type.BackColor = &HC0C0C0
		'    For i = 1 To 10
		'        form_no.w_etc(i).Enabled = False
		'        form_no.w_etc(i).BackColor = &HC0C0C0
		'    Next i
		
		Exit Sub
		
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
	
	Private Sub F_TMP_ETC3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_w_str As String
		Dim i As Short
		Dim ret As Short
		Dim w_ret As Short
		
        form_no = Me

		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		'フォント
        form_no.w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
                form_no.w_font.Items.Add(Tmp_font_word(i))
			End If
		Next i
		
		'タイプ
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_ETC3_ini)
		ret = set_read6(w_w_str, "etc3", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_ETC3()
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
		w_ret = RequestACAD("PICEMPTY")

        InitFlag = True '20100628追加コード
	End Sub
	
    'UPGRADE_WARNING: イベント w_font.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_font_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font.SelectedIndexChanged
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String
		Dim ret As Short

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_ETC3_ini)
				ret = set_read6(w_w_str, "etc3", i)
				If ret = False Then
                    MsgBox(Tmp_ETC3_ini & "File reading error.", 64, "BrandVB error")
					Exit Sub
				Else
					read_flg = 1
					Exit For
				End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_ETC3_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'タイプ
		w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
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

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        On Error Resume Next
		
		Err.Clear()
		
		w_text = w_hm_name.Text
		
		If Trim(w_text) = "" Then Exit Sub
		
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
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.TextChanged

        ' -> watanabe del VerUP(2011)
        'Dim flg As Short
        ' <- watanabe del VerUP(2011)

        Dim i As Short
        Dim j As Short

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If
		
		For i = 1 To MaxSelNum
			If (w_type.Text = Tmp_hm_word(i)) Then
				If (Tmp_prcs_code(i) = "ETC1") Then
                    form_no.w_etc(1).Enabled = True
                    form_no.w_etc(1).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
					For j = 2 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC2") Then 
					For j = 1 To 2
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    Next j
                    For j = 3 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    Next j
				ElseIf (Tmp_prcs_code(i) = "ETC3") Then 
					For j = 1 To 3
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 4 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC4") Then 
					For j = 1 To 4
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 5 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC5") Then 
					For j = 1 To 5
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 6 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC6") Then 
					For j = 1 To 6
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 7 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC7") Then 
					For j = 1 To 7
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 8 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC8") Then 
					For j = 1 To 8
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 9 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC9") Then 
					For j = 1 To 9
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
                    form_no.w_etc(j).Enabled = False
                    form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				ElseIf (Tmp_prcs_code(i) = "ETC10") Then 
					For j = 1 To 10
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
				End If
			End If
		Next i
		
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = w_type.Text Then
				w_hm_name.Text = Tmp_hm_code(i)
				Exit For
			End If
		Next i
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged

        ' -> watanabe del VerUP(2011)
        'Dim flg As Short
        ' <- watanabe del VerUP(2011)

        Dim i As Short
		Dim j As Short

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		For i = 1 To MaxSelNum
			If (w_type.Text = Tmp_hm_word(i)) Then
				If (Tmp_prcs_code(i) = "ETC1") Then
                    form_no.w_etc(1).Enabled = True
                    form_no.w_etc(1).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
					For j = 2 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC2") Then 
					For j = 1 To 2
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 3 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC3") Then 
					For j = 1 To 3
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 4 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC4") Then 
					For j = 1 To 4
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 5 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC5") Then 
					For j = 1 To 5
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 6 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC6") Then 
					For j = 1 To 6
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 7 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC7") Then 
					For j = 1 To 7
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 8 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC8") Then 
					For j = 1 To 8
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
					For j = 9 To 10
                        form_no.w_etc(j).Enabled = False
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
					Next j
				ElseIf (Tmp_prcs_code(i) = "ETC9") Then 
					For j = 1 To 9
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
                    form_no.w_etc(j).Enabled = False
                    form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				ElseIf (Tmp_prcs_code(i) = "ETC10") Then 
					For j = 1 To 10
                        form_no.w_etc(j).Enabled = True
                        form_no.w_etc(j).BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
					Next j
				End If
			End If
		Next i
		
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