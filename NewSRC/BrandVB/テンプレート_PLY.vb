Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_PLY
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_PLY()
		
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
                .HelpContext = 807
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim gm_no As Object
		Dim ZumenName As Object
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		Dim i As Object
		
		Dim w_mess As String
		Dim w_w_n As Short
		Dim w_w_gmcode As String
		
		Dim flg As Short
		Dim wstr As String


        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


		'20000124 追加
		flg = 0
        If form_no.w_sidewall.Text >= "2" Then
            flg = 1
        End If
		
		'20000124 追加
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = form_no.w_type.Text Then
				If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
					wstr = Tmp_prcs_code(i)
					Exit For
				ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then 
					wstr = Tmp_prcs_code(i)
					Exit For
				End If
			End If
		Next i
		
		
		'/* 入力チェック */
		w_w_n = 0
		
		'20000124 変更
		'   If Trim$(form_no.w_type.Text) = "POLYESTER+STEEL" Then
		'      w_w_n = 2
		'      If (w_tread1 = "") Or (w_tread2 = "") Then
		'         MsgBox "入力が違います"
		'         Exit Sub
		'      End If
		'   End If
		
		'   If Trim$(form_no.w_type.Text) = "POLYESTER+STEEL+NYLON" Then
		'      w_w_n = 3
		'      If (w_tread1 = "") Or (w_tread2 = "") Or (w_tread3 = "") Then
		'         MsgBox "入力が違います"
		'         Exit Sub
		'      End If
		'   End If
		
		'   If Trim$(form_no.w_type.Text) = "RAYON+STEEL" Then
		'      w_w_n = 2
		'      If (w_tread1 = "") Or (w_tread2 = "") Then
		'         MsgBox "入力が違います"
		'         Exit Sub
		'      End If
		'   End If
		
		'   If Trim$(form_no.w_type.Text) = "RAYON+STEEL+NYLON" Then
		'      w_w_n = 3
		'      If (w_tread1 = "") Or (w_tread2 = "") Or (w_tread3 = "") Then
		'         MsgBox "入力が違います"
		'         Exit Sub
		'      End If
		'   End If
		
		'   If Trim$(form_no.w_type.Text) = "NYLON" Then
		'      w_w_n = 1
		'      If w_tread1 = "" Then
		'         MsgBox "入力が違います"
		'         Exit Sub
		'      End If
		'   End If
		
		If (Trim(wstr) = "PLY1") Or (Trim(wstr) = "PLY1S") Then
			w_w_n = 1
			' -> watanabe edit 2007.03 -> ReEdit
			If w_tread1.Text = "" Then
				'      If chk_tread1.Value = 0 And w_tread1 = "" Then
				' <- watanabe edit 2007.03 -> ReEdit
                MsgBox("Input error.")
				Exit Sub
			End If
		End If
		
		If (Trim(wstr) = "PLY2") Or (Trim(wstr) = "PLY2S") Then
			w_w_n = 2
			' -> watanabe edit 2007.03 -> ReEdit
			If (w_tread1.Text = "") Or (w_tread2.Text = "") Then
				'      If (chk_tread1.Value = 0 And w_tread1 = "") Or (chk_tread2.Value = 0 And w_tread2 = "") Then
				' <- watanabe edit 2007.03 -> ReEdit
                MsgBox("Input error.")
				Exit Sub
			End If
		End If
		
		If (Trim(wstr) = "PLY3") Or (Trim(wstr) = "PLY3S") Then
			w_w_n = 3
			' -> watanabe edit 2007.03 -> ReEdit
			If (w_tread1.Text = "") Or (w_tread2.Text = "") Or (w_tread3.Text = "") Then
				'      If (chk_tread1.Value = 0 And w_tread1 = "") Or (chk_tread2.Value = 0 And w_tread2 = "") Or (chk_tread3.Value = 0 And w_tread3 = "") Then
				' <- watanabe edit 2007.03 -> ReEdit
                MsgBox("Input error.")
				Exit Sub
			End If
		End If

		'----- .NET 移行 -----
		'w_tread.Text = VB6.Format(Val(w_tread1.Text) + Val(w_tread2.Text) + Val(w_tread3.Text), "#")

		'If w_w_n = 1 Then w_tread.Text = VB6.Format(Val(w_tread1.Text), "#")
		'If w_w_n = 2 Then w_tread.Text = VB6.Format(Val(w_tread1.Text) + Val(w_tread2.Text), "#")

		w_tread.Text = (Val(w_tread1.Text) + Val(w_tread2.Text) + Val(w_tread3.Text)).ToString("#")

		If w_w_n = 1 Then w_tread.Text = Val(w_tread1.Text).ToString("#")
		If w_w_n = 2 Then w_tread.Text = (Val(w_tread1.Text) + Val(w_tread2.Text)).ToString("#")

		form_no.w_tread1.Text = Trim(form_no.w_tread1.Text)
		form_no.w_tread2.Text = Trim(form_no.w_tread2.Text)
		form_no.w_tread3.Text = Trim(form_no.w_tread3.Text)
		
		If check_F_TMP_PLY <> 0 Then
			Exit Sub
		End If
		
		form_no.Enabled = False
		F_MSG.Show()
		
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    If chk_tread1.Value = 0 Or chk_tread2.Value = 0 Or chk_tread3.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		'(Brand Ver.3 追加)
		'UPGRADE_ISSUE: Control w_tread は、汎用名前空間 Form 内にあるため、解決できませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' をクリックしてください。
		For i = 1 To Len(form_no.w_tread.Text)
			w_str = Mid(form_no.w_tread.Text, i, 1)
			If IsNumeric(w_str) Then
				If Val(w_str) >= 0 And Val(w_str) < 10 Then
					If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for calculated TREAD is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
				If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for calculated TREAD is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
					GoTo error_section
				End If
			End If
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    End If
		
		'    If chk_tread1.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		For i = 1 To Len(form_no.w_tread1.Text)
			w_str = Mid(form_no.w_tread1.Text, i, 1)
			If IsNumeric(w_str) Then
				If Val(w_str) >= 0 And Val(w_str) < 10 Then
					If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input ＴＲＥＡＤ1  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
				If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for input ＴＲＥＡＤ1  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
					GoTo error_section
				End If
			End If
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    End If
		' <- watanabe Add 2007.03 -> ReEdit
		
		If w_w_n >= 2 Then
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        If chk_tread2.Value = 0 Then
			' <- watanabe Add 2007.03 -> ReEdit
			
			For i = 1 To Len(form_no.w_tread2.Text)
				w_str = Mid(form_no.w_tread2.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＲＥＡＤ2  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input ＴＲＥＡＤ2  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        End If
			' <- watanabe Add 2007.03 -> ReEdit
			
		End If
		
		If w_w_n >= 3 Then
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        If chk_tread3.Value = 0 Then
			' <- watanabe Add 2007.03 -> ReEdit
			
			For i = 1 To Len(form_no.w_tread3.Text)
				w_str = Mid(form_no.w_tread3.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＲＥＡＤ3  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input ＴＲＥＡＤ3  is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        End If
			' <- watanabe Add 2007.03 -> ReEdit
			
		End If
		
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    If chk_sidewall.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		For i = 1 To Len(form_no.w_sidewall.Text)
			w_str = Mid(form_no.w_sidewall.Text, i, 1)
			If IsNumeric(w_str) Then
				If Val(w_str) >= 0 And Val(w_str) < 10 Then
					If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input SIDEWALL is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
				If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for input SIDEWALL is not set to the configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
					GoTo error_section
				End If
			End If
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    End If
		' <- watanabe Add 2007.03 -> ReEdit
		
		
		If FreePicNum < 1 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
			GoTo error_section
		End If
		
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
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    If chk_tread1.Value = 0 Or chk_tread2.Value = 0 Or chk_tread3.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		'[[ TREAD ]]
		For i = 1 To Len(form_no.w_tread.Text)
			gm_no = Val(Mid(form_no.w_tread.Text, i, 1))
			pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

			'----- .NET 移行 -----
			'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
			w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

			w_ret = PokeACAD("GMCODE1", w_mess)
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    Else
		'        w_mess = ""
		'        w_ret = PokeACAD("HOLDGM1", w_mess)
		'    End If
		'
		'    If chk_tread1.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		'[[ TREAD1 ]]
		For i = 1 To Len(form_no.w_tread1.Text)
			gm_no = Val(Mid(form_no.w_tread1.Text, i, 1))
			pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

			'----- .NET 移行 -----
			'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
			w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

			w_ret = PokeACAD("GMCODE2", w_mess)
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    Else
		'        w_mess = ""
		'        w_ret = PokeACAD("HOLDGM2", w_mess)
		'    End If
		' <- watanabe Add 2007.03 -> ReEdit
		
		If w_w_n >= 2 Then
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        If chk_tread2.Value = 0 Then
			' <- watanabe Add 2007.03 -> ReEdit
			
			'[[ TREAD2 ]]
			For i = 1 To Len(form_no.w_tread2.Text)
				gm_no = Val(Mid(form_no.w_tread2.Text, i, 1))
				pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
				If pic_no < 1 Then GoTo error_section
				ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE3", w_mess)
			Next i
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        Else
			'            w_mess = ""
			'            w_ret = PokeACAD("HOLDGM3", w_mess)
			'        End If
			' <- watanabe Add 2007.03 -> ReEdit
			
		End If
		
		If w_w_n >= 3 Then
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        If chk_tread3.Value = 0 Then
			' <- watanabe Add 2007.03 -> ReEdit
			
			'[[ TREAD3 ]]
			For i = 1 To Len(form_no.w_tread3.Text)
				gm_no = Val(Mid(form_no.w_tread3.Text, i, 1))
				pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
				If pic_no < 1 Then GoTo error_section
				ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE4", w_mess)
			Next i
			
			' -> watanabe Add 2007.03 -> ReEdit
			'        Else
			'            w_mess = ""
			'            w_ret = PokeACAD("HOLDGM4", w_mess)
			'        End If
			' <- watanabe Add 2007.03 -> ReEdit
			
		End If
		
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    If chk_sidewall.Value = 0 Then
		' <- watanabe Add 2007.03 -> ReEdit
		
		w_w_gmcode = "GMCODE5"
		If w_w_n = 1 Then w_w_gmcode = "GMCODE3"
		If w_w_n = 2 Then w_w_gmcode = "GMCODE4"
		If w_w_n = 3 Then w_w_gmcode = "GMCODE5"
		
		'[[ SIDEWALL ]]
		For i = 1 To Len(form_no.w_sidewall.Text)
			gm_no = Val(Mid(form_no.w_sidewall.Text, i, 1))
			pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

			'----- .NET 移行 -----
			'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
			w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

			w_ret = PokeACAD(w_w_gmcode, w_mess)
		Next i
		
		' -> watanabe Add 2007.03 -> ReEdit
		'    Else
		'        w_w_gmcode = "HOLDGM5"
		'        If w_w_n = 1 Then w_w_gmcode = "HOLDGM3"
		'        If w_w_n = 2 Then w_w_gmcode = "HOLDGM4"
		'        If w_w_n = 3 Then w_w_gmcode = "HOLDGM5"
		'        w_mess = ""
		'        w_ret = PokeACAD(w_w_gmcode, w_mess)
		'    End If
		' <- watanabe Add 2007.03 -> ReEdit
		
		
		'// 終了の送信
        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG")

		'// 画面ロック
		form_no.Command2.Enabled = False
		form_no.Command4.Enabled = False
		form_no.Command6.Enabled = False
		form_no.w_tread1.Enabled = False
        form_no.w_tread1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		form_no.w_tread2.Enabled = False
        form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		form_no.w_tread3.Enabled = False
        form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		form_no.w_sidewall.Enabled = False
        form_no.w_sidewall.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		form_no.w_type.Enabled = False
        form_no.w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_PLY_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
		
		
		'フォント
		'(Brand Ver.3 追加)
        form_no.w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
                form_no.w_font.Items.Add(Tmp_font_word(i))
			End If
		Next i
		
		'タイプ
		'(Brand Ver.3 変更)
		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & Trim(Tmp_Ply1_ini)
		ret = set_read5(w_w_str, "ply1", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "PLY1" Then
				'             form_no.w_type.AddItem "POLYESTER+STEEL"
				'          ElseIf Tmp_hm_word(i) = "PLY2" Then
				'             form_no.w_type.AddItem "POLYESTER+STEEL+NYLON"
				'          ElseIf Tmp_hm_word(i) = "PLY3" Then
				'             form_no.w_type.AddItem "RAYON+STEEL"
				'          ElseIf Tmp_hm_word(i) = "PLY4" Then
				'             form_no.w_type.AddItem "RAYON+STEEL+NYLON"
				'          ElseIf Tmp_hm_word(i) = "PLY5" Then
				'             form_no.w_type.AddItem "NYLON"
				'          End If
				If Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    form_no.w_type.Items.Add(Tmp_hm_word(i))
				End If
			End If
		Next i
		
        Call Clear_F_TMP_PLY()

        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")

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
                w_w_str = Trim(w_w_str) & Trim(Tmp_Ply1_ini)
                ret = set_read5(w_w_str, "ply1", i)
                If ret = False Then
                    MsgBox(Tmp_Ply1_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Ply1_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'タイプ
		'(Brand Ver.3 変更)
		w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "PLY1" Then
				'             form_no.w_type.AddItem "POLYESTER+STEEL"
				'          ElseIf Tmp_hm_word(i) = "PLY2" Then
				'             form_no.w_type.AddItem "POLYESTER+STEEL+NYLON"
				'          ElseIf Tmp_hm_word(i) = "PLY3" Then
				'             form_no.w_type.AddItem "RAYON+STEEL"
				'          ElseIf Tmp_hm_word(i) = "PLY4" Then
				'             form_no.w_type.AddItem "RAYON+STEEL+NYLON"
				'          ElseIf Tmp_hm_word(i) = "PLY5" Then
				'             form_no.w_type.AddItem "NYLON"
				'          End If
				If Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    form_no.w_type.Items.Add(Tmp_hm_word(i))
				End If
			End If
		Next i
		
        form_no.w_type.Text = ""
        form_no.w_hm_name.Text = ""
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
        form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
		Dim w_file As Object
		Dim TiffFile As Object
		Dim w_text As Object

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        On Error Resume Next
		Err.Clear()
		
        w_text = w_hm_name.Text
		
        If Trim(w_text) = "" Then Exit Sub
		
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
            form_no.ImgThumbnail1.Width = 457 '20100701コード変更
            form_no.ImgThumbnail1.Height = 193 '20100701コード変更
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
        End If
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_sidewall.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_sidewall_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_sidewall.TextChanged
		
		Dim flg As Short

        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        Dim i As Short

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		If (w_sidewall.Text < "1") Or (w_sidewall.Text > "9") Then
			Exit Sub
		End If
		
		
		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "POLYESTER+STEEL" Then
		'       w_str = "PLY1"
		'   ElseIf w_type.Text = "POLYESTER+STEEL+NYLON" Then
		'       w_str = "PLY2"
		'   ElseIf w_type.Text = "RAYON+STEEL" Then
		'       w_str = "PLY3"
		'   ElseIf w_type.Text = "RAYON+STEEL+NYLON" Then
		'       w_str = "PLY4"
		'   ElseIf w_type.Text = "NYLON" Then
		'       w_str = "PLY5"
		'   End If
		
		flg = 0
        If form_no.w_sidewall.Text >= "2" Then
            '20000124 修正
            '      w_str = w_str & "S"
            flg = 1
        End If
		
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                '20000124 追加
                If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    Exit For
                    '20000124 追加
                ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    Exit For
                End If
            End If
		Next i
		
	End Sub
	
	Private Sub w_sidewall_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_sidewall.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_name.Text) <> "" Then
				w_hm_name.Text = ""
				'Brand Ver.5 TIFF->BMP 変更 start
				'         ImgThumbnail1.Image = ""
				ImgThumbnail1.Image = Nothing
				'Brand Ver.5 TIFF->BMP 変更 end
			End If
		End If
		
	End Sub
	
	Private Sub w_sidewall_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_sidewall.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		Dim flg As Short

        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        Dim i As Short
		
		
		If Trim(w_type.Text) = "" Then
			GoTo EventExitSub
		End If
		
		If (KeyAscii < 49) Or (KeyAscii > 57) Then
			GoTo EventExitSub
		End If
		
		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "POLYESTER+STEEL" Then
		'       w_str = "PLY1"
		'   ElseIf w_type.Text = "POLYESTER+STEEL+NYLON" Then
		'       w_str = "PLY2"
		'   ElseIf w_type.Text = "RAYON+STEEL" Then
		'       w_str = "PLY3"
		'   ElseIf w_type.Text = "RAYON+STEEL+NYLON" Then
		'       w_str = "PLY4"
		'   ElseIf w_type.Text = "NYLON" Then
		'       w_str = "PLY5"
		'   End If
		
		flg = 0
		If form_no.w_sidewall.Text >= "2" Then
			'20000124 修正
			'      w_str = w_str & "S"
			flg = 1
		End If
		
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
			If Tmp_hm_word(i) = form_no.w_type.Text Then
				'20000124 追加
				If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
					form_no.w_hm_name.Text = Tmp_hm_code(i)
					Exit For
					'20000124 追加
				ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then 
					form_no.w_hm_name.Text = Tmp_hm_code(i)
					Exit For
				End If
			End If
		Next i
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_sidewall_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_sidewall.Leave
		
		If (Val(w_sidewall.Text) < 1) Or (Val(w_sidewall.Text) > 9) Then
			If w_hm_name.Text <> "" Then
				w_hm_name.Text = ""
				'Brand Ver.5 TIFF->BMP 変更 start
				'         ImgThumbnail1.Image = ""
				ImgThumbnail1.Image = Nothing
				'Brand Ver.5 TIFF->BMP 変更 end
			End If
		End If
		
	End Sub
	
	
    'UPGRADE_WARNING: イベント w_type.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.TextChanged
		
		Dim flg As Short
		Dim wstr As String
		Dim i As Short

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "POLYESTER+STEEL" Then
		'       w_str = "PLY1"
		'   ElseIf w_type.Text = "POLYESTER+STEEL+NYLON" Then
		'       w_str = "PLY2"
		'   ElseIf w_type.Text = "RAYON+STEEL" Then
		'       w_str = "PLY3"
		'   ElseIf w_type.Text = "RAYON+STEEL+NYLON" Then
		'       w_str = "PLY4"
		'   ElseIf w_type.Text = "NYLON" Then
		'       w_str = "PLY5"
		'   End If
		
		
		flg = 0
		If form_no.w_sidewall.Text >= "2" Then
			'20000124 修正
			'      w_str = w_str & "S"
			flg = 1
		End If
		
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                '20000124 追加
                If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    wstr = Tmp_prcs_code(i)
                    Exit For
                    '20000124 追加
                ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    wstr = Tmp_prcs_code(i)
                    Exit For
                End If
            End If
		Next i
		
		
		'20000124 修正
		'   If (w_type.Text = "POLYESTER+STEEL") Or (w_type.Text = "RAYON+STEEL") Then
		'      form_no.w_tread3.Enabled = False
		'      form_no.w_tread3.BackColor = &HC0C0C0
		'   ElseIf w_type.Text = "NYLON" Then
		'      form_no.w_tread2.Enabled = False
		'      form_no.w_tread2.BackColor = &HC0C0C0
		'      form_no.w_tread3.Enabled = False
		'      form_no.w_tread3.BackColor = &HC0C0C0
		'   Else
		'      form_no.w_tread2.Enabled = True
		'      form_no.w_tread2.BackColor = &H80000005
		'      form_no.w_tread3.Enabled = True
		'      form_no.w_tread3.BackColor = &H80000005
		'   End If
		
		If (Trim(wstr) = "PLY1") Or (Trim(wstr) = "PLY1S") Then
            form_no.w_tread2.Enabled = False
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_tread3.Enabled = False
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = False
			'        form_no.chk_tread3.Enabled = False
			'        form_no.chk_tread2.Value = 2
			'        form_no.chk_tread3.Value = 2
			' <- watanabe add 2007.03 -> del
			
		ElseIf (Trim(wstr) = "PLY2") Or (Trim(wstr) = "PLY2S") Then 
            form_no.w_tread2.Enabled = True
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
            form_no.w_tread3.Enabled = False
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = True
			'        form_no.chk_tread3.Enabled = False
			'        form_no.chk_tread3.Value = 2
			' <- watanabe add 2007.03 -> del
			
		Else
            form_no.w_tread2.Enabled = True
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
            form_no.w_tread3.Enabled = True
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = True
			'        form_no.chk_tread3.Enabled = True
			' <- watanabe add 2007.03 -> del
			
		End If
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
		
		Dim flg As Short
		Dim wstr As String
		Dim i As Short

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "POLYESTER+STEEL" Then
		'       w_str = "PLY1"
		'   ElseIf w_type.Text = "POLYESTER+STEEL+NYLON" Then
		'       w_str = "PLY2"
		'   ElseIf w_type.Text = "RAYON+STEEL" Then
		'       w_str = "PLY3"
		'   ElseIf w_type.Text = "RAYON+STEEL+NYLON" Then
		'       w_str = "PLY4"
		'   ElseIf w_type.Text = "NYLON" Then
		'       w_str = "PLY5"
		'   End If
		
		
		flg = 0
        If form_no.w_sidewall.Text >= "2" Then
            '20000124 修正
            '      w_str = w_str & "S"
            flg = 1
        End If
		
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                '20000124 追加
                If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    wstr = Tmp_prcs_code(i)
                    Exit For
                    '20000124 追加
                ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then
                    form_no.w_hm_name.Text = Tmp_hm_code(i)
                    wstr = Tmp_prcs_code(i)
                    Exit For
                End If
            End If
		Next i
		
		
		'20000124 修正
		'   If (w_type.Text = "POLYESTER+STEEL") Or (w_type.Text = "RAYON+STEEL") Then
		'      form_no.w_tread3.Enabled = False
		'      form_no.w_tread3.BackColor = &HC0C0C0
		'   ElseIf w_type.Text = "NYLON" Then
		'      form_no.w_tread2.Enabled = False
		'      form_no.w_tread2.BackColor = &HC0C0C0
		'      form_no.w_tread3.Enabled = False
		'      form_no.w_tread3.BackColor = &HC0C0C0
		'   Else
		'      form_no.w_tread2.Enabled = True
		'      form_no.w_tread2.BackColor = &H80000005
		'      form_no.w_tread3.Enabled = True
		'      form_no.w_tread3.BackColor = &H80000005
		'   End If
		
		If (Trim(wstr) = "PLY1") Or (Trim(wstr) = "PLY1S") Then
            form_no.w_tread2.Enabled = False
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_tread3.Enabled = False
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = False
			'        form_no.chk_tread3.Enabled = False
			'        form_no.chk_tread2.Value = 2
			'        form_no.chk_tread3.Value = 2
			' <- watanabe add 2007.03 -> del
			
		ElseIf (Trim(wstr) = "PLY2") Or (Trim(wstr) = "PLY2S") Then 
            form_no.w_tread2.Enabled = True
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
            form_no.w_tread3.Enabled = False
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = True
			'        form_no.chk_tread3.Enabled = False
			'        form_no.chk_tread3.Value = 2
			' <- watanabe add 2007.03 -> del
			
		Else
            form_no.w_tread2.Enabled = True
            form_no.w_tread2.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
            form_no.w_tread3.Enabled = True
            form_no.w_tread3.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629コード変更
			
			' -> watanabe add 2007.03 -> del
			'        form_no.chk_tread2.Enabled = True
			'        form_no.chk_tread3.Enabled = True
			' <- watanabe add 2007.03 -> del
			
		End If
		
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