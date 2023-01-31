Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_UTQG
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_UTQG()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		
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
                .HelpContext = 805
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim gm_alph As Object
		Dim gm_no As Object
		Dim ZumenName As Object
		Dim pic_no As Object
		Dim w_ret As Object
		Dim i As Object
		
		Dim w_mess As String
		Dim w_str As String
		
		'/* 入力チェック */
        form_no.w_treadwear.Text = Trim(form_no.w_treadwear.Text)
		form_no.w_traction.Text = Trim(form_no.w_traction.Text)
		form_no.w_temperature.Text = Trim(form_no.w_temperature.Text)
		
		If check_F_TMP_UTQG <> 0 Then Exit Sub
		
		form_no.Enabled = False
		F_MSG.Show()
		
		' -> watanabe Add 2007.03
		If chk_treadwear.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'(Brand Ver.3 追加)
			For i = 1 To Len(form_no.w_treadwear.Text)
				w_str = Mid(form_no.w_treadwear.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input TREADWEAR is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input TREADWEAR  is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03
		End If
		
		If chk_traction.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			For i = 1 To Len(form_no.w_traction.Text)
				w_str = Mid(form_no.w_traction.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＲＡＣＴＩＯＮ is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input ＴＲＡＣＴＩＯＮ is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03
		End If
		
		If chk_temperature.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			For i = 1 To Len(form_no.w_temperature.Text)
				w_str = Mid(form_no.w_temperature.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input ＴＥＭＰＥＲＡＴＵＲＥ is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input TEMPERATURE  is not set to the configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03
		End If
		' <- watanabe Add 2007.03
		
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
		
		' -> watanabe Add 2007.03
		If chk_treadwear.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'[[ TREADWEAR ]]
			For i = 1 To Len(form_no.w_treadwear.Text)
				gm_no = Val(Mid(form_no.w_treadwear.Text, i, 1))
				pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
				If pic_no < 1 Then GoTo error_section
				ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE1", w_mess)
			Next i
			
			' -> watanabe Add 2007.03
		Else
			w_mess = ""
			w_ret = PokeACAD("HOLDGM1", w_mess)
		End If
		
		If chk_traction.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'[[ TRACTION ]]
			For i = 1 To Len(form_no.w_traction.Text)
				gm_alph = Mid(form_no.w_traction.Text, i, 1)
				pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
				If pic_no < 1 Then GoTo error_section
				ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE2", w_mess)
			Next i
			
			' -> watanabe Add 2007.03
		Else
			w_mess = ""
			w_ret = PokeACAD("HOLDGM2", w_mess)
		End If
		
		If chk_temperature.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'[[ TEMPERATURE ]]
			For i = 1 To Len(form_no.w_temperature.Text)
				gm_alph = Mid(form_no.w_temperature.Text, i, 1)
				pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
				If pic_no < 1 Then GoTo error_section
				ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE3", w_mess)
			Next i
			
			' -> watanabe Add 2007.03
		Else
			w_mess = ""
			w_ret = PokeACAD("HOLDGM3", w_mess)
		End If
		' <- watanabe Add 2007.03
		
		'// 終了の送信
		w_ret = RequestACAD("TMPCHANG")
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_UTQG_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_ret As Object
		Dim ret As Object
		
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
		w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg1_ini)
		ret = set_read5(w_w_str, "utqg1", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "U1" Then
				'             form_no.w_type.AddItem "一段"
				'          ElseIf Tmp_hm_word(i) = "U2" Then
				'             form_no.w_type.AddItem "二段"
				'          End If
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		
		Call Clear_F_TMP_UTQG()
		
		form_main.Text2.Text = ""
		CommunicateMode = comFreePic
		w_ret = RequestACAD("PICEMPTY")
		
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
				w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg1_ini)
				ret = set_read5(w_w_str, "utqg1", i)
				If ret = False Then
                    MsgBox(Tmp_Utqg1_ini & "File reading error.", 64, "BrandVB error")
					Exit Sub
				Else
					read_flg = 1
					Exit For
				End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Utqg1_ini & ")", 64, "Configuration file error")
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
				'          If Tmp_hm_word(i) = "U1" Then
				'             form_no.w_type.AddItem "一段"
				'          ElseIf Tmp_hm_word(i) = "U2" Then
				'             form_no.w_type.AddItem "二段"
				'          End If
                form_no.w_type.Items.Add(Tmp_hm_word(i))
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
        Dim w_file As String '20100708 修正
        Dim TiffFile As String
        Dim w_text As String
		
        On Error Resume Next
		Err.Clear()
		
        w_text = form_no.w_hm_name.Text
		
        If Trim(w_text) = "" Then Exit Sub
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'       TiffFile = TIFFDir & form_no.w_hm_name.Text & ".tif"
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
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Private Sub w_temperature_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_temperature.Leave
		
		form_no.w_temperature.Text = UCase(Trim(form_no.w_temperature.Text))
		
	End Sub
	
	Private Sub w_traction_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_traction.Leave
		
		form_no.w_traction.Text = UCase(Trim(form_no.w_traction.Text))
		
	End Sub
	
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
		Dim i As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "一段" Then
		'       w_str = "U1"
		'   ElseIf w_type.Text = "二段" Then
		'       w_str = "U2"
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