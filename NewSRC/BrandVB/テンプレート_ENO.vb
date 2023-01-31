Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_ENO
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_ENO()
		
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
                .HelpContext = 804
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
		Dim w_mess As String
		Dim i As Short
		
        form_no.w_shonin.Text = Trim(form_no.w_shonin.Text)
		
		'/* 入力チェック */
		If check_F_TMP_ENO <> 0 Then Exit Sub
		
		form_no.Enabled = False
		F_MSG.Show()
		
		
		' -> watanabe Add 2007.03
		If chk_shonin.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'(Brand Ver.3 追加)
			For i = 1 To Len(form_no.w_shonin.Text)
				w_str = Mid(form_no.w_shonin.Text, i, 1)
				If IsNumeric(w_str) Then
					If Val(w_str) >= 0 And Val(w_str) < 10 Then
						If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input Load index is not set to the configuration file (" & Tmp_E_no1_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
					If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input ply  is not set to the configuration file (" & Tmp_E_no1_ini & ")", 64, "Configuration file error")
						GoTo error_section
					End If
				End If
			Next i
			
			' -> watanabe Add 2007.03
		End If
		' <- watanabe Add 2007.03

        ' 2011/12/08 uriu added start
        'S番号
        If w_type.SelectedIndex >= 5 And w_type.SelectedIndex < 11 Then
            If chk_s.CheckState = 0 Then
                w_str = w_s.Text
                If IsNumeric(w_str) Then
                    If Val(w_str) >= 0 And Val(w_str) < 10 Then
                        If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input S number is not set to the configuration file (" & Tmp_E_no1_ini & ")", 64, "Configuration file error")
                            GoTo error_section
                        End If
                    End If
                Else
                    MsgBox("Please input numerical value into S number.", 64, "Input error")
                    GoTo error_section
                End If
            End If
        End If
        'R番号
        If w_type.SelectedIndex = 5 Or w_type.SelectedIndex = 6 Then
            If chk_r.CheckState = 0 Then
                w_str = w_r.Text
                If IsNumeric(w_str) Then
                    If Val(w_str) >= 0 And Val(w_str) < 10 Then
                        If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input R number is not set to the configuration file (" & Tmp_E_no1_ini & ")", 64, "Configuration file error")
                            GoTo error_section
                        End If
                    End If
                Else
                    MsgBox("Please input a numerical value into R number.", 64, "Input error")
                    GoTo error_section
                End If
            End If
        End If
        ' 2011/12/08 uriu added end

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
		If chk_shonin.CheckState = 0 Then
			' <- watanabe Add 2007.03
			
			'[[ 承認番号 ]]
			For i = 1 To Len(form_no.w_shonin.Text)
				gm_no = Val(Mid(form_no.w_shonin.Text, i, 1))
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
		' <- watanabe Add 2007.03

        ' 2011/12/08 uriu added start
        'S番号
        If chk_s.CheckState = 0 Then
            If w_s.Text <> "" Then
                gm_no = Val(w_s.Text)
                pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                If pic_no < 1 Then GoTo error_section
                ZumenName = "GM-" & Mid$(GensiNUM(gm_no), 1, 6)

                ' -> watanabe edit 2013.05.29
                'w_mess = Format(Val(pic_no), "00") & GensiDir & ZumenName
                w_mess = Format(Val(pic_no), "000") & GensiDir & ZumenName
                ' <- watanabe edit 2013.05.29

                w_ret = PokeACAD("GMCODE2", w_mess)
            End If
        Else
            w_mess = ""
            w_ret = PokeACAD("HOLDGM2", w_mess)
        End If

        'R番号
        If chk_r.CheckState = 0 Then
            If w_r.Text <> "" Then
                gm_no = Val(w_r.Text)
                pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                If pic_no < 1 Then GoTo error_section
                ZumenName = "GM-" & Mid$(GensiNUM(gm_no), 1, 6)

                ' -> watanabe edit 2013.05.29
                'w_mess = Format(Val(pic_no), "00") & GensiDir & ZumenName
                w_mess = Format(Val(pic_no), "000") & GensiDir & ZumenName
                ' <- watanabe edit 2013.05.29

                w_ret = PokeACAD("GMCODE3", w_mess)
            End If
        Else
            w_mess = ""
            w_ret = PokeACAD("HOLDGM3", w_mess)
        End If
        ' 2011/12/08 uriu added end

        '// 終了の送信
        CommunicateMode = comNone
		w_ret = RequestACAD("TMPCHANG")
		
		'// 画面ロック
		form_no.Command2.Enabled = False
		form_no.Command4.Enabled = False
		form_no.Command6.Enabled = False
		form_no.w_type.Enabled = False
        form_no.w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		form_no.w_shonin.Enabled = False
        form_no.w_shonin.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_ENO_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
                form_no.w_font.Items.Add(Tmp_font_word(i)) '20100624コード変更
			End If
		Next i
		
		'タイプ
		'(Brand Ver.3 変更)
		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & Trim(Tmp_E_no1_ini)
		ret = set_read5(w_w_str, "e_no1", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i)) '20100624コード変更
			End If
		Next i
		
        Call Clear_F_TMP_ENO()

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
                w_w_str = Trim(w_w_str) & Trim(Tmp_E_no1_ini)
                ret = set_read5(w_w_str, "e_no1", i)
                If ret = False Then
                    MsgBox(Tmp_E_no1_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_E_no1_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'タイプ
		'(Brand Ver.3 変更)
		w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		w_type.Text = ""
		w_hm_name.Text = ""
		'Brand Ver.5 TIFF->BMP 変更 start
		'   ImgThumbnail1.Image = ""
		ImgThumbnail1.Image = Nothing
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
        If Trim(w_text) = "" Then
            'Brand Ver.5 TIFF->BMP 変更 start
            '       form_no.ImgThumbnail1.Image = ""
            form_no.ImgThumbnail1.Image = Nothing
            'Brand Ver.5 TIFF->BMP 変更 end
            Exit Sub
        End If
		
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
		
		'Tiffﾌｧｲﾙ表示
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
	
    Private Sub w_shonin_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        form_no.w_shonin.Text = UCase(Trim(form_no.w_shonin.Text))

    End Sub
	
    'UPGRADE_WARNING: イベント w_type.TextChanged は、フォームが初期化されたときに発生します。
    Private Sub w_type_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.TextChanged
        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        If Trim(w_type.Text) = "" Then
            form_no.w_hm_name.Text = ""
            'Brand Ver.5 TIFF->BMP 変更 start
            '      form_no.ImgThumbnail1.Image = ""
            form_no.ImgThumbnail1.Image = Nothing
            'Brand Ver.5 TIFF->BMP 変更 end
        End If

    End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
		
        Dim i As Short
        Dim prcs_code As String
		
		'(Brand Ver.3 追加)
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = w_type.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                prcs_code = Tmp_prcs_code(i)
				Exit For
			End If
		Next i
		
        'If w_type.SelectedIndex >= 5 And w_type.SelectedIndex < 11 Then
        '    w_s.BackColor = SystemColors.Window
        '    w_s.Enabled = True
        '    chk_s.Enabled = True
        'Else
        '    w_s.BackColor = SystemColors.Control
        '    w_s.Enabled = False
        '    chk_s.Enabled = False
        'End If

        'If w_type.SelectedIndex = 5 Or w_type.SelectedIndex = 6 Then
        '    w_r.BackColor = SystemColors.Window
        '    w_r.Enabled = True
        '    chk_r.Enabled = True
        'Else
        '    w_r.BackColor = SystemColors.Control
        '    w_r.Enabled = False
        '    chk_r.Enabled = False
        'End If

        '2014/4/11 uriu changed
        Select Case w_type.SelectedIndex
            Case 0, 1, 2, 3, 4, 11, 12, 13, 14
                'SW入力なし
                w_s.BackColor = SystemColors.Control
                w_s.Enabled = False
                chk_s.Enabled = False
                w_r.BackColor = SystemColors.Control
                w_r.Enabled = False
                chk_r.Enabled = False
            Case 5, 6, 15, 16
                'SW入力あり
                w_s.BackColor = SystemColors.Window
                w_s.Enabled = True
                chk_s.Enabled = True
                w_r.BackColor = SystemColors.Window
                w_r.Enabled = True
                chk_r.Enabled = True
            Case 7, 8, 9, 10
                'Sのみ入力あり
                w_s.BackColor = SystemColors.Window
                w_s.Enabled = True
                chk_s.Enabled = True
                w_r.BackColor = SystemColors.Control
                w_r.Enabled = False
                chk_r.Enabled = False
        End Select

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