Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_ETC
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_ETC()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        InitFlag = False '20100628追加コード
		form_no.Close()
		End
		
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        'Dim CommonDialog1 As Object'20100616移植削除
		
        'With CommonDialog1
        '    .HelpCommand = MSComDlg.HelpConstants.cdlHelpContext
        '    .HelpFile = HelpFileName
        '    .HelpContext = 811
        '    .ShowHelp()
        'End With

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
		Dim gm_no As Object
		Dim ZumenName As Object
		Dim pic_no As Object
		Dim w_ret As Object
		Dim j As Object
		Dim i As Object
		
		Dim w_mess As String
		Dim w_w_n As Short
		Dim w_w_gmcode As String
		Dim w_cmd As String
		Dim w_str As String
		
		'/* 入力チェック */
		w_w_n = 0
		
		For i = 1 To MaxSelNum
            If (w_type.Text = Tmp_hm_word(i)) Then
                If (Tmp_prcs_code(i) = "ETC1") Then
                    If (w_etc(1).Text = "") Then
                        MsgBox("Input error.")
                        Exit Sub
                    End If
                    w_w_n = 1
                ElseIf (Tmp_prcs_code(i) = "ETC2") Then
                    For j = 1 To 2
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 2
                ElseIf (Tmp_prcs_code(i) = "ETC3") Then
                    For j = 1 To 3
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 3
                ElseIf (Tmp_prcs_code(i) = "ETC4") Then
                    For j = 1 To 4
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 4
                ElseIf (Tmp_prcs_code(i) = "ETC5") Then
                    For j = 1 To 5
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 5
                ElseIf (Tmp_prcs_code(i) = "ETC6") Then
                    For j = 1 To 6
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 6
                ElseIf (Tmp_prcs_code(i) = "ETC7") Then
                    For j = 1 To 7
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 7
                ElseIf (Tmp_prcs_code(i) = "ETC8") Then
                    For j = 1 To 8
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 8
                ElseIf (Tmp_prcs_code(i) = "ETC9") Then
                    For j = 1 To 9
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
                            Exit Sub
                        End If
                    Next j
                    w_w_n = 9
                ElseIf (Tmp_prcs_code(i) = "ETC10") Then
                    For j = 1 To 10
                        If w_etc(j).Text = "" Then
                            MsgBox("Input error.")
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
		
		
		'(Brand Ver.3 追加)
		For j = 1 To w_w_n
            For i = 1 To Len(form_no.w_etc(j).Text)
                w_str = Mid(form_no.w_etc(j).Text, i, 1)
                If IsNumeric(w_str) Then
                    If Val(w_str) >= 0 And Val(w_str) < 10 Then
                        If GensiNUM(Val(w_str)) = "" Then
                            MsgBox("A substituted primitive letter for input " & w_str & " is not set to the configuration file (" & Tmp_ETC_ini & ")", 64, "Configuration file error")
                            GoTo error_section
                        End If
                    End If
                ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                    If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & " is not set to the configuration file (" & Tmp_ETC_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                ElseIf Asc("a") <= Asc(w_str) And Asc(w_str) <= Asc("z") Then
                    If GensiALPHS(Asc(w_str) - Asc("a")) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & " is not set to the configuration file (" & Tmp_ETC_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                ElseIf 33 <= Asc(w_str) And Asc(w_str) <= 126 Then
                    If GensiKIGO(Asc(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input " & w_str & " is not set to the configuration file (" & Tmp_ETC_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            Next i
		Next j
		
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
		
		'[[ TYPE(1～10) ]]
		For j = 1 To w_w_n
            For i = 1 To Len(form_no.w_etc(j).Text)
                w_str = Mid(form_no.w_etc(j).Text, i, 1)
                If IsNumeric(w_str) Then
                    gm_no = Val(w_str)
                    pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                    If pic_no < 1 Then GoTo error_section
                    ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                    '----- .NET 移行 -----
                    'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                    w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                    w_cmd = "GMCODE" & j
                    w_ret = PokeACAD(w_cmd, w_mess)
                ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                    gm_no = Asc(w_str) - Asc("A")
                    pic_no = what_pic_from_gmcode(GensiALPH(gm_no))
                    If pic_no < 1 Then GoTo error_section
                    ZumenName = "GM-" & Mid(GensiALPH(gm_no), 1, 6)

                    '----- .NET 移行 -----
                    'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                    w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                    w_cmd = "GMCODE" & j
                    w_ret = PokeACAD(w_cmd, w_mess)
                ElseIf Asc("a") <= Asc(w_str) And Asc(w_str) <= Asc("z") Then
                    gm_no = Asc(w_str) - Asc("a")
                    pic_no = what_pic_from_gmcode(GensiALPHS(gm_no))
                    If pic_no < 1 Then GoTo error_section
                    ZumenName = "GM-" & Mid(GensiALPHS(gm_no), 1, 6)

                    '----- .NET 移行 -----
                    'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                    w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                    w_cmd = "GMCODE" & j
                    w_ret = PokeACAD(w_cmd, w_mess)
                ElseIf 33 <= Asc(w_str) And Asc(w_str) <= 126 Then
                    gm_no = Asc(w_str)
                    pic_no = what_pic_from_gmcode(GensiKIGO(gm_no))
                    If pic_no < 1 Then GoTo error_section
                    ZumenName = "GM-" & Mid(GensiKIGO(gm_no), 1, 6)

                    '----- .NET 移行 -----
                    'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                    w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                    w_cmd = "GMCODE" & j
                    w_ret = PokeACAD(w_cmd, w_mess)
                End If
            Next i
		Next j
		
		'// 終了の送信
        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG")

		'// 画面ロック
        form_no.Command2.Enabled = False
        form_no.Command4.Enabled = False
        form_no.Command6.Enabled = False
        form_no.w_type.Enabled = False
        form_no.w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		For i = 1 To 10
            form_no.w_etc(i).Enabled = False
            form_no.w_etc(i).BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		Next i
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_ETC_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_ETC_ini)
        ret = set_read5(w_w_str, "etc", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_ETC()
		
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

		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_ETC_ini)
                ret = set_read5(w_w_str, "etc", i)
                If ret = False Then
                    MsgBox(Tmp_ETC_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_ETC_ini & ")", 64, "Configuration file error")
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

		On Error Resume Next ' エラーのトラップを留保します。
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
            form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
            form_no.ImgThumbnail1.Height = 193 '200 '20100701コード変更
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
        End If
		'Brand Ver.5 TIFF->BMP 変更 end
		
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