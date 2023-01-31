Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_PTNCODE
    Inherits System.Windows.Forms.Form

	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		w_type.SelectedIndex = 0
		w_ptncode.Text = ""
		
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
                .HelpContext = 809
                .ShowHelp()
            End With
        End If
	End Sub
	
    Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
        Dim gm_no As Object
        Dim ZumenName As String
        Dim pic_no As Object
        Dim w_ret As Object
        Dim i As Object

        Dim w_mess As String
        Dim w_str As String

        w_ptncode.Text = Trim(w_ptncode.Text)
        '/ 入力チェック /
        If check_F_TMP_PTNCODE() <> 0 Then Exit Sub
        form_no.Enabled = False
        F_MSG.Show()

        '(Brand Ver.3 追加)
        For i = 2 To Len(form_no.w_ptncode.Text)
            w_str = Mid(form_no.w_ptncode.Text, i, 1)
            If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input Pattern code is not set to the configuration file (" & Tmp_Pattern1_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf w_str = "+" Then
                If GensiKIGO(Asc(w_str)) = "" Then
                    MsgBox("A substituted primitive letter for input Pattern code is not set to the configuration file (" & Tmp_Pattern1_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
            ElseIf w_str = "-" Then
                If GensiKIGO(Asc(w_str)) = "" Then
                    MsgBox("A substituted primitive letter for input Pattern code is not set to the configuration file (" & Tmp_Pattern1_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for input Pattern code is not set to the configuration file (" & Tmp_Pattern1_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
            End If
        Next i


        If FreePicNum < 1 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
            GoTo error_section
        End If


       
        w_mess = "" '初期化
        'temp_bz.pattern = LSet(Trim(w_ptncode.Text),Len(temp_bz.pattern))
        temp_bz.pattern = LSet(Trim(w_ptncode.Text), 6) '20100708コード変更　エラー回避策
        Call bz_spec_set(w_mess)
        w_ret = PokeACAD("SPECADD", w_mess)
        w_ret = RequestACAD("SPECADD")

        '// 置換モードの送信
        w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
        w_ret = RequestACAD("CHNGMODE")

        '（図面名）送信
        pic_no = what_pic_from_hmcode(form_no.w_hm_name.Text)
        If pic_no < 1 Then
            MsgBox("From the database, you did not get to edit letter data.", 64, "SQL error")
            GoTo error_section
        End If
        ZumenName = "HM-" & VB.Left(Trim(form_no.w_hm_name.Text), 6)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
        w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("HMCODE", w_mess)

        '[[ パターンコード ]]
        For i = 2 To Len(form_no.w_ptncode.Text)
            w_str = Mid(form_no.w_ptncode.Text, i, 1)

            '----- 3/6 1998 yamamoto update -----
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
            ElseIf w_str = "+" Then
                pic_no = what_pic_from_gmcode(GensiKIGO(Asc(w_str)))
                If pic_no < 1 Then
                    MsgBox("From the database, it was not possible to get data primitive letter (" & GensiKIGO(Asc(w_str)) & ")", 64, "SQL error")
                    GoTo error_section
                End If
                ZumenName = "GM-" & Mid(GensiKIGO(Asc(w_str)), 1, 6)

                '----- .NET 移行 -----
                'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                w_ret = PokeACAD("GMCODE1", w_mess)

            ElseIf w_str = "-" Then
                pic_no = what_pic_from_gmcode(GensiKIGO(Asc(w_str)))
                If pic_no < 1 Then
                    MsgBox("From the database, it was not possible to get data primitive letter (" & GensiKIGO(Asc(w_str)) & ")", 64, "SQL error")
                    GoTo error_section
                End If
                ZumenName = "GM-" & Mid(GensiKIGO(Asc(w_str)), 1, 6)

                '----- .NET 移行 -----
                'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
                w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

                w_ret = PokeACAD("GMCODE1", w_mess)

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
        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG")


        '// 画面ロック
        form_no.Command2.Enabled = False
        form_no.Command4.Enabled = False
        form_no.Command6.Enabled = False
        form_no.w_type.Enabled = False
        'form_no.w_type.BackColor = &HC0C0C0
        form_no.w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.w_ptncode.Enabled = False
        'form_no.w_ptncode.BackColor = &HC0C0C0
        form_no.w_ptncode.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

        Exit Sub

error_section:
        MsgBox("Finished in ErrorSection.")
        On Error Resume Next
        F_MSG.Close()
        form_no.Enabled = True

    End Sub
	
	Private Sub F_TMP_PTNCODE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        ' -> watanabe del VerUP(2011)
        'Dim commnad6 As Object
        'Dim commnad4 As Object
        'Dim commnad2 As Object
        ' <- watanabe del VerUP(2011)

        Dim error_no As Object
		Dim ret As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_ret As Short
		Dim w_w_str As String
		Dim time_start As Date
		Dim time_now As Date
		Dim w_msg As String
		Dim i As Short
		
        form_no = Me
        temp_bz.Initilize() '20100708追加コード

		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		'タイプ
		'(Brand Ver.3 変更)
		w_type.Items.Clear()
		For i = 1 To Tmp_font_cnt + 1
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
				w_type.Items.Add(Tmp_font_word(i))
			End If
		Next i
		
        w_type.SelectedIndex = 0 '20100629コード保留
        w_ptncode.Text = ""
        w_hm_name.Text = ""
		
		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & Trim(Tmp_Pattern1_ini)
		ret = set_read3(w_w_str, "pattern1", 1)
		
        form_main.Text2.Text = ""
        CommunicateMode = comPTNCODE
        w_ret = RequestACAD("PICEMPTY")

		time_start = Now
		Do 
			time_now = Now
            If Trim(form_main.Text2.Text) = "" Then
                If System.DateTime.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.DateTime.FromOADate(timeOutSecond) Then
                    MsgBox("Time-out error.", 64, "ERROR")
                    w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                    w_ret = RequestACAD("ERROR")
                    GoTo communicate_err_section
                End If
            ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                MsgBox("Communication error.", MsgBoxStyle.Critical, "Communicate Error")
                GoTo communicate_err_section

            ElseIf (VB.Left(form_main.Text2.Text, 8) = "PICEMPTY") Then

                ' -> watanabe edit 2013.06.03
                'FreePicNum = Val(Mid(form_main.Text2.Text, 9, 2))
                'If FreePicNum > 50 Then FreePicNum = 50
                FreePicNum = Val(Mid(form_main.Text2.Text, 9, 3))
                If FreePicNum > 130 Then FreePicNum = 130
                ' <- watanabe edit 2013.06.03

                form_main.Text2.Text = ""
                GoTo LOOP_EXIT
            Else
                w_msg = form_main.Text2.Text
                MsgBox("Not a free picture information." & w_msg)
                GoTo communicate_err_section

            End If

			System.Windows.Forms.Application.DoEvents()
		Loop 



LOOP_EXIT: 
        InitFlag = True '20100628追加コード

        CommunicateMode = comSpecData
        RequestACAD("SPECDATA")

		Exit Sub
		
communicate_err_section:
        ' -> watanabe edit VerUP(2011)
        'commnad2.Enabled = False
        'commnad4.Enabled = False
        'commnad6.Enabled = False
        Command2.Enabled = False
        Command4.Enabled = False
        Command6.Enabled = False
        ' <- watanabe edit VerUP(2011)

        w_type.Enabled = False
        w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        w_ptncode.Enabled = False
        w_ptncode.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
		Dim w_file As Object
        Dim TiffFile As String
        Dim w_text As String

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
        '    TiffFile = TIFFDir & w_hm_name.Text & ".tif"
        '
        '    'Tiffﾌｧｲﾙ表示
        '    w_file = Dir(TiffFile)
        '    If w_file <> "" Then
        '        form_no.ImgThumbnail1.Image = TiffFile
        '        form_no.ImgThumbnail1.ThumbWidth = 500
        '        form_no.ImgThumbnail1.ThumbHeight = 200
        '    Else
        '        MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
        '        form_no.ImgThumbnail1.Image = ""
        '    End If
        TiffFile = TIFFDir & w_hm_name.Text & ".bmp"

        'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
            form_no.ImgThumbnail1.Height = 193 '200 '20100701コード変更
        Else
            MsgBox("BMPﾌｧｲﾙが見つかりません", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
        End If
        'Brand Ver.5 TIFF->BMP 変更 end

        form_no.w_ptncode.Focus() 'SetFocus()

    End Sub
	
    'UPGRADE_WARNING: イベント w_ptncode.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_ptncode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_ptncode.TextChanged
		Dim i As Object

        ' -> watanabe del VerUP(2011)
        'Dim TxtNo As Short
        ' <- watanabe del VerUP(2011)

        Dim mach_flg As Short

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        If Len(Trim(form_no.w_ptncode.Text)) = 1 Then
            form_no.w_ptncode.Text = UCase(Trim(form_no.w_ptncode.Text))
            'form_no.w_ptncode.SelStart = Len(Trim(form_no.w_ptncode.Text))
            form_no.w_ptncode.SelectionStart = Len(Trim(form_no.w_ptncode.Text)) '20100628コード変更
            'form_no.w_ptncode.SelLength = Len(Trim(form_no.w_ptncode.Text))
            form_no.w_ptncode.SelectionLength = Len(Trim(form_no.w_ptncode.Text)) '20100628コード変更
        End If
		
		
		If Trim(w_ptncode.Text) = "" Then
			form_no.w_hm_name.Text = ""
			'Brand Ver.5 TIFF->BMP 変更 start
			'    form_no.ImgThumbnail1.Image = ""
            form_no.ImgThumbnail1.Image = Nothing
			'Brand Ver.5 TIFF->BMP 変更 end
			Exit Sub
		End If
		
		mach_flg = 0
		
		'(Brand Ver.3 変更)
		If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", VB.Left(w_ptncode.Text, 1), 0) > 0 Then
			For i = 1 To MaxSelNum
                If Trim(Tmp_hm_word(i)) = "" Then
                    Exit For
                ElseIf Trim(Tmp_hm_word(i)) = VB.Left(Trim(w_ptncode.Text), 1) Then
                    w_hm_name.Text = Tmp_hm_code(i)
                    mach_flg = 1
                    Exit For
                End If
			Next i
			
		ElseIf InStr(1, "0123456789", VB.Left(w_ptncode.Text, 1), 0) > 0 Then 
			For i = 1 To MaxSelNum
                If Trim(Tmp_hm_word(i)) = "" Then
                    Exit For
                ElseIf Trim(Tmp_hm_word(i)) = VB.Left(Trim(w_ptncode.Text), 1) Then
                    w_hm_name.Text = Tmp_hm_code(i)
                    mach_flg = 1
                    Exit For
                End If
			Next i
			
		Else
			MsgBox("入力が間違っています", 64, "パターンコードエラー")
			Exit Sub
		End If
		
		If mach_flg = 0 Then
            MsgBox("入力されたパターンの編集文字は、設定ファイル(" & Tmp_Pattern1_ini & ")に設定されていません", 64, "パターンコードエラー")
		End If
		
	End Sub
	
	Private Sub w_ptncode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_ptncode.Leave
		
		form_no.w_ptncode.Text = UCase(Trim(form_no.w_ptncode.Text))
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
		Dim ret As Object
		
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String
		
		'(Brand Cad System Ver.3 UP)
		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_type.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_Pattern1_ini)
                ret = set_read3(w_w_str, "pattern1", i)
                If ret = False Then
                    MsgBox(Tmp_Pattern1_ini & "ファイル読込エラー", 64, "BrandVB エラー")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("選択されたフォントタイプのデータは、設定ファイル(" & Tmp_Pattern1_ini & ")に設定されていません", 64, "Configuration file error")
			Exit Sub
		End If
		
	End Sub
End Class