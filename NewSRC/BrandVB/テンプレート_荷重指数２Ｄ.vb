Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_KAJUU2D
	Inherits System.Windows.Forms.Form
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
        'form_no.w_font.ListIndex = 0
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0)) '20100624変更コード
        form_no.w_load_index1.Text = ""
        form_no.w_load_index2.Text = ""
        form_no.w_sokudo.Text = ""
		
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
                .HelpContext = 801
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim gm_no As Object
		Dim ZumenName As Object
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		
		Dim w_mess As String
		Dim w_w_str As String
		Dim i As Short
		Dim load_data As String
		
		If check_F_TMP_KAJUU2D <> 0 Then Exit Sub
		
		load_data = Trim(w_load_index1.Text) & "/" & Trim(w_load_index2.Text) & Trim(w_sokudo.Text)
		
		form_no.Enabled = False
		F_MSG.Show()
		
		'(Brand Ver.3 追加)
		For i = 1 To Len(load_data)
            w_str = Mid(load_data, i, 1)
			If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input Load index is not set to the configuration file (" & Tmp_Load2D_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf w_str = "/" Then
                If GensiKIGO(Asc(w_str)) = "" Then
                    MsgBox("A substituted primitive letter for input Load index is not set to the configuration file (" & Tmp_Load2D_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for input Load index is not set to the configuration file (" & Tmp_Load2D_ini & ")", 64, "Configuration file error")
                    GoTo error_section
                End If
			End If
		Next i
		
		If FreePicNum < 1 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
			GoTo error_section
		End If
		
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
				
            ElseIf w_str = "/" Then
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
        w_mess = Tmp_Load2D_ini
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
        form_no.Command4.Enabled = False
        form_no.w_font.Enabled = False
        form_no.w_font.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.w_load_index1.Enabled = False
        form_no.w_load_index1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_load_index2.Enabled = False
        form_no.w_load_index2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_sokudo.Enabled = False
        form_no.w_sokudo.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_KAJUU2D_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_ret As Object
		Dim ret As Object
		Dim i As Object
		
		Dim w_w_str As String
		
		form_no = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
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
		
        'form_no.w_font.ListIndex = 0
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0)) '20100624変更コード
        form_no.w_load_index1.Text = ""
        form_no.w_load_index2.Text = ""
        form_no.w_sokudo.Text = ""
		
		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Load2D_ini)
        ret = set_read4(w_w_str, "load2D", 1)
		
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
                w_w_str = Trim(w_w_str) & Trim(Tmp_Load2D_ini)
                ret = set_read4(w_w_str, "load2D", i)
                If ret = False Then
                    MsgBox(Tmp_Load2D_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Load2D_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
	End Sub
	
	Private Sub w_sokudo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_sokudo.Leave
		
        form_no.w_sokudo.Text = UCase(Trim(form_no.w_sokudo.Text))
		
	End Sub
End Class