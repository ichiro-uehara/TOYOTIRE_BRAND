Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_LT2
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim gm_no As Object
		Dim ZumenName As Object
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		
		Dim w_mess As String
		Dim w_w_str As String
		Dim i As Short
		Dim lt_data As String
		
		If Option1.Checked = True Then
			lt_data = w_lt.Text
		Else
			lt_data = w_c.Text
		End If
		
		form_no.Enabled = False
		F_MSG.Show()
		
		'(Brand Ver.3 追加)
		For i = 1 To Len(lt_data)
            w_str = Mid(lt_data, i, 1)
			If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("Replacement for the primitive character string that is selected has not been set in the configuration file (" & Tmp_Lt2_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("Replacement for the primitive character string that is selected has not been set in the configuration file (" & Tmp_Lt2_ini & ")", 64, "Configuration file error")
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
		
		'[ＬＴ]
		For i = 1 To Len(lt_data)
            w_str = Mid(lt_data, i, 1)
			
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
        w_mess = Tmp_Lt2_ini
        w_ret = PokeACAD("TMPNAME", w_mess)
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_mess = "TYPE" & i
                w_ret = PokeACAD("TMPDATANO", w_mess)
				Exit For
			End If
		Next i
		w_mess = Trim(lt_data)
        w_ret = PokeACAD("TMPSPELL", w_mess)

        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG3")

        form_no.Command1.Enabled = False
        form_no.Command2.Enabled = False
        form_no.Command4.Enabled = False
        form_no.w_font.Enabled = False
        form_no.w_font.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.Option1.Checked = False
        form_no.Option1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.Option2.Checked = False
        form_no.Option2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
		
		Exit Sub
		
error_section: 
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
        'form_no.w_font.ListIndex = 0
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0)) '20100624コード変更
        'form_no.Option1.Value = 1
        form_no.Option1.Checked = True
        'form_no.Option2.Value = 0
        form_no.Option2.Checked = False
		
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
                .HelpContext = 801
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_TMP_LT2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0)) '20100624コード変更
        'form_no.Option1.Value = 1
        form_no.Option1.Checked = True
        'form_no.Option2.Value = 0
        form_no.Option2.Checked = False

		w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_Lt2_ini)
        ret = set_read4(w_w_str, "lt2", 1)
		
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
                w_w_str = Trim(w_w_str) & Trim(Tmp_Lt2_ini)
                ret = set_read4(w_w_str, "lt2", i)
                If ret = False Then
                    MsgBox(Tmp_Lt2_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Lt2_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
	End Sub
End Class