Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_MORUDO
	Inherits System.Windows.Forms.Form
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_MORUDO()
		
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
                .HelpContext = 803
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim gm_no As Object
		Dim gm_alph As Object
        Dim ZumenName As String
		Dim pic_no As Object
		Dim w_ret As Object
		Dim w_str As Object
		Dim i As Object
		
		Dim w_mess As String
		Dim w_char As New VB6.FixedLengthString(1)
		Dim wstr As String

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


		'/* 入力チェック */
        form_no.w_kubun.Text = Trim(form_no.w_kubun.Text)
        form_no.w_no.Text = Trim(form_no.w_no.Text)
		If check_F_TMP_MOLD <> 0 Then
			Exit Sub
		End If
		
		form_no.Enabled = False
		F_MSG.Show()
		
		'(Brand Ver.3 追加)
        For i = 1 To Len(form_no.w_no.Text)
            w_str = Mid(form_no.w_no.Text, i, 1)
            If IsNumeric(w_str) Then
                If Val(w_str) >= 0 And Val(w_str) < 10 Then
                    If GensiNUM(Val(w_str)) = "" Then
                        MsgBox("A substituted primitive letter for input Mold number is not set to the configuration file (" & Tmp_Mold_no1_ini & ")", 64, "Configuration file error")
                        GoTo error_section
                    End If
                End If
            ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then
                If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                    MsgBox("A substituted primitive letter for input Mold number is not set to the configuration file (" & Tmp_Mold_no1_ini & ")", 64, "Configuration file error")
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
		
        pic_no = what_pic_from_hmcode(form_no.w_hm_name.Text)
        If pic_no < 1 Then GoTo error_section
        ZumenName = "HM-" & VB.Left(Trim(form_no.w_hm_name.Text), 6)

		'----- .NET 移行 -----
		'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
		w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

		w_ret = PokeACAD("HMCODE", w_mess)
		
		
		'20000124 追加
		For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                wstr = Tmp_prcs_code(i)
                Exit For
            End If
		Next i
		
		
		'20000124 修正
		'   If form_no.w_type.Text = "番号のみ" Then
		If wstr = "MNO1" Then
			
			'[[ 番号 ]]
            For i = 1 To Len(form_no.w_no.Text)
                w_char.Value = Mid(form_no.w_no.Text, i, 1)
                '2005.05.09 O.Kawaguchi 変更 A 追加
                '2010.12.21 T.Uriu 変更 M 追加
                '2011.04.14 T.Uriu 変更 Z 追加
                '2012.08.31 T.Uriu 変更 B 追加
                '       If w_char = "T" Then
                If w_char.Value = "T" Or w_char.Value = "A" Or _
                   w_char.Value = "M" Or w_char.Value = "Z" Or _
                   w_char.Value = "B" Then
                    gm_alph = Mid(form_no.w_no.Text, i, 1)
                    pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
                    ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)
                Else
                    gm_no = Val(Mid(form_no.w_no.Text, i, 1))
                    pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                    ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)
                End If
                If pic_no < 1 Then GoTo error_section

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE1", w_mess)
            Next i
			
		Else
			
			'[[ 区分 ]]
            For i = 1 To Len(form_no.w_kubun.Text)
                gm_alph = Mid(form_no.w_kubun.Text, i, 1)
                pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
                If pic_no < 1 Then GoTo error_section
                ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE1", w_mess)
            Next i
			
			'[[ 番号 ]]
            For i = 1 To Len(form_no.w_no.Text)
                w_char.Value = Mid(form_no.w_no.Text, i, 1)
                '2005.05.09 O.Kawaguchi 変更 A 追加
                '2010.12.21 T.Uriu 変更 M 追加
                '2011.04.14 T.Uriu 変更 Z 追加
                '       If w_char = "T" Then
                If w_char.Value = "T" Or w_char.Value = "A" Or _
                   w_char.Value = "M" Or w_char.Value = "Z" Or _
                   w_char.Value = "B" Then
                    gm_alph = Mid(form_no.w_no.Text, i, 1)
                    pic_no = what_pic_from_gmcode(GensiALPH(Asc(gm_alph) - Asc("A")))
                    ZumenName = "GM-" & Mid(GensiALPH(Asc(gm_alph) - Asc("A")), 1, 6)
                Else
                    gm_no = Val(Mid(form_no.w_no.Text, i, 1))
                    pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
                    ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)
                End If
                If pic_no < 1 Then GoTo error_section

				'----- .NET 移行 -----
				'w_mess = VB6.Format(Val(pic_no), "000") & GensiDir & ZumenName
				w_mess = Val(pic_no).ToString("000") & GensiDir & ZumenName

				w_ret = PokeACAD("GMCODE2", w_mess)
            Next i
		End If
		
		'// 終了の送信
        CommunicateMode = comNone
        w_ret = RequestACAD("TMPCHANG")

		'画面ロック
        form_no.Command2.Enabled = False
        form_no.Command4.Enabled = False
        form_no.Command6.Enabled = False
        form_no.w_type.Enabled = False
        form_no.w_type.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        form_no.w_kubun.Enabled = False
        form_no.w_kubun.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_no.Enabled = False
        form_no.w_no.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		
		Exit Sub
		
error_section: 
		'MsgBox "ｴﾗｰが発生しました", , "ｴﾗｰ"
		On Error Resume Next
		F_MSG.Close()
		form_no.Enabled = True
		
	End Sub
	
	Private Sub F_TMP_MORUDO_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        w_w_str = Trim(w_w_str) & Trim(Tmp_Mold_no1_ini)
        ret = set_read5(w_w_str, "mold1", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
				'20000124 修正
				'          If Tmp_hm_word(i) = "0" Then
				'             form_no.w_type.AddItem "番号のみ"
				'          ElseIf Tmp_hm_word(i) = "1" Then
				'             form_no.w_type.AddItem "区分(1桁)＋番号"
				'          ElseIf Tmp_hm_word(i) = "2" Then
				'             form_no.w_type.AddItem "区分(2桁)＋番号"
				'          End If
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_MORUDO()
		
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
                w_w_str = Trim(w_w_str) & Trim(Tmp_Mold_no1_ini)
                ret = set_read5(w_w_str, "mold_no1", i)
                If ret = False Then
                    MsgBox(Tmp_Mold_no1_ini & "File reading error.", 64, "BrandVB error")
                    Exit Sub
                Else
                    read_flg = 1
                    Exit For
                End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Mold_no1_ini & ")", 64, "Configuration file error")
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
				'          If Tmp_hm_word(i) = "0" Then
				'             form_no.w_type.AddItem "番号のみ"
				'          ElseIf Tmp_hm_word(i) = "1" Then
				'             form_no.w_type.AddItem "区分(1桁)＋番号"
				'          ElseIf Tmp_hm_word(i) = "2" Then
				'             form_no.w_type.AddItem "区分(2桁)＋番号"
				'          End If
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
        Dim w_file As String
        Dim TiffFile As String
        Dim w_text As String

        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

        On Error Resume Next
		Err.Clear()
		
        w_text = w_hm_name.Text
        If w_text = "" Then Exit Sub
		
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
	
	
	Private Sub w_kubun_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kubun.Leave
		
        form_no.w_kubun.Text = UCase(Trim(form_no.w_kubun.Text))
		
	End Sub
	
	Private Sub w_no_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no.Leave
		
        form_no.w_no.Text = UCase(Trim(form_no.w_no.Text))
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_type.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged
        Dim wstr As Object
		Dim i As Object

        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


        If InitFlag = False Then '20100628追加コード
            Exit Sub
        End If

		'(Brand Ver.3 変更)
		'20000124 修正
		'   If w_type.Text = "番号のみ" Then
		'       w_str = "0"
		'   ElseIf w_type.Text = "区分(1桁)＋番号" Then
		'       w_str = "1"
		'   ElseIf w_type.Text = "区分(2桁)＋番号" Then
		'       w_str = "2"
		'   End If
		
		For i = 1 To MaxSelNum
			'20000124 修正
			'       If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = w_type.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                '20000124 追加
                wstr = Tmp_prcs_code(i)
                Exit For
            End If
		Next i
		
		'20000124 修正
		'   If wstr = "0" Then
        If wstr = "MNO1" Then
            form_no.w_kubun.Enabled = False
            form_no.w_kubun.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        Else
            form_no.w_kubun.Enabled = True
            form_no.w_kubun.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
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