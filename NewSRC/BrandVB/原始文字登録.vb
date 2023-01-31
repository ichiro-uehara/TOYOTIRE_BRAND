Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_GMSAVE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim result As Object
		
		Dim w_ret As Short
		Dim w_mess As String
		Dim ZumenName As String
		Dim TiffFile As String
		Dim tmpTiffFile As String
		Dim w_file As String
		
        On Error Resume Next
		Err.Clear()

        init_sql()

        If check_F_GMSAVE() <> 0 Then
            end_sql()
            Exit Sub

        Else
            If open_mode = "NEW" Then
                result = gm_insert()
            Else
                result = gm_update()
            End If

            If result = FAIL Then
                MsgBox("Failed to register the primitive character.", 64, "registration error")
            Else
                MsgBox("Registered the primitive character.")
                ZumenName = "GM-" & Trim(form_no.w_font_name.Text)

                TiffFile = TIFFDir & Trim(form_no.w_font_name.Text) & VB.Left(Trim(form_no.w_font_class1.Text), 1) & VB.Left(Trim(form_no.w_font_class2.Text), 1) & VB.Left(Trim(form_no.w_name1.Text), 1) & VB.Left(Trim(form_no.w_name2.Text), 1) & ".bmp"
                tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
                FileCopy(tmpTiffFile, TiffFile)
                If Err.Number <> 0 Then
                    MsgBox("error_no: " & Str(Err.Number) & Err.Description, , "file error")
                End If

                'BMPﾌｧｲﾙ表示
                w_file = Dir(TiffFile)
                If w_file <> "" Then
                    form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
                    form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
                    form_no.ImgThumbnail1.Height = 193 ' 200 '20100701コード変更
                Else
                    MsgBox("BMP file can not be found.", MsgBoxStyle.Critical, "File not found")
                End If
                'Brand Ver.5 TIFF->BMP 変更 end

                '特性データ送信
                w_ret = temp_gm_get()

                '（図面名、配置PIC）送信

                '----- .NET 移行 -----
                'w_mess = VB6.Format(VB.Left(form_no.w_haiti_pic.Text, 3), "000") & GensiDir & ZumenName
                w_mess = String.Format("000", Strings.Left(form_no.w_haiti_pic.Text, 3)) & GensiDir & ZumenName

                w_ret = PokeACAD("ACADSAVE", w_mess)
                w_ret = RequestACAD("ACADSAVE")

                '画面ロック
                form_no.Command1.Enabled = False
                form_no.Command2.Enabled = False
                form_no.Command4.Enabled = False
                form_no.w_font_name.Enabled = False
                form_no.w_font_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
                form_no.w_font_class1.Enabled = False
                form_no.w_font_class1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_name1.Enabled = False
                form_no.w_name1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_name2.Enabled = False
                form_no.w_name2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_comment.Enabled = False
                form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_dep_name.Enabled = False
                form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_entry_name.Enabled = False
                form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_base_r.Enabled = False
                form_no.w_base_r.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_hem_width.Enabled = False
                form_no.w_hem_width.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_old_font_name.Enabled = False
                form_no.w_old_font_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_old_font_class.Enabled = False
                form_no.w_old_font_class.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                form_no.w_old_name.Enabled = False
                form_no.w_old_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

            End If
        End If
		
		end_sql()
		
		Exit Sub
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_GMSAVE()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click

        Me.Close()
        '  form_main.Show
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
                .HelpContext = 300
                .ShowHelp()
            End With
        End If
	End Sub

    Private Sub F_GMSAVE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim w_file As Object
        Dim tmpTiffFile As String
        Dim w_ret As Object

        Dim aa As String
        Dim TiffFileName As String

        On Error Resume Next
        Err.Clear()

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)

        Call Clear_F_GMSAVE()

        If open_mode = "NEW" Then
            w_ret = PokeACAD("SAVEMODE", "FRESH")
            RequestACAD("SAVEMODE")

            'フォント区分１
            w_font_class1.Items.Clear()
            w_font_class1.Items.Add("A:Solid")
            w_font_class1.Items.Add("F:Hemming letter")
            w_font_class1.Items.Add("H:Hutchings letter")
            w_font_class1.Items.Add("B:Edge & Hutchings")
            w_font_class1.Items.Add("D:Dummy letter")
            w_font_class1.Items.Add("N:Screw")
            w_font_class1.Items.Add("P:Plate")
            '----- .NET移行  -----
            'w_font_class1.Text = VB6.GetItemString(w_font_class1, 0)
            w_font_class1.SelectedIndex = 0

            '文字名１
            w_name1.Items.Clear()
            w_name1.Items.Add("A:an alphabetic character")
            w_name1.Items.Add("B:Number")
            w_name1.Items.Add("C:Hiragana letter")
            w_name1.Items.Add("D:Katakana letter")
            w_name1.Items.Add("E:kanji letter")
            w_name1.Items.Add("F:Etc")
            '----- .NET移行  -----
            'w_name1.Text = VB6.GetItemString(w_name1, 0)
            w_name1.SelectedIndex = 0


            '縁取り幅ロック (Brand CAD System Ver.3 UP)
            w_hem_width.Enabled = False
            w_hem_width.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

            'Bmpﾌｧｲﾙ
            tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
            w_file = Dir(tmpTiffFile)
            If w_file <> "" Then
                form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(tmpTiffFile)
                form_no.ImgThumbnail1.Width = 457 '500 '20100701コード変更
                form_no.ImgThumbnail1.Height = 193 ' 200 '20100701コード変更
            Else
                MsgBox("BMP file can not be found.", MsgBoxStyle.Critical, "File not found")
            End If

            Call true_date(aa)
            form_no.w_entry_date.Text = aa
        Else
            w_ret = PokeACAD("SAVEMODE", "MODIFY")
            RequestACAD("SAVEMODE")

            tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
            w_file = Dir(tmpTiffFile)
            If w_file <> "" Then
                form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(tmpTiffFile)
                form_no.ImgThumbnail1.Width = System.Drawing.Image.FromFile(tmpTiffFile).Width '500 '20100701コード変更
                form_no.ImgThumbnail1.Height = System.Drawing.Image.FromFile(tmpTiffFile).Height ' 200 '20100701コード変更
            Else
                MsgBox("BMP file can not be found.", MsgBoxStyle.Critical, "File not found")
            End If

            Call true_date(aa)
            form_no.w_entry_date.Text = aa

            form_no.w_font_name.Enabled = False
            form_no.w_font_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_font_class1.Enabled = False
            form_no.w_font_class1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_font_class2.Enabled = False
            form_no.w_font_class2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_name1.Enabled = False
            form_no.w_name1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_name2.Enabled = False
            form_no.w_name2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If

        '----- .NET移行 (StartPositionプロパティをCenterScreenで対応) -----
        'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
        'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 4) ' フォームを画面の縦方向にセンタリングします。

        form_no.Text1.Text = open_mode

        ' -> watanabe mov 20110204 ACAD通信前にセットする様に変更
        CommunicateMode = comSpecData
        ' <- watanabe mov 20110204 ACAD通信前にセットする様に変更

        RequestACAD("SPECDATA")

    End Sub

    Private Sub w_comment_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_comment.Leave
		
        form_no.w_comment.Text = apos_check(form_no.w_comment.Text)
		
	End Sub
	
	Private Sub w_dep_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_dep_name.Leave
		
        form_no.w_dep_name.Text = UCase(Trim(form_no.w_dep_name.Text))
		
	End Sub
	
	
    'UPGRADE_WARNING: イベント w_font_class1.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_font_class1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_class1.SelectedIndexChanged

        If TypeOf form_no.ActiveControl Is System.Windows.Forms.Button Then
            Exit Sub
        End If

        '----- .NET移行  -----
        'If w_font_class1.Text = VB6.GetItemString(w_font_class1, 4) Or w_font_class1.Text = VB6.GetItemString(w_font_class1, 5) Or w_font_class1.Text = VB6.GetItemString(w_font_class1, 6) Then
        '    w_name1.Text = VB6.GetItemString(w_name1, 5)
        If w_font_class1.Text = w_font_class1.Items(4).ToString() Or w_font_class1.Text = w_font_class1.Items(5).ToString() Or w_font_class1.Text = w_font_class1.Items(6).ToString() Then
            w_name1.Text = w_name1.Items(5).ToString()
            w_name1.Enabled = False
            w_name2.Text = ""
        Else
            w_name1.Enabled = True
		End If
		
	End Sub
	
	Private Sub w_font_class1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_font_class1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_font_class1, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub w_font_class1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_class1.Leave
		
		'フォント区分1 ｢縁取り｣ OR ｢縁＆ハッチング｣ チェック  (Brand CAD System Ver.3 UP)
		If VB.Left(Trim(w_font_class1.Text), 1) = "F" Or VB.Left(Trim(w_font_class1.Text), 1) = "B" Then
			'縁取り幅ロック解除
			w_hem_width.Enabled = True
			w_hem_width.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
		Else
			w_hem_width.Text = "0.00"
			w_hem_width.Enabled = False
			w_hem_width.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
		End If
		
	End Sub
	
	Private Sub w_font_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_name.Leave
		
        form_no.w_font_name.Text = UCase(Trim(form_no.w_font_name.Text))
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_name1.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_name1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name1.SelectedIndexChanged
		
		If TypeOf form_no.ActiveControl Is System.Windows.Forms.Button Then
			Exit Sub
		End If
		
        'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。
		If w_name1.Text Is System.DBNull.Value Or w_name1.Text <> dummy_text Then
			w_name2.Text = ""
		End If
		
	End Sub
	
	Private Sub w_name1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name1.Enter
		
		dummy_text = w_name1.Text
		
	End Sub
	
	Private Sub w_name1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_name1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_name1, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_name2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name2.Leave
		
        form_no.w_name2.Text = UCase(Trim(form_no.w_name2.Text))
		
	End Sub
	
	Private Sub w_old_font_class_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_old_font_class.Leave
		
        form_no.w_old_font_class.Text = UCase(Trim(form_no.w_old_font_class.Text))
		
	End Sub
	
	Private Sub w_old_font_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_old_font_name.Leave
		
        form_no.w_old_font_name.Text = UCase(Trim(form_no.w_old_font_name.Text))
		
	End Sub
	
	Private Sub w_old_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_old_name.Leave
		
        form_no.w_old_name.Text = UCase(Trim(form_no.w_old_name.Text))
		
	End Sub
End Class