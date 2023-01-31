Option Strict Off
Option Explicit On
Friend Class F_GZSAVE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
        Dim ZumenName As String
		Dim result As Object
		
		Dim w_mess As String


        ' -> watanabe add VerUP(2011)
        result = FAIL
        ' <- watanabe add VerUP(2011)


		init_sql()
        If open_mode = "NEW" Then
            If check_F_GZSAVE() <> 0 Then
                Exit Sub
            Else
                result = gz_insert()
            End If
        ElseIf open_mode = "Revision number" Then
            If check_F_GZSAVE() <> 0 Then
                Exit Sub
            Else
                result = gz_addnum()
            End If
        ElseIf open_mode = "modify" Then
            If check_F_GZSAVE() <> 0 Then
                Exit Sub
            Else
                result = gz_update()
            End If
        End If
		
        If result = FAIL Then
            MsgBox("Failed to register a Stamp drawing.", 64, "registration error")
        Else
            MsgBox("Registered the carved seal drawing.")

            '（図面名）送信
            ZumenName = "KO-" & Trim(form_no.w_no1.Text) & "-" & Trim(form_no.w_no2.Text)
            w_mess = KokuinDir & ZumenName
            w_ret = PokeACAD("MDLSAVE", w_mess)
            w_ret = RequestACAD("MDLSAVE")

            '画面ロック
            form_no.Command1.Enabled = False
            form_no.Command2.Enabled = False
            form_no.Command4.Enabled = False
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_comment.Enabled = False
            form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_dep_name.Enabled = False
            form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If
		end_sql()
		
	End Sub
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_GZSAVE()
		
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
                .HelpContext = 500
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_GZSAVE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        ' -> watanabe del VerUP(2011)
        'Dim i As Object
        'Dim lp As Object
        ' <- watanabe del VerUP(2011)

        Dim w_ret As Object
		
		Dim aa As String


        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)


        form_no = Me
        temp_gz.Initilize() '20100702追加コード
		
		form_no.Text1.Text = open_mode
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		Call Clear_F_GZSAVE()
		
		form_no.w_id.Text = "KO" '固定
		
        If Text1.Text = "NEW" Then
            w_ret = PokeACAD("SAVEMODE", "FRESH")
            RequestACAD("SAVEMODE")
            temp_gz.id = "KO"
            temp_gz.no1 = ""
            temp_gz.no2 = "00"
            temp_gz.comment = ""
            temp_gz.dep_name = ""
            temp_gz.entry_name = ""
            Call true_date(aa)
            temp_gz.entry_date = aa
            temp_gz.gm_num = 0
            CommunicateMode = comSpecData
            RequestACAD("GMCODE")
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        ElseIf Text1.Text = "Revision number" Then
            w_ret = PokeACAD("SAVEMODE", "CHANGE")
            RequestACAD("SAVEMODE")
            CommunicateMode = comSpecData
            RequestACAD("ZMNNAME")
            temp_gz.id = "KO"
            temp_gz.no1 = ""
            temp_gz.no2 = ""
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        ElseIf Text1.Text = "modify" Then
            w_ret = PokeACAD("SAVEMODE", "MODIFY")
            RequestACAD("SAVEMODE")
            CommunicateMode = comSpecData
            RequestACAD("ZMNNAME")
            temp_gz.id = "KO"
            temp_gz.no1 = ""
            temp_gz.no2 = ""
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If


        ' -> watanabe del VerUP(2011)   フォームロード時だとグリッドの行が更新されない
        ''初期設定 <- TEST ->
        'MSFlexGrid1.Rows = 2
        'MSFlexGrid1.Cols = 5
        '
        '' 行高さの設定
        'For lp = 0 To MSFlexGrid1.Rows - 1
        '    MSFlexGrid1.set_RowHeight(lp, 300)
        'Next lp
        '
        '' 列幅の設定
        'MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 1)
        'MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 2)
        'MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 6)
        'MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 2)
        'MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 6)
        'For i = 0 To 4
        '    MSFlexGrid1.set_FixedAlignment(i, 2)
        'Next i
        '
        'w_ret = Set_Grid_Data(MSFlexGrid1, "NO", 0, 0)
        'w_ret = Set_Grid_Data(MSFlexGrid1, "ｴﾗｰ", 0, 1)
        'w_ret = Set_Grid_Data(MSFlexGrid1, "原始文字ｺｰﾄﾞ", 0, 2)
        'w_ret = Set_Grid_Data(MSFlexGrid1, "ｴﾗｰ", 0, 3)
        'w_ret = Set_Grid_Data(MSFlexGrid1, "原始文字ｺｰﾄﾞ", 0, 4)
        ' <- watanabe del VerUP(2011)

	End Sub

    '----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
#If False Then
    Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		
        MsgBox("You can not change the key input.", 64)
		eventArgs.KeyAscii = 0
		
	End Sub
#End If

    Private Sub w_comment_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_comment.Leave
		
        form_no.w_comment.Text = apos_check(form_no.w_comment.Text)
		
	End Sub
	
	Private Sub w_dep_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_dep_name.Leave
		
        form_no.w_dep_name.Text = UCase(Trim(form_no.w_dep_name.Text))
		
	End Sub
	
	Private Sub w_no1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_no1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim irt As Object
		
		Dim f As System.Windows.Forms.Control
		
		If KeyAscii = 9 Or KeyAscii = 10 Or KeyAscii = 13 Then
			
			w_no1.Text = Trim(w_no1.Text)
			
            f = form_no.w_no1
            irt = check_0((w_no1.Text), 4, 0, f)
            If irt <> 0 Then
                MsgBox("Code is invalid.", 64, "Input error")
                f.Focus()
            End If
			
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_no1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no1.Leave
		
        form_no.w_no1.Text = UCase(Trim(form_no.w_no1.Text))
		
	End Sub
	
	Private Sub w_no2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_no2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim irt As Object
		
		Dim f As System.Windows.Forms.Control
		
		If KeyAscii = 9 Or KeyAscii = 10 Or KeyAscii = 13 Then
			
			w_no2.Text = Trim(w_no2.Text)
            f = form_no.w_no1
            irt = check_0((w_no2.Text), 2, 0, f)
            If irt <> 0 Then
                MsgBox("Code is invalid.", 64, "Input error")
                f.Focus()
            End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class