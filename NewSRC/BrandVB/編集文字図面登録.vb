Option Strict Off
Option Explicit On
Friend Class F_HZSAVE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
		Dim ZumenName As Object
		Dim result As Object
		
		Dim w_mess As String


        ' -> watanabe add VerUP(2011)
        result = FAIL
        ' <- watanabe add VerUP(2011)


		init_sql()
        If open_mode = "NEW" Then
            If check_F_HZSAVE() <> 0 Then
                Exit Sub
            Else
                result = hz_insert()
            End If
        ElseIf open_mode = "Revision number" Then
            If check_F_HZSAVE() <> 0 Then
                Exit Sub
            Else
                result = hz_addnum()
            End If
        ElseIf open_mode = "modify" Then
            If check_F_HZSAVE() <> 0 Then
                Exit Sub
            Else
                result = hz_update()
            End If
        End If
		
        If result = FAIL Then
            MsgBox("Failed to register the Editing characters drawing.", 64, "registration error")
        Else
            MsgBox("Registered the Editing characters drawing.")

            '�i�}�ʖ��j���M
            ZumenName = "HE-" & Trim(form_no.w_no1.Text) & "-" & Trim(form_no.w_no2.Text)
            w_mess = HensyuZumenDir & ZumenName
            w_ret = PokeACAD("MDLSAVE", w_mess)
            w_ret = RequestACAD("MDLSAVE")

            '��ʃ��b�N
            form_no.Command1.Enabled = False
            form_no.Command2.Enabled = False
            form_no.Command4.Enabled = False
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_comment.Enabled = False
            form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_dep_name.Enabled = False
            form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_date.Enabled = False
            form_no.w_entry_date.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If
		end_sql()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_HZSAVE()
		
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
                .HelpContext = 600
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_HZSAVE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        temp_hz.Initilize() '20100702�ǉ��R�[�h
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B
		
		Text1.Text = open_mode
		
		Call Clear_F_HZSAVE()
		
		Call true_date(aa)
		w_entry_date.Text = aa
		
		w_id.Text = "HE"

        form_no.w_entry_date.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        form_no.w_hm_num.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

        If Text1.Text = "NEW" Then
            w_ret = PokeACAD("SAVEMODE", "FRESH")
            RequestACAD("SAVEMODE")
            temp_hz.id = "HE"
            temp_hz.no1 = ""
            temp_hz.no2 = "00"
            temp_hz.comment = ""
            temp_hz.dep_name = ""
            temp_hz.entry_name = ""
            Call true_date(aa)
            temp_hz.entry_date = aa
            temp_hz.hm_num = 0
            CommunicateMode = comSpecData
            RequestACAD("HMCODE")
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
        ElseIf Text1.Text = "Revision number" Then
            w_ret = PokeACAD("SAVEMODE", "CHANGE")
            RequestACAD("SAVEMODE")
            CommunicateMode = comSpecData
            RequestACAD("ZMNNAME")
            temp_hz.id = "HE"
            temp_hz.no1 = ""
            temp_hz.no2 = ""
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        ElseIf Text1.Text = "modify" Then
            w_ret = PokeACAD("SAVEMODE", "MODIFY")
            RequestACAD("SAVEMODE")
            CommunicateMode = comSpecData
            RequestACAD("ZMNNAME")
            temp_hz.id = "HE"
            temp_hz.no1 = ""
            temp_hz.no2 = ""
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If
		
        ' -> watanabe del VerUP(2011)
        ''�����ݒ� <- TEST ->
        'MSFlexGrid1.Rows = 2
        'MSFlexGrid1.Cols = 5
        '
        'For lp = 0 To MSFlexGrid1.Rows - 1
        '          MSFlexGrid1.set_RowHeight(lp, 300)
        'Next lp
        '
        'MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 1)
        'MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 2)
        'MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 6)
        'MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 2)
        'MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 18 * 6)
        '
        'For i = 0 To 4
        '          MSFlexGrid1.set_FixedAlignment(i, 2)
        'Next i
        '
        '      w_ret = Set_Grid_Data(MSFlexGrid1, "NO", 0, 0)
        '      w_ret = Set_Grid_Data(MSFlexGrid1, "�װ", 0, 1)
        '      w_ret = Set_Grid_Data(MSFlexGrid1, "�ҏW��������", 0, 2)
        '      w_ret = Set_Grid_Data(MSFlexGrid1, "�װ", 0, 3)
        '      w_ret = Set_Grid_Data(MSFlexGrid1, "�ҏW��������", 0, 4)
        ' <- watanabe del VerUP(2011)

	End Sub

    '----- .NET�ڍs (ToDo:DataGridView�̃C�x���g�ɕύX) -----
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