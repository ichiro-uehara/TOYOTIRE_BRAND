Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class F_BZDELE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
		Dim result As Object
		Dim ret As Short
		Dim w_mess As String
		
		ret = check_0((w_no1.Text), 4, 0, w_no1)
		If ret <> 0 Then
			Exit Sub
		End If
		ret = check_0((w_no2.Text), 2, 0, w_no2)
		If ret <> 0 Then
			Exit Sub
		End If
		
        If Trim(form_no.w_no1.Text) = "" Then
            MsgBox("Number is non-entry.", 64, "InputError")
            Exit Sub
        End If
        If Trim(form_no.w_no2.Text) = "" Then
            MsgBox("Revision number is non-input.", 64, "InputError")
            Exit Sub
        End If
		
		ret = init_sql
		If ret = False Then Exit Sub
		
        result = bz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
        If result = FAIL Then
            MsgBox("There is no brand drawing corresponding.", 64, "Search error")
            end_sql()
            Exit Sub
        Else
            dataset_F_BZDELE()
        End If
		end_sql()
		
        w_ret = MsgBox("Do you delete it?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, "Confirmation")
		If w_ret = MsgBoxResult.Yes Then
			init_sql()
            result = bz_delete(Trim(w_no1.Text), Trim(w_no2.Text))
            If result = FAIL Then
                MsgBox("Failed to delete the brand drawing.", 64, "Delete error")
            Else
                'POKE送信->ACAD（図面名）
                w_mess = "AT-B-" & Trim(w_no1.Text) & "-" & Trim(w_no2.Text)
                w_ret = PokeACAD("MDLDELE", w_mess)
                w_ret = RequestACAD("MDLDELE")
                Clear_F_BZDELE()
                MsgBox("Delete the brand drawing.")
            End If
			end_sql()
		Else
            MsgBox("can not delete it.", , "Cancellation")
		End If
		end_sql()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_BZDELE()
		
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
                .HelpContext = 701
                .ShowHelp()
            End With
        End If	
	End Sub
	
	Private Sub F_BZDELE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim lp As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        form_no = Me
        temp_bz.Initilize() '20100702追加コード
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 4) ' フォームを画面の縦方向にセンタリングします。
		
		Call Clear_F_BZDELE()
		
		w_id.Text = "AT-B"
		
		MSFlexGrid1.Rows = 9 ' 列と行の総数を設定します。
		MSFlexGrid1.Cols = 5
		
		w_hm_num.Text = CStr((MSFlexGrid1.Rows - 1) * (MSFlexGrid1.Cols - 1))
		
		For lp = 0 To MSFlexGrid1.Rows - 1
			MSFlexGrid1.set_RowHeight(lp, 400)
		Next lp
		
		MSFlexGrid1.set_ColWidth(0, 1000)
		For lp = 1 To MSFlexGrid1.Cols - 1
			MSFlexGrid1.set_ColWidth(lp, 1900)
		Next lp
		
		For lp = 1 To MSFlexGrid1.Rows - 1
			MSFlexGrid1.Row = lp
			MSFlexGrid1.Col = 0
			MSFlexGrid1.Text = Str(lp)
		Next lp
		
		For lp = 1 To MSFlexGrid1.Cols - 1
			MSFlexGrid1.Row = 0
			MSFlexGrid1.Col = lp
			MSFlexGrid1.Text = Str(lp)
		Next lp
		
	End Sub

	'----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
#If False Then
	Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		
        MsgBox("You can not change the key input.", 64)
		eventArgs.KeyAscii = 0
		
	End Sub
#End If

	Private Sub w_kikaku1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku1.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_kikaku2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku2.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_kikaku3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku3.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku3_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku3.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_kikaku4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku4.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku4_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku4.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_kikaku5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku5.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku5_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku5.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_kikaku6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku6.Click
		
		Dim f As System.Windows.Forms.Control
		
		w_kikaku.Text = ""
		
		f = form_no.w_kikaku1
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku2
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku3
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku4
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku5
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		f = form_no.w_kikaku6
		If f.Text <> "" Then
			If Mid(f.Text, 1, 1) <> " " Then
				w_kikaku.Text = w_kikaku.Text & Mid(f.Text, 1, 1)
			End If
		End If
		
	End Sub
	
	Private Sub w_kikaku6_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_kikaku6.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_nasiji_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_nasiji.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	
	
	Private Sub w_no1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no1.Leave
		
		form_no.w_no1.Text = UCase(Trim(form_no.w_no1.Text))
		
	End Sub
	
	Private Sub w_peak_mark_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_peak_mark.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_plant_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_plant.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_side_kenti_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_side_kenti.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_side_moyou_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_side_moyou.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_size1.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
	End Sub
	
    'UPGRADE_WARNING: イベント w_size2.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size3.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size4.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size4_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size5.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size5_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size6.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size6_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size7.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size7_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size7.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_size8.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub w_size8_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size8.TextChanged
		w_size.Text = w_size1.Text & w_size2.Text & w_size3.Text & w_size4.Text & w_size5.Text & w_size6.Text & w_size7.Text & w_size8.Text
		
	End Sub
	Private Sub w_syubetu_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syubetu.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_syurui_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syurui.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_tos_moyou_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_tos_moyou.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
        MsgBox("You can not change the key input.", 64)
		KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class