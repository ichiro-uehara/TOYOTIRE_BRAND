Option Strict Off
Option Explicit On
Friend Class F_GZDELE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
		Dim result As Object
		Dim ret As Object
		
		Dim w_mess As String
		
        If Trim(form_no.w_no1.Text) = "" Then
            MsgBox("Drawing number is non-input.", 64, "InputError")
            Exit Sub
        ElseIf Trim(form_no.w_no2.Text) = "" Then
            MsgBox("Revision number is non-input.", 64, "InputError")
            Exit Sub
        End If
		
        ret = init_sql()
        If ret = False Then Exit Sub
		
        result = gz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
        If result = FAIL Then
            MsgBox("There is no  carved seal drawing corresponding.", 64, "Search error")
            end_sql()
            Exit Sub
        Else
            dataset_F_GZDELE()
        End If
		end_sql()
		
        w_ret = MsgBox("Do you delete it?", MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation + MsgBoxStyle.DefaultButton2, "Confirmation")
		If w_ret = MsgBoxResult.Yes Then
			init_sql()
            result = gz_delete(Trim(w_no1.Text), Trim(w_no2.Text))
            If result = FAIL Then
                MsgBox("Failed to delete the Stamp drawing.", 64, "Delete error")
            Else
                MsgBox("Delete the carved seal drawing.")
                'POKE送信->ACAD（図面名）
                w_mess = KokuinDir & "KO-" & Trim(w_no1.Text) & "-" & Trim(w_no2.Text)
                w_ret = PokeACAD("MDLDELE", w_mess)
                w_ret = RequestACAD("MDLDELE")
                Clear_F_GZDELE()
            End If
			end_sql()
		Else
            MsgBox("can not delete it.", , "Cancellation")
		End If
		end_sql()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_GZDELE()
		
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
                .HelpContext = 501
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_GZDELE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_ret As Object
		Dim i As Object
		Dim lp As Object
		
		Dim aa As String

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)


        form_no = Me
        temp_gz.Initilize() '20100702追加コード
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		Call Clear_F_GZDELE()
		
		temp_gz.id = "KO"
		temp_gz.no1 = ""
		temp_gz.no2 = ""
		temp_gz.comment = ""
		temp_gz.dep_name = ""
		temp_gz.entry_name = ""
		Call true_date(aa)
		temp_gz.entry_date = aa
		temp_gz.gm_num = 0
		
		'初期設定 (GRID)
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
		
		' 行高さの設定
		For lp = 0 To MSFlexGrid1.Rows - 1
            form_no.MSFlexGrid1.RowHeight(lp) = 300
		Next lp
		
		' 列幅の設定
        '20100705コード変更
        form_no.MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 1)
        form_no.MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 21 * 4)
        For i = 0 To 5
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
		
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "NO", 0, 0)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "1", 0, 1)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "2", 0, 2)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "3", 0, 3)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "4", 0, 4)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "5", 0, 5)
		
		'***** 12/8 1997 yamamoto start *****
		'focus 奪い合い問題を防ぐためのﾌﾗｸﾞ
		w_no1_flg = 0
		w_no2_flg = 0
		'***** 12/8 1997 yamamoto end *****
	End Sub

	'----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
#If False Then
	Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		
        MsgBox("You can not change the key input.", 64)
		eventArgs.KeyAscii = 0
		
	End Sub
#End If

	Private Sub w_no1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_no1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim result As Object
		Dim ret As Object
		Dim irt As Object
		
		Dim f As System.Windows.Forms.Control
		
		If KeyAscii = 9 Or KeyAscii = 10 Or KeyAscii = 13 Then
			'       tab                          return
			
			w_no1.Text = Trim(w_no1.Text)
			
            f = form_no.w_no1
            irt = check_0((w_no1.Text), 4, 0, f)
            If irt <> 0 Then
                MsgBox("Code is invalid.", 64, "Input error")
                f.Focus()
            End If
			
            If Trim(form_no.w_no1.Text) = "" Then GoTo EventExitSub
            If Trim(form_no.w_no2.Text) = "" Then GoTo EventExitSub
			
            ret = init_sql()
            If ret = False Then GoTo EventExitSub
			
            result = gz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
            If result = False Then
                MsgBox("There is no Stamp drawing.", MsgBoxStyle.Critical, "ｴﾗｰ")
            Else
                dataset_F_GZDELE()
            End If
			
			end_sql()
			
		End If
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_no1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no1.Leave
		Dim result As Object
		Dim ret As Object
		Dim irt As Object
		
		Dim f As System.Windows.Forms.Control
		
        form_no.w_no1.Text = UCase(Trim(form_no.w_no1.Text))
		
		If TypeOf form_no.ActiveControl Is System.Windows.Forms.Button Then
			Exit Sub
		End If
		
		'***** 12/8 1997 yamamoto start *****
		
        If Trim(form_no.w_no1.Text) = "" Then Exit Sub
		
		If w_no2_flg = 1 Then Exit Sub
		
        f = form_no.w_no1
		w_no1.Text = Trim(w_no1.Text)
        irt = check_0((w_no1.Text), 4, 0, f)
        If irt <> 0 Then
            w_no1_flg = 1
            MsgBox("Code is invalid.", 64, "Input error")
            Exit Sub
        Else
            w_no1_flg = 0
        End If
		
        If Trim(form_no.w_no2.Text) = "" Then Exit Sub
		
		'***** 12/8 1997 yamamoto end *****
		
        ret = init_sql()
        If ret = False Then Exit Sub
		
        result = gz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
        If result = False Then
            MsgBox("There is no Stamp drawing.", MsgBoxStyle.Critical, "ｴﾗｰ")
        Else
            dataset_F_GZDELE()
        End If
		
		end_sql()
		
	End Sub
	
	Private Sub w_no2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_no2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim result As Object
		Dim ret As Object
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
			
            If Trim(form_no.w_no1.Text) = "" Then GoTo EventExitSub
            If Trim(form_no.w_no2.Text) = "" Then GoTo EventExitSub
			
            ret = init_sql()
            If ret = False Then GoTo EventExitSub
			
            result = gz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
            If result = False Then
                MsgBox("There is no Stamp drawing.", MsgBoxStyle.Critical, "ｴﾗｰ")
            Else
                dataset_F_GZDELE()
            End If
			
			end_sql()
			
		End If
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_no2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no2.Leave
		Dim result As Object
		Dim ret As Object
		Dim irt As Object
		
		Dim f As System.Windows.Forms.Control
		
		If TypeOf form_no.ActiveControl Is System.Windows.Forms.Button Then
			Exit Sub
		End If
		
		'***** 12/8 1997 yamamoto start *****
        If Trim(form_no.w_no2.Text) = "" Then Exit Sub
		
		If w_no1_flg = 1 Then Exit Sub
		
        f = form_no.w_no2
		w_no2.Text = Trim(w_no2.Text)
        irt = check_0((w_no2.Text), 2, 0, f)
        If irt <> 0 Then
            w_no2_flg = 1
            MsgBox("Code is invalid.", 64, "Input error")
            Exit Sub
        Else
            w_no2_flg = 0
        End If
		
        If Trim(form_no.w_no1.Text) = "" Then Exit Sub
		
		'***** 12/8 1997 yamamoto end *****
		
        ret = init_sql()
        If ret = False Then Exit Sub
		
        result = gz_search(form_no.w_no1.Text, form_no.w_no2.Text, 0)
        If result = False Then
            MsgBox("There is no Stamp drawing.", MsgBoxStyle.Critical, "ｴﾗｰ")
        Else
            dataset_F_GZDELE()
        End If
		
		end_sql()
		
	End Sub
End Class