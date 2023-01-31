Option Strict Off
Option Explicit On
Friend Class F_ZMNCALL
	Inherits System.Windows.Forms.Form

	Private Sub w_no1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no1.Leave
		
		form_no.w_no1.Text = UCase(Trim(form_no.w_no1.Text))
		
	End Sub
	
	Private Sub z_end_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles z_end.Click
		form_no.Close()
		End
	End Sub
	
    'UPGRADE_WARNING: イベント w_taisho.TextChanged は、フォームが初期化されたときに発生します。
    Private Sub w_taisho_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_taisho.TextChanged

        If (Trim(w_taisho.Text) <> "Stamp drawing") And (Trim(w_taisho.Text) <> "Editing characters drawing") And (Trim(w_taisho.Text) <> "Brand drawing") Then
            w_id.Text = ""
        End If
    End Sub
	
    'UPGRADE_WARNING: イベント w_taisho.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_taisho_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_taisho.SelectedIndexChanged
		
		Select Case w_taisho.SelectedIndex
			Case 0
				w_id.Text = "KO"
			Case 1
				w_id.Text = "HE"
			Case 2
				w_id.Text = "AT-B"
		End Select
		
	End Sub
	
	Private Sub w_taisho_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_taisho.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_taisho, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub z_read_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles z_read.Click
		Dim ret As Short
		
		If Trim(w_id.Text) = "" Then
            MsgBox("Please select your search drawing.")
			w_taisho.Focus()
			Exit Sub '----- 12/11 1997 yamamoto add -----
		End If
		If Trim(w_no1.Text) = "" Then
            MsgBox("Please enter the number.")
			w_no1.Focus()
			Exit Sub '----- 12/11 1997 yamamoto add -----
		End If
		
		' -> watanabe Edit 2007.03
		'   If Len(w_no1.Text) <> 4 Then
		'      MsgBox "番号の桁数が不正です。"
		'      w_no1.SetFocus
		'      Exit Sub  '----- 12/11 1997 yamamoto add -----
		'   End If
		If Trim(w_id.Text) = "AT-B" Then
			' -> watanabe Edit 2007.06
			'        If Len(w_no1.Text) <> 5 Then
			If Len(w_no1.Text) <> 4 And Len(w_no1.Text) <> 5 Then
				' <- watanabe Edit 2007.06
                MsgBox("Number of digits in the number is wrong.")
				w_no1.Focus()
				Exit Sub '----- 12/11 1997 yamamoto add -----
			End If
		Else
			If Len(w_no1.Text) <> 4 Then
                MsgBox("Number of digits in the number is wrong.")
				w_no1.Focus()
				Exit Sub '----- 12/11 1997 yamamoto add -----
			End If
		End If
		' -> watanabe Edit 2007.03
		
		If Trim(w_no2.Text) = "" Then
            MsgBox("Enter the revision number.")
			w_no2.Focus()
			Exit Sub '----- 12/11 1997 yamamoto add -----
		End If
		If Len(w_no2.Text) <> 2 Then
            MsgBox("Number of digits of revision number is wrong.")
			w_no2.Focus()
			Exit Sub '----- 12/11 1997 yamamoto add -----
		End If
		If (w_no2.Text < "0") Or (w_no2.Text > "99") Then
            MsgBox("Please enter a　numerical value.")
			w_no2.Focus()
			Exit Sub '----- 12/11 1997 yamamoto add -----
		End If
		
		init_sql()
		Select Case Trim(w_id.Text)
			Case "KO"
				' Brand Ver.3 変更
				'       DBTableName = DBName & "..gz_kanri"
				DBTableName = DBName & "..gz_kanri1"
				DBTableName2 = DBName & "..gz_kanri2"
				ret = gz_read((w_id.Text), (w_no1.Text), (w_no2.Text))
			Case "HE"
				' Brand Ver.3 変更
				'       DBTableName = DBName & "..hz_kanri"
				DBTableName = DBName & "..hz_kanri1"
				DBTableName2 = DBName & "..hz_kanri2"
				ret = hz_read((w_id.Text), (w_no1.Text), (w_no2.Text))
			Case "AT-B"
				' Brand Ver.3 変更
				'       DBTableName = DBName & "..bz_kanri"
				DBTableName = DBName & "..bz_kanri1"
				DBTableName2 = DBName & "..bz_kanri2"
				ret = bz_read((w_id.Text), (w_no1.Text), (w_no2.Text))
		End Select
		end_sql()
		If ret = FAIL Then
			w_taisho.Focus()
		Else

            ' -> watanabe edit VerUP(2011)
            'Me.Enabled = False
            w_taisho.Enabled = False
            w_id.Enabled = False
            w_no1.Enabled = False
            w_no2.Enabled = False
            z_read.Enabled = False
            ' <- watanabe edit VerUP(2011)

        End If
		
	End Sub
	
	Private Sub F_ZMNCALL_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
        form_no = Me

        '20100702追加コード
        temp_gz.Initilize()
        temp_hz.Initilize()
        temp_bz.Initilize()


		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		'ｺﾝﾎﾞﾎﾞｯｸｽ
		w_taisho.Items.Clear()
        w_taisho.Items.Add("Stamp drawing")
        w_taisho.Items.Add("Editing characters drawing")
        w_taisho.Items.Add("Brand drawing")
		
		'項目クリア
		w_taisho.Text = ""
		w_id.Text = ""
		w_no1.Text = ""
		w_no2.Text = ""
		
	End Sub
End Class