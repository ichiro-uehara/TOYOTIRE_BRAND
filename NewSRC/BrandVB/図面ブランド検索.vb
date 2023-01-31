Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_ZSEARCH_BRAND
	Inherits System.Windows.Forms.Form
	
	'概要：ボタンクリック処理
	'説明：キャンセルフラグを立てる
	'----- 1/28 1997 by yamamoto -----
	Private Sub cmd_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Cancel.Click
		
		GL_cancel_flg = 1
		
	End Sub
	
	'概要：ボタンクリック処理
	'説明：追加項目：ロックセット、キャンセルが選択されると、処理を中止する（ グリッドはクリア ）
	Private Sub cmd_Search_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Search.Click
		Dim lp As Object
		Dim j As Object
		Dim num As String

        ' -> watanabe del VerUP(2011)
        'Dim result As Object
        ' <- watanabe del VerUP(2011)

		Dim search_word(12) As String
		Dim L_DAT(20) As String
		Dim i_cnt, i As Short
		Dim ic As Object
		Dim ir As Short
		Dim srh_cnt As Integer
		Dim w_ret As Short
		Dim index_row As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)


		GL_cancel_flg = 0
		srh_cnt = 0
		
		If check_F_ZSEARCH_BRAND <> 0 Then Exit Sub
		MSFlexGrid1.Redraw = False
		
		'  MsgBox "検索します"
		MSFlexGrid1.Rows = 2
		For i = 0 To MSFlexGrid1.Cols - 1
			w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
		Next i
		MSFlexGrid1.Enabled = False
		' Brand Ver.3 変更
		'   DBTableName = DBName & "..bz_kanri"
		DBTableName = DBName & "..bz_kanri1"
		DBTableName2 = DBName & "..bz_kanri2"
		
		init_sql()
		
		i_cnt = 0
		search_word(i_cnt) = " WHERE flag_delete = 0 "
		If Trim(w_pattern.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND pattern LIKE '" & Trim(w_pattern.Text) & "%' "
		End If
		If Trim(w_size1.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size1 = '" & Trim(w_size1.Text) & "' "
		End If
		If Trim(w_size2.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size2 = '" & Trim(w_size2.Text) & "' "
		End If
		If Trim(w_size3.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size3 = '" & Trim(w_size3.Text) & "' "
		End If
		If Trim(w_size4.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size4 = '" & Trim(w_size4.Text) & "' "
		End If
		If Trim(w_size5.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size5 = '" & Trim(w_size5.Text) & "' "
		End If
		If Trim(w_size6.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND size6 = '" & Trim(w_size6.Text) & "' "
		End If
		If Trim(w_kanri_no.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND kanri_no LIKE '" & Trim(w_kanri_no.Text) & "%' "
		End If
		If Trim(w_entry_name.Text) <> "" Then
			i_cnt = i_cnt + 1
			search_word(i_cnt) = "AND entry_name = '" & Trim(w_entry_name.Text) & "' "
		End If
        If Trim(w_entry_date_0.Text) <> "" Then
            i_cnt = i_cnt + 1
            search_word(i_cnt) = "AND entry_date >= '" & Trim(w_entry_date_0.Text) & " 00:00" & "' "
        End If
        If Trim(w_entry_date_1.Text) <> "" Then
            i_cnt = i_cnt + 1
            search_word(i_cnt) = "AND entry_date <= '" & Trim(w_entry_date_1.Text) & " 23:59" & "' "
        End If

        'ﾃﾞｰﾀﾍﾞｰｽ該当件数を表示


        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT COUNT(*) FROM " & DBTableName)
        'For i = 0 To i_cnt
        '          result = SqlCmd(SqlConn, search_word(i))
        'Next i
        '
        '      result = SqlExec(SqlConn)
        '      result = SqlResults(SqlConn)
        '
        '      If result = SUCCEED Then
        '          Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '              num = SqlData(SqlConn, 1)
        '              w_total.Text = num
        '          Loop
        '      Else
        '          MsgBox("検索に失敗しました")
        '          GoTo error_section
        '      End If


        '検索コマンド作成
        sqlcmd = "SELECT COUNT(*) FROM " & DBTableName
        For i = 0 To i_cnt
            sqlcmd = sqlcmd & search_word(i)
        Next i

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        w_total.Text = "-1"
        Do Until Rs.EOF

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                num = CStr(Val(Rs.rdoColumns(0).Value))
            Else
                num = "-1"
            End If
            w_total.Text = num

            Rs.MoveNext()
        Loop

        Rs.Close()

        If w_total.Text = "-1" Then
            MsgBox("Failed to find.")
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        If CDbl(w_total.Text) = 0 Then
            MsgBox("There is no brand drawing corresponding.")
            GoTo error_section
        Else
            If w_total.Text > AskNum Then
                w_ret = MsgBox("There is " & w_total.Text & " data. Would you like to view?", MsgBoxStyle.YesNo, "Confirmation")
                If w_ret = MsgBoxResult.No Then
                    MsgBox("Canceled the search.", , "Cancel")
                    w_total.Text = ""
                    GoTo error_section
                End If
            End If
        End If
        w_ret = co_rockset_F_ZSEARCH(1, 1)
		
		'ｸﾞﾘｯﾄﾞに検索内容表示
		If CDbl(w_total.Text) > 0 Then
			MSFlexGrid1.Rows = Int((CDbl(w_total.Text) - 1) / 2) + 2
		Else
			MSFlexGrid1.Rows = 2
			For i = 0 To MSFlexGrid1.Cols - 1
				w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
			Next i
		End If
		
		MSFlexGrid1.set_RowHeight(-1, 300)
		index_row = "; NO "


		'検索コマンド作成
		sqlcmd = "SELECT id, no1, no2"
        sqlcmd = sqlcmd & " FROM " & DBTableName
        For i = 0 To i_cnt
            sqlcmd = sqlcmd & search_word(i)
        Next i
        sqlcmd = sqlcmd & " ORDER BY id, no1, no2"

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        i = 0
        ic = 0
        ir = 0
        Do Until Rs.EOF

            System.Windows.Forms.Application.DoEvents()
            If GL_cancel_flg = 1 Then GoTo cancel_end_section

            i = i + 1
            ic = (i - 1) - Int((i - 1) / 2) * 2 + 1
            ir = Int((i - 1) / 2) + 1
            w_ret = Set_Grid_Data(MSFlexGrid1, "", ir, (ic - 1) * 3 + 1)
            w_ret = Set_Grid_Data(MSFlexGrid1, "◇", ir, (ic - 1) * 3 + 2)

            For j = 1 To 3
                If IsDBNull(Rs.rdoColumns(j - 1).Value) = False Then
                    L_DAT(j) = Rs.rdoColumns(j - 1).Value
                Else
                    L_DAT(j) = ""
                End If
            Next j

            w_ret = Set_Grid_Data(MSFlexGrid1, L_DAT(1) & "-" & L_DAT(2) & "-" & L_DAT(3), ir, 3 + (ic - 1) * 3)

            srh_cnt = srh_cnt + 1
            w_total.Text = CStr(srh_cnt)
			If (srh_cnt + 1) Mod 2 = 0 Then
				'----- .NET 移行 -----
				'index_row = index_row & "|" & VB6.Format((srh_cnt + 1) / 2)
				index_row = index_row & "|" & ((srh_cnt + 1) / 2).ToString()
			End If

			Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        MSFlexGrid1.FormatString = index_row
        MSFlexGrid1.set_FixedAlignment(0, 4)
        MSFlexGrid1.Redraw = True
        MSFlexGrid1.Enabled = True

error_section:

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        MSFlexGrid1.Redraw = True
        w_ret = co_rockset_F_ZSEARCH(1, 0)

        Exit Sub

cancel_end_section:

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        MSFlexGrid1.Rows = 2
        MSFlexGrid1.Cols = 7
        For lp = 0 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Row = 1
            MSFlexGrid1.Col = lp
            MSFlexGrid1.Text = ""
        Next lp
        MSFlexGrid1.Redraw = True
        w_total.Text = ""
        w_ret = co_rockset_F_ZSEARCH(1, 0)
        MsgBox("Search has been canceled.", 64, "Cancel")

	End Sub
	
	
	Private Sub cmd_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Clear.Click
		
		Call Clear_F_ZSEARCH_BRAND()
		
	End Sub
	
	Private Sub cmd_End_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_End.Click
		
		form_no.Close()
		'  form_main.Show
		End
	End Sub
	
	Private Sub cmd_Help_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Help.Click
        On Error Resume Next
        Err.Clear()
        Dim oCommonDialog As Object
        oCommonDialog = CreateObject("MSComDlg.CommonDialog")

        If Err.Number = 0 Then
            With oCommonDialog
                .HelpCommand = cdlHelpContext
                .HelpFile = "c:\VBhelp\BRAND.HLP"
                .HelpContext = 902
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub cmd_ZumenRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_ZumenRead.Click
		Dim w_ret As Object
		
		Dim syori_cnt As Short
		Dim w_str As String
		Dim w_mess As String
		Dim ZumenName As String
		Dim error_no As String
		Dim i, j As Short
		Dim time_start As Date
		Dim time_now As Date
		Dim w_name As String
		Dim w_err As String


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w_str = ""
        w_name = ""
        ' <- watanabe add VerUP(2011)


		For i = 1 To MSFlexGrid1.Rows - 1
			For j = 1 To 2
				If syori_cnt >= FreePicNum Then
                    MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures=" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
					Exit Sub
				End If
                w_ret = Get_Grid_Data(MSFlexGrid1, w_err, i, 3 * (j - 1) + 1) 'エラーフラグ
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 3 * (j - 1) + 2) '読込フラグ
                w_ret = Get_Grid_Data(MSFlexGrid1, w_name, i, 3 * (j - 1) + 3) '図面名
				If w_str = "◆" And w_err = "" Then
					ZumenName = w_name
					' -> watanabe edit 2007.03
					'            w_mess = ZumenName
                    w_mess = BrandDir & ZumenName
					' <- watanabe edit 2007.03
					
                    w_ret = PokeACAD("MDLREAD", w_mess)
                    w_ret = RequestACAD("MDLREAD")
					
					time_start = Now
					Do 
						time_now = Now
                        If Trim(form_main.Text2.Text) = "" Then
                            If System.DateTime.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.DateTime.FromOADate(timeOutSecond) Then
                                MsgBox("Time-out error.", 64, "ERROR")
                                w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                                w_ret = RequestACAD("ERROR")
                                Exit Sub
                            End If
                        ElseIf VB.Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then
                            w_ret = Set_Grid_Data(MSFlexGrid1, "0", i, 3 * (j - 1) + 1)
                            form_main.Text2.Text = ""
                            FreePicNum = FreePicNum - 1
                            Exit Do
                        ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                            error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                            w_ret = Set_Grid_Data(MSFlexGrid1, error_no, i, 3 * (j - 1) + 1)
                            form_main.Text2.Text = ""
                            Exit Do
                        Else
                            MsgBox("Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
                            w_ret = Set_Grid_Data(MSFlexGrid1, "?", i, 3 * (j - 1) + 1)
                            form_main.Text2.Text = ""
                            Exit Sub
                        End If
					Loop 
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 3 * (j - 1) + 1)
					'図面読み込みＯＫ
					If Val(w_str) = 0 Then
						cmd_Search.Enabled = False
						cmd_Clear.Enabled = False
						cmd_Help.Enabled = False
						cmd_ZumenRead.Enabled = False
						w_pattern.Enabled = False
						w_pattern.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size1.Enabled = False
						w_size1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size2.Enabled = False
						w_size2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size3.Enabled = False
						w_size3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size4.Enabled = False
						w_size4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size5.Enabled = False
						w_size5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_size6.Enabled = False
						w_size6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_kanri_no.Enabled = False
						w_kanri_no.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_entry_name.Enabled = False
						w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                        w_entry_date_0.Enabled = False
                        w_entry_date_0.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                        w_entry_date_1.Enabled = False
                        w_entry_date_1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_total.Enabled = False
						w_total.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						MSFlexGrid1.Enabled = False
					End If
					syori_cnt = syori_cnt + 1
				End If
			Next j
		Next i
		
		If syori_cnt = 0 Then
            MsgBox("Data to be read is not selected.")
		Else
            MsgBox("CAD reading completion.")
		End If
		
	End Sub
	'概要：フォームロード
	'説明：追加項目：ロックセット
	Private Sub F_ZSEARCH_BRAND_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim w_ret As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim index_col As String
		
        form_no = Me
        temp_bz.Initilize() '20100702追加コード
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		Call Clear_F_ZSEARCH_BRAND()
		
		MSFlexGrid1.Rows = 2
		MSFlexGrid1.Cols = 7
		
		' 行高さの設定
		MSFlexGrid1.set_RowHeight(-1, 300)
		
        index_col = "^NO|^error|^Read|^Drawing name|^error|^Read|^Drawing name"
		MSFlexGrid1.FormatString = index_col
		
		' 列幅の設定
		MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 4)
		MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(6, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 4)
		
		MSFlexGrid1.Enabled = False
        w_ret = co_rockset_F_ZSEARCH(1, 0)
		
        form_main.Text2.Text = ""
		CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")
		
	End Sub

	'----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
#If False Then
	Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		Dim w_ret As Object
		
		Dim w_col As Short
		Dim w_row As Short
		Dim w_err As String
		Dim w1 As String
		Dim w2 As String
		Dim i As Short


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w1 = ""
        w2 = ""
        ' <- watanabe add VerUP(2011)


		If eventArgs.KeyAscii <> 32 Then Exit Sub
		
		w_col = Val(CStr(MSFlexGrid1.Col))
		w_row = Val(CStr(MSFlexGrid1.Row))
		
		MSFlexGrid1.Redraw = False
		
		If w_col = 2 Or w_col = 5 Then
            w_ret = Get_Grid_Data(MSFlexGrid1, w_err, w_row, w_col - 1)
			If w_err = "" Then
				If MSFlexGrid1.Text = "◆" Then
					MSFlexGrid1.Text = "◇"
				ElseIf MSFlexGrid1.Text = "◇" Then 
					For i = 1 To MSFlexGrid1.Rows - 1
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
						If w1 = "" And w2 = "◆" Then
                            w_ret = Set_Grid_Data(MSFlexGrid1, "◇", i, 2)
						End If
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 4)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 5)
						If w1 = "" And w2 = "◆" Then
                            w_ret = Set_Grid_Data(MSFlexGrid1, "◇", i, 5)
						End If
					Next i
					MSFlexGrid1.Text = "◆"
				End If
			End If
		End If
		
		MSFlexGrid1.Redraw = True
		
	End Sub
#End If

	'----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
#If False Then
	Private Sub MSFlexGrid1_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent) Handles MSFlexGrid1.MouseDownEvent
		Dim w_select As Object
		Dim w_ret As Object
		
		Dim w_col As Short
		Dim w_row As Short
		Dim w_err As String
		Dim w1 As String
		Dim w2 As String
		Dim i As Short


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w1 = ""
        w2 = ""
        ' <- watanabe add VerUP(2011)


		w_col = Val(CStr(MSFlexGrid1.Col))
		w_row = Val(CStr(MSFlexGrid1.Row))
		
		MSFlexGrid1.Redraw = False
		
		If w_col = 2 Or w_col = 5 Then
            w_ret = Get_Grid_Data(MSFlexGrid1, w_err, w_row, w_col - 1)
			If w_err = "" Then
				If MSFlexGrid1.Text = "◆" Then
					MSFlexGrid1.Text = "◇"
				ElseIf MSFlexGrid1.Text = "◇" Then 
                    w_select = 0
					For i = 1 To MSFlexGrid1.Rows - 1
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
						If w1 = "" And w2 = "◆" Then
                            w_ret = Set_Grid_Data(MSFlexGrid1, "◇", i, 2)
						End If
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 4)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 5)
						If w1 = "" And w2 = "◆" Then
							w_ret = Set_Grid_Data(MSFlexGrid1, "◇", i, 5)
						End If
					Next i
					MSFlexGrid1.Text = "◆"
				End If
			End If
		End If
		
		MSFlexGrid1.Redraw = True
		
	End Sub
#End If
	Private Sub w_pattern_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_pattern.Leave
		
		form_no.w_pattern.Text = UCase(Trim(form_no.w_pattern.Text))
		
	End Sub
	
	Private Sub w_size1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.Leave
		
		form_no.w_size1.Text = UCase(Trim(form_no.w_size1.Text))
		
	End Sub
	
	Private Sub w_size2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.Leave
		
		form_no.w_size2.Text = UCase(Trim(form_no.w_size2.Text))
		
	End Sub
	
	Private Sub w_size3_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.Leave
		
		form_no.w_size3.Text = UCase(Trim(form_no.w_size3.Text))
		
	End Sub
	
	Private Sub w_size4_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.Leave
		
		form_no.w_size4.Text = UCase(Trim(form_no.w_size4.Text))
		
	End Sub
	
	Private Sub w_size5_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size5.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii > 32 Then
			If (KeyAscii = CDbl("100")) Or (KeyAscii = CDbl("114")) Then
				KeyAscii = KeyAscii - 32
			ElseIf (KeyAscii <> CDbl("45")) And (KeyAscii <> CDbl("68")) And (KeyAscii <> CDbl("82")) Then 
				KeyAscii = 0
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size5_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.Leave
		
		form_no.w_size5.Text = UCase(Trim(form_no.w_size5.Text))
		
	End Sub
	
	Private Sub w_size6_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.Leave
		
		form_no.w_size6.Text = UCase(Trim(form_no.w_size6.Text))
		
	End Sub
End Class