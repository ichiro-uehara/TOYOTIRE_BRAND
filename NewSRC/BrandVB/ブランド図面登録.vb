Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class F_BZSAVE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim ZumenName As Object
		Dim w_ret As Object
		Dim henban As Object
		Dim result As Object
        Dim wk_size As String
		Dim w_mess As String
		Dim hex_data As New VB6.FixedLengthString(109)

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


        '------< 登録ボタン >------------------------------


		'規格チェック用 特性データ 送信
		'     Call temp_bz_get(4)
		'     Call bz_spec_set(hex_data)
		
		'     w_ret = PokeACAD("SPECADD", hex_data)
		'     w_ret = RequestACAD("SPECADD")
		'    CommunicateMode = comSpecData
		
		'規格チェック
		'    If KIKAKU_CHK <> 0 Then
		'       form_no.Enabled = False
		'       F_MSG3.Show
		'     CommunicateMode = comNone
		'       w_ret = RequestACAD("BZKIKAKU")
		'    End If
		
		If Trim(w_no2.Text) = "" Then
			w_no2.Text = "00"
		End If
		
		F_MSG3.Close()
		form_no.Enabled = True
		
		w_size.Text = Trim(w_size1.Text) & Trim(w_size2.Text) & Trim(w_size3.Text) & Trim(w_size4.Text) & Trim(w_size5.Text) & Trim(w_size6.Text) & Trim(w_size7.Text) & Trim(w_size8.Text)
		
		Call init_sql()


        ' -> watanabe edit VerUP(2011)
        '      SqlFreeBuf((SqlConn))
        'result = SqlCmd(SqlConn, "SELECT size_code ")
        'result = SqlCmd(SqlConn, " FROM " & STANDARD_DBName & "..size_code")
        'result = SqlCmd(SqlConn, " WHERE ")
        'result = SqlCmd(SqlConn, " syurui = '" & Trim(w_syurui.Text) & "' AND")
        'result = SqlCmd(SqlConn, " size1 = '" & w_size1.Text & "' AND")
        'result = SqlCmd(SqlConn, " size2 = '" & w_size2.Text & "' AND")
        'result = SqlCmd(SqlConn, " size3 = '" & w_size3.Text & "' AND")
        'result = SqlCmd(SqlConn, " size5 = '" & w_size5.Text & "' AND")
        'result = SqlCmd(SqlConn, " size6 = '" & w_size6.Text & "'")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		wk_size = SqlData(SqlConn, 1)
        '	Else
        '		'1999/6/10 yamamoto.f syusei MS対応（ｻｲｽﾞｴﾗｰﾁｪｯｸを外す） Start
        '		'          MsgBox "サイズが登録されていません"
        '		'          Exit Sub
        '		wk_size = ""
        '		'1999/6/10 yamamoto.f syusei MS対応（ｻｲｽﾞｴﾗｰﾁｪｯｸを外す） End
        '	End If
        'Else
        '	Exit Sub
        'End If


        '検索キーセット
        key_code = " syurui = '" & Trim(w_syurui.Text) & "' AND"
        key_code = key_code & " size1 = '" & w_size1.Text & "' AND"
        key_code = key_code & " size2 = '" & w_size2.Text & "' AND"
        key_code = key_code & " size3 = '" & w_size3.Text & "' AND"
        key_code = key_code & " size5 = '" & w_size5.Text & "' AND"
        key_code = key_code & " size6 = '" & w_size6.Text & "'"

        '検索コマンド作成
        sqlcmd = "SELECT size_code FROM " & STANDARD_DBName & "..size_code WHERE " & key_code

        'ヒット数チェック
        cnt = VBRDO_Count(GL_T_RDO, STANDARD_DBName & "..size_code", key_code)
        If cnt = 0 Then
            wk_size = ""

        ElseIf cnt = -1 Then
            Exit Sub

        Else
            '検索
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                wk_size = Rs.rdoColumns(0).Value
            Else
                wk_size = ""
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        w_size_code.Text = wk_size
        Call end_sql()


        ' -> watanabe add VerUP(2011)
        result = FAIL
        ' <- watanabe add VerUP(2011)


        init_sql()
        If open_mode = "NEW" Then
            If check_F_BZSAVE() <> 0 Then
                end_sql()
                Exit Sub
            Else
                result = bz_insert()
            End If

        ElseIf open_mode = "Revision number" Then
            If check_F_BZSAVE() <> 0 Then
                end_sql()
                Exit Sub
            Else
                If SqlConn = 0 Then
                    MsgBox("Can not access the database.", MsgBoxStyle.Critical, "SQL error")
                    Exit Sub
                End If

                '変番を自動連番します
                henban = what_no2_BZ(temp_bz.no1)

                If henban = -1 Then
                    MsgBox("Failed to auto sequence number of a variable number.")
                    Exit Sub
                ElseIf henban = 0 Then
                    MsgBox("Drawing is not registered. Please sign up for new.", MsgBoxStyle.Critical, "Drawing Unregistered")
                    Exit Sub
                End If

				'----- .NET 移行 -----
				'form_no.w_no2.Text = VB6.Format(henban, "00")
				form_no.w_no2.Text = henban.ToString("00")

				temp_bz.no2 = form_no.w_no2.Text

                result = bz_insert()
            End If

        ElseIf open_mode = "modify" Then
            If check_F_BZSAVE() <> 0 Then
                end_sql()
                Exit Sub
            Else
                result = bz_update()
            End If
        End If

        If result = FAIL Then
            MsgBox("Failed D / B register of brand drawing.", 64, "registration error")
        Else

            '登録用 特性データ 送信
            Call temp_bz_get(4)
            Call bz_spec_set(hex_data.Value)

            w_ret = PokeACAD("SPECADD", hex_data.Value)
            w_ret = RequestACAD("SPECADD")

            '（図面名）送信 ＆ 図面セーブ
            '     ZumenName = "AT-B" & Trim(form_no.w_no1) & "-" & Trim(form_no.w_no2)
            ZumenName = "AT-B-" & Trim(form_no.w_no1.Text) & "-" & Trim(form_no.w_no2.Text)
            w_mess = BrandDir & ZumenName
            w_ret = PokeACAD("MDLSAVE", w_mess)
            If w_ret = 1 Then MsgBox("Request handle error occurred.") '20100705追加コード
            w_ret = RequestACAD("MDLSAVE")


            '画面ロック
            form_no.Command1.Enabled = False
            form_no.Command2.Enabled = False
            form_no.Command4.Enabled = False
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kanri_no.Enabled = False
            form_no.w_kanri_no.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_comment.Enabled = False
            form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_dep_name.Enabled = False
            form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_syurui.Enabled = False
            form_no.w_syurui.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_pattern.Enabled = False
            form_no.w_pattern.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_syubetu.Enabled = False
            form_no.w_syubetu.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size1.Enabled = False
            form_no.w_size1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size2.Enabled = False
            form_no.w_size2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size3.Enabled = False
            form_no.w_size3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size4.Enabled = False
            form_no.w_size4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size5.Enabled = False
            form_no.w_size5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size6.Enabled = False
            form_no.w_size6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size7.Enabled = False
            form_no.w_size7.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_size8.Enabled = False
            form_no.w_size8.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_plant.Enabled = False
            form_no.w_plant.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku1.Enabled = False
            form_no.w_kikaku1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku2.Enabled = False
            form_no.w_kikaku2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku3.Enabled = False
            form_no.w_kikaku3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku4.Enabled = False
            form_no.w_kikaku4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku5.Enabled = False
            form_no.w_kikaku5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_kikaku6.Enabled = False
            form_no.w_kikaku6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_tos_moyou.Enabled = False
            form_no.w_tos_moyou.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_peak_mark.Enabled = False
            form_no.w_peak_mark.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_side_moyou.Enabled = False
            form_no.w_side_moyou.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_side_kenti.Enabled = False
            form_no.w_side_kenti.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_nasiji.Enabled = False
            form_no.w_nasiji.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

            MsgBox("Registered the brand drawing. (D / B · drawing)")

            '     Call temp_bz_get(4)
            '     Call bz_spec_set(hex_data)

            '     w_ret = PokeACAD("SPECADD", hex_data)
            '     w_ret = RequestACAD("SPECADD")

        End If
        end_sql()


        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        MsgBox(Err.Description, MsgBoxStyle.Critical, "System error")

        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

    End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_BZSAVE()
		
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
                .HelpContext = 700
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_BZSAVE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim i As Object
		Dim lp As Object
		Dim w_ret As Object
		
		'--------< ブランド図面 登録 画面ＬＯＡＤ >----------------------
		
		Dim aa As String

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)

        form_no = Me
        temp_bz.Initilize() '20100702追加コード
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 4) ' フォームを画面の縦方向にセンタリングします。
		
		Text1.Text = open_mode
		
		Call Clear_F_BZSAVE()
		
		w_id.Text = "AT-B"
        If open_mode = "NEW" Then
            w_ret = PokeACAD("SAVEMODE", "FRESH")
            RequestACAD("SAVEMODE")

            form_no.w_no1.Text = ""
            form_no.w_no2.Text = "00"
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_no2.Enabled = False
            form_no.w_kanri_no.Text = ""
            Call true_date(aa)
            form_no.w_entry_date.Text = aa

        ElseIf open_mode = "modify" Then
            w_ret = PokeACAD("SAVEMODE", "MODIFY")
            RequestACAD("SAVEMODE")
            '     RequestACAD ("ZMNNAME")
            CommunicateMode = comSpecData

            Call true_date(aa)
            w_entry_date.Text = aa
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

        ElseIf open_mode = "Revision number" Then
            w_ret = PokeACAD("SAVEMODE", "CHANGE")
            RequestACAD("SAVEMODE")
            '     RequestACAD ("ZMNNAME")
            CommunicateMode = comSpecData

            Call true_date(aa)
            w_entry_date.Text = aa
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        Else
            MsgBox("Open mode error! !")
            Exit Sub
        End If
		
		'タイヤ種類
		w_syurui.Items.Clear()
		w_syurui.Items.Add("PC")
		w_syurui.Items.Add("LT")
		w_syurui.Items.Add("TB")
		w_syurui.Text = VB6.GetItemString(w_syurui, 0)
		
		'パターン種別
		w_syubetu.Items.Clear()
		w_syubetu.Items.Add("SM")
		w_syubetu.Items.Add("AS")
		w_syubetu.Items.Add("M+S")
		w_syubetu.Items.Add("STL")
		w_syubetu.Items.Add("SNW")
		w_syubetu.Items.Add("TP")
		w_syubetu.Items.Add("NP")
		w_syubetu.Text = VB6.GetItemString(w_syubetu, 0)
		
		'工場
		w_plant.Items.Clear()
        w_plant.Items.Add("Sendai(TT)")
        w_plant.Items.Add("Kuwana(KW)")
        w_plant.Items.Add("Cheng shin(CS)")
        w_plant.Items.Add("Shanghai(CH)")
		w_plant.Text = VB6.GetItemString(w_plant, 0)
		
		'規格
		w_kikaku1.Items.Clear()
        w_kikaku1.Items.Add(" :Blank")
		w_kikaku1.Items.Add("J:JIS")
		w_kikaku1.Items.Add("E:ECE")
		w_kikaku1.Items.Add("F:FMVSS")
		w_kikaku1.Items.Add("I:INMETRO")
		w_kikaku1.Items.Add("G:GULF")
		w_kikaku1.Items.Add("A:ADR")
		w_kikaku1.Text = VB6.GetItemString(w_kikaku1, 0)
		
		w_kikaku2.Items.Clear()
        w_kikaku2.Items.Add(" :Blank")
		w_kikaku2.Items.Add("J:JIS")
		w_kikaku2.Items.Add("E:ECE")
		w_kikaku2.Items.Add("F:FMVSS")
		w_kikaku2.Items.Add("I:INMETRO")
		w_kikaku2.Items.Add("G:GULF")
		w_kikaku2.Items.Add("A:ADR")
		w_kikaku2.Text = VB6.GetItemString(w_kikaku2, 0)
		
		w_kikaku3.Items.Clear()
        w_kikaku3.Items.Add(" :Blank")
		w_kikaku3.Items.Add("J:JIS")
		w_kikaku3.Items.Add("E:ECE")
		w_kikaku3.Items.Add("F:FMVSS")
		w_kikaku3.Items.Add("I:INMETRO")
		w_kikaku3.Items.Add("G:GULF")
		w_kikaku3.Items.Add("A:ADR")
		w_kikaku3.Text = VB6.GetItemString(w_kikaku3, 0)
		
		w_kikaku4.Items.Clear()
        w_kikaku4.Items.Add(" :Blank")
		w_kikaku4.Items.Add("J:JIS")
		w_kikaku4.Items.Add("E:ECE")
		w_kikaku4.Items.Add("F:FMVSS")
		w_kikaku4.Items.Add("I:INMETRO")
		w_kikaku4.Items.Add("G:GULF")
		w_kikaku4.Items.Add("A:ADR")
		w_kikaku4.Text = VB6.GetItemString(w_kikaku4, 0)
		
		w_kikaku5.Items.Clear()
        w_kikaku5.Items.Add(" :Blank")
		w_kikaku5.Items.Add("J:JIS")
		w_kikaku5.Items.Add("E:ECE")
		w_kikaku5.Items.Add("F:FMVSS")
		w_kikaku5.Items.Add("I:INMETRO")
		w_kikaku5.Items.Add("G:GULF")
		w_kikaku5.Items.Add("A:ADR")
		w_kikaku5.Text = VB6.GetItemString(w_kikaku5, 0)
		
		w_kikaku6.Items.Clear()
        w_kikaku6.Items.Add(" :Blank")
		w_kikaku6.Items.Add("J:JIS")
		w_kikaku6.Items.Add("E:ECE")
		w_kikaku6.Items.Add("F:FMVSS")
		w_kikaku6.Items.Add("I:INMETRO")
		w_kikaku6.Items.Add("G:GULF")
		w_kikaku6.Items.Add("A:ADR")
		w_kikaku6.Text = VB6.GetItemString(w_kikaku6, 0)
		
		'TOS対応模様
		w_tos_moyou.Items.Clear()
        w_tos_moyou.Items.Add("0:Null")
        w_tos_moyou.Items.Add("1:YES")
		w_tos_moyou.Text = VB6.GetItemString(w_tos_moyou, 0)
		
		'サイド凹凸模様
		w_side_moyou.Items.Clear()
        w_side_moyou.Items.Add("0:Null")
        w_side_moyou.Items.Add("1:YES")
		w_side_moyou.Text = VB6.GetItemString(w_side_moyou, 0)
		
		'サイド凹凸検地
		w_side_kenti.Items.Clear()
        w_side_kenti.Items.Add("0:Null")
        w_side_kenti.Items.Add("1:YES")
		w_side_kenti.Text = VB6.GetItemString(w_side_kenti, 0)
		
		'ピークマーク
		w_peak_mark.Items.Clear()
        w_peak_mark.Items.Add("0:Null")
        w_peak_mark.Items.Add("1:YES")
		w_peak_mark.Text = VB6.GetItemString(w_peak_mark, 0)
		
		'梨地
		w_nasiji.Items.Clear()
        w_nasiji.Items.Add("0:Null")
        w_nasiji.Items.Add("1:Electric discharge (N20)")
        w_nasiji.Items.Add("2:Electric discharge (N30)")
        w_nasiji.Items.Add("3:Sand blast")
		w_nasiji.Text = VB6.GetItemString(w_nasiji, 0)
		
		MSFlexGrid1.Rows = 2
        MSFlexGrid1.Cols = 5
		
		' 行高さの設定
		For lp = 0 To MSFlexGrid1.Rows - 1
			MSFlexGrid1.set_RowHeight(lp, 300)
		Next lp
		
		' 列幅の設定
		MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 13 * 1)
		MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 13 * 3)
		MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 13 * 3)
		MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 13 * 3)
        MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 13 * 3)
        For i = 0 To 4
            MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
		
        w_ret = Set_Grid_Data(MSFlexGrid1, "NO", 0, 0)
        w_ret = Set_Grid_Data(MSFlexGrid1, "1", 0, 1)
        w_ret = Set_Grid_Data(MSFlexGrid1, "2", 0, 2)
        w_ret = Set_Grid_Data(MSFlexGrid1, "3", 0, 3)
        w_ret = Set_Grid_Data(MSFlexGrid1, "4", 0, 4)

        CommunicateMode = comSpecData
        RequestACAD("SPECDATA")

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
	
    'UPGRADE_WARNING: イベント w_kikaku1.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku1.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku1, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_kikaku2.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku2.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku2, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_kikaku3.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku3_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku3.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku3, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_kikaku4.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku4_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku4.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku4, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_kikaku5.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku5_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku5.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku5, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_kikaku6.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_kikaku6_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku6.SelectedIndexChanged
		
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
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_kikaku6, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_nasiji_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_nasiji.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_nasiji, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	
	
	
	Private Sub w_no1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_no1.Leave
		
        form_no.w_no1.Text = UCase(Trim(form_no.w_no1.Text))
		
	End Sub
	
	
	Private Sub w_pattern_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_pattern.Leave
		
        form_no.w_pattern.Text = UCase(Trim(form_no.w_pattern.Text))
		
	End Sub
	
	Private Sub w_peak_mark_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_peak_mark.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_peak_mark, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
    'UPGRADE_WARNING: イベント w_plant.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_plant_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_plant.SelectedIndexChanged
		
        If w_plant.Text = "Sendai(TT)" Then
            w_plant_code.Text = "CX"
        ElseIf w_plant.Text = "Kuwana(KW)" Then
            w_plant_code.Text = "N3"
        ElseIf w_plant.Text = "Cheng shin(CS)" Then
            w_plant_code.Text = "UY"
        ElseIf w_plant.Text = "Shanghai(CH)" Then
            w_plant_code.Text = "9T"
        End If
		
	End Sub
	
	Private Sub w_plant_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_plant.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_plant, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_side_kenti_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_side_kenti.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_side_kenti, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_side_moyou_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_side_moyou.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_side_moyou, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
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
			ElseIf (KeyAscii <> CDbl("48")) Or (KeyAscii <> CDbl("68")) Or (KeyAscii <> CDbl("82")) Then 
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
	
	Private Sub w_size7_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size7.Leave
		
        form_no.w_size7.Text = UCase(Trim(form_no.w_size7.Text))
		
	End Sub
	
	Private Sub w_size8_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size8.Leave
		
        form_no.w_size8.Text = UCase(Trim(form_no.w_size8.Text))
		
	End Sub
	
	Private Sub w_syubetu_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syubetu.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_syubetu, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_syurui_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syurui.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_syurui, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_tos_moyou_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_tos_moyou.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_tos_moyou, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class