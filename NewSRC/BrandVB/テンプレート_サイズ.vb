Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_TMP_SIZE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim w_ret As Object
        Dim ss3 As String
        Dim ss2 As String
        Dim ss1 As String
        Dim i As Long
		Dim result As Object
		
		Dim w_text As String
		Dim w_text1 As String
		Dim w_text2 As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        On Error Resume Next
		Err.Clear()
		
		'入力チェック
		If w_size_chk(1) <> 0 Then Exit Sub

        '   MsgBox "検索します。"
		
		'// ｺﾝﾎﾞﾎﾞｯｸｽ初期ｸﾘｱ
		w_hm_name.Items.Clear()
		
		'// 検索
        w_text = Trim(form_no.w_size1.Text) & Trim(form_no.w_size2.Text) & Trim(form_no.w_size3.Text) & Trim(form_no.w_size4.Text) & Trim(form_no.w_size5.Text) & Trim(form_no.w_size6.Text)
		
		'MsgBox "検索するｽﾍﾟﾙ[" & w_text & "]", , "DEBUG"
		
		init_sql()
		

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT font_name, no")
        ''Brand Ver.3 変更
        ''   result = SqlCmd(SqlConn, " FROM " & DBName & "..hm_kanri")
        '      result = SqlCmd(SqlConn, " FROM " & DBName & "..hm_kanri1")
        '      result = SqlCmd(SqlConn, " WHERE ( flag_delete = 0) AND ( spell LIKE '" & w_text & "%')")
        '      result = SqlCmd(SqlConn, " ORDER BY font_name, no")
        '      result = SqlExec(SqlConn)
        '      result = SqlResults(SqlConn)
        '
        '      ' -> watanabe mov VerUP(2011)
        '      i = 0
        '      ' <- watanabe mov VerUP(2011)
        '
        '      If result = SUCCEED Then
        '
        '          ' -> watanabe mov VerUP(2011)
        '          'i = 0
        '          ' -> watanabe mov VerUP(2011)
        '
        '          Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '              i = i + 1
        '              ss1 = SqlData(SqlConn, 1)
        '              ss2 = SqlData(SqlConn, 2)
        '              w_hm_name.Items.Add(ss1 & ss2)
        '          Loop
        '      Else
        '          MsgBox("ﾃﾞｰﾀﾍﾞｰｽSELECTｴﾗｰ", MsgBoxStyle.Critical)
        '      End If


        '検索コマンド作成
        sqlcmd = "SELECT font_name, no"
        sqlcmd = sqlcmd & " FROM " & DBName & "..hm_kanri1"
        sqlcmd = sqlcmd & " WHERE ( flag_delete = 0) AND ( spell LIKE '" & w_text & "%')"
        sqlcmd = sqlcmd & " ORDER BY font_name, no"

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        i = 0

        Do Until Rs.EOF
            i = i + 1

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                ss1 = Rs.rdoColumns(0).Value
            Else
                ss1 = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                ss2 = Rs.rdoColumns(1).Value
            Else
                ss2 = ""
            End If

            w_hm_name.Items.Add(ss1 & ss2)

            Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        '  w_hm_name.Text = w_hm_name.List(0)
        w_hm_num.Text = i
        w_hm_name.SelectedIndex = 0


        '(Brand Cad System Ver.3 UP )
        init_sql()
        w_text1 = VB.Left(Trim(w_hm_name.Text), 6)
        w_text2 = Mid(Trim(w_hm_name.Text), 7, 2)


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "SELECT high, width, ang")
        'result = sqlcmd(SqlConn, " FROM " & DBName & "..hm_kanri1")
        'result = sqlcmd(SqlConn, " WHERE ( flag_delete = 0 ) AND ")
        'result = sqlcmd(SqlConn, " ( font_name  = '" & w_text1 & "') AND ")
        'result = sqlcmd(SqlConn, " ( no  = '" & w_text2 & "')")
        'result = sqlcmd(SqlConn, " ORDER BY high, width, ang")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        'Brand System Ver.3 追加
        '        ss1 = SqlData(SqlConn, 1)
        '        w_hight.Text = ss1
        '        ss2 = SqlData(SqlConn, 2)
        '        w_width.Text = ss2
        '        ss3 = SqlData(SqlConn, 3)
        '        w_ang.Text = ss3
        '    Loop
        'Else
        '    MsgBox("ﾃﾞｰﾀﾍﾞｰｽSELECTｴﾗｰ", MsgBoxStyle.Critical)
        'End If


        '検索コマンド作成
        sqlcmd = "SELECT high, width, ang"
        sqlcmd = sqlcmd & " FROM " & DBName & "..hm_kanri1"
        sqlcmd = sqlcmd & " WHERE ( flag_delete = 0 ) AND "
        sqlcmd = sqlcmd & " ( font_name  = '" & w_text1 & "') AND "
        sqlcmd = sqlcmd & " ( no  = '" & w_text2 & "')"
        sqlcmd = sqlcmd & " ORDER BY high, width, ang"

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        Do Until Rs.EOF

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                ss1 = Rs.rdoColumns(0).Value
            Else
                ss1 = ""
            End If
            w_hight.Text = ss1

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                ss2 = Rs.rdoColumns(1).Value
            Else
                ss2 = ""
            End If
            w_width.Text = ss2

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                ss3 = Rs.rdoColumns(2).Value
            Else
                ss3 = ""
            End If
            w_ang.Text = ss3

            Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        If Val(w_hm_num.Text) = 0 Then
            form_no.w_size1.Focus() 'SetFocus()
        Else
            form_main.Text2.Text = ""
            CommunicateMode = comFreePic
            w_ret = RequestACAD("PICEMPTY")
            form_no.w_hm_name.Focus() 'SetFocus()
        End If


        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        Rs.Close()
        end_sql()
        ' <- watanabe add VerUP(2011)

    End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_TMP_SIZE(0)
		
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
                .HelpContext = 800
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		Dim ZumenName As Object
		Dim w_ret As Object
		Dim result As Object
        Dim pic_no As Integer
		
		Dim w_mess As String

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        ' -> watanabe add VerUP(2011)
        w_mess = ""
        ' <- watanabe add VerUP(2011)


		'ピクチャのチェック
		'Brand Ver.5 TIFF->BMP 変更 start
		'    If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Can not read.", 64)
            w_size1.Focus()
            Exit Sub
        End If
		
		init_sql()
		' Brand Ver.3 変更
		'    DBTableName = DBName & "..hm_kanri"
		DBTableName = DBName & "..hm_kanri1"
		
		form_no.Enabled = False
		F_MSG.Show()
		
		If FreePicNum < 1 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ピクチャ数が足りません" & Chr(13) & "空きピクチャ数 =" & FreePicNum)
            ErrMsg = "The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum
            ErrTtl = "Template reading"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		
        pic_no = -1


        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT haiti_pic")
        '      result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE (")
        'result = SqlCmd(SqlConn, " font_name = '" & Mid(w_hm_name.Text, 1, 6) & "' AND")
        'result = SqlCmd(SqlConn, " no = '" & Mid(w_hm_name.Text, 7, 2) & "')")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '              pic_no = Val(SqlData(SqlConn, 1))
        '	Loop 
        'Else
        '	MsgBox("SQLｴﾗｰ", 64, "SQLｴﾗｰ")
        '	GoTo error_section
        'End If


        '検索コマンド作成
        sqlcmd = "SELECT haiti_pic"
        sqlcmd = sqlcmd & " FROM " & DBTableName
        sqlcmd = sqlcmd & " WHERE ("
        sqlcmd = sqlcmd & " font_name = '" & Mid(w_hm_name.Text, 1, 6) & "' AND"
        sqlcmd = sqlcmd & " no = '" & Mid(w_hm_name.Text, 7, 2) & "')"

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        Do Until Rs.EOF
            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                pic_no = Val(Rs.rdoColumns(0).Value)
            Else
                pic_no = -1
            End If

            Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        '** 画面情報を送信 **
        Call temp_bz_get(1)
        Call bz_spec_set(w_mess)
        w_ret = PokeACAD("SPECADD", w_mess)
        w_ret = RequestACAD("SPECADD")

        ZumenName = "HM-" & VB.Left(Trim(form_no.w_hm_name.Text), 6)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
        w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("ACADREAD", w_mess)
        w_ret = RequestACAD("TMPREAD")
        '   time_start = Now
        '   Do
        '     time_now = Now
        '     If Trim(form_main.text2.Text) = "" Then
        '      If time_now - time_start > timeOutSecond Then
        '        MsgBox "タイムアウトエラー", 64, "ERROR"
        '        w_ret = PokeACAD("ERROR", "TIMEOUT" & timeOutSecond & "秒が経過しました。")
        '        w_ret = RequestACAD("ERROR")
        '        GoTo LOOP_EXIT
        '      End If
        '     ElseIf Left$(Trim(form_main.text2.Text), 7) = "OK-DATA" Then
        '      form_main.text2.Text = ""
        '      GoTo LOOP_EXIT
        '     ElseIf Left(Trim(form_main.text2.Text), 5) = "ERROR" Then
        '      error_no = Mid$(Trim(form_main.text2.Text), 6, 3)
        '      MsgBox "エラーが返りました   ERROR_NO=" & error_no, 64, "ACAD戻り値ｴﾗｰ"
        '      form_main.text2.Text = ""
        '      GoTo LOOP_EXIT
        '     Else
        '      MsgBox "ﾘﾀｰﾝｺｰﾄﾞが不正です" & Chr(13) & Trim(form_main.text2.Text), 64, "ACAD戻り値ｴﾗｰ"
        '      form_main.text2.Text = ""
        '      GoTo LOOP_EXIT
        '     End If
        '   Loop
        'LOOP_EXIT:


        end_sql()
        Exit Sub

error_section:

        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)
        ' <- watanabe add VerUP(2011)

        On Error Resume Next

        ' -> watanabe add VerUP(2011)
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        F_MSG.Close()
        form_no.Enabled = True
    End Sub
	
	Private Sub F_TMP_SIZE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

		form_no = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
		Call Clear_F_TMP_SIZE(0)

        CommunicateMode = comSpecData
		RequestACAD("SPECDATA")
		
	End Sub
	
    'UPGRADE_WARNING: イベント w_hm_name.SelectedIndexChanged は、フォームが初期化されたときに発生します。
	Private Sub w_hm_name_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.SelectedIndexChanged
		Dim w_file As Object
		Dim TiffFile As Object
		Dim w_text As Object
        Dim ss3 As String
        Dim ss2 As String
        Dim ss1 As String
		Dim result As Object
		
		Dim w_text1 As String
		Dim w_text2 As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)


		On Error Resume Next ' エラーのトラップを留保します。
		Err.Clear()
		
		'(Brand Cad System Ver.3 UP )
		init_sql()
		w_text1 = VB.Left(Trim(w_hm_name.Text), 6)
		w_text2 = Mid(Trim(w_hm_name.Text), 7, 2)


        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT high, width, ang")
        'result = SqlCmd(SqlConn, " FROM " & DBName & "..hm_kanri1")
        'result = SqlCmd(SqlConn, " WHERE ( flag_delete = 0 ) AND ")
        'result = SqlCmd(SqlConn, " ( font_name  = '" & w_text1 & "') AND ")
        'result = SqlCmd(SqlConn, " ( no  = '" & w_text2 & "')")
        'result = SqlCmd(SqlConn, " ORDER BY high, width, ang")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		'Brand System Ver.3 追加
        '              ss1 = SqlData(SqlConn, 1)
        '		w_hight.Text = ss1
        '		ss2 = SqlData(SqlConn, 2)
        '		w_width.Text = ss2
        '		ss3 = SqlData(SqlConn, 3)
        '		w_ang.Text = ss3
        '	Loop 
        'Else
        '	MsgBox("ﾃﾞｰﾀﾍﾞｰｽSELECTｴﾗｰ", MsgBoxStyle.Critical)
        'End If


        '検索コマンド作成
        sqlcmd = "SELECT high, width, ang"
        sqlcmd = sqlcmd & " FROM " & DBName & "..hm_kanri1"
        sqlcmd = sqlcmd & " WHERE ( flag_delete = 0 ) AND "
        sqlcmd = sqlcmd & " ( font_name  = '" & w_text1 & "') AND "
        sqlcmd = sqlcmd & " ( no  = '" & w_text2 & "')"
        sqlcmd = sqlcmd & " ORDER BY high, width, ang"

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        Do Until Rs.EOF

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                ss1 = Rs.rdoColumns(0).Value
            Else
                ss1 = ""
            End If
            w_hight.Text = ss1

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                ss2 = Rs.rdoColumns(1).Value
            Else
                ss2 = ""
            End If
            w_width.Text = ss2

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                ss3 = Rs.rdoColumns(2).Value
            Else
                ss3 = ""
            End If
            w_ang.Text = ss3

            Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        w_text = w_hm_name.Text

        'Brand Ver.5 TIFF->BMP 変更 start
        '    TiffFile = TIFFDir & w_hm_name.Text & ".tif"
        '
        '    'Tiffﾌｧｲﾙ表示
        '    w_file = Dir(TiffFile)
        '    If w_file <> "" Then
        '       form_no.ImgThumbnail1.Image = TiffFile
        '       form_no.ImgThumbnail1.ThumbWidth = 500
        '       form_no.ImgThumbnail1.ThumbHeight = 200
        '    Else
        '       MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
        '       form_no.ImgThumbnail1.Image = ""
        '    End If
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


        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        Rs.Close()
        end_sql()
        ' <- watanabe add VerUP(2011)

	End Sub
	
	Private Sub w_hm_name_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_hm_name.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		'  MsgBox "キー入力は出来ません。", 64
		'  KeyAscii = 0
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size1.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_num.Text) <> "" Then
				Call Clear_F_TMP_SIZE(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_num.Text) <> "" Then
			Call Clear_F_TMP_SIZE(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.Leave
		
		form_no.w_size1.Text = UCase(Trim(form_no.w_size1.Text))
		
	End Sub
	
	Private Sub w_size2_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size2.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_num.Text) <> "" Then
				Call Clear_F_TMP_SIZE(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_num.Text) <> "" Then
			Call Clear_F_TMP_SIZE(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.Leave
		
		form_no.w_size2.Text = UCase(Trim(form_no.w_size2.Text))
		
	End Sub
	
	Private Sub w_size3_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size3.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_num.Text) <> "" Then
				Call Clear_F_TMP_SIZE(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size3_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size3.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_num.Text) <> "" Then
			Call Clear_F_TMP_SIZE(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size3_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.Leave
		
		form_no.w_size3.Text = UCase(Trim(form_no.w_size3.Text))
		
	End Sub
	
	
	Private Sub w_size4_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.Leave
		
		form_no.w_size4.Text = UCase(Trim(form_no.w_size4.Text))
		
	End Sub
	
	Private Sub w_size5_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size5.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_num.Text) <> "" Then
				Call Clear_F_TMP_SIZE(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size5_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size5.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii > 32 Then
			If (KeyAscii = CDbl("100")) Or (KeyAscii = CDbl("114")) Then
				KeyAscii = KeyAscii - 32
			ElseIf (KeyAscii <> CDbl("45")) And (KeyAscii <> CDbl("68")) And (KeyAscii <> CDbl("82")) And (KeyAscii <> CDbl("42")) Then 
				KeyAscii = 0
			End If
			GoTo EventExitSub
		End If
		If Trim(w_hm_num.Text) <> "" Then
			Call Clear_F_TMP_SIZE(1)
		End If
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size5_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.Leave
		
		form_no.w_size5.Text = UCase(Trim(form_no.w_size5.Text))
		
	End Sub
	
	Private Sub w_size6_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles w_size6.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If KeyCode = 46 Then
			If Trim(w_hm_num.Text) <> "" Then
				Call Clear_F_TMP_SIZE(1)
			End If
		End If
		
	End Sub
	
	Private Sub w_size6_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size6.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If Trim(w_hm_num.Text) <> "" Then
			Call Clear_F_TMP_SIZE(1)
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub w_size6_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.Leave
		
		form_no.w_size6.Text = UCase(Trim(form_no.w_size6.Text))
		
	End Sub
End Class