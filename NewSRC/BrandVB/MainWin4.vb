Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_MAIN4
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MAIN4_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Short
		Dim w_w_str As String
		
		form_main = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。

#If DEBUG Then
        '20100623移植変更
        w_w_str = "C:\ACAD19_02\BrandV5\uenv\BR_Set.ini"
        ret = set_read(w_w_str)

#Else
		'97.04.23 n.matsumi update start ...............................
		w_w_str = Environ("ACAD_SET")
		'    MsgBox "初期設定ﾌｧｲﾙ1:" & w_w_str, 64
		w_w_str = Trim(w_w_str) & "BR_Set.ini"
		ret = set_read(w_w_str)
		'    MsgBox "初期設定ﾌｧｲﾙ2:" & w_w_str, 64
		
		'ret = config_read("..\Files\BrandVB.cfg")
        'n.m    ret = set_read("..\Files\BrandVB.set")
		'97.04.23 n.matsumi update ended ...............................

#End If

		If ret = False Then
            MsgBox("Error reading initialization file (BR_Set.ini)", MsgBoxStyle.Information, "error")
			GoTo error_section
		End If
		'***** 12/8 1997 yamamoto start*****
		'    ret = env_get()
		'    If ret = False Then
		'         GoTo error_section
		'    End If
		'***** 12/8 1997 yamamoto end*****
		'   text2.LinkTimeout = timeOutSecond * 10
		'    ret = init_sql
		'    If ret = False Then
		'      MsgBox "SQLｻｰﾊﾞｰと接続出来ませんでした", vbInformation
		''      GoTo error_section
		'    End If
		
		ret = init_cad
		Select Case ret
			Case -1
                MsgBox("Fail to connect with the AdvanceCad.", MsgBoxStyle.Information)
				GoTo error_section
			Case errNoAppResponded
                MsgBox("AdvanceCad has not been started.", MsgBoxStyle.Information)
                MsgBox("It is a communication error. It is finished.")
				GoTo error_section
		End Select

        CommunicateMode = comCodeHyo
		ret = RequestACAD("CODEHE")
		DBTableName = DBName & "..hm_kanri" '編集文字管理
		
		F_MSG.Show()
		
		Exit Sub
		
error_section: 
		
        MsgBox("To exit", MsgBoxStyle.Critical, "Error end")
		End
		
	End Sub
	
	Private Sub Form_Terminate_Renamed()
		'SQL接続をｸﾛｰｽﾞします

        ' -> watanabe edit VerUP(2011)
        'SqlExit()
        Call end_sql()
        ' <- watanabe edit VerUP(2011)

    End Sub
	
	Private Sub LINK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles LINK.Click
		
		'   ret = init_cad
		'   Select Case ret
		'       Case False
		'           MsgBox "AdvanceCadとの接続に失敗しました", 64
		'       Case errNoAppResponded
		'           MsgBox "AdvanceCadは起動されていません"
		'   End Select
		
	End Sub
	
	Private Sub POKE_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles POKE.Click
		
		'On Error Resume Next
		' form_main.text2.LinkPoke
		' If Err Then MsgBox Error
		
	End Sub
	
	Private Sub REQUEST_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles REQUEST.Click
		
		'/// TEST VERSION
		'On Error Resume Next
		'  Text2.LinkItem = "WINNAME"
		' text2.LinkItem = form_main.text2.Text
		' text2.LinkRequest
		' NotifyFlag = False
		
	End Sub
	
    'UPGRADE_WARNING: イベント Text2.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
        Dim result As Integer '20100707 修正
		
		Dim w_ret As Short
		Dim Command_Line As String
		Dim he_code As String

        ' -> watanabe del VerUP(2011)
        'Static hIndex As Short
        'Dim w_w_str As String
        ' <- watanabe del VerUP(2011)

        Dim ret As Short
		Dim w_mess As String
		Dim w_serch_font As String
		Dim w_serch_no As String
		Dim w_gm_name001 As String
		Dim w_font_name As String
		Dim w_no As String
		Dim w_hz_no1 As String
		Dim w_hz_no2 As String
		
		Dim gm_code As String
		Dim w_s_font As String
		Dim w_s_cls1 As String
		Dim w_s_cls2 As String
		Dim w_s_name1 As String
		Dim w_s_name2 As String
		Dim w_old_font_name As String
		Dim w_gz_id As String
		Dim w_gz_no1 As String
		Dim w_gz_no2 As String

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


		If form_main.Text2.Text = "" Then Exit Sub
		
		Command_Line = Trim(form_main.Text2.Text)
		'output_command_line (Command_Line) '----- 12/11 1997 yamamoto add(debug)-----
		
		If VB.Left(Command_Line, 5) = "ERROR" Then
			MsgBox(Command_Line, MsgBoxStyle.Critical, "ERROR FROM ACAD")
			Exit Sub
		End If
		
		'Brand Ver.3 追加
		If VB.Left(Command_Line, 7) = "VBKILL1" Then
            CommunicateMode = comCodeHyo2
            w_ret = RequestACAD("CODEKO")
            MsgBox("DEBUG: 1")
			Exit Sub
		End If
		
		'Brand Ver.3 変更
		'    If Left$(Command_Line, 6) = "VBKILL" Then
		If VB.Left(Command_Line, 7) = "VBKILL2" Then
			F_MSG.Close()
			'      MsgBox "VBKILL受信しました" & Chr(13) & "BrandVBを終了します"
			End
		End If
		
        If Trim(form_main.SRflag.Text) = "SEND" Then Exit Sub
		
		Select Case CommunicateMode
			'編集文字コードデータ到着待ち時
			Case comCodeHyo
				
				If VB.Left(Command_Line, 6) = "CODEHE" Then
					''          CommunicateMode = comNone
					
					' Brand Ver.3 変更
					'             DBTableName = DBName & "..hm_kanri"  '編集文字管理
					DBTableName = DBName & "..hm_kanri1" '編集文字管理(基本部)
					DBTableName2 = DBName & "..hm_kanri2" '編集文字管理(文字部)
					
					ret = init_sql
					
					If ret = False Then
                        ' -> watanabe edit VerUP(2011)
                        'MsgBox("SQLｻｰﾊﾞｰと接続出来ませんでした(HM_kanri)", MsgBoxStyle.Information)
                        ErrMsg = "Cannot be connected to the SQL server.(HM_kanri)"
                        ErrTtl = "editing characters code acquisition"
                        ' <- watanabe edit VerUP(2011)
                        GoTo error_section
					End If
					
					he_code = Mid(Command_Line, 9, Len(Command_Line) - 8)
					w_font_name = Mid(he_code, 1, 6)
					w_no = Mid(he_code, 7, 2)
					
					w_serch_font = " font_name = '" & w_font_name & "'"
					w_serch_no = " no = '" & w_no & "'"

                    '編集文字D/B検索 編集文字図面名 取得
					

                    ' -> watanabe edit VerUP(2011)
                    '               'Brand Ver.3 変更
                    ''              result = SqlCmd(SqlConn, "SELECT hz_no1, hz_no2, gm_name001 FROM " & DBTableName)
                    'result = SqlCmd(SqlConn, "SELECT hz_no1, hz_no2 FROM " & DBTableName)
                    '
                    'result = SqlCmd(SqlConn, " WHERE " & w_serch_font & " AND" & w_serch_no)
                    '
                    'result = SqlExec(SqlConn)
                    'result = SqlResults(SqlConn)
                    '
                    'result = SqlNextRow(SqlConn)
                    '
                    ''編集文字図面名 POKE
                    'w_hz_no1 = SqlData(SqlConn, 1)
                    'If Trim(w_hz_no1) = "" Then
                    '	w_hz_no1 = "----"
                    'End If
                    '
                    'w_hz_no2 = SqlData(SqlConn, 2)
                    'If Trim(w_hz_no2) = "" Then
                    '	w_hz_no2 = "--"
                    'End If
					

                    '検索コマンド作成
                    sqlcmd = "SELECT hz_no1, hz_no2 FROM " & DBTableName
                    sqlcmd = sqlcmd & " WHERE " & w_serch_font & " AND" & w_serch_no

                    '検索
                    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                    Rs.MoveFirst()

                    '編集文字図面名 POKE
                    w_hz_no1 = ""
                    w_hz_no2 = ""
                    If GL_T_RDO.Con.RowsAffected() > 0 Then
                        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                            w_hz_no1 = Rs.rdoColumns(0).Value
                        Else
                            w_hz_no1 = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                            w_hz_no2 = Rs.rdoColumns(1).Value
                        Else
                            w_hz_no2 = ""
                        End If
                    End If

                    If Trim(w_hz_no1) = "" Then
                        w_hz_no1 = "----"
                    End If

                    If Trim(w_hz_no2) = "" Then
                        w_hz_no2 = "--"
                    End If

                    Rs.Close()
                    ' <- watanabe edit VerUP(2011)


                    ' Brand Ver.3 変更
                    '                 w_gm_name001 = SqlData$(SqlConn, 3)
                    end_sql()

                    init_sql()


                    ' -> watanabe edit VerUP(2011)
                    'result = sqlcmd(SqlConn, "SELECT gm_name FROM " & DBTableName2)
                    'result = sqlcmd(SqlConn, " WHERE " & w_serch_font & " AND" & w_serch_no)
                    'result = SqlExec(SqlConn)
                    'result = SqlResults(SqlConn)
                    'result = SqlNextRow(SqlConn)
                    'w_gm_name001 = SqlData(SqlConn, 1)


                    '検索コマンド作成
                    sqlcmd = "SELECT gm_name FROM " & DBTableName2
                    sqlcmd = sqlcmd & " WHERE " & w_serch_font & " AND" & w_serch_no

                    '検索
                    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                    Rs.MoveFirst()

                    '編集文字図面名 POKE
                    w_gm_name001 = ""
                    If GL_T_RDO.Con.RowsAffected() > 0 Then
                        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                            w_gm_name001 = Rs.rdoColumns(0).Value
                        Else
                            w_gm_name001 = ""
                        End If
                    End If

                    Rs.Close()
                    ' <- watanabe edit VerUP(2011)


                    '                w_mess = "HE-" & Left(Trim(w_hz_no1), 4)
                    '                w_mess = w_mess & "-" & Left(Trim(w_hz_no2), 2)
                    '                Mid$(w_mess, 1, 10) = "HE-" & Left(Trim(w_hz_no1), 4) & "-" & Left(Trim(w_hz_no2), 2)
                    w_mess = "HE-" & VB.Left(Trim(w_hz_no1), 4) & "-" & VB.Left(Trim(w_hz_no2), 2)

                    end_sql()

                    '原始文字検索
                    DBTableName = DBName & "..gm_kanri" '編集文字管理

                    ret = init_sql()

                    If ret = False Then
                        ' -> watanabe edit VerUP(2011)
                        'MsgBox("SQLｻｰﾊﾞｰと接続出来ませんでした(GM_kanri)", MsgBoxStyle.Information)
                        ErrMsg = "Cannot be connected to the SQL server.(GM_kanri)"
                        ErrTtl = "Primitive character code acquisition"
                        ' <- watanabe edit VerUP(2011)
                        GoTo error_section
                    End If

                    gm_code = w_gm_name001
                    w_s_font = "font_name = '" & Mid(gm_code, 1, 6) & "' AND "
                    w_s_cls1 = "font_class1 = '" & Mid(gm_code, 7, 1) & "' AND "
                    w_s_cls2 = "font_class2 = '" & Mid(gm_code, 8, 1) & "' AND "
                    w_s_name1 = "name1 = '" & Mid(gm_code, 9, 1) & "' AND "
                    w_s_name2 = "name2 = '" & Mid(gm_code, 10, 1) & "'"

                    '  MsgBox " w_font = " & w_s_font & " " & w_s_cls1 & " " & w_s_cls2 & " " & w_s_name1 & " " & w_s_name2

                    '編集文字D/B検索 編集文字図面名 取得


                    ' -> watanabe edit VerUP(2011)
                    'result = sqlcmd(SqlConn, "SELECT old_font_name, gz_id, gz_no1, gz_no2 FROM " & DBTableName)
                    '
                    'result = sqlcmd(SqlConn, " WHERE " & w_s_font & w_s_cls1 & w_s_cls2 & w_s_name1 & w_s_name2)
                    '
                    'result = SqlExec(SqlConn)
                    'result = SqlResults(SqlConn)
                    '
                    'result = SqlNextRow(SqlConn)
                    '
                    ''編集文字図面名 POKE
                    'w_old_font_name = SqlData(SqlConn, 1)
                    'If Trim(w_old_font_name) = "" Then
                    '    w_old_font_name = "------"
                    'End If
                    '
                    'w_gz_id = SqlData(SqlConn, 2)
                    'If Trim(w_gz_id) = "" Then
                    '    w_gz_id = "KO"
                    'End If
                    '
                    'w_gz_no1 = SqlData(SqlConn, 3)
                    'If Trim(w_gz_no1) = "" Then
                    '    w_gz_no1 = "----"
                    'End If
                    '
                    'w_gz_no2 = SqlData(SqlConn, 4)
                    'If Trim(w_gz_no2) = "" Then
                    '    w_gz_no2 = "--"
                    'End If


                    '検索コマンド作成
                    sqlcmd = "SELECT old_font_name, gz_id, gz_no1, gz_no2 FROM " & DBTableName
                    sqlcmd = sqlcmd & " WHERE " & w_s_font & w_s_cls1 & w_s_cls2 & w_s_name1 & w_s_name2

                    '検索
                    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                    Rs.MoveFirst()

                    '原始文字図面名 POKE
                    w_old_font_name = ""
                    w_gz_id = ""
                    w_gz_no1 = ""
                    w_gz_no2 = ""
                    If GL_T_RDO.Con.RowsAffected() > 0 Then
                        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                            w_old_font_name = Rs.rdoColumns(0).Value
                        Else
                            w_old_font_name = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                            w_gz_id = Rs.rdoColumns(1).Value
                        Else
                            w_gz_id = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                            w_gz_no1 = Rs.rdoColumns(2).Value
                        Else
                            w_gz_no1 = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                            w_gz_no2 = Rs.rdoColumns(3).Value
                        Else
                            w_gz_no2 = ""
                        End If
                    End If

                    If Trim(w_old_font_name) = "" Then
                        w_old_font_name = "------"
                    End If

                    If Trim(w_gz_id) = "" Then
                        w_gz_id = "KO"
                    End If

                    If Trim(w_gz_no1) = "" Then
                        w_gz_no1 = "----"
                    End If

                    If Trim(w_gz_no2) = "" Then
                        w_gz_no2 = "--"
                    End If

                    Rs.Close()
                    ' <- watanabe edit VerUP(2011)


                    'MsgBox "old font =[" & w_old_font_name & "]"
                    'MsgBox "gz_id    =[" & w_gz_id & "]"
                    'MsgBox "gz_no1   =[" & w_gz_no1 & "]"
                    'MsgBox "gz_no2   =[" & w_gz_no2 & "]"

                    w_mess = w_mess & VB.Left(Trim(w_gz_id), 2) & "-" & VB.Left(Trim(w_gz_no1), 4) & "-" & VB.Left(Trim(w_gz_no2), 2)
                    w_mess = w_mess & VB.Left(Trim(w_old_font_name), 6)
                    '                 Mid$(w_mess, 11, 10) = Left(Trim(w_gz_id), 2) & "-" & Left(Trim(w_gz_no1), 4) & "-" & Left(Trim(w_gz_no2), 2)
                    '                 Mid$(w_mess, 21, 6) = Left(Trim(w_old_font_name), 6)

                    'MsgBox " w_mess = " & w_mess

                    end_sql()

                    w_ret = PokeACAD("SAVEHE", w_mess)

                    '編集文字図面名 REQUEST

                    CommunicateMode = comCodeHyo
                    w_ret = RequestACAD("CODEHE")
                    MsgBox("DEBUG: 0")
                Else
                    MsgBox("It is not a editing characters code [" & Command_Line & "]")
                End If

                ''           form_main.text2.Text = ""


                'Brand Ver.3 追加
            Case comCodeHyo2

                If VB.Left(Command_Line, 6) = "CODEKO" Then

                    '原始文字検索
                    DBTableName = DBName & "..gm_kanri" '原始文字管理

                    ret = init_sql()

                    If ret = False Then
                        ' -> watanabe edit VerUP(2011)
                        'MsgBox("SQLｻｰﾊﾞｰと接続出来ませんでした(GM_kanri)", MsgBoxStyle.Information)
                        ErrMsg = "Cannot be connected to the SQL server.(GM_kanri)"
                        ErrTtl = "Primitive character code acquisition"
                        ' <- watanabe edit VerUP(2011)
                        GoTo error_section
                    End If

                    gm_code = Mid(Command_Line, 9, Len(Command_Line) - 8)
                    w_s_font = "font_name = '" & Mid(gm_code, 1, 6) & "' AND "
                    w_s_cls1 = "font_class1 = '" & Mid(gm_code, 7, 1) & "' AND "
                    w_s_cls2 = "font_class2 = '" & Mid(gm_code, 8, 1) & "' AND "
                    w_s_name1 = "name1 = '" & Mid(gm_code, 9, 1) & "' AND "
                    w_s_name2 = "name2 = '" & Mid(gm_code, 10, 1) & "'"

                    'MsgBox " w_font = " & w_s_font & " " & w_s_cls1 & " " & w_s_cls2 & " " & w_s_name1 & " " & w_s_name2

                    '原始文字D/B検索 刻印図面名 取得


                    ' -> watanabe edit VerUP(2011)
                    'result = sqlcmd(SqlConn, "SELECT old_font_name, gz_id, gz_no1, gz_no2 FROM " & DBTableName)
                    '
                    'result = sqlcmd(SqlConn, " WHERE " & w_s_font & w_s_cls1 & w_s_cls2 & w_s_name1 & w_s_name2)
                    '
                    'result = SqlExec(SqlConn)
                    'result = SqlResults(SqlConn)
                    '
                    'result = SqlNextRow(SqlConn)
                    '
                    ''刻印図面名 POKE
                    'w_old_font_name = SqlData(SqlConn, 1)
                    'If Trim(w_old_font_name) = "" Then
                    '    w_old_font_name = "------"
                    'End If
                    '
                    'w_gz_id = SqlData(SqlConn, 2)
                    'If Trim(w_gz_id) = "" Then
                    '    w_gz_id = "KO"
                    'End If
                    '
                    'w_gz_no1 = SqlData(SqlConn, 3)
                    'If Trim(w_gz_no1) = "" Then
                    '    w_gz_no1 = "----"
                    'End If
                    '
                    'w_gz_no2 = SqlData(SqlConn, 4)
                    'If Trim(w_gz_no2) = "" Then
                    '    w_gz_no2 = "--"
                    'End If


                    '検索コマンド作成
                    sqlcmd = "SELECT old_font_name, gz_id, gz_no1, gz_no2 FROM " & DBTableName
                    sqlcmd = sqlcmd & " WHERE " & w_s_font & w_s_cls1 & w_s_cls2 & w_s_name1 & w_s_name2

                    '検索
                    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                    Rs.MoveFirst()

                    '刻印図面名 POKE
                    w_old_font_name = ""
                    w_gz_id = ""
                    w_gz_no1 = ""
                    w_gz_no2 = ""
                    If GL_T_RDO.Con.RowsAffected() > 0 Then
                        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                            w_old_font_name = Rs.rdoColumns(0).Value
                        Else
                            w_old_font_name = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                            w_gz_id = Rs.rdoColumns(1).Value
                        Else
                            w_gz_id = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                            w_gz_no1 = Rs.rdoColumns(2).Value
                        Else
                            w_gz_no1 = ""
                        End If

                        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                            w_gz_no2 = Rs.rdoColumns(3).Value
                        Else
                            w_gz_no2 = ""
                        End If
                    End If

                    If Trim(w_old_font_name) = "" Then
                        w_old_font_name = "------"
                    End If

                    If Trim(w_gz_id) = "" Then
                        w_gz_id = "KO"
                    End If

                    If Trim(w_gz_no1) = "" Then
                        w_gz_no1 = "----"
                    End If

                    If Trim(w_gz_no2) = "" Then
                        w_gz_no2 = "--"
                    End If

                    Rs.Close()
                    ' <- watanabe edit VerUP(2011)


                    'MsgBox "old font =[" & w_old_font_name & "]"
                    'MsgBox "gz_id    =[" & w_gz_id & "]"
                    'MsgBox "gz_no1   =[" & w_gz_no1 & "]"
                    'MsgBox "gz_no2   =[" & w_gz_no2 & "]"

                    w_mess = VB.Left(Trim(w_gz_id), 2) & "-" & VB.Left(Trim(w_gz_no1), 4) & "-" & VB.Left(Trim(w_gz_no2), 2)
                    w_mess = w_mess & VB.Left(Trim(w_old_font_name), 6)

                    'MsgBox " w_mess = " & w_mess

                    end_sql()

                    w_ret = PokeACAD("NAMEKO", w_mess)

                    '原始文字コード REQUEST

                    CommunicateMode = comCodeHyo2
                    w_ret = RequestACAD("CODEKO")
                    MsgBox("DEBUG: 2")


                Else
                    MsgBox("It is not a primitive character code [" & Command_Line & "]")
                End If

			Case Else
                MsgBox("communicateMode error")

		End Select
		
		
		Exit Sub
		
error_section: 

        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()
        Rs.Close()
        end_sql()
        ' <- watanabe add VerUP(2011)

        MsgBox("To exit", MsgBoxStyle.Critical, "Error end")
		End
		
    End Sub

    Private Sub Text2_LinkClose()
        Dim Connected As Object

        Connected = False

    End Sub
	
	Private Sub Text2_LinkError(ByRef LinkErr As Short)
		Dim Msg As Object
        Msg = "DDE communication error"
		MsgBox(Msg)
	End Sub
	
	Private Sub Text2_LinkNotify()
        'Dim NotifyFlag As Object
		If Not NotifyFlag Then
            MsgBox("Can get the new data from the DDE source.")
			NotifyFlag = True
		End If
	End Sub
	
	Private Sub Vbsql1_Error(ByVal SqlConn As Integer, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal ErrorStr As String, ByVal OSErrorNum As Integer, ByVal OSErrorStr As String, ByRef RetCode As Integer)
		MsgBox("DB-Library Error: " & Str(ErrorNum) & " " & ErrorStr)
	End Sub
	
	Private Sub Vbsql1_Message(ByVal SqlConn As Integer, ByVal Message As Integer, ByVal State As Integer, ByVal Severity As Integer, ByVal MsgStr As String, ByVal ServerNameStr As String, ByVal ProcNameStr As String, ByVal Line As Integer)
		' If Severity > 1 Then
		'   MsgBox ("SQL Server Error: " + Str$(Message&) + " " + MsgStr$)
		' End If
	End Sub
End Class