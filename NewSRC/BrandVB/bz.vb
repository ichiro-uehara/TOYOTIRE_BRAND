Option Strict Off
Option Explicit On
Module MJ_BZ
	Function bz_insert() As Short
        Dim fno As Integer
        Dim j As Integer
        Dim i As Integer
        Dim now_time As String
        Dim result As Integer
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim wcommand As String
        'Dim w_command As String
        'Dim kubun As Short
        ' <- watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        '--------------< ブランド図面 登録 新規 >-----------------------------


		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		

        ' -> watanabe edit VerUP(2011)
        '      SqlFreeBuf((SqlConn))
        '
        'result = SqlCmd(SqlConn, "SELECT * ")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ")
        'result = SqlCmd(SqlConn, " id = 'AT-B' AND")
        '      result = SqlCmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "' AND")
        '      result = SqlCmd(SqlConn, " no2 = '" & form_no.w_no2.Text & "'")
        'result = SqlExec(SqlConn)
        '' If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		Loop 
        '		MsgBox("図面番号が既にブランド図面に存在します。" & Chr(13) & "新規での登録は出来ません", MsgBoxStyle.Critical, "number exist error")
        '		GoTo error_section
        '	End If
        'Else
        '	GoTo error_section
        'End If


        '検索キーセット
        key_code = " id = 'AT-B' AND"
        key_code = key_code & " no1 = '" & form_no.w_no1.Text & "' AND"
        key_code = key_code & " no2 = '" & form_no.w_no2.Text & "'"

        '検索コマンド作成
        sqlcmd = "SELECT *  FROM " & DBTableName & " WHERE " & key_code

        'ヒット数チェック
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt > 0 Then
            ErrMsg = "Drawing number exists in the already brand drawing." & Chr(13) & "It is not possible to register a new."
            ErrTtl = "number exist error"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Brand drawing registration"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        w_str(1) = "0" '削除フラグ
        w_str(2) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
        w_str(3) = "'" & form_no.w_no1.Text & "'" '図面番号
        w_str(4) = "'" & form_no.w_no2.Text & "'" '変番
        w_str(5) = "'" & form_no.w_kanri_no.Text & "'" '業務管理番号
        w_str(6) = "'" & form_no.w_syurui.Text & "'" 'タイヤ種類
        w_str(7) = "'" & form_no.w_syubetu.Text & "'" 'パターン種別
        w_str(8) = "'" & form_no.w_pattern.Text & "'" 'パターン
        w_str(9) = "'" & form_no.w_size.Text & "'" 'サイズ
        w_str(10) = "'" & form_no.w_size1.Text & "'" 'サイズ1(外径)
        w_str(11) = "'" & form_no.w_size2.Text & "'" 'サイズ2(継)
        w_str(12) = "'" & form_no.w_size3.Text & "'" 'サイズ3(断面幅)
        w_str(13) = "'" & form_no.w_size4.Text & "'" 'サイズ4(速度)
        w_str(14) = "'" & form_no.w_size5.Text & "'" 'サイズ5(構造)
        w_str(15) = "'" & form_no.w_size6.Text & "'" 'サイズ6(リム径)
        w_str(16) = "'" & form_no.w_size7.Text & "'" 'サイズ7(接尾)
        w_str(17) = "'" & form_no.w_size8.Text & "'" 'サイズ8(プライ)
        w_str(18) = "'" & form_no.w_size_code.Text & "'" 'サイズコード
        w_str(19) = "'" & form_no.w_kikaku.Text & "'" '規格
        w_str(20) = "'" & Mid(form_no.w_plant.Text, 4, 2) & "'" '工場
        w_str(21) = "'" & form_no.w_plant_code.Text & "'" '工場コード
        w_str(22) = CStr(Val(Mid(form_no.w_tos_moyou.Text, 1, 1))) 'ＴＯＳ対応模様
        w_str(23) = CStr(Val(Mid(form_no.w_side_moyou.Text, 1, 1))) 'サイド凹凸模様
        w_str(24) = CStr(Val(Mid(form_no.w_side_kenti.Text, 1, 1))) 'サイド凹凸検知
        w_str(25) = CStr(Val(Mid(form_no.w_peak_mark.Text, 1, 1))) 'ピークマーク
        w_str(26) = CStr(Val(Mid(form_no.w_nasiji.Text, 1, 1))) '梨地加工
        w_str(27) = "'" & form_no.w_comment.Text & "'" 'コメント
        w_str(28) = "'" & form_no.w_dep_name.Text & "'" '部署コード
        w_str(29) = "'" & form_no.w_entry_name.Text & "'" '登録者

        If Len(Hour(TimeOfDay)) = 1 Then
            now_time = "0" & Hour(TimeOfDay)
        Else
            now_time = Hour(TimeOfDay)
        End If

        If Len(Minute(TimeOfDay)) = 1 Then
            now_time = Trim(now_time) & ":0" & Minute(TimeOfDay)
        Else
            now_time = Trim(now_time) & ":" & Minute(TimeOfDay)
        End If

        w_str(30) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '登録日

        w_str(31) = CStr(Val(form_no.w_hm_num.Text)) '編集文字数


        'ブランド図面ﾌｧｲﾙに登録

        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        '
        '' -> watanabe add VerUP(2011)
        'wcommand = ""
        '' <- watanabe add VerUP(2011)
        '
        'wcommand = wcommand & "INSERT INTO " & DBTableName & " VALUES("
        '
        '' Brand Ver.3 変更
        '' For i = 1 To 31
        'For i = 1 To 30
        '    result = sqlcmd(SqlConn, w_str(i) & ",")
        '    wcommand = wcommand & w_str(i) & ","
        'Next i
        '
        'result = sqlcmd(SqlConn, w_str(31))
        'wcommand = wcommand & w_str(31)
        '
        '' Brand Ver.3 変更
        '' For i = 1 To temp_bz.hm_num - 1
        ''     result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(i) & "',")
        ''           wcommand = wcommand & "'" & temp_bz.hm_name(i) & "',"
        '' Next i
        '' result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(temp_bz.hm_num) & "'")
        ''            wcommand = wcommand & "'" & temp_bz.hm_name(temp_bz.hm_num) & "'"
        '' If temp_bz.hm_num < 100 Then
        ''    For i = 1 To 100 - temp_bz.hm_num
        ''        result = SqlCmd(SqlConn, ",'" & Space$(8) & "'")
        ''        wcommand = wcommand & ",'" & Space$(8) & "'"
        ''    Next i
        '' End If
        '
        'result = sqlcmd(SqlConn, ")")
        'wcommand = wcommand & ")"
        '
        ''MsgBox "sql-[" & wcommand & "]"
        'result = SqlExec(SqlConn)
        ''ブランド図面の登録に失敗した時
        'If result = FAIL Then
        '    GoTo error_section
        'End If
        'result = SqlResults(SqlConn)


        sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        For i = 1 To 30
            sqlcmd = sqlcmd & w_str(i) & ","
        Next i
        sqlcmd = sqlcmd & w_str(31)
        sqlcmd = sqlcmd & ")"

        'ｺﾏﾝﾄﾞ実行
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        For i = 1 To Val(Trim(form_no.w_hm_num.Text))

            init_sql()

            w_str(1) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '図面番号
            w_str(3) = "'" & form_no.w_no2.Text & "'" '変番
            w_str(4) = i '編集文字番号


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '    result = sqlcmd(SqlConn, w_str(j) & ",")
            'Next j
            'result = sqlcmd(SqlConn, "'" & temp_bz.hm_name(i) & "'")
            'result = sqlcmd(SqlConn, " )")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then
            '    GoTo error_section
            'End If
            'result = SqlResults(SqlConn)


            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            For j = 1 To 4
                sqlcmd = sqlcmd & w_str(j) & ","
            Next j
            sqlcmd = sqlcmd & "'" & temp_bz.hm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            'ｺﾏﾝﾄﾞ実行
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()

        Next i
        ' <- watanabe add VerUP(2011)


        ' -> watanabe del VerUP(2011)   現在未使用、最初期の残骸？
        'INSERT_ERR:
        '        FileClose(fno)
        ' <- watanabe del VerUP(2011)

        bz_insert = True

        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If

        On Error Resume Next
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        bz_insert = FAIL
    End Function
	
	Function bz_read(ByRef wk_id As String, ByRef wk_no1 As String, ByRef wk_no2 As String) As Short
		Dim w_ret As Object
		Dim bz_code As Object
		Dim result As Object
        Dim w_mess As String
		Dim wk_entry_name As String

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        'ﾋﾟｸﾁｬ番号

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT entry_name")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        '
        'result = SqlCmd(SqlConn, " WHERE ( flag_delete = 0 AND")
        'result = SqlCmd(SqlConn, " id = '" & wk_id & "' AND")
        'result = SqlCmd(SqlConn, " no1 = '" & wk_no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & wk_no2 & "')")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		wk_entry_name = CStr(Val(SqlData(SqlConn, 1)))
        '          Else
        '
        '              ' -> watanabe add VerUP(2011)
        '              bz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
        '              ' <- watanabe add VerUP(2011)
        '
        '              MsgBox("指定されたブランド図面が見つかりません" & Chr(13) & bz_code, MsgBoxStyle.Critical, "data not found")
        '              bz_read = FAIL
        '              Exit Function
        '	End If
        'Else
        '	MsgBox("SQL エラー")
        '	bz_read = FAIL
        '	Exit Function
        'End If


        '検索キーセット
        key_code = "flag_delete = 0 AND"
        key_code = key_code & " id = '" & wk_id & "' AND"
        key_code = key_code & " no1 = '" & wk_no1 & "' AND"
        key_code = key_code & " no2 = '" & wk_no2 & "'"

        '検索コマンド作成
        sqlcmd = "SELECT entry_name FROM " & DBTableName & " WHERE ( " & key_code & ")"

        'ヒット数チェック
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            bz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
            ErrMsg = "There is no brand drawings specified." & Chr(13) & bz_code
            ErrTtl = "data not found"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Brand drawing reading"
            GoTo error_section
        End If

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            wk_entry_name = CStr(Val(Rs.rdoColumns(0).Value))
        Else
            wk_entry_name = ""
        End If
        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        '   w_mess = BrandDir & wk_id & "-" & wk_no1 & "-" & wk_no2
		w_mess = BrandDir & wk_id & wk_no1 & "-" & wk_no2
		w_ret = PokeACAD("MDLREAD", w_mess)
		w_ret = RequestACAD("MDLREAD")


        ' -> watanabe add VerUP(2011)
        bz_read = SUCCEED
        ' <- watanabe add VerUP(2011)


        ' -> watanabe add VerUP(2011)
        Exit Function

error_section:
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()
        Rs.Close()

        bz_read = FAIL
        ' <- watanabe add VerUP(2011)

    End Function
	
	Function bz_update() As Short
		Dim i As Object
        Dim result As Integer '20100707 修正
		Dim now_time As Object
        Dim w_str(140) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' -> watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        '--------------< ブランド図面 登録 修正 >-----------------------------

		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		
		w_str(1) = "0" '削除フラグ
		w_str(2) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
		w_str(3) = "'" & form_no.w_no1.Text & "'" '図面番号
		w_str(4) = "'" & form_no.w_no2.Text & "'" '変番
		w_str(5) = "'" & form_no.w_kanri_no.Text & "'" '業務管理番号
		w_str(6) = "'" & form_no.w_syurui.Text & "'" 'タイヤ種類
		w_str(7) = "'" & form_no.w_syubetu.Text & "'" 'パターン種別
		w_str(8) = "'" & form_no.w_pattern.Text & "'" 'パターン
		w_str(9) = "'" & form_no.w_size.Text & "'" 'サイズ
		w_str(10) = "'" & form_no.w_size1.Text & "'" 'サイズ1(外径)
		w_str(11) = "'" & form_no.w_size2.Text & "'" 'サイズ2(継)
		w_str(12) = "'" & form_no.w_size3.Text & "'" 'サイズ3(断面幅)
		w_str(13) = "'" & form_no.w_size4.Text & "'" 'サイズ4(速度)
		w_str(14) = "'" & form_no.w_size5.Text & "'" 'サイズ5(構造)
		w_str(15) = "'" & form_no.w_size6.Text & "'" 'サイズ6(リム径)
		w_str(16) = "'" & form_no.w_size7.Text & "'" 'サイズ7(接尾)
		w_str(17) = "'" & form_no.w_size8.Text & "'" 'サイズ8(プライ)
		w_str(18) = "'" & form_no.w_size_code.Text & "'" 'サイズコード
		w_str(19) = "'" & form_no.w_kikaku.Text & "'" '規格
		w_str(20) = "'" & Mid(form_no.w_plant.Text, 4, 2) & "'" '工場
		w_str(21) = "'" & form_no.w_plant_code.Text & "'" '工場コード
		w_str(22) = CStr(Val(Mid(form_no.w_tos_moyou.Text, 1, 1))) 'ＴＯＳ対応模様
		w_str(23) = CStr(Val(Mid(form_no.w_side_moyou.Text, 1, 1))) 'サイド凹凸模様
		w_str(24) = CStr(Val(Mid(form_no.w_side_kenti.Text, 1, 1))) 'サイド凹凸検知
		w_str(25) = CStr(Val(Mid(form_no.w_peak_mark.Text, 1, 1))) 'ピークマーク
		w_str(26) = CStr(Val(Mid(form_no.w_nasiji.Text, 1, 1))) '梨地加工
		w_str(27) = "'" & form_no.w_comment.Text & "'" 'コメント
		w_str(28) = "'" & form_no.w_dep_name.Text & "'" '部署コード
		w_str(29) = "'" & form_no.w_entry_name.Text & "'" '登録者
		
		If Len(Hour(TimeOfDay)) = 1 Then
			now_time = "0" & Hour(TimeOfDay)
		Else
			now_time = Hour(TimeOfDay)
		End If
		
		If Len(Minute(TimeOfDay)) = 1 Then
			now_time = Trim(now_time) & ":0" & Minute(TimeOfDay)
		Else
			now_time = Trim(now_time) & ":" & Minute(TimeOfDay)
		End If
		
		w_str(30) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '登録日
		
        w_str(31) = CStr(Val(Trim(form_no.w_hm_num.Text))) '編集文字数
		
		
		'Brand Ver.3 変更
		' For i = 1 To 100
		'     w_str(i + 31) = "'" & temp_bz.hm_name(i) & "'"
		' Next i
		
		
		'テーブル検索

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_bz.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_bz.no2 & "' )")
        '
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		temp_bz.comment = SqlData(SqlConn, 1)
        '		temp_bz.dep_name = SqlData(SqlConn, 2)
        '		temp_bz.entry_name = SqlData(SqlConn, 3)
        '		temp_bz.entry_date = SqlData(SqlConn, 4)
        '	Loop 
        'Else
        '	MsgBox("ブランド図面がありません。", MsgBoxStyle.Critical, "number is exist")
        '	GoTo error_section
        'End If


        '検索キーセット
        key_code = "no1 = '" & temp_bz.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_bz.no2 & "'"

        '検索コマンド作成
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE ( " & key_code & " )"

        'ヒット数チェック
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            ErrMsg = "There is no brand drawings specified"
            ErrTtl = "Brand drawing update registration"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Brand drawing update registration"
            GoTo error_section
        End If

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            temp_bz.comment = Rs.rdoColumns(0).Value
        Else
            temp_bz.comment = ""
        End If

        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
            temp_bz.dep_name = Rs.rdoColumns(1).Value
        Else
            temp_bz.dep_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
            temp_bz.entry_name = Rs.rdoColumns(2).Value
        Else
            temp_bz.entry_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
            temp_bz.entry_date = Rs.rdoColumns(3).Value
        Else
            temp_bz.entry_date = ""
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)

		
		'テーブル更新

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "UPDATE " & DBTableName)
        'result = SqlCmd(SqlConn, " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ",")
        'result = SqlCmd(SqlConn, " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ",")
        'result = SqlCmd(SqlConn, " kanri_no = " & w_str(5) & ", syurui = " & w_str(6) & ",")
        'result = SqlCmd(SqlConn, " syubetu = " & w_str(7) & ", pattern = " & w_str(8) & ",")
        'result = SqlCmd(SqlConn, " size = " & w_str(9) & ", size1 = " & w_str(10) & ",")
        'result = SqlCmd(SqlConn, " size2 = " & w_str(11) & ", size3 = " & w_str(12) & ",")
        'result = SqlCmd(SqlConn, " size4 = " & w_str(13) & ", size5 = " & w_str(14) & ",")
        'result = SqlCmd(SqlConn, " size6 = " & w_str(15) & ", size7 = " & w_str(16) & ",")
        'result = SqlCmd(SqlConn, " size8 = " & w_str(17) & ", size_code = " & w_str(18) & ",")
        'result = SqlCmd(SqlConn, " kikaku = " & w_str(19) & ", plant = " & w_str(20) & ",")
        'result = SqlCmd(SqlConn, " plant_code = " & w_str(21) & ", tos_moyou = " & w_str(22) & ",")
        'result = SqlCmd(SqlConn, " side_moyou = " & w_str(23) & ", side_kenti = " & w_str(24) & ",")
        'result = SqlCmd(SqlConn, " peak_mark = " & w_str(25) & ", nasiji = " & w_str(26) & ",")
        'result = SqlCmd(SqlConn, " comment = " & w_str(27) & ", dep_name = " & w_str(28) & ",")
        'result = SqlCmd(SqlConn, " entry_name = " & w_str(29) & ", entry_date = " & w_str(30) & ",")
        ''Brand Ver.3 変更
        '' result = SqlCmd(SqlConn, " hm_num = " & w_str(31) & ",")
        'result = SqlCmd(SqlConn, " hm_num = " & w_str(31))
        '
        '' Brand Ver.3 変更
        '' For i = 1 To 99
        ''     hname = " hm_name" & Format(i, "000")
        ''     result = SqlCmd(SqlConn, hname & " = " & w_str(31 + i) & ",")
        '' Next i
        '' hname = " hm_name" & Format(100, "000")
        '' result = SqlCmd(SqlConn, hname & " = " & w_str(131))
        '
        'result = SqlCmd(SqlConn, " From " & DBTableName & "(PAGLOCK)")
        'result = SqlCmd(SqlConn, " WHERE ( ")
        'result = SqlCmd(SqlConn, " id = 'AT-B' AND")
        '      result = SqlCmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "' AND")
        '      result = SqlCmd(SqlConn, " no2 = '" & form_no.w_no2.Text & "' )")
        '
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)


        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ","
        sqlcmd = sqlcmd & " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ","
        sqlcmd = sqlcmd & " kanri_no = " & w_str(5) & ", syurui = " & w_str(6) & ","
        sqlcmd = sqlcmd & " syubetu = " & w_str(7) & ", pattern = " & w_str(8) & ","
        sqlcmd = sqlcmd & " size = " & w_str(9) & ", size1 = " & w_str(10) & ","
        sqlcmd = sqlcmd & " size2 = " & w_str(11) & ", size3 = " & w_str(12) & ","
        sqlcmd = sqlcmd & " size4 = " & w_str(13) & ", size5 = " & w_str(14) & ","
        sqlcmd = sqlcmd & " size6 = " & w_str(15) & ", size7 = " & w_str(16) & ","
        sqlcmd = sqlcmd & " size8 = " & w_str(17) & ", size_code = " & w_str(18) & ","
        sqlcmd = sqlcmd & " kikaku = " & w_str(19) & ", plant = " & w_str(20) & ","
        sqlcmd = sqlcmd & " plant_code = " & w_str(21) & ", tos_moyou = " & w_str(22) & ","
        sqlcmd = sqlcmd & " side_moyou = " & w_str(23) & ", side_kenti = " & w_str(24) & ","
        sqlcmd = sqlcmd & " peak_mark = " & w_str(25) & ", nasiji = " & w_str(26) & ","
        sqlcmd = sqlcmd & " comment = " & w_str(27) & ", dep_name = " & w_str(28) & ","
        sqlcmd = sqlcmd & " entry_name = " & w_str(29) & ", entry_date = " & w_str(30) & ","
        sqlcmd = sqlcmd & " hm_num = " & w_str(31)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ( "
        sqlcmd = sqlcmd & " id = 'AT-B' AND"
        sqlcmd = sqlcmd & " no1 = '" & form_no.w_no1.Text & "' AND"
        sqlcmd = sqlcmd & " no2 = '" & form_no.w_no2.Text & "' )"

        'ｺﾏﾝﾄﾞ実行
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


		end_sql()
		
        'Brand Ver.3 追加
		'現データ削除
		init_sql()


        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "DELETE FROM " & DBTableName2 & " WHERE ( ")
        'result = SqlCmd(SqlConn, "id = 'AT-B' AND ")
        'result = SqlCmd(SqlConn, "no1 = '" & form_no.w_no1.Text & "' AND ")
        'result = SqlCmd(SqlConn, "no2 = '" & form_no.w_no2.Text & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)


        sqlcmd = "DELETE FROM " & DBTableName2 & " WHERE ( "
        sqlcmd = sqlcmd & "id = 'AT-B' AND "
        sqlcmd = sqlcmd & "no1 = '" & form_no.w_no1.Text & "' AND "
        sqlcmd = sqlcmd & "no2 = '" & form_no.w_no2.Text & "' )"

        'ｺﾏﾝﾄﾞ実行
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not delete the existing data from the database.(" & DBTableName2 & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()
		
		'新規登録
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            init_sql()

            w_str(1) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '図面番号
            w_str(3) = "'" & form_no.w_no2.Text & "'" '変番
            w_str(4) = i '編集文字番号
            w_str(5) = "'" & temp_bz.hm_name(i) & "'" '編集文字コード


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'result = SqlCmd(SqlConn, w_str(1) & ", ")
            'result = SqlCmd(SqlConn, w_str(2) & ", ")
            'result = SqlCmd(SqlConn, w_str(3) & ", ")
            'result = SqlCmd(SqlConn, w_str(4) & ", ")
            'result = SqlCmd(SqlConn, w_str(5) & " )")
            '
            'result = SqlExec(SqlConn)
            'If result = FAIL Then GoTo error_section
            'result = SqlResults(SqlConn)


            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            sqlcmd = sqlcmd & w_str(1) & ", "
            sqlcmd = sqlcmd & w_str(2) & ", "
            sqlcmd = sqlcmd & w_str(3) & ", "
            sqlcmd = sqlcmd & w_str(4) & ", "
            sqlcmd = sqlcmd & w_str(5) & " )"

            'ｺﾏﾝﾄﾞ実行
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i
		
		
		bz_update = True
		
		Exit Function
		
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
        ' <- watanabe add VerUP(2011)

        bz_update = FAIL
    End Function

	Function bz_addnum() As Short
		Dim j As Object
		Dim i As Object
        Dim result As Integer '20100707 修正
		Dim now_time As Object
		Dim henban As Object
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        '--------------< ブランド図面 登録 変番 >-----------------------------

		
		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		
		'変番を自動連番します
		henban = what_no2_BZ(temp_bz.no1)
		
		If henban = -1 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("変番の自動連番に失敗しました")
            ErrMsg = "Failed to auto sequence number of a variable number."
            ErrTtl = "Brand drawing revision number registration"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        ElseIf henban = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("図面が登録されていません。新規で登録して下さい。", MsgBoxStyle.Critical, "図面未登録")
            ErrMsg = "Drawing is not registered. Please sign up for new."
            ErrTtl = "Drawing Unregistered"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If

        '----- .NET 移行 -----
        'form_no.w_no2.Text = VB6.Format(henban, "00")
        form_no.w_no2.Text = henban.ToString("00")

        temp_bz.no2 = form_no.w_no2.Text
		
		MsgBox("addnum:no1=[" & temp_bz.no1 & temp_bz.no2 & "]")
		
		
		w_str(1) = "0" '削除フラグ
		w_str(2) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
		w_str(3) = "'" & form_no.w_no1.Text & "'" '図面番号
		w_str(4) = "'" & form_no.w_no2.Text & "'" '変番
		w_str(5) = "'" & form_no.w_kanri_no.Text & "'" '業務管理番号
		w_str(6) = "'" & form_no.w_syurui.Text & "'" 'タイヤ種類
		w_str(7) = "'" & form_no.w_syubetu.Text & "'" 'パターン種別
		w_str(8) = "'" & form_no.w_pattern.Text & "'" 'パターン
		w_str(9) = "'" & form_no.w_size.Text & "'" 'サイズ
		w_str(10) = "'" & form_no.w_size1.Text & "'" 'サイズ1(外径)
		w_str(11) = "'" & form_no.w_size2.Text & "'" 'サイズ2(継)
		w_str(12) = "'" & form_no.w_size3.Text & "'" 'サイズ3(断面幅)
		w_str(13) = "'" & form_no.w_size4.Text & "'" 'サイズ4(速度)
		w_str(14) = "'" & form_no.w_size5.Text & "'" 'サイズ5(構造)
		w_str(15) = "'" & form_no.w_size6.Text & "'" 'サイズ6(リム径)
		w_str(16) = "'" & form_no.w_size7.Text & "'" 'サイズ7(接尾)
		w_str(17) = "'" & form_no.w_size8.Text & "'" 'サイズ8(プライ)
		w_str(18) = "'" & form_no.w_size_code.Text & "'" 'サイズコード
		w_str(19) = "'" & form_no.w_kikaku.Text & "'" '規格
		w_str(20) = "'" & Mid(form_no.w_plant.Text, 4, 2) & "'" '工場
		w_str(21) = "'" & form_no.w_plant_code.Text & "'" '工場コード
		w_str(22) = CStr(Val(Mid(form_no.w_tos_moyou.Text, 1, 1))) 'ＴＯＳ対応模様
		w_str(23) = CStr(Val(Mid(form_no.w_side_moyou.Text, 1, 1))) 'サイド凹凸模様
		w_str(24) = CStr(Val(Mid(form_no.w_side_kenti.Text, 1, 1))) 'サイド凹凸検知
		w_str(25) = CStr(Val(Mid(form_no.w_peak_mark.Text, 1, 1))) 'ピークマーク
		w_str(26) = CStr(Val(Mid(form_no.w_nasiji.Text, 1, 1))) '梨地加工
		' w_str(22) = Mid$(form_no.w_tos_moyou.Text, 1, 1)          'ＴＯＳ対応模様
		' w_str(23) = Mid$(form_no.w_side_moyou.Text, 1, 1)        'サイド凹凸模様
		' w_str(24) = Mid$(form_no.w_side_kenti.Text, 1, 1)        'サイド凹凸検知
		' w_str(25) = Mid$(form_no.w_peak_mark.Text, 1, 1)         'ピークマーク
		' w_str(26) = Mid$(form_no.w_nasiji.Text, 1, 1)            '梨地加工
		w_str(27) = "'" & form_no.w_comment.Text & "'" 'コメント
		w_str(28) = "'" & form_no.w_dep_name.Text & "'" '部署コード
		w_str(29) = "'" & form_no.w_entry_name.Text & "'" '登録者
		
		If Len(Hour(TimeOfDay)) = 1 Then
			now_time = "0" & Hour(TimeOfDay)
		Else
			now_time = Hour(TimeOfDay)
		End If
		
		If Len(Minute(TimeOfDay)) = 1 Then
			now_time = Trim(now_time) & ":0" & Minute(TimeOfDay)
		Else
			now_time = Trim(now_time) & ":" & Minute(TimeOfDay)
		End If
		
		w_str(30) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '登録日
		
        w_str(31) = CStr(Val(Trim(form_no.w_hm_num.Text))) '編集文字数
		' w_str(31) = form_no.w_hm_num.Text                         '編集文字数
		
		'ブランド図面ﾌｧｲﾙに登録
		' result = SqlCmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES( ")
		'  w_command = "INSERT INTO " & DBTableName & " VALUES("
        '
		'For i = 1 To 31
		'     result = SqlCmd(SqlConn, w_str(i) & ",")
		'          w_command = w_command & w_str(i) & ","
		'
		' Next i
        '
		' For i = 1 To temp_bz.hm_num - 1
		' For i = 1 To 99
		'    result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(i) & "',")
		'          w_command = w_command & "'" & temp_bz.hm_name(i) & "',"
        '
		' Next i
		'      result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(temp_bz.hm_num) & "' )")
		'          w_command = w_command & "'" & temp_bz.hm_name(temp_bz.hm_num) & "')"
		'      result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(100) & "' )")
		'          w_command = w_command & "'" & temp_bz.hm_name(100) & "')"
        '
		' If temp_bz.hm_num < 100 Then
		'    For i = 1 To 100 - temp_bz.hm_num - 1
		'       result = SqlCmd(SqlConn, "'" & Space$(8) & "',")
		'           w_command = w_command & "'" & Space$(8) & "',"
		'   Next i
		' End If
        '
		'       result = SqlCmd(SqlConn, "'" & Space$(8) & "',")
		'           w_command = w_command & "'" & Space$(8) & "',"
		' result = SqlCmd(SqlConn, ")")
        '
		' result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(temp_bz.hm_num) & "' )")
		'          w_command = w_command & " )"
        '
		'MsgBox "SQL=[" & w_command & "]"
        '
		' For i = 1 To temp_bz.hm_num - 1
		'     result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(i) & "',")
		' Next i
        '
		' result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(temp_bz.hm_num) & "'")
		' If temp_bz.hm_num < 100 Then
		'    For i = 1 To 100 - temp_bz.hm_num
		'       result = SqlCmd(SqlConn, ",'" & Space$(8) & "'")
		'    Next i
		' End If
        '
		' result = SqlCmd(SqlConn, ")")
		

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        '' Brand Ver.3 変更
        '' For i = 1 To 31
        'For i = 1 To 30
        '	result = SqlCmd(SqlConn, w_str(i) & ",")
        'Next i
        'result = SqlCmd(SqlConn, w_str(i))
        '
        '' Brand Ver.3 変更
        '' For i = 1 To temp_bz.hm_num - 1
        ''     result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(i) & "',")
        '' Next i
        '' result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(temp_bz.hm_num) & "'")
        ''  If temp_bz.hm_num < 100 Then
        ''    For i = 1 To 100 - temp_bz.hm_num
        ''       result = SqlCmd(SqlConn, ",'" & Space$(8) & "'")
        ''    Next i
        '' End If
        '
        'result = SqlCmd(SqlConn, ")")
        'result = SqlExec(SqlConn)
        '
        ''ブランド図面の登録に失敗した時
        'If result = FAIL Then
        '	GoTo error_section
        'End If
        'result = SqlResults(SqlConn)
		

        sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        For i = 1 To 30
            sqlcmd = sqlcmd & w_str(i) & ","
        Next i
        sqlcmd = sqlcmd & w_str(i)
        sqlcmd = sqlcmd & ")"

        'ｺﾏﾝﾄﾞ実行
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()
		
		' Brand Ver.3 変更
		For i = 1 To temp_bz.hm_num
			init_sql()

            w_str(1) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
			w_str(2) = "'" & form_no.w_no1.Text & "'" '図面番号
			w_str(3) = "'" & form_no.w_no2.Text & "'" '変番
			w_str(4) = i '編集文字番号


            ' -> watanabe edit VerUP(2011)
            '         result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '	result = SqlCmd(SqlConn, w_str(i) & ",")
            'Next j
            'result = SqlCmd(SqlConn, "'" & temp_bz.hm_name(i) & "'")
            'result = SqlCmd(SqlConn, " )")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then
            '	GoTo error_section
            'End If
            'result = SqlResults(SqlConn)


            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            For j = 1 To 4
                sqlcmd = sqlcmd & w_str(i) & ","
            Next j
            sqlcmd = sqlcmd & "'" & temp_bz.hm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            'ｺﾏﾝﾄﾞ実行
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
		Next i
		

		'テーブル検索
		' result = SqlCmd(SqlConn, "INSERT INTO " & DBTableName)
		' result = SqlCmd(SqlConn, " VALUES( comment = '" & form_no.w_comment & "',")
		' result = SqlCmd(SqlConn, " dep_name ='" & form_no.w_dep_name & "',")
		' result = SqlCmd(SqlConn, " entry_name ='" & form_no.w_entry_name & "',")
		' result = SqlCmd(SqlConn, " entry_date ='" & form_no.w_entry_date & "',")
		' result = SqlCmd(SqlConn, " hm_num =" & form_no.w_hm_num & ",")
        '
		' For i = 1 To form_no.w_hm_num - 1
        '    result = SqlCmd(SqlConn, " hm_name" & Format(i, "000") & "='" & temp_bz.hm_name(i) & "',")
		' Next i
        '
		' result = SqlCmd(SqlConn, " hm_name" & Format(form_no.w_hm_num, "000") & "='" & temp_bz.hm_name(form_no.w_hm_num) & "'")
        '
		' result = SqlExec(SqlConn)
        ' If result = FAIL Then GoTo error_section
		' result = SqlResults(SqlConn)
		' MsgBox "addnum:444"
		
		bz_addnum = True
		
		Exit Function
		
error_section: 
        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        MsgBox("Brand registration (change number): D / B registration error! !")
        bz_addnum = FAIL
    End Function
	
	Function bz_search(ByRef bz_code1 As String, ByRef bz_code2 As String, ByRef flag As Short) As Short
		Dim i As Object
		Dim w_ret As Object
        Dim result As Integer '20100707 修正
        Dim w_str(42) As String
        Dim ww As String

        ' -> watanabe del VerUP(2011)
        'Dim df As DateInfo
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        Dim errflg As Integer
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        errflg = 0
        ' <- watanabe add VerUP(2011)


        'flag 0:削除フラグが0のデータのみ検索
        'flag 1:すべてのデータを検索


        If SqlConn = 0 Then
            MsgBox("Can not access the database.", MsgBoxStyle.Critical, "SQL error")
            ' -> watanabe add VerUP(2011)
            errflg = 1
            ' <- watanabe add VerUP(2011)
            GoTo error_section
        End If
		
		'BZ_KANRIテーブルより該当する原始文字データを求める
		temp_bz.no1 = bz_code1
		temp_bz.no2 = bz_code2
		

        ' -> watanabe edit VerUP(2011)
        '      w_command = "SELECT *"
        ''w_command = "SELECT comment, dep_name, entry_name, entry_date, hm_num"
        ''For i = 1 To 63
        ''   w_command = w_command & ", hm_name" & Format(i, "000")
        ''Next i
        'w_command = w_command & " FROM " & DBTableName
        'If flag = 0 Then
        '	w_command = w_command & " WHERE (flag_delete = 0 AND no1 = '" & temp_bz.no1 & "' AND"
        'Else
        '	w_command = w_command & " WHERE (no1 = '" & temp_bz.no1 & "' AND"
        'End If
        'w_command = w_command & " no2 = '" & temp_bz.no2 & "')"
        '
        'result = SqlCmd(SqlConn, w_command)
        '
        '
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	'   Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		temp_bz.flag_delete = Val(SqlData(SqlConn, 1))
        '
        '		' -> watanabe edit 2007.06
        '		'     temp_bz.id = Val(SqlData$(SqlConn, 2))
        '		temp_bz.id = SqlData(SqlConn, 2)
        '		' <- watanabe edit 2007.06
        '
        '		temp_bz.no1 = SqlData(SqlConn, 3)
        '		temp_bz.no2 = SqlData(SqlConn, 4)
        '		temp_bz.kanri_no = SqlData(SqlConn, 5)
        '		temp_bz.syurui = SqlData(SqlConn, 6)
        '		temp_bz.syubetu = SqlData(SqlConn, 7)
        '		temp_bz.pattern = SqlData(SqlConn, 8)
        '		temp_bz.Size = SqlData(SqlConn, 9)
        '		temp_bz.size1 = SqlData(SqlConn, 10)
        '		temp_bz.size2 = SqlData(SqlConn, 11)
        '		temp_bz.size3 = SqlData(SqlConn, 12)
        '		temp_bz.size4 = SqlData(SqlConn, 13)
        '		temp_bz.size5 = SqlData(SqlConn, 14)
        '		temp_bz.size6 = SqlData(SqlConn, 15)
        '		temp_bz.size7 = SqlData(SqlConn, 16)
        '		temp_bz.size8 = SqlData(SqlConn, 17)
        '		temp_bz.size_code = SqlData(SqlConn, 18)
        '		temp_bz.kikaku = SqlData(SqlConn, 19)
        '		temp_bz.plant = SqlData(SqlConn, 20)
        '		temp_bz.plant_code = SqlData(SqlConn, 21)
        '		temp_bz.tos_moyou = Val(SqlData(SqlConn, 22))
        '		temp_bz.side_moyou = Val(SqlData(SqlConn, 23))
        '		temp_bz.side_kenti = Val(SqlData(SqlConn, 24))
        '		temp_bz.peak_mark = Val(SqlData(SqlConn, 25))
        '		temp_bz.nasiji = Val(SqlData(SqlConn, 26))
        '		temp_bz.comment = SqlData(SqlConn, 27)
        '		temp_bz.dep_name = SqlData(SqlConn, 28)
        '		temp_bz.entry_name = SqlData(SqlConn, 29)
        '		ww = SqlData(SqlConn, 30)
        '		w_ret = SqlDateCrack(SqlConn, df, ww)
        '		temp_bz.entry_date = df.Year_Renamed & df.Month_Renamed & df.Day_Renamed
        '
        '		temp_bz.hm_num = Val(SqlData(SqlConn, 31))
        '
        '		' Brand Ver.3 変更
        '		'     For i = 1 To 100
        '		'       temp_bz.hm_name(i) = SqlData$(SqlConn, 31 + i)
        '		'     Next i
        '	Else
        '		GoTo error_section
        '	End If
        '
        '	'   Loop
        'Else
        '	GoTo error_section
        'End If


        '検索キーセット
        If flag = 0 Then
            key_code = "flag_delete = 0 AND no1 = '" & temp_bz.no1 & "' AND"
        Else
            key_code = "no1 = '" & temp_bz.no1 & "' AND"
        End If
        key_code = key_code & " no2 = '" & temp_bz.no2 & "'"

        '検索コマンド作成
        sqlcmd = "SELECT * FROM " & DBTableName & " WHERE (" & key_code & ")"

        'ヒット数チェック
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            errflg = 1
            GoTo error_section
        ElseIf cnt = -1 Then
            errflg = 1
            GoTo error_section
        End If

        '検索
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            temp_bz.flag_delete = Val(Rs.rdoColumns(0).Value)
        Else
            temp_bz.flag_delete = 0
        End If

        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
            temp_bz.id = Rs.rdoColumns(1).Value
        Else
            temp_bz.id = ""
        End If

        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
            temp_bz.no1 = Rs.rdoColumns(2).Value
        Else
            temp_bz.no1 = ""
        End If

        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
            temp_bz.no2 = Rs.rdoColumns(3).Value
        Else
            temp_bz.no2 = ""
        End If

        If IsDBNull(Rs.rdoColumns(4).Value) = False Then
            temp_bz.kanri_no = Rs.rdoColumns(4).Value
        Else
            temp_bz.kanri_no = ""
        End If

        If IsDBNull(Rs.rdoColumns(5).Value) = False Then
            temp_bz.syurui = Rs.rdoColumns(5).Value
        Else
            temp_bz.syurui = ""
        End If

        If IsDBNull(Rs.rdoColumns(6).Value) = False Then
            temp_bz.syubetu = Rs.rdoColumns(6).Value
        Else
            temp_bz.syubetu = ""
        End If

        If IsDBNull(Rs.rdoColumns(7).Value) = False Then
            temp_bz.pattern = Rs.rdoColumns(7).Value
        Else
            temp_bz.pattern = ""
        End If

        If IsDBNull(Rs.rdoColumns(8).Value) = False Then
            temp_bz.Size = Rs.rdoColumns(8).Value
        Else
            temp_bz.Size = ""
        End If

        If IsDBNull(Rs.rdoColumns(9).Value) = False Then
            temp_bz.size1 = Rs.rdoColumns(9).Value
        Else
            temp_bz.size1 = ""
        End If

        If IsDBNull(Rs.rdoColumns(10).Value) = False Then
            temp_bz.size2 = Rs.rdoColumns(10).Value
        Else
            temp_bz.size2 = ""
        End If

        If IsDBNull(Rs.rdoColumns(11).Value) = False Then
            temp_bz.size3 = Rs.rdoColumns(11).Value
        Else
            temp_bz.size3 = ""
        End If

        If IsDBNull(Rs.rdoColumns(12).Value) = False Then
            temp_bz.size4 = Rs.rdoColumns(12).Value
        Else
            temp_bz.size4 = ""
        End If

        If IsDBNull(Rs.rdoColumns(13).Value) = False Then
            temp_bz.size5 = Rs.rdoColumns(13).Value
        Else
            temp_bz.size5 = ""
        End If

        If IsDBNull(Rs.rdoColumns(14).Value) = False Then
            temp_bz.size6 = Rs.rdoColumns(14).Value
        Else
            temp_bz.size6 = ""
        End If

        If IsDBNull(Rs.rdoColumns(15).Value) = False Then
            temp_bz.size7 = Rs.rdoColumns(15).Value
        Else
            temp_bz.size7 = ""
        End If

        If IsDBNull(Rs.rdoColumns(16).Value) = False Then
            temp_bz.size8 = Rs.rdoColumns(16).Value
        Else
            temp_bz.size8 = ""
        End If

        If IsDBNull(Rs.rdoColumns(17).Value) = False Then
            temp_bz.size_code = Rs.rdoColumns(17).Value
        Else
            temp_bz.size_code = ""
        End If

        If IsDBNull(Rs.rdoColumns(18).Value) = False Then
            temp_bz.kikaku = Rs.rdoColumns(18).Value
        Else
            temp_bz.kikaku = ""
        End If

        If IsDBNull(Rs.rdoColumns(19).Value) = False Then
            temp_bz.plant = Rs.rdoColumns(19).Value
        Else
            temp_bz.plant = ""
        End If

        If IsDBNull(Rs.rdoColumns(20).Value) = False Then
            temp_bz.plant_code = Rs.rdoColumns(20).Value
        Else
            temp_bz.plant_code = ""
        End If

        If IsDBNull(Rs.rdoColumns(21).Value) = False Then
            temp_bz.tos_moyou = Val(Rs.rdoColumns(21).Value)
        Else
            temp_bz.tos_moyou = 0
        End If

        If IsDBNull(Rs.rdoColumns(22).Value) = False Then
            temp_bz.side_moyou = Val(Rs.rdoColumns(22).Value)
        Else
            temp_bz.side_moyou = 0
        End If

        If IsDBNull(Rs.rdoColumns(23).Value) = False Then
            temp_bz.side_kenti = Val(Rs.rdoColumns(23).Value)
        Else
            temp_bz.side_kenti = 0
        End If

        If IsDBNull(Rs.rdoColumns(24).Value) = False Then
            temp_bz.peak_mark = Val(Rs.rdoColumns(24).Value)
        Else
            temp_bz.peak_mark = 0
        End If

        If IsDBNull(Rs.rdoColumns(25).Value) = False Then
            temp_bz.nasiji = Val(Rs.rdoColumns(25).Value)
        Else
            temp_bz.nasiji = 0
        End If

        If IsDBNull(Rs.rdoColumns(26).Value) = False Then
            temp_bz.comment = Rs.rdoColumns(26).Value
        Else
            temp_bz.comment = ""
        End If

        If IsDBNull(Rs.rdoColumns(27).Value) = False Then
            temp_bz.dep_name = Rs.rdoColumns(27).Value
        Else
            temp_bz.dep_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(28).Value) = False Then
            temp_bz.entry_name = Rs.rdoColumns(28).Value
        Else
            temp_bz.entry_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(29).Value) = False Then
            Dim tmpstr As String
            tmpstr = Rs.rdoColumns(29).Value
            temp_bz.entry_date = Left(tmpstr, 4) & Mid(tmpstr, 6, 2) & Mid(tmpstr, 9, 2)
        Else
            temp_bz.entry_date = ""
        End If

        If IsDBNull(Rs.rdoColumns(30).Value) = False Then
            temp_bz.hm_num = Val(Rs.rdoColumns(30).Value)
        Else
            temp_bz.hm_num = 0
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 追加
        For i = 1 To temp_bz.hm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT hm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ( "
            'w_command = w_command & " no1 = '" & temp_bz.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_bz.no2 & "' AND"
            'w_command = w_command & " hm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_bz.hm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '検索キーセット
            key_code = " no1 = '" & temp_bz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_bz.no2 & "' AND"
            key_code = key_code & " hm_no = " & i

            '検索コマンド作成
            sqlcmd = "SELECT hm_name FROM " & DBTableName2 & " WHERE ( " & key_code & " )"

            'ヒット数チェック
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Then
                Exit For
            ElseIf cnt = -1 Then
                Exit For
            End If

            '検索
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_bz.hm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_bz.hm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i


        bz_search = True

        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        If errflg = 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "System error")
        End If

        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        bz_search = FAIL
    End Function
	
	Sub bz_spec_set(ByRef hexdata As String)
		Dim w_ret As Integer
		Dim www As New VB6.FixedLengthString(4)
		Dim ww2 As New VB6.FixedLengthString(2)


        ' -> watanabe edit VerUP(2011)
        'hexdata = ""
        ''========================================
        ''TEMP_BZデータをＨＥＸに変換します
        ''========================================
        '
        'hexdata = hexdata & temp_bz.id
        '
        '' -> watanabe Edit 2007.06
        ''    hexdata = hexdata & temp_bz.no1
        'If Len(Trim(temp_bz.no1)) = 4 Then
        '	hexdata = hexdata & Trim(temp_bz.no1) & " "
        'Else
        '	hexdata = hexdata & temp_bz.no1
        'End If
        '' <- watanabe Edit 2007.06
        '
        'hexdata = hexdata & temp_bz.no2
        '
        'hexdata = hexdata & temp_bz.kanri_no
        'hexdata = hexdata & temp_bz.syurui
        'hexdata = hexdata & temp_bz.syubetu
        '
        'hexdata = hexdata & temp_bz.pattern
        '
        'hexdata = hexdata & temp_bz.Size
        'hexdata = hexdata & temp_bz.size1
        'hexdata = hexdata & temp_bz.size2
        'hexdata = hexdata & temp_bz.size3
        'hexdata = hexdata & temp_bz.size4
        'hexdata = hexdata & temp_bz.size5
        'hexdata = hexdata & temp_bz.size6
        'hexdata = hexdata & temp_bz.size7
        'hexdata = hexdata & temp_bz.size8
        'hexdata = hexdata & temp_bz.size_code
        'hexdata = hexdata & temp_bz.kikaku
        '
        'Select Case temp_bz.plant
        '	Case "仙台"
        '		ww2.Value = "TT"
        '	Case "桑名"
        '		ww2.Value = "NW"
        '	Case "正新"
        '		ww2.Value = "CS"
        '	Case "上海"
        '		ww2.Value = "CH"
        'End Select
        '
        'hexdata = hexdata & ww2.Value
        'hexdata = hexdata & temp_bz.plant_code
        '
        'w_ret = ShttoHex(temp_bz.tos_moyou, www.Value)
        'hexdata = hexdata & www.Value
        '
        'w_ret = ShttoHex(temp_bz.side_moyou, www.Value)
        'hexdata = hexdata & www.Value
        '
        'w_ret = ShttoHex(temp_bz.side_kenti, www.Value)
        'hexdata = hexdata & www.Value
        '
        'w_ret = ShttoHex(temp_bz.peak_mark, www.Value)
        'hexdata = hexdata & www.Value
        '
        'w_ret = ShttoHex(temp_bz.nasiji, www.Value)
        'hexdata = hexdata & www.Value


        Dim ii As Integer

        ' 必要文字数分、スペースで初期化
        hexdata = ""
        For ii = 1 To 110
            hexdata = hexdata & " "
        Next ii

        ' 必要文字数分、スペースで初期化
        www.Value = ""
        For ii = 1 To 4
            www.Value = www.Value & " "
        Next ii

        ' 必要文字数分、スペースで初期化
        ww2.Value = ""
        For ii = 1 To 2
            ww2.Value = ww2.Value & " "
        Next ii


        '========================================
        'TEMP_BZデータをＨＥＸに変換します
        '========================================

        Mid(hexdata, 1, 4) = temp_bz.id
        Mid(hexdata, 5, 5) = temp_bz.no1
        Mid(hexdata, 10, 2) = temp_bz.no2

        Mid(hexdata, 12, 8) = temp_bz.kanri_no
        Mid(hexdata, 20, 2) = temp_bz.syurui
        Mid(hexdata, 22, 3) = temp_bz.syubetu

        Mid(hexdata, 25, 6) = temp_bz.pattern

        Mid(hexdata, 31, 21) = temp_bz.Size
        Mid(hexdata, 52, 5) = temp_bz.size1
        Mid(hexdata, 57, 1) = temp_bz.size2
        Mid(hexdata, 58, 5) = temp_bz.size3
        Mid(hexdata, 63, 1) = temp_bz.size4
        Mid(hexdata, 64, 1) = temp_bz.size5
        Mid(hexdata, 65, 4) = temp_bz.size6
        Mid(hexdata, 69, 2) = temp_bz.size7
        Mid(hexdata, 71, 2) = temp_bz.size8
        Mid(hexdata, 73, 2) = temp_bz.size_code
        Mid(hexdata, 75, 10) = temp_bz.kikaku

        Select Case temp_bz.plant
            Case "Sendai"
                ww2.Value = "TT"
            Case "Kuwana"
                ww2.Value = "NW"
            Case "Cheng shin"
                ww2.Value = "CS"
            Case "Shanghai"
                ww2.Value = "CH"
        End Select

        Mid(hexdata, 85, 2) = ww2.Value
        Mid(hexdata, 87, 2) = temp_bz.plant_code

        w_ret = ShttoHex(temp_bz.tos_moyou, www.Value)
        Mid(hexdata, 89, 4) = www.Value

        w_ret = ShttoHex(temp_bz.side_moyou, www.Value)
        Mid(hexdata, 93, 4) = www.Value

        w_ret = ShttoHex(temp_bz.side_kenti, www.Value)
        Mid(hexdata, 97, 4) = www.Value

        w_ret = ShttoHex(temp_bz.peak_mark, www.Value)
        Mid(hexdata, 101, 4) = www.Value

        w_ret = ShttoHex(temp_bz.nasiji, www.Value)
        Mid(hexdata, 105, 4) = www.Value
        ' <- watanabe edit VerUP(2011)

    End Sub
	
	Sub temp_bz_get(ByRef syori_flg As Short)
		
		'パラメータ
		'syori_flg :1 = サイズのみ
		'           2 = サイズとタイヤ種類
		'           3 = サイズ、タイヤ種類、工場
		'           4 = 全部
		
		
		If Trim(form_no.w_size1.Text) <> "" Then
			If Trim(form_no.w_size1.Text) <> Trim(temp_bz.size1) Then
				temp_bz.size1 = LSet(form_no.w_size1.Text, Len(temp_bz.size1))
			End If
		End If
		
		
		If Trim(form_no.w_size2.Text) <> "" Then
			If Trim(form_no.w_size2.Text) <> Trim(temp_bz.size2) Then
				temp_bz.size2 = LSet(form_no.w_size2.Text, Len(temp_bz.size2))
			End If
		End If
		
		If Trim(form_no.w_size3.Text) <> "" Then
			If Trim(form_no.w_size3.Text) <> Trim(temp_bz.size3) Then
				temp_bz.size3 = LSet(form_no.w_size3.Text, Len(temp_bz.size3))
			End If
		End If
		
		If Trim(form_no.w_size5.Text) <> "" Then
			If Trim(form_no.w_size5.Text) <> Trim(temp_bz.size5) Then
				temp_bz.size5 = LSet(form_no.w_size5.Text, Len(temp_bz.size5))
			End If
		End If
		
		If Trim(form_no.w_size6.Text) <> "" Then
			If Trim(form_no.w_size6.Text) <> Trim(temp_bz.size6) Then
				temp_bz.size6 = LSet(form_no.w_size6.Text, Len(temp_bz.size6))
			End If
		End If
		If syori_flg < 2 Then Exit Sub
		
		If Trim(form_no.w_syurui.Text) <> Trim(temp_bz.syurui) Then
			temp_bz.syurui = LSet(form_no.w_syurui.Text, Len(temp_bz.syurui))
		End If
		
		If syori_flg < 3 Then Exit Sub
		
		If Trim(form_no.w_plant_code.Text) <> Trim(temp_bz.plant_code) Then
			temp_bz.plant_code = LSet(form_no.w_plant_code.Text, Len(temp_bz.plant_code))
			Select Case temp_bz.plant_code
				Case "CX"
					temp_bz.plant = "TT"
				Case "N3"
					temp_bz.plant = "KW"
				Case "UY"
					temp_bz.plant = "CS"
				Case "9T"
					temp_bz.plant = "CH"
			End Select
		End If
		
		If syori_flg < 4 Then Exit Sub
		
		If Trim(form_no.w_id.Text) <> "" Then
			If Trim(form_no.w_id.Text) <> Trim(temp_bz.id) Then
				temp_bz.id = LSet(form_no.w_id.Text, Len(temp_bz.id))
			End If
		End If
		
		If Trim(form_no.w_no1.Text) <> "" Then
			If Trim(form_no.w_no1.Text) <> Trim(temp_bz.no1) Then
				temp_bz.no1 = LSet(form_no.w_no1.Text, Len(temp_bz.no1))
			End If
		End If
		
		If Trim(form_no.w_no2.Text) <> "" Then
			If Trim(form_no.w_no2.Text) <> Trim(temp_bz.no2) Then
				temp_bz.no2 = LSet(form_no.w_no2.Text, Len(temp_bz.no2))
			End If
		End If
		
		If Trim(form_no.w_kanri_no.Text) <> "" Then
			If Trim(form_no.w_kanri_no.Text) <> Trim(temp_bz.kanri_no) Then
				temp_bz.kanri_no = LSet(form_no.w_kanri_no.Text, Len(temp_bz.kanri_no))
			End If
		End If
		If Trim(form_no.w_syurui.Text) <> "" Then
			If Trim(form_no.w_syurui.Text) <> Trim(temp_bz.syurui) Then
				temp_bz.syurui = LSet(form_no.w_syurui.Text, Len(temp_bz.syurui))
			End If
		End If
		If Trim(form_no.w_syubetu.Text) <> "" Then
			If Trim(form_no.w_syubetu.Text) <> Trim(temp_bz.syubetu) Then
				temp_bz.syubetu = LSet(form_no.w_syubetu.Text, Len(temp_bz.syubetu))
			End If
		End If
		If Trim(form_no.w_pattern.Text) <> "" Then
			If Trim(form_no.w_pattern.Text) <> Trim(temp_bz.pattern) Then
				temp_bz.pattern = LSet(form_no.w_pattern.Text, Len(temp_bz.pattern))
			End If
		End If
		If Trim(form_no.w_size.Text) <> "" Then
			If Trim(form_no.w_size.Text) <> Trim(temp_bz.Size) Then
				temp_bz.Size = LSet(form_no.w_size.Text, Len(temp_bz.Size))
			End If
		End If
		If Trim(form_no.w_size_code.Text) <> "" Then
			If Trim(form_no.w_size_code.Text) <> Trim(temp_bz.size_code) Then
				temp_bz.size_code = LSet(form_no.w_size_code.Text, Len(temp_bz.size_code))
			End If
		End If
		If Trim(form_no.w_kikaku.Text) <> "" Then
			If Trim(form_no.w_kikaku.Text) <> Trim(temp_bz.kikaku) Then
				temp_bz.kikaku = LSet(form_no.w_kikaku.Text, Len(temp_bz.kikaku))
			End If
		End If
		If Trim(form_no.w_plant.Text) <> "" Then
			If Trim(form_no.w_plant.Text) <> Trim(temp_bz.plant) Then
				temp_bz.plant = LSet(form_no.w_plant.Text, Len(temp_bz.plant))
			End If
		End If
		If Trim(form_no.w_plant_code.Text) <> "" Then
			If Trim(form_no.w_plant_code.Text) <> Trim(temp_bz.plant_code) Then
				temp_bz.plant_code = LSet(form_no.w_plant_code.Text, Len(temp_bz.plant_code))
			End If
		End If

        '20100705コード変更
        'temp_bz.tos_moyou = Val(Mid(form_no.w_tos_moyou, 1, 1))
        'temp_bz.side_moyou = Val(Mid(form_no.w_side_moyou, 1, 1))
        'temp_bz.side_kenti = Val(Mid(form_no.w_side_kenti, 1, 1))
        'temp_bz.peak_mark = Val(Mid(form_no.w_peak_mark, 1, 1))
        'temp_bz.nasiji = Val(Mid(form_no.w_nasiji, 1, 1))
        temp_bz.tos_moyou = Val(Mid(form_no.w_tos_moyou.Text, 1, 1))
        temp_bz.side_moyou = Val(Mid(form_no.w_side_moyou.Text, 1, 1))
        temp_bz.side_kenti = Val(Mid(form_no.w_side_kenti.Text, 1, 1))
        temp_bz.peak_mark = Val(Mid(form_no.w_peak_mark.Text, 1, 1))
        temp_bz.nasiji = Val(Mid(form_no.w_nasiji.Text, 1, 1))
		
		
	End Sub
	
	Function temp_bz_set(ByRef flag As Short, ByRef hexdata As String) As Short
		Dim i As Object
		Dim w_ret As Object
		
		Dim aa As String
		
		'========================================
		'ブランド図面特性データをＨＥＸより変換します
		'========================================
		
		If flag = 0 Then
			temp_bz.flag_delete = 0
			temp_bz.id = Mid(hexdata, 1, 4)
			' -> watanabe edit 2007.03
			'    temp_bz.no1 = Mid$(hexdata, 5, 4)
			'    temp_bz.no2 = Mid$(hexdata, 9, 2)
			'    temp_bz.kanri_no = Mid$(hexdata, 11, 8)
			'    temp_bz.syurui = Mid$(hexdata, 19, 2)
			'    temp_bz.syubetu = Mid$(hexdata, 21, 3)
			'    temp_bz.pattern = Mid$(hexdata, 24, 6)
			'    temp_bz.Size = Mid$(hexdata, 30, 21)
			'    temp_bz.size1 = Mid$(hexdata, 51, 5)
			'    temp_bz.size2 = Mid$(hexdata, 56, 1)
			'    temp_bz.size3 = Mid$(hexdata, 57, 5)
			'    temp_bz.size4 = Mid$(hexdata, 62, 1)
			'    temp_bz.size5 = Mid$(hexdata, 63, 1)
			'    temp_bz.size6 = Mid$(hexdata, 64, 4)
			'    temp_bz.size7 = Mid$(hexdata, 68, 2)
			'    temp_bz.size8 = Mid$(hexdata, 70, 2)
			'    temp_bz.size_code = Mid$(hexdata, 72, 2)
			'    temp_bz.kikaku = Mid$(hexdata, 74, 10)
			'    temp_bz.plant = Mid$(hexdata, 84, 2)
			'    temp_bz.plant_code = Mid$(hexdata, 86, 2)
			'    w_ret = HextoSht(Mid$(hexdata, 88, 4), temp_bz.tos_moyou)
			'    w_ret = HextoSht(Mid$(hexdata, 92, 4), temp_bz.side_moyou)
			'    w_ret = HextoSht(Mid$(hexdata, 96, 4), temp_bz.side_kenti)
			'    w_ret = HextoSht(Mid$(hexdata, 100, 4), temp_bz.peak_mark)
			'    w_ret = HextoSht(Mid$(hexdata, 104, 4), temp_bz.nasiji)
			
			' -> watanabe edit 2007.06
			'    temp_bz.no1 = Mid$(hexdata, 5, 5)
			temp_bz.no1 = Trim(Mid(hexdata, 5, 5))
			' <- watanabe edit 2007.06
			
			temp_bz.no2 = Mid(hexdata, 10, 2)
			temp_bz.kanri_no = Mid(hexdata, 12, 8)
			temp_bz.syurui = Mid(hexdata, 20, 2)
			temp_bz.syubetu = Mid(hexdata, 22, 3)
			temp_bz.pattern = Mid(hexdata, 25, 6)
			temp_bz.Size = Mid(hexdata, 31, 21)
			temp_bz.size1 = Mid(hexdata, 52, 5)
			temp_bz.size2 = Mid(hexdata, 57, 1)
			temp_bz.size3 = Mid(hexdata, 58, 5)
			temp_bz.size4 = Mid(hexdata, 63, 1)
			temp_bz.size5 = Mid(hexdata, 64, 1)
			temp_bz.size6 = Mid(hexdata, 65, 4)
			temp_bz.size7 = Mid(hexdata, 69, 2)
			temp_bz.size8 = Mid(hexdata, 71, 2)
			temp_bz.size_code = Mid(hexdata, 73, 2)
			temp_bz.kikaku = Mid(hexdata, 75, 10)
			temp_bz.plant = Mid(hexdata, 85, 2)
			temp_bz.plant_code = Mid(hexdata, 87, 2)
			w_ret = HextoSht(Mid(hexdata, 89, 4), temp_bz.tos_moyou)
			w_ret = HextoSht(Mid(hexdata, 93, 4), temp_bz.side_moyou)
			w_ret = HextoSht(Mid(hexdata, 97, 4), temp_bz.side_kenti)
			w_ret = HextoSht(Mid(hexdata, 101, 4), temp_bz.peak_mark)
			w_ret = HextoSht(Mid(hexdata, 105, 4), temp_bz.nasiji)
			' <- watanabe edit 2007.03

            ' -> watanabe add VerUP(2011)
            aa = ""
            ' <- watanabe add VerUP(2011)

			Call true_date(aa)
			temp_bz.entry_date = aa
			
            If open_mode = "NEW" Then
                temp_bz.id = Mid(hexdata, 1, 4)
                temp_bz.no1 = ""
                temp_bz.no2 = "00"
                temp_bz.comment = ""
                temp_bz.dep_name = ""
                temp_bz.entry_name = ""
                '      Call true_date(aa)
                '      temp_bz.entry_date = aa

            End If
		Else
			temp_bz.hm_num = Val(Mid(hexdata, 1, 3))
			
            ' -> watanabe edit 2013.05.29
            'For i = 1 To 100
            For i = 1 To 260
                ' <- watanabe edit 2013.05.29

                temp_bz.hm_name(i) = Mid(hexdata, (i - 1) * 8 + 4, 8)
            Next i
        End If

	End Function
	
    Function bz_delete(ByRef bz_code1 As String, ByRef bz_code2 As String) As Short
        Dim result As Integer '20100707 修正
        Dim w_str(42) As String
        'Dim w_command As String '20100616移植削除

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)


        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        w_str(1) = "1" '削除フラグ
        w_str(2) = "'" & "AT-B" & "'" 'ＩＤ(AT-B固定)
        w_str(3) = "'" & bz_code1 & "'" '図面番号(****)
        w_str(4) = "'" & bz_code2 & "'" '変番(00~99）


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "UPDATE " & DBTableName)
        'result = SqlCmd(SqlConn, " SET flag_delete = " & w_str(1))
        'result = SqlCmd(SqlConn, " From " & DBTableName & "(PAGLOCK)")
        'result = SqlCmd(SqlConn, " WHERE ( no1 = " & w_str(3) & " AND")
        'result = SqlCmd(SqlConn, " no2 = " & w_str(4) & ")")
        ''Send the command to SQL Server and start execution.
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        ''MsgBox "UPDATE Result = " & result
        'If result <> 1 Then GoTo error_section


        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ( no1 = " & w_str(3) & " AND"
        sqlcmd = sqlcmd & " no2 = " & w_str(4) & ")"

        'ｺﾏﾝﾄﾞ実行
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        bz_delete = True

        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        bz_delete = FAIL
    End Function
	
    Function zumen_no_set_bz(ByRef hexdata As String) As Short
        Dim t26 As String
        Dim t25 As String
        Dim t24 As String
        Dim t23 As String
        Dim t22 As Short
        Dim t21 As Short
        Dim t20 As Short
        Dim t19 As Short
        Dim t18 As Short
        Dim t17 As String
        Dim t16 As String
        Dim t15 As String
        Dim t14 As String
        Dim t13 As String
        Dim t12 As String
        Dim t11 As String
        Dim t10 As String
        Dim t9 As String
        Dim t8 As String
        Dim t7 As String
        Dim t6 As String
        Dim t5 As String
        Dim t4 As String
        Dim t3 As String
        Dim t2 As String
        Dim t1 As String
        Dim result As Integer '20100707 修正
        'Dim aa As String '20100616移植削除
        'Dim nn As Short '20100616移植削除

        ' -> watanabe add VerUP(2011)
        Dim errflg As Integer
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        errflg = 0
        ' <- watanabe add VerUP(2011)


        ' MsgBox "zumen_no:hexdata=[" & hexdata & "]"

        Call init_sql()

        If open_mode = "modify" Then

            temp_bz.id = "AT-B"
            ' -> watanabe Edit 2007.03
            '    temp_bz.no1 = Mid$(hexdata, 6, 4)
            '    temp_bz.no2 = Mid$(hexdata, 11, 2)
            If Mid(hexdata, 10, 1) = "-" Then
                temp_bz.no1 = Mid(hexdata, 6, 4)
                temp_bz.no2 = Mid(hexdata, 11, 2)
            Else
                temp_bz.no1 = Mid(hexdata, 6, 5)
                temp_bz.no2 = Mid(hexdata, 12, 2)
            End If
            ' <- watanabe Edit 2007.03


            ' -> watanabe add VerUP(2011)
            '         result = sqlcmd(SqlConn, "SELECT kanri_no, syurui, syubetu, pattern, size, ")
            'result = SqlCmd(SqlConn, "size1, size2, size3, size4, size5, size6, size7, size8, ")
            'result = SqlCmd(SqlConn, "size_code, kikaku, plant, plant_code, ")
            'result = SqlCmd(SqlConn, "tos_moyou, side_moyou, side_kenti, peak_mark, nasiji, ")
            'result = SqlCmd(SqlConn, "comment, dep_name, entry_name, entry_date ")
            'result = SqlCmd(SqlConn, " FROM " & DBTableName)
            'result = SqlCmd(SqlConn, " WHERE (")
            'result = SqlCmd(SqlConn, " flag_delete = 0 AND")
            'result = SqlCmd(SqlConn, " id = 'AT-B' AND")
            'result = SqlCmd(SqlConn, " no1 = '" & temp_bz.no1 & "' AND")
            'result = SqlCmd(SqlConn, " no2 = '" & temp_bz.no2 & "' )")
            'result = SqlExec(SqlConn)
            '
            'If result = FAIL Then GoTo error_section
            '
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
            '		temp_bz.kanri_no = SqlData(SqlConn, 1)
            '		temp_bz.syurui = SqlData(SqlConn, 2)
            '		temp_bz.syubetu = SqlData(SqlConn, 3)
            '		temp_bz.pattern = SqlData(SqlConn, 4)
            '		temp_bz.Size = SqlData(SqlConn, 5)
            '		temp_bz.size1 = SqlData(SqlConn, 6)
            '		temp_bz.size2 = SqlData(SqlConn, 7)
            '		temp_bz.size3 = SqlData(SqlConn, 8)
            '		temp_bz.size4 = SqlData(SqlConn, 9)
            '		temp_bz.size5 = SqlData(SqlConn, 10)
            '		temp_bz.size6 = SqlData(SqlConn, 11)
            '		temp_bz.size7 = SqlData(SqlConn, 12)
            '		temp_bz.size8 = SqlData(SqlConn, 13)
            '		temp_bz.size_code = SqlData(SqlConn, 14)
            '		temp_bz.kikaku = SqlData(SqlConn, 15)
            '		temp_bz.plant = SqlData(SqlConn, 16)
            '		temp_bz.plant_code = SqlData(SqlConn, 17)
            '		temp_bz.tos_moyou = Val(SqlData(SqlConn, 18))
            '		temp_bz.side_moyou = Val(SqlData(SqlConn, 19))
            '		temp_bz.side_kenti = Val(SqlData(SqlConn, 20))
            '		temp_bz.peak_mark = Val(SqlData(SqlConn, 21))
            '		temp_bz.nasiji = Val(SqlData(SqlConn, 22))
            '		temp_bz.comment = SqlData(SqlConn, 23)
            '		temp_bz.dep_name = SqlData(SqlConn, 24)
            '		temp_bz.entry_name = SqlData(SqlConn, 25)
            '		'          temp_bz.entry_date = SqlData$(SqlConn, 26)
            '	Loop 
            '	If temp_bz.entry_name = "" Then
            '		MsgBox("ブランド図面データがありません" & Chr(13) & "修正処理は出来ません", MsgBoxStyle.Critical, "ｴﾗｰ")
            '		GoTo error_section
            '	End If
            'Else
            '	GoTo error_section
            'End If


            '検索キーセット
            key_code = "flag_delete = 0 AND"
            key_code = key_code & " id = 'AT-B' AND"
            key_code = key_code & " no1 = '" & temp_bz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_bz.no2 & "'"

            '検索コマンド作成
            sqlcmd = "SELECT kanri_no, syurui, syubetu, pattern, size, "
            sqlcmd = sqlcmd & "size1, size2, size3, size4, size5, size6, size7, size8, "
            sqlcmd = sqlcmd & "size_code, kikaku, plant, plant_code, "
            sqlcmd = sqlcmd & "tos_moyou, side_moyou, side_kenti, peak_mark, nasiji, "
            sqlcmd = sqlcmd & "comment, dep_name, entry_name, entry_date "
            sqlcmd = sqlcmd & " FROM " & DBTableName
            sqlcmd = sqlcmd & " WHERE ( " & key_code & " )"

            'ヒット数チェック
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = 0 Then
                MsgBox("There is no brand drawing data." & Chr(13) & "Can not modify processing.", MsgBoxStyle.Critical, "error")
                errflg = 1
                GoTo error_section
            ElseIf cnt = -1 Then
                MsgBox("An error occurred on the existing record during the database search.", MsgBoxStyle.Critical, "error")
                errflg = 1
                GoTo error_section
            End If

            '検索
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_bz.kanri_no = Rs.rdoColumns(0).Value
            Else
                temp_bz.kanri_no = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_bz.syurui = Rs.rdoColumns(1).Value
            Else
                temp_bz.syurui = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_bz.syubetu = Rs.rdoColumns(2).Value
            Else
                temp_bz.syubetu = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_bz.pattern = Rs.rdoColumns(3).Value
            Else
                temp_bz.pattern = ""
            End If

            If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                temp_bz.Size = Rs.rdoColumns(4).Value
            Else
                temp_bz.Size = ""
            End If

            If IsDBNull(Rs.rdoColumns(5).Value) = False Then
                temp_bz.size1 = Rs.rdoColumns(5).Value
            Else
                temp_bz.size1 = ""
            End If

            If IsDBNull(Rs.rdoColumns(6).Value) = False Then
                temp_bz.size2 = Rs.rdoColumns(6).Value
            Else
                temp_bz.size2 = ""
            End If

            If IsDBNull(Rs.rdoColumns(7).Value) = False Then
                temp_bz.size3 = Rs.rdoColumns(7).Value
            Else
                temp_bz.size3 = ""
            End If

            If IsDBNull(Rs.rdoColumns(8).Value) = False Then
                temp_bz.size4 = Rs.rdoColumns(8).Value
            Else
                temp_bz.size4 = ""
            End If

            If IsDBNull(Rs.rdoColumns(9).Value) = False Then
                temp_bz.size5 = Rs.rdoColumns(9).Value
            Else
                temp_bz.size5 = ""
            End If

            If IsDBNull(Rs.rdoColumns(10).Value) = False Then
                temp_bz.size6 = Rs.rdoColumns(10).Value
            Else
                temp_bz.size6 = ""
            End If

            If IsDBNull(Rs.rdoColumns(11).Value) = False Then
                temp_bz.size7 = Rs.rdoColumns(11).Value
            Else
                temp_bz.size7 = ""
            End If

            If IsDBNull(Rs.rdoColumns(12).Value) = False Then
                temp_bz.size8 = Rs.rdoColumns(12).Value
            Else
                temp_bz.size8 = ""
            End If

            If IsDBNull(Rs.rdoColumns(13).Value) = False Then
                temp_bz.size_code = Rs.rdoColumns(13).Value
            Else
                temp_bz.size_code = ""
            End If

            If IsDBNull(Rs.rdoColumns(14).Value) = False Then
                temp_bz.kikaku = Rs.rdoColumns(14).Value
            Else
                temp_bz.kikaku = ""
            End If

            If IsDBNull(Rs.rdoColumns(15).Value) = False Then
                temp_bz.plant = Rs.rdoColumns(15).Value
            Else
                temp_bz.plant = ""
            End If

            If IsDBNull(Rs.rdoColumns(16).Value) = False Then
                temp_bz.plant_code = Rs.rdoColumns(16).Value
            Else
                temp_bz.plant_code = ""
            End If

            If IsDBNull(Rs.rdoColumns(17).Value) = False Then
                temp_bz.tos_moyou = Val(Rs.rdoColumns(17).Value)
            Else
                temp_bz.tos_moyou = 0
            End If

            If IsDBNull(Rs.rdoColumns(18).Value) = False Then
                temp_bz.side_moyou = Val(Rs.rdoColumns(18).Value)
            Else
                temp_bz.side_moyou = 0
            End If

            If IsDBNull(Rs.rdoColumns(19).Value) = False Then
                temp_bz.side_kenti = Val(Rs.rdoColumns(19).Value)
            Else
                temp_bz.side_kenti = 0
            End If

            If IsDBNull(Rs.rdoColumns(20).Value) = False Then
                temp_bz.peak_mark = Val(Rs.rdoColumns(20).Value)
            Else
                temp_bz.peak_mark = 0
            End If

            If IsDBNull(Rs.rdoColumns(21).Value) = False Then
                temp_bz.nasiji = Val(Rs.rdoColumns(21).Value)
            Else
                temp_bz.nasiji = 0
            End If

            If IsDBNull(Rs.rdoColumns(22).Value) = False Then
                temp_bz.comment = Rs.rdoColumns(22).Value
            Else
                temp_bz.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(23).Value) = False Then
                temp_bz.dep_name = Rs.rdoColumns(23).Value
            Else
                temp_bz.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(24).Value) = False Then
                temp_bz.entry_name = Rs.rdoColumns(24).Value
            Else
                temp_bz.entry_name = ""
            End If

            Rs.Close()
            ' <- watanabe add VerUP(2011)


        ElseIf open_mode = "Revision number" Then
            temp_bz.id = "AT-B"

            ' -> watanabe Edit 2007.06
            '    temp_bz.no1 = Mid$(hexdata, 6, 4)
            '    temp_bz.no2 = Mid$(hexdata, 11, 2)
            If Mid(hexdata, 10, 1) = "-" Then
                temp_bz.no1 = Mid(hexdata, 6, 4)
                temp_bz.no2 = Mid(hexdata, 11, 2)
            Else
                temp_bz.no1 = Mid(hexdata, 6, 5)
                temp_bz.no2 = Mid(hexdata, 12, 2)
            End If

            '検索キーセット
            key_code = "flag_delete = 0 AND"
            key_code = key_code & " id = 'AT-B' AND"
            key_code = key_code & " no1 = '" & temp_bz.no1 & "'"

            '検索コマンド作成
            sqlcmd = "SELECT no2, kanri_no, syurui, syubetu, pattern, size, "
            sqlcmd = sqlcmd & "size1, size2, size3, size4, size5, size6, size7, size8, "
            sqlcmd = sqlcmd & "size_code, kikaku, plant, plant_code, "
            sqlcmd = sqlcmd & "tos_moyou, side_moyou, side_kenti, peak_mark, nasiji, "
            sqlcmd = sqlcmd & "comment, dep_name, entry_name, entry_date "
            sqlcmd = sqlcmd & " FROM " & DBTableName
            sqlcmd = sqlcmd & " WHERE ( " & key_code & " )"

            'ヒット数チェック
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = -1 Then
                errflg = 1
                GoTo error_section

            ElseIf cnt > 0 Then

                '検索
                Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                Rs.MoveFirst()

                Do Until Rs.EOF

                    If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                        temp_bz.no2 = Rs.rdoColumns(0).Value
                    End If

                    If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                        t1 = Rs.rdoColumns(1).Value
                    Else
                        t1 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                        t2 = Rs.rdoColumns(2).Value
                    Else
                        t2 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                        t3 = Rs.rdoColumns(3).Value
                    Else
                        t3 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                        t4 = Rs.rdoColumns(4).Value
                    Else
                        t4 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(5).Value) = False Then
                        t5 = Rs.rdoColumns(5).Value
                    Else
                        t5 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(6).Value) = False Then
                        t6 = Rs.rdoColumns(6).Value
                    Else
                        t6 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(7).Value) = False Then
                        t7 = Rs.rdoColumns(7).Value
                    Else
                        t7 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(8).Value) = False Then
                        t8 = Rs.rdoColumns(8).Value
                    Else
                        t8 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(9).Value) = False Then
                        t9 = Rs.rdoColumns(9).Value
                    Else
                        t9 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(10).Value) = False Then
                        t10 = Rs.rdoColumns(10).Value
                    Else
                        t10 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(11).Value) = False Then
                        t11 = Rs.rdoColumns(11).Value
                    Else
                        t11 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(12).Value) = False Then
                        t12 = Rs.rdoColumns(12).Value
                    Else
                        t12 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(13).Value) = False Then
                        t13 = Rs.rdoColumns(13).Value
                    Else
                        t13 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(14).Value) = False Then
                        t14 = Rs.rdoColumns(14).Value
                    Else
                        t14 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(15).Value) = False Then
                        t15 = Rs.rdoColumns(15).Value
                    Else
                        t15 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(16).Value) = False Then
                        t16 = Rs.rdoColumns(16).Value
                    Else
                        t16 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(17).Value) = False Then
                        t17 = Rs.rdoColumns(17).Value
                    Else
                        t17 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(18).Value) = False Then
                        t18 = CShort(Rs.rdoColumns(18).Value)
                    Else
                        t18 = 0
                    End If

                    If IsDBNull(Rs.rdoColumns(19).Value) = False Then
                        t19 = CShort(Rs.rdoColumns(19).Value)
                    Else
                        t19 = 0
                    End If

                    If IsDBNull(Rs.rdoColumns(20).Value) = False Then
                        t20 = CShort(Rs.rdoColumns(20).Value)
                    Else
                        t20 = 0
                    End If

                    If IsDBNull(Rs.rdoColumns(21).Value) = False Then
                        t21 = CShort(Rs.rdoColumns(21).Value)
                    Else
                        t21 = 0
                    End If

                    If IsDBNull(Rs.rdoColumns(22).Value) = False Then
                        t22 = CShort(Rs.rdoColumns(22).Value)
                    Else
                        t22 = 0
                    End If

                    If IsDBNull(Rs.rdoColumns(23).Value) = False Then
                        t23 = Rs.rdoColumns(23).Value
                    Else
                        t23 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(24).Value) = False Then
                        t24 = Rs.rdoColumns(24).Value
                    Else
                        t24 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(25).Value) = False Then
                        t25 = Rs.rdoColumns(25).Value
                    Else
                        t25 = ""
                    End If

                    If IsDBNull(Rs.rdoColumns(26).Value) = False Then
                        t26 = Rs.rdoColumns(26).Value
                    Else
                        t26 = ""
                    End If

                    temp_bz.kanri_no = t1
                    temp_bz.syurui = t2
                    temp_bz.syubetu = t3
                    temp_bz.pattern = t4
                    temp_bz.Size = t5
                    temp_bz.size1 = t6
                    temp_bz.size2 = t7
                    temp_bz.size3 = t8
                    temp_bz.size4 = t9
                    temp_bz.size5 = t10
                    temp_bz.size6 = t11
                    temp_bz.size7 = t12
                    temp_bz.size8 = t13
                    temp_bz.size_code = t14
                    temp_bz.kikaku = t15
                    temp_bz.plant = t16
                    temp_bz.plant_code = t17
                    temp_bz.tos_moyou = t18
                    temp_bz.side_moyou = t19
                    temp_bz.side_kenti = t20
                    temp_bz.peak_mark = t21
                    temp_bz.nasiji = t22
                    temp_bz.comment = t23
                    temp_bz.dep_name = t24
                    temp_bz.entry_name = t25

                    Rs.MoveNext()
                Loop

                Rs.Close()
            End If

            If open_mode = "Revision number" Then
                '----- .NET 移行 -----
                'temp_bz.no2 = VB6.Format(Val(temp_bz.no2) + 1, "00")

                temp_bz.no2 = (Val(temp_bz.no2) + 1).ToString("00")
            End If
            ' <- watanabe edit VerUP(2011)


        End If

        Call end_sql()
        zumen_no_set_bz = True
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        If errflg = 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "System error")
        End If

        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        Call end_sql()
        zumen_no_set_bz = FAIL
    End Function
End Module