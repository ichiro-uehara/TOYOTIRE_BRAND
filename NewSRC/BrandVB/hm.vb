Option Strict Off
Option Explicit On

Imports System.Collections.Generic

Module MJ_HM
	
    Function hm_read(ByRef hm_code As String) As Short '20100706 引数をObj->Strに修正
        Dim error_no As String '20100706 修正
        Dim time_now As Object
        Dim time_start As Object
        Dim w_ret As Object
        Dim pic_no As Integer
        Dim result As Integer
        Dim ZumenName As String
        Dim no As String
        Dim font_name As String
        Dim w_mess As String

        ' -> watanabe add VerUP(2011)
        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET 移行(一旦コメント化) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""
        ' <- watanabe add VerUP(2011)

        If FreePicNum < 1 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ピクチャ数が足りません" & Chr(13) & "空きピクチャ数 =" & FreePicNum)
            ErrMsg = "The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum
            ErrTtl = "Editing characters read"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        font_name = Left(hm_code, 6)
        no = Mid(hm_code, 7, 2)

        '図面名
        ZumenName = "HM-" & font_name


        ' -> watanabe edit VerUP(2011)
        ''ﾋﾟｸﾁｬ番号
        'result = SqlCmd(SqlConn, "SELECT haiti_pic")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( flag_delete = 0 AND")
        'result = SqlCmd(SqlConn, " font_name = '" & font_name & "' AND")
        'result = SqlCmd(SqlConn, " no = '" & no & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '    'Retrieve and print the result rows.
        '    If SqlNextRow(SqlConn) = REGROW Then
        '        pic_no = Val(SqlData(SqlConn, 1))
        '    Else
        '        MsgBox("指定された編集文字が見つかりません" & Chr(13) & hm_code, MsgBoxStyle.Critical, "data not found")
        '        GoTo error_section
        '    End If
        '
        'Else
        '    MsgBox("指定された編集文字が見つかりません" & Chr(13) & hm_code, MsgBoxStyle.Critical, "data not found")
        '    GoTo error_section
        'End If


        '検索キーセット
        key_code = "flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & font_name & "' AND"
        key_code = key_code & " no = '" & no & "' "

        '検索コマンド作成
        sqlcmd = "SELECT haiti_pic FROM " & DBTableName & " WHERE ( " & key_code & ")"

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            ErrMsg = "Editing characters specified was not found." & Chr(13) & hm_code
            ErrTtl = "Editing characters read"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Editing characters read"
            GoTo error_section
        End If

        '検索
        '----- .NET 移行(一旦コメント化) -----
        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'Rs.MoveFirst()

        'ﾋﾟｸﾁｬ番号
        '----- .NET 移行(一旦コメント化) -----
        'If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '    pic_no = Val(Rs.rdoColumns(0).Value)
        'Else
        '    pic_no = 0
        'End If

        'Rs.Close()
        ' <- watanabe edit VerUP(2011)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(pic_no, "000") & HensyuDir & ZumenName
        w_mess = pic_no.ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("ACADREAD", w_mess)
        w_ret = RequestACAD("ACADREAD")

        time_start = Now
        Do
            time_now = Now
            If Trim(form_main.Text2.Text) = "" Then
                If time_now - time_start > timeOutSecond Then
                    ' -> watanabe edit VerUP(2011)
                    'MsgBox("タイムアウトエラー", 64, "ERROR")
                    ErrMsg = "Time-out error."
                    ErrTtl = "ERROR"
                    ' <- watanabe edit VerUP(2011)
                    w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                    w_ret = RequestACAD("ERROR")
                    GoTo error_section
                End If

            ElseIf Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then
                MsgBox("CAD reading end.")
                FreePicNum = FreePicNum - 1
                GoTo LOOP_EXIT

            ElseIf Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                ' -> watanabe edit VerUP(2011)
                'MsgBox("ＣＡＤ読込みに失敗しました", MsgBoxStyle.Critical, "CAD読込みｴﾗｰ")
                ErrMsg = "Failed to read CAD."
                ErrTtl = "CAD reading error"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section

            Else
                ' -> watanabe edit VerUP(2011)
                'MsgBox("ﾘﾀｰﾝｺｰﾄﾞが不正です" & Chr(13) & Trim(form_main.Text2.Text), 64, "ACAD戻り値ｴﾗｰ")
                ErrMsg = "Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text)
                ErrTtl = "Error of the return value of the ACAD"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
            End If
        Loop

LOOP_EXIT:

        hm_read = True
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
        '----- .NET 移行(一旦コメント化) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        hm_read = FAIL
    End Function

	Function hm_insert() As Short
        Dim j As Integer '20100706 型修正
        Dim result As Integer
        Dim i As Integer
		Dim now_time As Object
		Dim pic_no As Object
        Dim w_str(150) As String

        Dim kubun_no As Short

        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String

        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""

        If SqlConn = 0 Then
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '------------ hm_kanri1 テーブル 登録 -------------

        '----- .NET 移行(文字列の「' '」を削除) -----

        w_str(1) = "0" '削除フラグ
        w_str(2) = "HM" 'ＩＤ(HM固定)
        w_str(3) = Trim(form_no.w_font_name.Text) 'ﾌｫﾝﾄ名(HE****)
        w_str(4) = Left(form_no.w_no.Text, 2) '区分番号（00～99の自動連番）
        w_str(5) = Trim(form_no.w_spell.Text) 'スペル
        w_str(6) = CStr(temp_hm.haiti_sitei) '配置方法
		w_str(7) = Trim(form_no.w_gm_num.Text) '原始文字数
		w_str(8) = Trim(form_no.w_width.Text) '幅
		w_str(9) = Trim(form_no.w_high.Text) '高さ
		w_str(10) = Trim(form_no.w_ang.Text) '角度

        '区分番号の自動連番処理
		kubun_no = what_no_HM(form_no.w_font_name.Text)
        If kubun_no = -1 Then
            ErrMsg = "Failed to auto sequence number of the Category number."
            ErrTtl = "Editing characters registration"
            GoTo error_section
        End If

        '----- .NET 移行 -----
        'form_no.w_no.Text = VB6.Format(kubun_no, "00")
        form_no.w_no.Text = kubun_no.ToString("00")

        w_str(4) = Left(form_no.w_no.Text, 2)

        pic_no = what_pic_no("HM", form_no.w_font_name.Text)
		If pic_no = -1 Then
            ErrMsg = "Could not picture number set." & Chr(13) & "Please change the character name registration."
            ErrTtl = "Editing characters registration"
            GoTo error_section
        End If

        '----- .NET 移行 -----
        'form_no.w_haiti_pic.Text = VB6.Format(pic_no, "000")
        form_no.w_haiti_pic.Text = pic_no.ToString("000")

        w_str(11) = form_no.w_haiti_pic.Text '配置PIC
        w_str(12) = "  " '刻印図面ID(w_gz_id)
        w_str(13) = "    " '刻印図面番号(w_gz_no1)
        w_str(14) = "  " '刻印図面変番(w_gz_no2)
        w_str(15) = Trim(form_no.w_comment.Text) 'コメント
        w_str(16) = Trim(form_no.w_dep_name.Text) '部署コード
        w_str(17) = Trim(form_no.w_entry_name.Text) '登録者

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

        '----- .NET 移行 -----
        'w_str(18) = Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) '登録日
        w_str(18) = Left(form_no.w_entry_date.Text, 4) & "-" & Mid(form_no.w_entry_date.Text, 5, 2) & "-" & Mid(form_no.w_entry_date.Text, 7, 2) & " " &
                    Trim(now_time)

        '----- .NET 移行 -----

        'sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        'For i = 1 To 17
        '    sqlcmd = sqlcmd & Trim(w_str(i)) & ","
        'Next i
        'sqlcmd = sqlcmd & Trim(w_str(18)) & ")"

        ''ｺﾏﾝﾄﾞ実行
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If

        '--------------------------------------------------

        '登録パラメータ作成
        Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
        Dim param As ADO_PARAM_Struct

        param.DataSize = 0
        param.Sign = ""

        param.ColumnName = "flag_delete"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(1)
        paramList.Add(param)

        param.ColumnName = "id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(2)
        paramList.Add(param)

        param.ColumnName = "font_name"
        param.Value = w_str(3)
        paramList.Add(param)

        param.ColumnName = "no"
        param.Value = w_str(4)
        paramList.Add(param)

        param.ColumnName = "spell"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(5)
        param.DataSize = 255
        paramList.Add(param)

        param.ColumnName = "haiti_sitei"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(6)
        param.DataSize = 0
        paramList.Add(param)

        param.ColumnName = "gm_num"
        param.SqlDbType = SqlDbType.SmallInt
        param.Value = w_str(7)
        paramList.Add(param)

        param.ColumnName = "width"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(8)
        paramList.Add(param)

        param.ColumnName = "high"
        param.Value = w_str(9)
        paramList.Add(param)

        param.ColumnName = "ang"
        param.Value = w_str(10)
        paramList.Add(param)

        param.ColumnName = "haiti_pic"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(11)
        paramList.Add(param)

        param.ColumnName = "hz_id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(12)
        paramList.Add(param)

        param.ColumnName = "hz_no1"
        param.Value = w_str(13)
        paramList.Add(param)

        param.ColumnName = "hz_no2"
        param.Value = w_str(14)
        paramList.Add(param)

        param.ColumnName = "comment"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(15)
        param.DataSize = 255
        paramList.Add(param)

        param.ColumnName = "dep_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(16)
        param.DataSize = 0
        paramList.Add(param)

        param.ColumnName = "entry_name"
        param.Value = w_str(17)
        paramList.Add(param)

        param.ColumnName = "entry_date"
        param.SqlDbType = SqlDbType.SmallDateTime
        param.Value = w_str(18)
        paramList.Add(param)

        If VBADO_Insert(GL_T_ADO, DBTableName, paramList) <> 1 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '----- .NET 移行 -----


        ' エラーが発生した時用のメッセージをクリア
        ErrMsg = ""
        ErrTtl = ""

        '----- .NET 移行 -----
        ''DB接続終了
        'end_sql()

        '------------ hm_kanri2 テーブル 登録 -------------

        For i = 1 To Val(Trim(form_no.w_gm_num.Text))

            '----- .NET 移行 -----
            ''DB接続開始
            'init_sql()

            w_str(1) = "HM" 'ＩＤ(HM固定)
            w_str(2) = Trim(form_no.w_font_name.Text) 'ﾌｫﾝﾄ名(HE****)
            w_str(3) = Left(form_no.w_no.Text, 2) '区分番号（00～99の自動連番）
            w_str(4) = i.ToString() '原始文字番号
            w_str(5) = Trim(temp_hm.gm_name(i)) '原始文字コード

            '----- .NET 移行 -----
            'sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            'For j = 1 To 4
            '    sqlcmd = sqlcmd & Trim(w_str(j)) & ","
            'Next j
            'sqlcmd = sqlcmd & Trim(w_str(5)) & ")"

            ''ｺﾏﾝﾄﾞ実行
            'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            'If GL_T_RDO.Con.RowsAffected() = 0 Then
            '    ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
            '    ErrTtl = "SQL error"
            '    GoTo error_section
            'End If

            '' エラーが発生した時用のメッセージをクリア
            'ErrMsg = ""
            'ErrTtl = ""

            ''DB接続終了
            'end_sql()
            '---------------------------------------------
            paramList.Clear()

            param.ColumnName = "id"
            param.SqlDbType = SqlDbType.Char
            param.Value = w_str(1)
            paramList.Add(param)

            param.ColumnName = "font_name"
            param.Value = w_str(2)
            paramList.Add(param)

            param.ColumnName = "no"
            param.Value = w_str(3)
            paramList.Add(param)

            param.ColumnName = "gm_no"
            param.SqlDbType = SqlDbType.SmallInt
            param.Value = w_str(4)
            paramList.Add(param)

            param.ColumnName = "gm_name"
            param.SqlDbType = SqlDbType.Char
            param.Value = w_str(5)
            paramList.Add(param)

            If VBADO_Insert(GL_T_ADO, DBTableName2, paramList) <> 1 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If

            '----- .NET 移行 -----
        Next i

        'DB接続終了
        end_sql()

        hm_insert = True
		Exit Function
		
error_section:
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If

        On Error Resume Next
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)
        Err.Clear()

        hm_insert = FAIL

    End Function

    Function hm_update() As Short
        Dim result As Integer '20100706 修正
        Dim i As Integer '20100706 修正
        Dim now_time As Object
        Dim w_str(150) As String

        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String

        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""

        'MsgBox "編集文字データをUPDATEします"

        If SqlConn = 0 Then
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '----- .NET 移行 -----

        'w_str(1) = "flag_delete = 0" '削除フラグ
        'w_str(2) = "id = '" & "HM" & "'" 'ＩＤ(HM固定)
        'w_str(3) = "font_name = '" & Trim(form_no.w_font_name.Text) & "'" 'ﾌｫﾝﾄ名(HE****)
        'w_str(4) = "no ='" & Left(form_no.w_no.Text, 2) & "'" '区分番号（00～99の自動連番）
        'w_str(5) = "spell ='" & form_no.w_spell.Text & "'" 'スペル
        'w_str(6) = "haiti_sitei=" & Str(temp_hm.haiti_sitei) '配置方法
        'w_str(7) = "gm_num =" & form_no.w_gm_num.Text '原始文字数
        'w_str(8) = "width =" & form_no.w_width.Text '幅
        'w_str(9) = "high =" & form_no.w_high.Text '高さ
        'w_str(10) = "ang =" & form_no.w_ang.Text '角度
        'w_str(11) = "haiti_pic =" & form_no.w_haiti_pic.Text '配置PIC
        'w_str(12) = "hz_id ='" & "      " & "'" '刻印図面ID(w_gz_id)
        'w_str(13) = "hz_no1 ='" & "  " & "'" '刻印図面番号(w_gz_no1)
        'w_str(14) = "hz_no2 ='" & "    " & "'" '刻印図面変番(w_gz_no2)
        'w_str(15) = "comment ='" & form_no.w_comment.Text & "'" 'コメント
        'w_str(16) = "dep_name ='" & form_no.w_dep_name.Text & "'" '部署コード
        'w_str(17) = "entry_name ='" & form_no.w_entry_name.Text & "'" '登録者

        'If Len(Hour(TimeOfDay)) = 1 Then
        '    now_time = "0" & Hour(TimeOfDay)
        'Else
        '    now_time = Hour(TimeOfDay)
        'End If

        'If Len(Minute(TimeOfDay)) = 1 Then
        '    now_time = Trim(now_time) & ":0" & Minute(TimeOfDay)
        'Else
        '    now_time = Trim(now_time) & ":" & Minute(TimeOfDay)
        'End If

        'w_str(18) = "entry_date ='" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '登録日


        'sqlcmd = "UPDATE " & DBTableName & " SET "
        'For i = 5 To 17
        '    sqlcmd = sqlcmd & Trim(w_str(i)) & " , "
        'Next i
        'sqlcmd = sqlcmd & Trim(w_str(18))
        'sqlcmd = sqlcmd & " WHERE ( " & w_str(1) & " AND "
        'sqlcmd = sqlcmd & w_str(2) & " AND "
        'sqlcmd = sqlcmd & w_str(3) & " AND "
        'sqlcmd = sqlcmd & w_str(4) & " )"

        'ｺﾏﾝﾄﾞ実行
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If

        ''DB接続終了
        'end_sql()


        ''DB接続開始
        'init_sql()

        ''現データ削除
        'sqlcmd = "DELETE FROM " & DBTableName2 & " WHERE ( "
        'sqlcmd = sqlcmd & "id = 'HM' AND "
        'sqlcmd = sqlcmd & "font_name = '" & Trim(form_no.w_font_name.Text) & "' AND "
        'sqlcmd = sqlcmd & "no ='" & Left(form_no.w_no.Text, 2) & "' )"

        'ｺﾏﾝﾄﾞ実行
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "It is not possible to delete the current data from the database.(" & DBTableName2 & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If

        ''DB接続開始
        'end_sql()


        ''新規登録
        'For i = 1 To Val(Trim(form_no.w_gm_num.Text))

        '    'DB接続開始
        '    init_sql()

        '    w_str(1) = "'" & "HM" & "'" 'ＩＤ(HM固定)
        '    w_str(2) = "'" & Trim(form_no.w_font_name.Text) & "'" 'ﾌｫﾝﾄ名(HE****)
        '    w_str(3) = "'" & Left(form_no.w_no.Text, 2) & "'" '区分番号（00～99の自動連番）
        '    w_str(4) = i '原始文字番号
        '    w_str(5) = "'" & Trim(temp_hm.gm_name(i)) & "'" '原始文字コード

        '    sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
        '    sqlcmd = sqlcmd & w_str(1) & ", "
        '    sqlcmd = sqlcmd & w_str(2) & ", "
        '    sqlcmd = sqlcmd & w_str(3) & ", "
        '    sqlcmd = sqlcmd & w_str(4) & ", "
        '    sqlcmd = sqlcmd & w_str(5) & " )"

        '    'ｺﾏﾝﾄﾞ実行
        '    'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        '    'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    '    ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
        '    '    ErrTtl = "SQL error"
        '    '    GoTo error_section
        '    'End If

        '    'DB接続開始
        '    end_sql()
        'Next i

        '-----------------------------------------------------------------------

        '------------ hm_kanri1 テーブル 更新 -------------

        w_str(1) = "0" '削除フラグ
        w_str(2) = "HM" 'ＩＤ(HM固定)
        w_str(3) = Trim(form_no.w_font_name.Text) 'ﾌｫﾝﾄ名(HE****)
        w_str(4) = Left(form_no.w_no.Text, 2) '区分番号（00～99の自動連番）
        w_str(5) = form_no.w_spell.Text 'スペル
        w_str(6) = Str(temp_hm.haiti_sitei) '配置方法
        w_str(7) = form_no.w_gm_num.Text '原始文字数
        w_str(8) = form_no.w_width.Text '幅
        w_str(9) = form_no.w_high.Text '高さ
        w_str(10) = form_no.w_ang.Text '角度
        w_str(11) = form_no.w_haiti_pic.Text '配置PIC
        w_str(12) = "      " '刻印図面ID(w_gz_id)
        w_str(13) = "  " '刻印図面番号(w_gz_no1)
        w_str(14) = "    " '刻印図面変番(w_gz_no2)
        w_str(15) = form_no.w_comment.Text 'コメント
        w_str(16) = form_no.w_dep_name.Text '部署コード
        w_str(17) = form_no.w_entry_name.Text '登録者

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

        w_str(18) = Left(form_no.w_entry_date.Text, 4) & "-" & Mid(form_no.w_entry_date.Text, 5, 2) & "-" & Mid(form_no.w_entry_date.Text, 7, 2) & " " &
                    Trim(now_time)

        '登録パラメータ作成
        Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
        Dim param As ADO_PARAM_Struct

        param.DataSize = 0
        param.Sign = ""

        param.ColumnName = "flag_delete"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(1)
        param.Sign = "="
        paramList.Add(param)

        param.ColumnName = "id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(2)
        paramList.Add(param)

        param.ColumnName = "font_name"
        param.Value = w_str(3)
        paramList.Add(param)

        param.ColumnName = "no"
        param.Value = w_str(4)
        paramList.Add(param)

        param.ColumnName = "spell"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(5)
        param.DataSize = 255
        param.Sign = ""
        paramList.Add(param)

        param.ColumnName = "haiti_sitei"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(6)
        param.DataSize = 0
        paramList.Add(param)

        param.ColumnName = "gm_num"
        param.SqlDbType = SqlDbType.SmallInt
        param.Value = w_str(7)
        paramList.Add(param)

        param.ColumnName = "width"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(8)
        paramList.Add(param)

        param.ColumnName = "high"
        param.Value = w_str(9)
        paramList.Add(param)

        param.ColumnName = "ang"
        param.Value = w_str(10)
        paramList.Add(param)

        param.ColumnName = "haiti_pic"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(11)
        paramList.Add(param)

        param.ColumnName = "hz_id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(12)
        paramList.Add(param)

        param.ColumnName = "hz_no1"
        param.Value = w_str(13)
        paramList.Add(param)

        param.ColumnName = "hz_no2"
        param.Value = w_str(14)
        paramList.Add(param)

        param.ColumnName = "comment"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(15)
        param.DataSize = 255
        paramList.Add(param)

        param.ColumnName = "dep_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(16)
        param.DataSize = 0
        paramList.Add(param)

        param.ColumnName = "entry_name"
        param.Value = w_str(17)
        paramList.Add(param)

        param.ColumnName = "entry_date"
        param.SqlDbType = SqlDbType.SmallDateTime
        param.Value = w_str(18)
        paramList.Add(param)

        If VBADO_Update(GL_T_ADO, DBTableName, paramList) <> 1 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '------------ hm_kanri2 テーブル 更新 -------------

        '現データ削除
        paramList.Clear()

        param.DataSize = 0

        param.ColumnName = "id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(2)
        param.Sign = "="
        paramList.Add(param)

        param.ColumnName = "font_name"
        param.Value = w_str(3)
        paramList.Add(param)

        param.ColumnName = "no"
        param.Value = w_str(4)
        paramList.Add(param)

        If VBADO_Delete(GL_T_ADO, DBTableName2, paramList) <> 1 Then
            ErrMsg = "It is not possible to delete the current data from the database.(" & DBTableName2 & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '新規登録

        param.DataSize = 0
        param.Sign = ""

        For i = 1 To Val(Trim(form_no.w_gm_num.Text))

            w_str(1) = "HM" 'ＩＤ(HM固定)
            w_str(2) = Trim(form_no.w_font_name.Text) 'ﾌｫﾝﾄ名(HE****)
            w_str(3) = Left(form_no.w_no.Text, 2) '区分番号（00～99の自動連番）
            w_str(4) = i.ToString() '原始文字番号
            w_str(5) = Trim(temp_hm.gm_name(i)) '原始文字コード

            paramList.Clear()

            param.ColumnName = "id"
            param.SqlDbType = SqlDbType.Char
            param.Value = w_str(1)
            paramList.Add(param)

            param.ColumnName = "font_name"
            param.Value = w_str(2)
            paramList.Add(param)

            param.ColumnName = "no"
            param.Value = w_str(3)
            paramList.Add(param)

            param.ColumnName = "gm_no"
            param.SqlDbType = SqlDbType.SmallInt
            param.Value = w_str(4)
            paramList.Add(param)

            param.ColumnName = "gm_name"
            param.SqlDbType = SqlDbType.Char
            param.Value = w_str(5)
            paramList.Add(param)

            If VBADO_Insert(GL_T_ADO, DBTableName2, paramList) <> 1 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If

        Next

        'DB接続終了
        end_sql()

        '----- .NET 移行 -----

        hm_update = True
        Exit Function

error_section:

        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()

        hm_update = FAIL

    End Function

	Function hm_search(ByRef hm_code As String) As Short
        Dim i As Integer '20100706 修正
		Dim w_ret As Object
        Dim result As Integer '20100706 修正
        Dim w_str(14) As String
        Dim ww As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        'Dim df As DateInfo
        ' <- watanabe del VerUP(2011)

        ' -> watanabe add VerUP(2011)
        Dim errflg As Integer
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET 移行(一旦コメント化) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        errflg = 0
        ' <- watanabe add VerUP(2011)

		If SqlConn = 0 Then
            MsgBox("Can not access the database.", MsgBoxStyle.Critical, "SQL error")
            ' -> watanabe add VerUP(2011)
            errflg = 1
            ' <- watanabe add VerUP(2011)
            GoTo error_section
		End If
		
        'HM_KANRIテーブルより該当する編集文字データを求める
		temp_hm.font_name = Mid(hm_code, 1, 6)
		temp_hm.no = Mid(hm_code, 7, 2)


        '検索キーセット
        key_code = "flag_delete = 0 AND font_name = '" & temp_hm.font_name & "' AND"
        key_code = key_code & " no = '" & temp_hm.no & "' "

        '検索コマンド作成
        sqlcmd = "SELECT id, spell, haiti_sitei, gm_num, width, high, ang,"
        sqlcmd = sqlcmd & " haiti_pic, hz_id, hz_no1, hz_no2,"
        sqlcmd = sqlcmd & " comment, dep_name, entry_name, entry_date"
        sqlcmd = sqlcmd & " FROM " & DBTableName
        sqlcmd = sqlcmd & " WHERE ( " & key_code & ")"

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            errflg = 1
            GoTo error_section
        ElseIf cnt = -1 Then
            errflg = 1
            GoTo error_section
        End If

        '検索
        '----- .NET 移行(一旦コメント化) -----
        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'Rs.MoveFirst()

        'If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '    temp_hm.id = Rs.rdoColumns(0).Value
        'Else
        '    temp_hm.id = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(1).Value) = False Then
        '    temp_hm.spell = Rs.rdoColumns(1).Value
        'Else
        '    temp_hm.spell = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(2).Value) = False Then
        '    temp_hm.haiti_sitei = Val(Rs.rdoColumns(2).Value)
        'Else
        '    temp_hm.haiti_sitei = 0
        'End If

        'If IsDBNull(Rs.rdoColumns(3).Value) = False Then
        '    temp_hm.gm_num = Val(Rs.rdoColumns(3).Value)
        'Else
        '    temp_hm.gm_num = 0
        'End If

        'If IsDBNull(Rs.rdoColumns(4).Value) = False Then
        '    temp_hm.width = Val(Rs.rdoColumns(4).Value)
        'Else
        '    temp_hm.width = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(5).Value) = False Then
        '    temp_hm.high = Val(Rs.rdoColumns(5).Value)
        'Else
        '    temp_hm.high = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(6).Value) = False Then
        '    temp_hm.ang = Val(Rs.rdoColumns(6).Value)
        'Else
        '    temp_hm.ang = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(7).Value) = False Then
        '    temp_hm.haiti_pic = Val(Rs.rdoColumns(7).Value)
        'Else
        '    temp_hm.haiti_pic = 0
        'End If

        'If IsDBNull(Rs.rdoColumns(8).Value) = False Then
        '    temp_hm.hz_id = Rs.rdoColumns(8).Value
        'Else
        '    temp_hm.hz_id = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(9).Value) = False Then
        '    temp_hm.hz_no1 = Rs.rdoColumns(9).Value
        'Else
        '    temp_hm.hz_no1 = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(10).Value) = False Then
        '    temp_hm.hz_no2 = Rs.rdoColumns(10).Value
        'Else
        '    temp_hm.hz_no2 = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(11).Value) = False Then
        '    temp_hm.comment = Rs.rdoColumns(11).Value
        'Else
        '    temp_hm.comment = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(12).Value) = False Then
        '    temp_hm.dep_name = Rs.rdoColumns(12).Value
        'Else
        '    temp_hm.dep_name = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(13).Value) = False Then
        '    temp_hm.entry_name = Rs.rdoColumns(13).Value
        'Else
        '    temp_hm.entry_name = ""
        'End If

        '登録日編集
        '----- .NET 移行(一旦コメント化) -----
        'If IsDBNull(Rs.rdoColumns(14).Value) = False Then
        '    Dim tmpstr As String
        '    tmpstr = Rs.rdoColumns(14).Value
        '    temp_hm.entry_date = Left(tmpstr, 4) & Mid(tmpstr, 6, 2) & Mid(tmpstr, 9, 2)
        'Else
        '    temp_hm.entry_date = ""
        'End If

        'Rs.Close()

        'DB接続終了
        end_sql()


        For i = 1 To temp_hm.gm_num

            'DB接続開始
            init_sql()

            '検索キーセット
            key_code = "font_name = '" & temp_hm.font_name & "' AND"
            key_code = key_code & " no = '" & temp_hm.no & "' AND" & " gm_no = " & i & " "

            '検索コマンド作成
            sqlcmd = "SELECT gm_name"
            sqlcmd = sqlcmd & " FROM " & DBTableName2
            sqlcmd = sqlcmd & " WHERE ( " & key_code & ")"

            'ヒット数チェック
            '----- .NET 移行(一旦コメント化) -----
            'cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Then
                errflg = 1
                GoTo error_section
            ElseIf cnt = -1 Then
                errflg = 1
                GoTo error_section
            End If

            '検索
            '----- .NET 移行(一旦コメント化) -----
            'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            'Rs.MoveFirst()

            'If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            '    temp_hm.gm_name(i) = Rs.rdoColumns(0).Value
            'Else
            '    temp_hm.gm_name(i) = ""
            'End If

            'Rs.Close()

            'DB接続終了
            end_sql()
        Next i
        ' <- watanabe add VerUP(2011)

        hm_search = True
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        If errflg = 0 Then
            MsgBox(Err.Description, MsgBoxStyle.Critical, "System error")
        End If

        On Error Resume Next
        Err.Clear()
        '----- .NET 移行(一旦コメント化) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        hm_search = FAIL
    End Function

	Function temp_hm_set(ByRef flag As Short, ByRef hexdata As String) As Short

        Dim de As Object
        Dim gmnum1 As Object
		Dim spelnum As Object
        Dim i As Integer '20100706 修正
		Dim result As Object
        Dim w_ret As Integer '20100706 修正
        Dim aa As String
		Dim ss(50) As String
        Dim Wstr As String
		Dim Wname0 As String
		Dim Wname1 As String
		Dim Wname2 As String
		Dim ww As String

        Dim rq_name As String

        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer

        '----- .NET 移行(コメント化) -----
        'Dim Rs As RDO.rdoResultset

        On Error Resume Next 'エラーのトラップを留保します。
		Err.Clear()


        aa = ""
        rq_name = ""


        'MsgBox "record_no=" & flag
        '==============================
        '編集文字データをＨＥＸより変換
        '==============================
        '１レコード目
        If flag = 0 Then '編集文字特性
			temp_hm.id = Mid(hexdata, 1, 2)
			temp_hm.font_name = Mid(hexdata, 3, 6)
			temp_hm.no = Mid(hexdata, 9, 2)
			temp_hm.spell = Mid(hexdata, 11, 255)
			w_ret = HextoSht(Mid(hexdata, 266, 4), temp_hm.haiti_sitei)
			w_ret = HextoSht(Mid(hexdata, 270, 4), temp_hm.gm_num)
			w_ret = HextoDbl(Mid(hexdata, 274, 16), temp_hm.width)
			w_ret = HextoDbl(Mid(hexdata, 290, 16), temp_hm.high)
			w_ret = HextoDbl(Mid(hexdata, 306, 16), temp_hm.ang)
			w_ret = HextoSht(Mid(hexdata, 322, 4), temp_hm.haiti_pic)
			
            If open_mode = "Change" Then

                init_sql()


                '検索キーセット
                key_code = "font_name = '" & temp_hm.font_name & "' AND"
                key_code = key_code & " no = '" & temp_hm.no & "' "

                '----- .NET 移行-----

                ''検索コマンド作成
                'sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE ( " & key_code & ")"

                ''ヒット数チェック
                'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
                'If cnt = 0 Then
                '    MsgBox("Editing characters specified was not found.")

                'ElseIf cnt = -1 Then
                '    MsgBox("An error occurred on the existing record during the database search.")

                'Else
                '    '検索
                '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                '    Rs.MoveFirst()

                '    If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                '        temp_hm.comment = Rs.rdoColumns(0).Value
                '    Else
                '        temp_hm.comment = ""
                '    End If

                '    If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                '        temp_hm.dep_name = Rs.rdoColumns(1).Value
                '    Else
                '        temp_hm.dep_name = ""
                '    End If

                '    If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                '        temp_hm.entry_name = Rs.rdoColumns(2).Value
                '    Else
                '        temp_hm.entry_name = ""
                '    End If

                '    Rs.Close()
                'End If

                '--------------------------------------------------------------------------------

                temp_hm.comment = ""
                temp_hm.dep_name = ""
                temp_hm.entry_name = ""

                'テーブルレコード数チェック
                cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

                If cnt = 0 Then
                    MsgBox("Editing characters specified was not found.")

                ElseIf cnt = -1 Then
                    MsgBox("An error occurred on the existing record during the database search.")

                Else
                    '検索

                    Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
                    Dim param As ADO_PARAM_Struct

                    param.DataSize = 0
                    param.Value = Nothing
                    param.Sign = ""

                    param.ColumnName = "comment"
                    param.SqlDbType = SqlDbType.VarChar
                    paramList.Add(param)

                    param.ColumnName = "dep_name"
                    param.SqlDbType = SqlDbType.Char
                    paramList.Add(param)

                    param.ColumnName = "entry_name"
                    paramList.Add(param)

                    'Databaseレコード検索処理
                    Dim dataList As List(Of List(Of String)) = New List(Of List(Of String))
                    If VBADO_Search(GL_T_ADO, DBTableName, key_code, paramList, dataList) = 1 Then
                        temp_hm.comment = dataList(0)(0)
                        temp_hm.dep_name = dataList(0)(1)
                        temp_hm.entry_name = dataList(0)(2)
                    Else
                        MsgBox("Editing characters specified was not found.")
                    End If

                End If
                '----- .NET 移行-----

                end_sql()

            End If

			'    MsgBox "gm_ num = " & temp_hm.gm_num
			
            If open_mode = "NEW" Then
                temp_hm.font_name = ""
                'temp_hm.no = ""
                temp_hm.haiti_pic = 0
                temp_hm.comment = ""
                For i = 1 To 500
                    temp_hm.gm_name(i) = ""
                Next i
            End If
			
			Call true_date(aa)
			temp_hm.entry_date = aa
			
			'デバック
			For i = 1 To 500
				temp_hm.gm_name(i) = ""
			Next i

            '次のデータ要求
            CommunicateMode = comSpecData

            RequestACAD("SPEC2011")

            '２レコード目以降

        Else '原始文字コード
			'５０ずつ原始文字コードデータを受け取ります
			spelnum = temp_hm.gm_num
			If (temp_hm.gm_num > 255) Then spelnum = 255
			
			gmnum1 = temp_hm.gm_num

            de = 50
			If flag * 50 > gmnum1 Then
				de = gmnum1 Mod 50
			End If

            For i = 1 To de
                ss(i) = Mid(hexdata, (i * 10) - 9, 10)
                temp_hm.gm_name((flag - 1) * 50 + i) = ss(i)
            Next i

            If flag * 50 > gmnum1 Then
                '        If open_mode = "新規" Then
                Wstr = ""
                For i = 1 To gmnum1
                    Wname0 = Mid(temp_hm.gm_name(i), 7, 1)
                    Wname1 = Mid(temp_hm.gm_name(i), 9, 1)
                    Wname2 = Mid(temp_hm.gm_name(i), 10, 1)
                    If Wname1 = "A" Or Wname1 = "B" Then
                        Wstr = Wstr & Wname2
                    ElseIf Wname0 = "D" Then
                        Wstr = Wstr & "_"
                    Else
                        Wstr = Wstr & "#"
                    End If
                Next i
                If CDbl(Wstr) > 255 Then
                    Wstr = Left(Wstr, 255)
                End If
                temp_hm.spell = Wstr
                '        End If
                CommunicateMode = comNone
                dataset_F_HMSAVE()

            Else
                '次のデータ要求

                '----- .NET 移行 -----
                If flag + 1 < 10 Then
                    '----- .NET 移行 -----
                    'rq_name = "SPEC201" & Left(VB6.Format(flag + 1, "0"), 1)
                    rq_name = "SPEC201" & Left((flag + 1).ToString("0"), 1)
                ElseIf flag + 1 = 10 Then
                    rq_name = "SPEC201" & "A"
                End If

                CommunicateMode = comSpecData
                RequestACAD(rq_name)
            End If
        End If
		
	End Function

    Function temp_hm_get() As Short
        Dim Msg As String '20100706 修正
        Dim w_ret As Integer '20100706 修正

        '----- .NET 移行 -----
        'Dim hexdata As New VB6.FixedLengthString(325)
        'Dim www As New VB6.FixedLengthString(16)
        Dim hexdata As String = New String(" "c, 325)
        Dim www As String = New String(" "c, 16)

        Err.Clear()
        On Error GoTo error_section


        '----- .NET 移行(コメント化) -----
        'Dim ii As Integer

        '' 必要文字数分、スペースで初期化
        'hexdata.Value = ""
        'For ii = 1 To 325
        '    hexdata.Value = hexdata.Value & " "
        'Next ii

        '' 必要文字数分、スペースで初期化
        'www.Value = ""
        'For ii = 1 To 16
        '    www.Value = www.Value & " "
        'Next ii


        '========================================
        '原始文字データをＨＥＸに変換して送信します
        '========================================

        temp_hm.id = "HM"
        temp_hm.font_name = form_no.w_font_name.Text
        temp_hm.no = form_no.w_no.Text
        temp_hm.haiti_pic = form_no.w_haiti_pic.Text
        temp_hm.spell = form_no.w_spell.Text
        If Len(Trim(temp_hm.spell)) <= 255 Then
            temp_hm.spell = Trim(temp_hm.spell) & Space(255 - Len(Trim(temp_hm.spell)))
        Else
            temp_hm.spell = Left(Trim(form_no.w_spell.Text), 255)
        End If

        temp_hm.comment = form_no.w_comment.Text
        If Len(Trim(temp_hm.comment)) <= 255 Then
            temp_hm.comment = Trim(temp_hm.comment) & Space(255 - Len(Trim(temp_hm.comment)))
        Else
            temp_hm.comment = Left(Trim(form_no.w_comment.Text), 255)
        End If

        temp_hm.dep_name = form_no.w_dep_name.Text
        temp_hm.entry_name = form_no.w_entry_name.Text
        temp_hm.entry_date = form_no.w_entry_date.Text
        temp_hm.width = form_no.w_width.Text
        temp_hm.high = form_no.w_high.Text
        temp_hm.ang = form_no.w_ang.Text
        temp_hm.gm_num = form_no.w_gm_num.Text

        '----- .NET 移行 -----
        'hexdata.Value = Space(325)
        'Mid(hexdata.Value, 1, 2) = temp_hm.id
        'Mid(hexdata.Value, 3, 6) = temp_hm.font_name
        'Mid(hexdata.Value, 9, 2) = temp_hm.no
        '--------------------------------------------
        Mid(hexdata, 1, 2) = temp_hm.id
        Mid(hexdata, 3, 6) = temp_hm.font_name
        Mid(hexdata, 9, 2) = temp_hm.no
        '----- .NET 移行 -----

        '----- .NET 移行 -----
        'w_ret = ShttoHex(temp_hm.haiti_sitei, www.Value)
        'Mid(hexdata.Value, 11, 4) = www.Value
        'w_ret = ShttoHex(temp_hm.gm_num, www.Value)
        'Mid(hexdata.Value, 15, 4) = www.Value
        'w_ret = DbltoHex(temp_hm.width, www.Value)
        'Mid(hexdata.Value, 19, 16) = www.Value
        'w_ret = DbltoHex(temp_hm.high, www.Value)
        'Mid(hexdata.Value, 35, 16) = www.Value
        'w_ret = DbltoHex(temp_hm.ang, www.Value)
        'Mid(hexdata.Value, 51, 16) = www.Value
        'w_ret = ShttoHex(temp_hm.haiti_pic, www.Value)
        'Mid(hexdata.Value, 67, 4) = www.Value

        'Mid(hexdata.Value, 71, 255) = temp_hm.spell
        '--------------------------------------------
        w_ret = ShttoHex(temp_hm.haiti_sitei, www)
        Mid(hexdata, 11, 4) = www
        w_ret = ShttoHex(temp_hm.gm_num, www)
        Mid(hexdata, 15, 4) = www
        w_ret = DbltoHex(temp_hm.width, www)
        Mid(hexdata, 19, 16) = www
        w_ret = DbltoHex(temp_hm.high, www)
        Mid(hexdata, 35, 16) = www
        w_ret = DbltoHex(temp_hm.ang, www)
        Mid(hexdata, 51, 16) = www
        w_ret = ShttoHex(temp_hm.haiti_pic, www)
        Mid(hexdata, 67, 4) = www

        Mid(hexdata, 71, 255) = temp_hm.spell
        '----- .NET 移行 -----

        '１レコード目送信
        w_ret = PokeACAD("SPECADD", hexdata)
        w_ret = RequestACAD("SPECADD")

        Exit Function

error_section:
        If Err.Number <> 0 Then
            Msg = "There was an error in the error number [" & Str(Err.Number) & ":" & Err.Source & "]." & Chr(13) & Err.Description
            MsgBox(Msg, , "error")
            Resume Next
        End If

    End Function
	
    Function hm_delete(ByRef hm_code As String) As Short
        Dim result As Integer '20100706 修正
        Dim i As Integer '20100706 修正
        Dim w_str(150) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

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

        ''''' Dim wk_spell As String
        ''''' LSet wk_spell = Trim(form_no.w_plant_code.Text)

        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        w_str(1) = " flag_delete = 1" '削除フラグ
        w_str(2) = " id = '" & "HM" & "'" 'ＩＤ(HM固定)
        w_str(3) = " font_name = '" & Mid(hm_code, 1, 6) & "'" 'ﾌｫﾝﾄ名(KO****)
        w_str(4) = " no = '" & Mid(hm_code, 7, 2) & "'" '区分番号
        w_str(5) = " spell = '" & form_no.w_spell.Text & "'" 'スペル
        ' w_str(6) = form_no.w_haiti_sitei                                             '配置方法
        w_str(6) = " haiti_sitei = " & temp_hm.haiti_sitei '配置方法
        w_str(7) = " gm_num = " & temp_hm.gm_num '原始文字数
        w_str(8) = " width = " & form_no.w_width.Text '全幅
        w_str(9) = " high = " & form_no.w_high.Text '基準高さ
        w_str(10) = " ang =" & form_no.w_ang.Text '基準角度
        w_str(11) = " haiti_pic = " & form_no.w_haiti_pic.Text '配置PIC
        w_str(12) = " hz_id = '" & temp_hm.hz_id & "'" '刻印図面ID(w_hz_id)
        w_str(13) = " hz_no1 = '" & temp_hm.hz_no1 & "'" '刻印図面番号(w_hz_no1)
        w_str(14) = " hz_no2 = '" & temp_hm.hz_no2 & "'" '刻印図面変番(w_hz_no2)
        w_str(15) = " comment = '" & form_no.w_comment.Text & "'" 'コメント
        w_str(16) = " dep_name = '" & form_no.w_dep_name.Text & "'" '部署コード
        w_str(17) = " entry_name = '" & form_no.w_entry_name.Text & "'" '登録者
        ' w_str(18) = " entry_date = '" & form_no.w_entry_date.Text & "'"              '登録日

        'Brand Ver.3 変更
        ' For i = 1 To 127
        '   w_str(i + 17) = "gm_name" & Format(i, "000") & " = '" & Trim(temp_hm.gm_name(i)) & "'"          '原始文字コード
        ' Next i


        ' -> watanabe edit VerUP(2011)
        '      w_command = "UPDATE " & DBTableName & " SET "
        '
        ''Brand Ver.3 変更
        '' For i = 1 To 126 + 17
        'For i = 1 To 16
        '	w_command = w_command & Trim(w_str(i)) & " , "
        'Next i
        ''Brand Ver.3 変更
        '' w_command = w_command & Trim(w_str(127 + 17))
        'w_command = w_command & Trim(w_str(17))
        '
        ''YAMAOKA MOD
        '
        'w_command = w_command & " WHERE ( "
        'w_command = w_command & Trim(w_str(3)) & " AND "
        'w_command = w_command & Trim(w_str(4)) & ")"
        '
        ' '' MsgBox "UPDATA=" & w_command
        '
        ''yamaoka MOd
        '
        'result = SqlCmd(SqlConn, w_command)
        ''Send the command to SQL Server and start execution.
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        ''MsgBox "UPDATE Result = " & result
        'If result <> 1 Then GoTo error_section


        sqlcmd = "UPDATE " & DBTableName & " SET "
        For i = 1 To 16
            sqlcmd = sqlcmd & Trim(w_str(i)) & " , "
        Next i
        sqlcmd = sqlcmd & Trim(w_str(17))
        sqlcmd = sqlcmd & " WHERE ( "
        sqlcmd = sqlcmd & Trim(w_str(3)) & " AND "
        sqlcmd = sqlcmd & Trim(w_str(4)) & ")"

        'ｺﾏﾝﾄﾞ実行
        '----- .NET 移行(一旦コメント化) -----
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)

        hm_delete = True
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

        hm_delete = FAIL
    End Function
End Module