Option Strict Off
Option Explicit On

Imports System.Collections.Generic

Module MJ_GM
	Function gm_insert() As Short
		Dim result As Object
		Dim now_time As Object
		Dim pic_no As Object
        Dim w_str(42) As String
        'Dim w_command As String'20100616移植削除
		Dim kubun As Short

        Dim ErrMsg As String
        Dim ErrTtl As String
        Dim sqlcmd As String

        On Error GoTo error_section
        Err.Clear()
        ErrMsg = ""
        ErrTtl = ""

        If SqlConn = 0 Then
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            GoTo error_section
        End If


        '----- .NET 移行(文字列の「' '」を削除) -----

        w_str(1) = "0" '削除フラグ
        w_str(2) = "GM" 'ＩＤ(GM固定)
        w_str(3) = form_no.w_font_name.Text 'ﾌｫﾝﾄ名(KO****)
        w_str(4) = Left(form_no.w_font_class1.Text, 1) 'ﾌｫﾝﾄ区分1(A,F,H,B,D,P,N）
        w_str(5) = Left(form_no.w_font_class2.Text, 1) 'ﾌｫﾝﾄ区分2(0～9: 自動連番）
        w_str(6) = Left(form_no.w_name1.Text, 1) '文字名1
        w_str(7) = Left(form_no.w_name2.Text, 1) '文字名2
        w_str(8) = form_no.w_high.Text '高さ
		w_str(9) = form_no.w_width.Text '幅
		w_str(10) = form_no.w_ang.Text '角度
		w_str(11) = form_no.w_moji_high.Text '実高さ
		w_str(12) = form_no.w_moji_shift.Text 'ずれ量
        Select Case form_no.w_org_hor.Text '水平原点位置(現在Cに固定) '20100706コード変更
            Case "Center"
                w_str(13) = "C"
            Case "Left end"
                w_str(13) = "L"
            Case "Right end"
                w_str(13) = "R"
            Case Else
                ' -> watanabe edit VerUP(2011)
                'MsgBox("水平原点位置エラー", MsgBoxStyle.Critical, "原始文字登録")
                ErrMsg = "Horizontal origin position error."
                ErrTtl = "Primitive character registration"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
		
        Select Case form_no.w_org_ver.Text '垂直原点位置(現在Bに固定)
            Case "Center"
                w_str(14) = "C"
            Case "Top"
                w_str(14) = "T"
            Case "Lower end"
                w_str(14) = "B"
            Case Else
                ' -> watanabe edit VerUP(2011)
                'MsgBox("垂直原点位置エラー", MsgBoxStyle.Critical, "原始文字登録")
                ErrMsg = "Vertical origin position error."
                ErrTtl = "Primitive character registration"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
		
		w_str(15) = CStr(temp_gm.org_x) '文字原点座標X
		w_str(16) = CStr(temp_gm.org_y) '文字原点座標Y
		w_str(17) = CStr(temp_gm.left_bottom_x) '枠左下座標X
		w_str(18) = CStr(temp_gm.left_bottom_y) '枠左下座標Y
		w_str(19) = CStr(temp_gm.right_bottom_x) '枠右下座標X
		w_str(20) = CStr(temp_gm.right_bottom_y) '枠右下座標Y
		w_str(21) = CStr(temp_gm.right_top_x) '枠右上座標X
		w_str(22) = CStr(temp_gm.right_top_y) '枠右上座標Y
		w_str(23) = CStr(temp_gm.left_top_x) '枠左上座標X
		w_str(24) = CStr(temp_gm.left_top_y) '枠左上座標Y
		
		w_str(25) = form_no.w_hem_width.Text '縁取り幅
		w_str(26) = form_no.w_hatch_ang.Text 'ﾊｯﾁﾝｸﾞ角度
		w_str(27) = form_no.w_hatch_width.Text 'ﾊｯﾁﾝｸﾞ幅
		w_str(28) = form_no.w_hatch_space.Text 'ﾊｯﾁﾝｸﾞ間隔
		w_str(29) = form_no.w_hatch_x.Text 'ﾊｯﾁﾝｸﾞ始点X
		w_str(30) = form_no.w_hatch_y.Text 'ﾊｯﾁﾝｸﾞ始点Y
		w_str(31) = form_no.w_base_r.Text '基準Ｒ
        w_str(32) = form_no.w_old_font_name.Text '旧ﾌｫﾝﾄ名
        w_str(33) = form_no.w_old_font_class.Text '旧ﾌｫﾝﾄ区分
        w_str(34) = form_no.w_old_name.Text '旧文字名

        pic_no = what_pic_no("GM", form_no.w_font_name.Text)

        If pic_no = -1 Then
            'MsgBox("ピクチャ番号設定できませんでした" & Chr(13) & "登録文字名を変更してください", MsgBoxStyle.Critical, "原始文字登録")
            ErrMsg = "Could not picture number set." & Chr(13) & "Please change the character name registration."
            ErrTtl = "Primitive character registration"
            GoTo error_section
        End If

        '----- .NET 移行 -----
        'form_no.w_haiti_pic.Text = VB6.Format(pic_no, "000")
        form_no.w_haiti_pic.Text = pic_no.ToString("000")

        w_str(35) = form_no.w_haiti_pic.Text '配置PIC
        w_str(36) = "  " '刻印図面ID(w_gz_id)
        w_str(37) = "    " '刻印図面番号(w_gz_no1)
        w_str(38) = "  " '刻印図面変番(w_gz_no2)
        w_str(39) = form_no.w_comment.Text 'コメント
        w_str(40) = form_no.w_dep_name.Text '部署コード
        w_str(41) = form_no.w_entry_name.Text '登録者

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
        'w_str(42) = "'" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '登録日
        w_str(42) = Left(form_no.w_entry_date.Text, 4) & "-" & Mid(form_no.w_entry_date.Text, 5, 2) & "-" & Mid(form_no.w_entry_date.Text, 7, 2) & " " &
                    Trim(now_time)

        w_str(3) = form_no.w_font_name.Text 'ﾌｫﾝﾄ名(KO****)
        w_str(4) = Left(form_no.w_font_class1.Text, 1) 'ﾌｫﾝﾄ区分1(A,F,H,B,D,P,N）
        w_str(5) = Left(form_no.w_font_class2.Text, 1) 'ﾌｫﾝﾄ区分2(0～9: 自動連番）
        w_str(6) = Left(form_no.w_name1.Text, 1) '文字名1
        w_str(7) = Left(form_no.w_name2.Text, 1) '文字名2

        kubun = what_font_class2_GM(form_no.w_font_name.Text, Left(form_no.w_font_class1.Text, 1), Left(form_no.w_name1.Text, 1), Left(form_no.w_name2.Text, 1))

        If kubun < 0 Then
            ErrMsg = "It was not possible to set the font Category."
            ErrTtl = "Primitive character registration"
            GoTo error_section
        End If

        form_no.w_font_class2.Text = kubun
        w_str(5) = form_no.w_font_class2.Text

        '----- .NET 移行(文字列の「' '」を削除) -----


        '----- .NET 移行 -----

        'sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        'sqlcmd = sqlcmd & w_str(1) & "," & w_str(2) & "," & w_str(3) & "," & w_str(4) & "," & w_str(5) & "," & w_str(6) & ","
        'sqlcmd = sqlcmd & w_str(7) & "," & w_str(8) & "," & w_str(9) & "," & w_str(10) & "," & w_str(11) & "," & w_str(12) & ","
        'sqlcmd = sqlcmd & w_str(13) & "," & w_str(14) & "," & w_str(15) & "," & w_str(16) & "," & w_str(17) & "," & w_str(18) & ","
        'sqlcmd = sqlcmd & w_str(19) & "," & w_str(20) & "," & w_str(21) & "," & w_str(22) & "," & w_str(23) & "," & w_str(24) & ","
        'sqlcmd = sqlcmd & w_str(25) & "," & w_str(26) & "," & w_str(27) & "," & w_str(28) & "," & w_str(29) & "," & w_str(30) & ","
        'sqlcmd = sqlcmd & w_str(31) & "," & w_str(32) & "," & w_str(33) & "," & w_str(34) & "," & w_str(35) & "," & w_str(36) & ","
        'sqlcmd = sqlcmd & w_str(37) & "," & w_str(38) & "," & w_str(39) & "," & w_str(40) & "," & w_str(41) & "," & w_str(42)
        'sqlcmd = sqlcmd & ")"


        ''ｺﾏﾝﾄﾞ実行
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If

        '-------------------------------------------------------

        '登録パラメータ作成
        Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
        Dim param As ADO_PARAM_Struct

        '----- .NET 移行（文字列の「' '」を削除）---------
        For i As Integer = 1 To 42
            w_str(i) = w_str(i).Trim("'"c)
        Next

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

        param.ColumnName = "font_class1"
        param.Value = w_str(4)
        paramList.Add(param)

        param.ColumnName = "font_class2"
        param.Value = w_str(5)
        paramList.Add(param)

        param.ColumnName = "name1"
        param.Value = w_str(6)
        paramList.Add(param)

        param.ColumnName = "name2"
        param.Value = w_str(7)
        paramList.Add(param)

        param.ColumnName = "high"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(8)
        paramList.Add(param)

        param.ColumnName = "width"
        param.Value = w_str(9)
        paramList.Add(param)

        param.ColumnName = "ang"
        param.Value = w_str(10)
        paramList.Add(param)

        param.ColumnName = "moji_high"
        param.Value = w_str(11)
        paramList.Add(param)

        param.ColumnName = "moji_shift"
        param.Value = w_str(12)
        paramList.Add(param)

        param.ColumnName = "org_hor"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(13)
        paramList.Add(param)

        param.ColumnName = "org_ver"
        param.Value = w_str(14)
        paramList.Add(param)

        param.ColumnName = "org_x"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(15)
        paramList.Add(param)

        param.ColumnName = "org_y"
        param.Value = w_str(16)
        paramList.Add(param)

        param.ColumnName = "left_bottom_x"
        param.Value = w_str(17)
        paramList.Add(param)

        param.ColumnName = "left_bottom_y"
        param.Value = w_str(18)
        paramList.Add(param)

        param.ColumnName = "right_bottom_x"
        param.Value = w_str(19)
        paramList.Add(param)

        param.ColumnName = "right_bottom_y"
        param.Value = w_str(20)
        paramList.Add(param)

        param.ColumnName = "right_top_x"
        param.Value = w_str(21)
        paramList.Add(param)

        param.ColumnName = "right_top_y"
        param.Value = w_str(22)
        paramList.Add(param)

        param.ColumnName = "left_top_x"
        param.Value = w_str(23)
        paramList.Add(param)

        param.ColumnName = "left_top_y"
        param.Value = w_str(24)
        paramList.Add(param)

        param.ColumnName = "hem_width"
        param.Value = w_str(25)
        paramList.Add(param)

        param.ColumnName = "hatch_ang"
        param.Value = w_str(26)
        paramList.Add(param)

        param.ColumnName = "hatch_width"
        param.Value = w_str(27)
        paramList.Add(param)

        param.ColumnName = "hatch_space"
        param.Value = w_str(28)
        paramList.Add(param)

        param.ColumnName = "hatch_x"
        param.Value = w_str(29)
        paramList.Add(param)

        param.ColumnName = "hatch_y"
        param.Value = w_str(30)
        paramList.Add(param)

        param.ColumnName = "base_r"
        param.Value = w_str(31)
        paramList.Add(param)

        param.ColumnName = "old_font_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(32)
        paramList.Add(param)

        param.ColumnName = "old_font_class"
        param.Value = w_str(33)
        paramList.Add(param)

        param.ColumnName = "old_name"
        param.Value = w_str(34)
        paramList.Add(param)

        param.ColumnName = "haiti_pic"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(35)
        paramList.Add(param)

        param.ColumnName = "gz_id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(36)
        paramList.Add(param)

        param.ColumnName = "gz_no1"
        param.Value = w_str(37)
        paramList.Add(param)

        param.ColumnName = "gz_no2"
        param.Value = w_str(38)
        paramList.Add(param)

        param.ColumnName = "comment"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(39)
        param.DataSize = 255
        paramList.Add(param)

        param.ColumnName = "dep_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(40)
        paramList.Add(param)

        param.ColumnName = "entry_name"
        param.Value = w_str(41)
        paramList.Add(param)

        param.ColumnName = "entry_date"
        param.SqlDbType = SqlDbType.SmallDateTime
        param.Value = w_str(42)
        paramList.Add(param)

        If VBADO_Insert(GL_T_ADO, DBTableName, paramList) <> 1 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
        End If

        '----- .NET 移行 -----

        gm_insert = True
		
		Exit Function
		
error_section:
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
            GoTo error_section
        End If

        On Error Resume Next
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)
        Err.Clear()

        gm_insert = FAIL

    End Function


    Function gm_read(ByRef gm_code As String) As Short '20100706 引数をObj->Strに修正
        Dim error_no As Object
        Dim time_now As Object
        Dim time_start As Object
        Dim w_ret As Object
        Dim pic_no As Integer
        Dim result As Object
        '20100706 型修正
        Dim ZumenName As String
        Dim name2 As String
        Dim name1 As String
        Dim class2 As String
        Dim class1 As String
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
            ErrTtl = "Primitive character reading"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        font_name = Left(gm_code, 6)
        class1 = Mid(gm_code, 7, 1)
        class2 = Mid(gm_code, 8, 1)
        name1 = Mid(gm_code, 9, 1)
        name2 = Mid(gm_code, 10, 1)

        '図面名
        ZumenName = "GM-" & font_name


        '検索キーセット
        key_code = "flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & font_name & "' AND"
        key_code = key_code & " font_class1 = '" & class1 & "' AND"
        key_code = key_code & " font_class2 = '" & class2 & "' AND"
        key_code = key_code & " name1 = '" & name1 & "' AND"
        key_code = key_code & " name2 = '" & name2 & "'"

        '検索コマンド作成
        '----- .NET 移行(一旦コメント化) -----
        'sqlcmd = "SELECT haiti_pic FROM " & DBTableName & " WHERE ( " & key_code & ")"

        ''ヒット数チェック
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    ErrMsg = "Primitive character specified was not found." & Chr(13) & gm_code
        '    ErrTtl = "Primitive character reading"
        '    GoTo error_section
        'ElseIf cnt = -1 Then
        '    ErrMsg = "An error occurred on the existing record during the database search."
        '    ErrTtl = "Primitive character reading"
        '    GoTo error_section
        'End If

        '検索
        '----- .NET 移行(一旦コメント化) -----
        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'Rs.MoveFirst()

        ''ﾋﾟｸﾁｬ番号
        'If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '    pic_no = Val(Rs.rdoColumns(0).Value)
        'Else
        '    pic_no = 0
        'End If
        '----- .NET 移行(一旦コメント化) -----
        'Rs.Close()

        ' <- watanabe edit VerUP(2011)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(pic_no, "000") & GensiDir & ZumenName
        w_mess = pic_no.ToString("000") & GensiDir & ZumenName

        w_ret = PokeACAD("ACADREAD", w_mess)
        w_ret = RequestACAD("ACADREAD")

        time_start = Now
        Do
            time_now = Now
            If Trim(form_main.Text2.Text) = "" Then
                If time_now - time_start > timeOutSecond Then
                    ' -> watanabe edit VerUP(2011)
                    'MsgBox("タイムアウトエラー", 64, "ERROR")
                    ErrMsg = "Time-out error"
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

        gm_read = True
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

        gm_read = FAIL
    End Function
	
	Function gm_update() As Short
		Dim result As Object
		Dim now_time As Object
        Dim w_str(42) As String

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

        '----- .NET 移行(文字列の「' '」を削除) -----

        w_str(1) = "0" '削除フラグ
        w_str(2) = "GM" 'ＩＤ(GM固定)
        w_str(3) = Left(form_no.w_font_name.Text, 6) 'ﾌｫﾝﾄ名(KO****)
        w_str(4) = Left(form_no.w_font_class1.Text, 1) 'ﾌｫﾝﾄ区分1(A,F,H,B,D,P,N）
        w_str(5) = Left(form_no.w_font_class2.Text, 1) 'ﾌｫﾝﾄ区分2(0～9: 自動連番）
        w_str(6) = Left(form_no.w_name1.Text, 1) '文字名1
        w_str(7) = Left(form_no.w_name2.Text, 1) '文字名2
        w_str(8) = form_no.w_high.Text '高さ
		w_str(9) = form_no.w_width.Text '幅
		w_str(10) = form_no.w_ang.Text '角度
		w_str(11) = form_no.w_moji_high.Text '実高さ
		w_str(12) = form_no.w_moji_shift.Text 'ずれ量
        Select Case form_no.w_org_hor.Text '水平原点位置(現在Cに固定)'20100706コード変更
            Case "Center"
                w_str(13) = "C"
            Case "Left end"
                w_str(13) = "L"
            Case "Right end"
                w_str(13) = "R"
            Case Else
                ' -> watanabe edit VerUP(2011)
                'MsgBox("水平原点位置エラー", MsgBoxStyle.Critical, "原始文字登録")
                ErrMsg = "Horizontal origin position error."
                ErrTtl = "Primitive character registration"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
        Select Case form_no.w_org_ver.Text '垂直原点位置(現在Bに固定)'20100706コード変更
            Case "Center"
                w_str(14) = "C"
            Case "Top"
                w_str(14) = "T"
            Case "Lower end"
                w_str(14) = "B"
            Case Else
                ' -> watanabe edit VerUP(2011)
                'MsgBox("垂直原点位置エラー", MsgBoxStyle.Critical, "原始文字登録")
                ErrMsg = "Vertical origin position error."
                ErrTtl = "Primitive character registration"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
		
		w_str(15) = CStr(temp_gm.org_x) '文字原点座標X
		w_str(16) = CStr(temp_gm.org_y) '文字原点座標Y
		w_str(17) = CStr(temp_gm.left_bottom_x) '枠左下座標X
		w_str(18) = CStr(temp_gm.left_bottom_y) '枠左下座標Y
		w_str(19) = CStr(temp_gm.right_bottom_x) '枠右下座標X
		w_str(20) = CStr(temp_gm.right_bottom_y) '枠右下座標Y
		w_str(21) = CStr(temp_gm.right_top_x) '枠右上座標X
		w_str(22) = CStr(temp_gm.right_top_y) '枠右上座標Y
		w_str(23) = CStr(temp_gm.left_top_x) '枠左上座標X
		w_str(24) = CStr(temp_gm.left_top_y) '枠左上座標Y
		w_str(25) = form_no.w_hem_width.Text '縁取り幅
		w_str(26) = form_no.w_hatch_ang.Text 'ﾊｯﾁﾝｸﾞ角度
		w_str(27) = form_no.w_hatch_width.Text 'ﾊｯﾁﾝｸﾞ幅
		w_str(28) = form_no.w_hatch_space.Text 'ﾊｯﾁﾝｸﾞ間隔
		w_str(29) = form_no.w_hatch_x.Text 'ﾊｯﾁﾝｸﾞ始点X
		w_str(30) = form_no.w_hatch_y.Text 'ﾊｯﾁﾝｸﾞ始点Y
		w_str(31) = form_no.w_base_r.Text '基準Ｒ
        w_str(32) = form_no.w_old_font_name.Text '旧ﾌｫﾝﾄ名
        w_str(33) = form_no.w_old_font_class.Text '旧ﾌｫﾝﾄ区分
        w_str(34) = form_no.w_old_name.Text '旧文字名
        w_str(35) = form_no.w_haiti_pic.Text '配置PIC
        w_str(36) = "  " '刻印図面ID(w_gz_id)
        w_str(37) = "    " '刻印図面番号(w_gz_no1)
        w_str(38) = "  " '刻印図面変番(w_gz_no2)
        w_str(39) = form_no.w_comment.Text 'コメント
        w_str(40) = form_no.w_dep_name.Text '部署コード
        w_str(41) = form_no.w_entry_name.Text '登録者

        '----- .NET 移行(文字列の「' '」を削除) -----

        If Len(Hour(TimeOfDay)) = 1 Then
			'UPGRADE_WARNING: オブジェクト now_time の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
			now_time = "0" & Hour(TimeOfDay)
		Else
			now_time = Hour(TimeOfDay)
		End If
		
		If Len(Minute(TimeOfDay)) = 1 Then
			now_time = Trim(now_time) & ":0" & Minute(TimeOfDay)
		Else
			now_time = Trim(now_time) & ":" & Minute(TimeOfDay)
		End If

        'w_str(42) = "'" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '登録日
        w_str(42) = Left(form_no.w_entry_date.Text, 4) & "-" & Mid(form_no.w_entry_date.Text, 5, 2) & "-" & Mid(form_no.w_entry_date.Text, 7, 2) & " " &
                    Trim(now_time)

        '----- .NET 移行 -----

        'sqlcmd = "UPDATE " & DBTableName
        'sqlcmd = sqlcmd & " SET  high = " & w_str(8) & ","
        'sqlcmd = sqlcmd & " width = " & w_str(9) & ", ang = " & w_str(10) & ","
        'sqlcmd = sqlcmd & " moji_high = " & w_str(11) & ", moji_shift = " & w_str(12) & ","
        'sqlcmd = sqlcmd & " org_hor = " & w_str(13) & ", org_ver = " & w_str(14) & ","
        'sqlcmd = sqlcmd & " org_x = " & w_str(15) & ", org_y = " & w_str(16) & ","
        'sqlcmd = sqlcmd & " left_bottom_x = " & w_str(17) & ", left_bottom_y = " & w_str(18) & ","
        'sqlcmd = sqlcmd & " right_bottom_x = " & w_str(19) & ", right_bottom_y = " & w_str(20) & ","
        'sqlcmd = sqlcmd & " right_top_x = " & w_str(21) & ", right_top_y = " & w_str(22) & ","
        'sqlcmd = sqlcmd & " left_top_x = " & w_str(23) & ", left_top_y = " & w_str(24) & ","
        'sqlcmd = sqlcmd & " hem_width = " & w_str(25) & ", hatch_ang = " & w_str(26) & ","
        'sqlcmd = sqlcmd & " hatch_width = " & w_str(27) & ", hatch_space = " & w_str(28) & ","
        'sqlcmd = sqlcmd & " hatch_x = " & w_str(29) & ", hatch_y = " & w_str(30) & ","
        'sqlcmd = sqlcmd & " base_r = " & w_str(31) & ", old_font_name = " & w_str(32) & ","
        'sqlcmd = sqlcmd & " old_font_class = " & w_str(33) & ", old_name = " & w_str(34) & ","
        'sqlcmd = sqlcmd & " haiti_pic = " & w_str(35) & ", gz_id = " & w_str(36) & ","
        'sqlcmd = sqlcmd & " gz_no1 = " & w_str(37) & ", gz_no2 = " & w_str(38) & ","
        'sqlcmd = sqlcmd & " comment = " & w_str(39) & ", dep_name = " & w_str(40) & ","
        'sqlcmd = sqlcmd & " entry_name = " & w_str(41) & ", entry_date = " & w_str(42)
        'sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        'sqlcmd = sqlcmd & " WHERE ( flag_delete = 0 AND font_name = " & w_str(3) & " AND"
        'sqlcmd = sqlcmd & " font_class1 = " & w_str(4) & " AND"
        'sqlcmd = sqlcmd & " font_class2 = " & w_str(5) & " AND"
        'sqlcmd = sqlcmd & " name1 = " & w_str(6) & " AND"
        'sqlcmd = sqlcmd & " name2 = " & w_str(7) & ")"

        'ｺﾏﾝﾄﾞ実行
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If

        '--------------------------------------------------------------------------

        Dim joken As String = "flag_delete = '0' AND font_name = '" & w_str(3) & "' AND " &
                              "font_class1 = '" & w_str(4) & "' AND font_class2 = '" & w_str(5) & "' AND " &
                              "name1 = '" & w_str(6) & "' AND name2 = '" & w_str(7) & "'"

        'テーブルレコード数チェック
        Dim count As Integer = VBADO_Count(GL_T_ADO, DBTableName, joken)

        If count = 0 Or count = -1 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If

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
        param.Sign = ""
        paramList.Add(param)

        param.ColumnName = "font_name"
        param.Value = w_str(3)
        param.Sign = "="
        paramList.Add(param)

        param.ColumnName = "font_class1"
        param.Value = w_str(4)
        paramList.Add(param)

        param.ColumnName = "font_class2"
        param.Value = w_str(5)
        paramList.Add(param)

        param.ColumnName = "name1"
        param.Value = w_str(6)
        paramList.Add(param)

        param.ColumnName = "name2"
        param.Value = w_str(7)
        paramList.Add(param)

        param.ColumnName = "high"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(8)
        param.Sign = ""
        paramList.Add(param)

        param.ColumnName = "width"
        param.Value = w_str(9)
        paramList.Add(param)

        param.ColumnName = "ang"
        param.Value = w_str(10)
        paramList.Add(param)

        param.ColumnName = "moji_high"
        param.Value = w_str(11)
        paramList.Add(param)

        param.ColumnName = "moji_shift"
        param.Value = w_str(12)
        paramList.Add(param)

        param.ColumnName = "org_hor"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(13)
        paramList.Add(param)

        param.ColumnName = "org_ver"
        param.Value = w_str(14)
        paramList.Add(param)

        param.ColumnName = "org_x"
        param.SqlDbType = SqlDbType.Float
        param.Value = w_str(15)
        paramList.Add(param)

        param.ColumnName = "org_y"
        param.Value = w_str(16)
        paramList.Add(param)

        param.ColumnName = "left_bottom_x"
        param.Value = w_str(17)
        paramList.Add(param)

        param.ColumnName = "left_bottom_y"
        param.Value = w_str(18)
        paramList.Add(param)

        param.ColumnName = "right_bottom_x"
        param.Value = w_str(19)
        paramList.Add(param)

        param.ColumnName = "right_bottom_y"
        param.Value = w_str(20)
        paramList.Add(param)

        param.ColumnName = "right_top_x"
        param.Value = w_str(21)
        paramList.Add(param)

        param.ColumnName = "right_top_y"
        param.Value = w_str(22)
        paramList.Add(param)

        param.ColumnName = "left_top_x"
        param.Value = w_str(23)
        paramList.Add(param)

        param.ColumnName = "left_top_y"
        param.Value = w_str(24)
        paramList.Add(param)

        param.ColumnName = "hem_width"
        param.Value = w_str(25)
        paramList.Add(param)

        param.ColumnName = "hatch_ang"
        param.Value = w_str(26)
        paramList.Add(param)

        param.ColumnName = "hatch_width"
        param.Value = w_str(27)
        paramList.Add(param)

        param.ColumnName = "hatch_space"
        param.Value = w_str(28)
        paramList.Add(param)

        param.ColumnName = "hatch_x"
        param.Value = w_str(29)
        paramList.Add(param)

        param.ColumnName = "hatch_y"
        param.Value = w_str(30)
        paramList.Add(param)

        param.ColumnName = "base_r"
        param.Value = w_str(31)
        paramList.Add(param)

        param.ColumnName = "old_font_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(32)
        paramList.Add(param)

        param.ColumnName = "old_font_class"
        param.Value = w_str(33)
        paramList.Add(param)

        param.ColumnName = "old_name"
        param.Value = w_str(34)
        paramList.Add(param)

        param.ColumnName = "haiti_pic"
        param.SqlDbType = SqlDbType.TinyInt
        param.Value = w_str(35)
        paramList.Add(param)

        param.ColumnName = "gz_id"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(36)
        paramList.Add(param)

        param.ColumnName = "gz_no1"
        param.Value = w_str(37)
        paramList.Add(param)

        param.ColumnName = "gz_no2"
        param.Value = w_str(38)
        paramList.Add(param)

        param.ColumnName = "comment"
        param.SqlDbType = SqlDbType.VarChar
        param.Value = w_str(39)
        param.DataSize = 255
        paramList.Add(param)

        param.ColumnName = "dep_name"
        param.SqlDbType = SqlDbType.Char
        param.Value = w_str(40)
        paramList.Add(param)

        param.ColumnName = "entry_name"
        param.Value = w_str(41)
        paramList.Add(param)

        param.ColumnName = "entry_date"
        param.SqlDbType = SqlDbType.SmallDateTime
        param.Value = w_str(42)
        paramList.Add(param)

        If VBADO_Update(GL_T_ADO, DBTableName, paramList) <> 1 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If

        '----- .NET 移行 -----

        gm_update = True
		
		Exit Function
		
error_section:
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()

        gm_update = FAIL

    End Function
	
	
	Function gm_search(ByRef gm_code As String) As Short
		Dim w_ret As Object
		Dim result As Object
        Dim w_str(42) As String
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
		

        'GM_KANRIテーブルより該当する原始文字データを求める
        temp_gm.font_name = Mid(gm_code, 1, 6)
        temp_gm.font_class1 = Mid(gm_code, 7, 1)
        temp_gm.font_class2 = Mid(gm_code, 8, 1)
        temp_gm.name1 = Mid(gm_code, 9, 1)
        temp_gm.name2 = Mid(gm_code, 10, 1)

        '検索キーセット
        key_code = "flag_delete = 0 AND font_name = '" & temp_gm.font_name & "' AND"
        key_code = key_code & " font_class1 = '" & temp_gm.font_class1 & "' AND"
        key_code = key_code & " font_class2 = '" & temp_gm.font_class2 & "' AND"
        key_code = key_code & " name1 = '" & temp_gm.name1 & "' AND"
        key_code = key_code & " name2 = '" & temp_gm.name2 & "'"

        '検索コマンド作成
        sqlcmd = "SELECT high, width, ang, moji_high, moji_shift, org_hor, org_ver,"
        sqlcmd = sqlcmd & " hem_width, hatch_ang, hatch_width, hatch_space, hatch_x,"
        sqlcmd = sqlcmd & " hatch_y, base_r, old_font_name, old_font_class, old_name, haiti_pic,"
        sqlcmd = sqlcmd & " comment, dep_name, entry_name, entry_date,"
        sqlcmd = sqlcmd & " flag_delete, id, org_x, org_y,"
        sqlcmd = sqlcmd & " left_bottom_x, left_bottom_y, right_bottom_x, right_bottom_y,"
        sqlcmd = sqlcmd & " right_top_x, right_top_y, left_top_x, left_top_y"
        sqlcmd = sqlcmd & " FROM " & DBTableName
        sqlcmd = sqlcmd & " WHERE (" & key_code & ")"

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    errflg = 1
        '    GoTo error_section
        'ElseIf cnt = -1 Then
        '    errflg = 1
        '    GoTo error_section
        'End If

        '検索
        '----- .NET 移行(一旦コメント化) -----
        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'Rs.MoveFirst()

        'If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '    temp_gm.high = Val(Rs.rdoColumns(0).Value)
        'Else
        '    temp_gm.high = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(1).Value) = False Then
        '    temp_gm.width = Val(Rs.rdoColumns(1).Value)
        'Else
        '    temp_gm.width = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(2).Value) = False Then
        '    temp_gm.ang = Val(Rs.rdoColumns(2).Value)
        'Else
        '    temp_gm.ang = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(3).Value) = False Then
        '    temp_gm.moji_high = Val(Rs.rdoColumns(3).Value)
        'Else
        '    temp_gm.moji_high = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(4).Value) = False Then
        '    temp_gm.moji_shift = Val(Rs.rdoColumns(4).Value)
        'Else
        '    temp_gm.moji_shift = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(5).Value) = False Then
        '    temp_gm.org_hor = Rs.rdoColumns(5).Value
        'Else
        '    temp_gm.org_hor = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(6).Value) = False Then
        '    temp_gm.org_ver = Rs.rdoColumns(6).Value
        'Else
        '    temp_gm.org_ver = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(7).Value) = False Then
        '    temp_gm.hem_width = Val(Rs.rdoColumns(7).Value)
        'Else
        '    temp_gm.hem_width = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(8).Value) = False Then
        '    temp_gm.hatch_ang = Val(Rs.rdoColumns(8).Value)
        'Else
        '    temp_gm.hatch_ang = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(9).Value) = False Then
        '    temp_gm.hatch_width = Val(Rs.rdoColumns(9).Value)
        'Else
        '    temp_gm.hatch_width = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(10).Value) = False Then
        '    temp_gm.hatch_space = Val(Rs.rdoColumns(10).Value)
        'Else
        '    temp_gm.hatch_space = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(11).Value) = False Then
        '    temp_gm.hatch_x = Val(Rs.rdoColumns(11).Value)
        'Else
        '    temp_gm.hatch_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(12).Value) = False Then
        '    temp_gm.hatch_y = Val(Rs.rdoColumns(12).Value)
        'Else
        '    temp_gm.hatch_y = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(13).Value) = False Then
        '    temp_gm.base_r = Val(Rs.rdoColumns(13).Value)
        'Else
        '    temp_gm.base_r = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(14).Value) = False Then
        '    temp_gm.old_font_name = Rs.rdoColumns(14).Value
        'Else
        '    temp_gm.old_font_name = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(15).Value) = False Then
        '    temp_gm.old_font_class = Rs.rdoColumns(15).Value
        'Else
        '    temp_gm.old_font_class = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(16).Value) = False Then
        '    temp_gm.old_name = Rs.rdoColumns(16).Value
        'Else
        '    temp_gm.old_name = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(17).Value) = False Then
        '    temp_gm.haiti_pic = Val(Rs.rdoColumns(17).Value)
        'Else
        '    temp_gm.haiti_pic = 0
        'End If

        'If IsDBNull(Rs.rdoColumns(18).Value) = False Then
        '    temp_gm.comment = Rs.rdoColumns(18).Value
        'Else
        '    temp_gm.comment = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(19).Value) = False Then
        '    temp_gm.dep_name = Rs.rdoColumns(19).Value
        'Else
        '    temp_gm.dep_name = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(20).Value) = False Then
        '    temp_gm.entry_name = Rs.rdoColumns(20).Value
        'Else
        '    temp_gm.entry_name = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(21).Value) = False Then
        '    Dim tmpstr As String
        '    tmpstr = Rs.rdoColumns(21).Value
        '    temp_gm.entry_date = Left(tmpstr, 4) & Mid(tmpstr, 6, 2) & Mid(tmpstr, 9, 2)
        'Else
        '    temp_gm.entry_date = ""
        'End If


        '以降は内容確認画面の項目
        '----- .NET 移行(一旦コメント化) -----
        'If IsDBNull(Rs.rdoColumns(22).Value) = False Then
        '    temp_gm.flag_delete = Val(Rs.rdoColumns(22).Value)
        'Else
        '    temp_gm.flag_delete = 0
        'End If

        'If IsDBNull(Rs.rdoColumns(23).Value) = False Then
        '    temp_gm.id = Rs.rdoColumns(23).Value
        'Else
        '    temp_gm.id = ""
        'End If

        'If IsDBNull(Rs.rdoColumns(24).Value) = False Then
        '    temp_gm.org_x = Val(Rs.rdoColumns(24).Value)
        'Else
        '    temp_gm.org_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(25).Value) = False Then
        '    temp_gm.org_y = Val(Rs.rdoColumns(25).Value)
        'Else
        '    temp_gm.org_y = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(26).Value) = False Then
        '    temp_gm.left_bottom_x = Val(Rs.rdoColumns(26).Value)
        'Else
        '    temp_gm.left_bottom_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(27).Value) = False Then
        '    temp_gm.left_bottom_y = Val(Rs.rdoColumns(27).Value)
        'Else
        '    temp_gm.left_bottom_y = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(28).Value) = False Then
        '    temp_gm.right_bottom_x = Val(Rs.rdoColumns(28).Value)
        'Else
        '    temp_gm.right_bottom_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(29).Value) = False Then
        '    temp_gm.right_bottom_y = Val(Rs.rdoColumns(29).Value)
        'Else
        '    temp_gm.right_bottom_y = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(30).Value) = False Then
        '    temp_gm.right_top_x = Val(Rs.rdoColumns(30).Value)
        'Else
        '    temp_gm.right_top_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(31).Value) = False Then
        '    temp_gm.right_top_y = Val(Rs.rdoColumns(31).Value)
        'Else
        '    temp_gm.right_top_y = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(32).Value) = False Then
        '    temp_gm.left_top_x = Val(Rs.rdoColumns(32).Value)
        'Else
        '    temp_gm.left_top_x = 0.0
        'End If

        'If IsDBNull(Rs.rdoColumns(33).Value) = False Then
        '    temp_gm.left_top_y = Val(Rs.rdoColumns(33).Value)
        'Else
        '    temp_gm.left_top_y = 0.0
        'End If

        'Rs.Close()
        '' <- watanabe add VerUP(2011)

        gm_search = True
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

        gm_search = FAIL
    End Function
	
	Function temp_gm_set(ByRef hexdata As String) As Short
		Dim result As Object
        Dim w_ret As Integer '20100706 修正
        Dim aa As String
		Dim ww As String

        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET 移行(コメント化) -----
        'Dim Rs As RDO.rdoResultset

        On Error Resume Next ' エラーのトラップを留保します。
		Err.Clear()

        aa = ""


        '========================================
        '原始文字データをＨＥＸより変換します
        '========================================
        temp_gm.id = Mid(hexdata, 1, 2)
		temp_gm.font_name = Mid(hexdata, 3, 6)
		temp_gm.font_class1 = Mid(hexdata, 9, 1)
		temp_gm.font_class2 = Mid(hexdata, 10, 1)
		temp_gm.name1 = Mid(hexdata, 11, 1)
		temp_gm.name2 = Mid(hexdata, 12, 1)
		w_ret = HextoDbl(Mid(hexdata, 13, 16), temp_gm.high)
		w_ret = HextoDbl(Mid(hexdata, 29, 16), temp_gm.width)
		w_ret = HextoDbl(Mid(hexdata, 45, 16), temp_gm.ang)
		w_ret = HextoDbl(Mid(hexdata, 61, 16), temp_gm.moji_high)
		w_ret = HextoDbl(Mid(hexdata, 77, 16), temp_gm.moji_shift)
		temp_gm.org_hor = Mid(hexdata, 93, 1)
		temp_gm.org_ver = Mid(hexdata, 94, 1)
		w_ret = HextoDbl(Mid(hexdata, 95, 16), temp_gm.org_x)
		w_ret = HextoDbl(Mid(hexdata, 111, 16), temp_gm.org_y)
		
		w_ret = HextoDbl(Mid(hexdata, 127, 16), temp_gm.left_bottom_x)
		w_ret = HextoDbl(Mid(hexdata, 143, 16), temp_gm.left_bottom_y)
		w_ret = HextoDbl(Mid(hexdata, 159, 16), temp_gm.right_bottom_x)
		w_ret = HextoDbl(Mid(hexdata, 175, 16), temp_gm.right_bottom_y)
		w_ret = HextoDbl(Mid(hexdata, 191, 16), temp_gm.right_top_x)
		w_ret = HextoDbl(Mid(hexdata, 207, 16), temp_gm.right_top_y)
		w_ret = HextoDbl(Mid(hexdata, 223, 16), temp_gm.left_top_x)
		w_ret = HextoDbl(Mid(hexdata, 239, 16), temp_gm.left_top_y)
		
		w_ret = HextoDbl(Mid(hexdata, 255, 16), temp_gm.hem_width)
		w_ret = HextoDbl(Mid(hexdata, 271, 16), temp_gm.hatch_ang)
		w_ret = HextoDbl(Mid(hexdata, 287, 16), temp_gm.hatch_width)
		w_ret = HextoDbl(Mid(hexdata, 303, 16), temp_gm.hatch_space)
		w_ret = HextoDbl(Mid(hexdata, 319, 16), temp_gm.hatch_x)
		w_ret = HextoDbl(Mid(hexdata, 335, 16), temp_gm.hatch_y)
		w_ret = HextoDbl(Mid(hexdata, 351, 16), temp_gm.base_r)

        If open_mode = "Change" Then

            init_sql()

            '検索キーセット
            key_code = "font_name = '" & temp_gm.font_name & "' AND"
            key_code = key_code & " font_class1 = '" & temp_gm.font_class1 & "' AND"
            key_code = key_code & " font_class2 = '" & temp_gm.font_class2 & "' AND"
            key_code = key_code & " name1 = '" & temp_gm.name1 & "' AND"
            key_code = key_code & " name2 = '" & temp_gm.name2 & "' "

            '----- .NET 移行-----

            '検索コマンド作成
            'sqlcmd = "SELECT comment, dep_name, entry_name, entry_date" & " FROM " & DBTableName & " WHERE ( " & key_code & ")"

            'ヒット数チェック
            'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            'If cnt = 0 Then
            '    MsgBox("Primitive character specified was not found.")

            'ElseIf cnt = -1 Then
            '    MsgBox("An error occurred on the existing record during the database search.")

            'Else
            '    '検索
            '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            '    Rs.MoveFirst()

            '    If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            '        temp_gm.comment = Rs.rdoColumns(0).Value
            '    Else
            '        temp_gm.comment = ""
            '    End If

            '    If IsDBNull(Rs.rdoColumns(1).Value) = False Then
            '        temp_gm.dep_name = Rs.rdoColumns(1).Value
            '    Else
            '        temp_gm.dep_name = ""
            '    End If

            '    If IsDBNull(Rs.rdoColumns(2).Value) = False Then
            '        temp_gm.entry_name = Rs.rdoColumns(2).Value
            '    Else
            '        temp_gm.entry_name = ""
            '    End If

            '    Rs.Close()
            'End If
            ' <- watanabe add VerUP(2011)

            '---------------------------------------------------------

            temp_gm.comment = ""
            temp_gm.dep_name = ""
            temp_gm.entry_name = ""

            'テーブルレコード数チェック
            cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

            If cnt = 0 Then
                MsgBox("Primitive character specified was not found.")

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
                    temp_gm.comment = dataList(0)(0)
                    temp_gm.dep_name = dataList(0)(1)
                    temp_gm.entry_name = dataList(0)(2)
                Else
                    MsgBox("Primitive character specified was not found.")
                End If

            End If


            '----- .NET 移行-----

            end_sql()

        End If

        temp_gm.old_font_name = Mid(hexdata, 367, 6)
		temp_gm.old_font_class = Mid(hexdata, 373, 2)
		temp_gm.old_name = Mid(hexdata, 375, 2)
		
		w_ret = HextoSht(Mid(hexdata, 377, 4), temp_gm.haiti_pic)

		Call true_date(aa)
		temp_gm.entry_date = aa
		
        If open_mode = "NEW" Then
            temp_gm.id = ""
            temp_gm.font_name = ""
            temp_gm.font_class1 = ""
            temp_gm.font_class2 = ""
            temp_gm.name1 = ""
            temp_gm.name2 = ""
            temp_gm.comment = ""
            temp_gm.old_font_name = ""
            temp_gm.old_font_class = ""
            temp_gm.old_name = ""
            temp_gm.base_r = 0.0#
            temp_gm.hem_width = 0.0#
            Call true_date(aa)
            temp_gm.entry_date = aa
        End If
		
	End Function
	
	Function temp_gm_get() As Short
        Dim w_ret As Object

        '----- .NET 移行 -----
        'Dim hexdata As New VB6.FixedLengthString(382)
        'Dim www As New VB6.FixedLengthString(16)
        Dim hexdata As String = New String(" "c, 382)
        Dim www As String = New String(" "c, 20)

        '----- .NET 移行(コメント化) -----
        'Dim ii As Integer

        ' 必要文字数分、スペースで初期化
        'hexdata.Value = ""
        'For ii = 1 To 382
        '    hexdata.Value = hexdata.Value & " "
        'Next ii

        '' 必要文字数分、スペースで初期化
        'www.Value = ""
        'For ii = 1 To 16
        '    www.Value = www.Value & " "
        'Next ii
        '----- .NET 移行(コメント化) -----

        '========================================
        '原始文字データをＨＥＸに変換して送信します
        '========================================

        '画面より変更項目内容の取り込み
        temp_gm.font_name = form_no.w_font_name.Text
        temp_gm.font_class1 = form_no.w_font_class1.Text
        temp_gm.font_class2 = form_no.w_font_class2.Text
        temp_gm.name1 = form_no.w_name1.Text
        temp_gm.name2 = form_no.w_name2.Text
        temp_gm.base_r = form_no.w_base_r.Text
        temp_gm.hem_width = form_no.w_hem_width.Text
        temp_gm.comment = form_no.w_comment.Text
        temp_gm.dep_name = form_no.w_dep_name.Text
        temp_gm.entry_name = form_no.w_entry_name.Text
        temp_gm.entry_date = form_no.w_entry_date.Text
        temp_gm.old_font_name = form_no.w_old_font_name.Text
        temp_gm.old_font_class = form_no.w_old_font_class.Text
        temp_gm.old_name = form_no.w_old_name.Text

        '----- .NET 移行 -----
        'Mid(hexdata.Value, 1, 2) = temp_gm.id
        'Mid(hexdata.Value, 3, 6) = temp_gm.font_name
        'Mid(hexdata.Value, 9, 1) = temp_gm.font_class1
        'Mid(hexdata.Value, 10, 1) = temp_gm.font_class2
        'Mid(hexdata.Value, 11, 1) = temp_gm.name1
        'Mid(hexdata.Value, 12, 1) = temp_gm.name2
        '-------------------------------------------------
        Mid(hexdata, 1, 2) = temp_gm.id
        Mid(hexdata, 3, 6) = temp_gm.font_name
        Mid(hexdata, 9, 1) = temp_gm.font_class1
        Mid(hexdata, 10, 1) = temp_gm.font_class2
        Mid(hexdata, 11, 1) = temp_gm.name1
        Mid(hexdata, 12, 1) = temp_gm.name2
        '----- .NET 移行 -----

        '----- .NET 移行 -----
        'w_ret = DbltoHex(temp_gm.high, www.Value)
        'Mid(hexdata.Value, 13, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.width, www.Value)
        'Mid(hexdata.Value, 29, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.ang, www.Value)
        'Mid(hexdata.Value, 45, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.moji_high, www.Value)
        'Mid(hexdata.Value, 61, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.moji_shift, www.Value)
        'Mid(hexdata.Value, 77, 16) = www.Value

        'Mid(hexdata.Value, 93, 1) = temp_gm.org_hor
        'Mid(hexdata.Value, 94, 1) = temp_gm.org_ver

        'w_ret = DbltoHex(temp_gm.org_x, www.Value)
        'Mid(hexdata.Value, 95, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.org_y, www.Value)
        'Mid(hexdata.Value, 111, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.left_bottom_x, www.Value)
        'Mid(hexdata.Value, 127, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.left_bottom_y, www.Value)
        'Mid(hexdata.Value, 143, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.right_bottom_x, www.Value)
        'Mid(hexdata.Value, 159, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.right_bottom_y, www.Value)
        'Mid(hexdata.Value, 175, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.right_top_x, www.Value)
        'Mid(hexdata.Value, 191, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.right_top_y, www.Value)
        'Mid(hexdata.Value, 207, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.left_top_x, www.Value)
        'Mid(hexdata.Value, 223, 16) = www.Value
        'w_ret = DbltoHex(temp_gm.left_top_y, www.Value)
        'Mid(hexdata.Value, 239, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hem_width, www.Value)
        'Mid(hexdata.Value, 255, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hatch_ang, www.Value)
        'Mid(hexdata.Value, 271, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hatch_width, www.Value)
        'Mid(hexdata.Value, 287, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hatch_space, www.Value)
        'Mid(hexdata.Value, 303, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hatch_x, www.Value)
        'Mid(hexdata.Value, 319, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.hatch_y, www.Value)
        'Mid(hexdata.Value, 335, 16) = www.Value

        'w_ret = DbltoHex(temp_gm.base_r, www.Value)
        'Mid(hexdata.Value, 351, 16) = www.Value

        'Mid(hexdata.Value, 367, 6) = temp_gm.old_font_name
        'Mid(hexdata.Value, 373, 2) = temp_gm.old_font_class
        'Mid(hexdata.Value, 375, 2) = temp_gm.old_name

        'w_ret = ShttoHex(form_no.w_haiti_pic.Text, www.Value)
        'Mid(hexdata.Value, 377, 4) = www.Value

        'w_ret = PokeACAD("SPECADD", hexdata.Value)
        '-------------------------------------------------

        w_ret = DbltoHex(temp_gm.high, www)
        Mid(hexdata, 13, 16) = www

        w_ret = DbltoHex(temp_gm.width, www)
        Mid(hexdata, 29, 16) = www

        w_ret = DbltoHex(temp_gm.ang, www)
        Mid(hexdata, 45, 16) = www

        w_ret = DbltoHex(temp_gm.moji_high, www)
        Mid(hexdata, 61, 16) = www

        w_ret = DbltoHex(temp_gm.moji_shift, www)
        Mid(hexdata, 77, 16) = www

        Mid(hexdata, 93, 1) = temp_gm.org_hor
        Mid(hexdata, 94, 1) = temp_gm.org_ver

        w_ret = DbltoHex(temp_gm.org_x, www)
        Mid(hexdata, 95, 16) = www
        w_ret = DbltoHex(temp_gm.org_y, www)
        Mid(hexdata, 111, 16) = www
        w_ret = DbltoHex(temp_gm.left_bottom_x, www)
        Mid(hexdata, 127, 16) = www
        w_ret = DbltoHex(temp_gm.left_bottom_y, www)
        Mid(hexdata, 143, 16) = www
        w_ret = DbltoHex(temp_gm.right_bottom_x, www)
        Mid(hexdata, 159, 16) = www
        w_ret = DbltoHex(temp_gm.right_bottom_y, www)
        Mid(hexdata, 175, 16) = www
        w_ret = DbltoHex(temp_gm.right_top_x, www)
        Mid(hexdata, 191, 16) = www
        w_ret = DbltoHex(temp_gm.right_top_y, www)
        Mid(hexdata, 207, 16) = www
        w_ret = DbltoHex(temp_gm.left_top_x, www)
        Mid(hexdata, 223, 16) = www
        w_ret = DbltoHex(temp_gm.left_top_y, www)
        Mid(hexdata, 239, 16) = www

        w_ret = DbltoHex(temp_gm.hem_width, www)
        Mid(hexdata, 255, 16) = www

        w_ret = DbltoHex(temp_gm.hatch_ang, www)
        Mid(hexdata, 271, 16) = www

        w_ret = DbltoHex(temp_gm.hatch_width, www)
        Mid(hexdata, 287, 16) = www

        w_ret = DbltoHex(temp_gm.hatch_space, www)
        Mid(hexdata, 303, 16) = www

        w_ret = DbltoHex(temp_gm.hatch_x, www)
        Mid(hexdata, 319, 16) = www

        w_ret = DbltoHex(temp_gm.hatch_y, www)
        Mid(hexdata, 335, 16) = www

        w_ret = DbltoHex(temp_gm.base_r, www)
        Mid(hexdata, 351, 16) = www

        Mid(hexdata, 367, 6) = temp_gm.old_font_name
        Mid(hexdata, 373, 2) = temp_gm.old_font_class
        Mid(hexdata, 375, 2) = temp_gm.old_name

        w_ret = ShttoHex(form_no.w_haiti_pic.Text, www)
        Mid(hexdata, 377, 4) = www

        w_ret = PokeACAD("SPECADD", hexdata)
        '----- .NET 移行 -----

        w_ret = RequestACAD("SPECADD")

    End Function
	
	
	Function gm_delete(ByRef gm_code As String) As Short
		Dim result As Object
		Dim now_time As Object
        Dim w_str(42) As String
        'Dim w_command As String'20100616移植削除

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
		w_str(2) = "'" & "GM" & "'" 'ＩＤ(GM固定)
		w_str(3) = "'" & Mid(gm_code, 1, 6) & "'" 'ﾌｫﾝﾄ名(KO****)
		w_str(4) = "'" & Mid(gm_code, 7, 1) & "'" 'ﾌｫﾝﾄ区分1(A,F,H,B,D,P,N）
		w_str(5) = "'" & Mid(gm_code, 8, 1) & "'" 'ﾌｫﾝﾄ区分2(0～9: 自動連番）
		w_str(6) = "'" & Mid(gm_code, 9, 1) & "'" '文字名1
		w_str(7) = "'" & Mid(gm_code, 10, 1) & "'" '文字名2
		w_str(8) = form_no.w_high.Text '高さ
		w_str(9) = form_no.w_width.Text '幅
		w_str(10) = form_no.w_ang.Text '角度
		w_str(11) = form_no.w_moji_high.Text '実高さ
		w_str(12) = form_no.w_moji_shift.Text 'ずれ量
        Select Case form_no.w_org_hor.Text '水平原点位置(現在Cに固定)'20100706コード変更
            Case "Center"
                w_str(13) = "'C'"
            Case "Left end"
                w_str(13) = "'L'"
            Case "Right end"
                w_str(13) = "'R'"
            Case Else
                'Debug.Print "水平原点位置ｴﾗｰ"
                ' -> watanabe edit VerUP(2011)
                ErrMsg = "Horizontal origin position error."
                ErrTtl = "Primitive character delete"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
        Select Case form_no.w_org_ver.Text '垂直原点位置(現在Bに固定)'20100706コード変更
            Case "Center"
                w_str(14) = "'C'"
            Case "Top"
                w_str(14) = "'T'"
            Case "Lower end"
                w_str(14) = "'B'"
            Case Else
                'Debug.Print "垂直原点位置ｴﾗｰ"
                ' -> watanabe edit VerUP(2011)
                ErrMsg = "Vertical origin position error."
                ErrTtl = "Primitive character delete"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
		
		w_str(15) = CStr(temp_gm.org_x) '文字原点座標X
		w_str(16) = CStr(temp_gm.org_y) '文字原点座標Y
		w_str(17) = CStr(temp_gm.left_bottom_x) '枠左下座標X
		w_str(18) = CStr(temp_gm.left_bottom_y) '枠左下座標Y
		w_str(19) = CStr(temp_gm.right_bottom_x) '枠右下座標X
		w_str(20) = CStr(temp_gm.right_bottom_y) '枠右下座標Y
		w_str(21) = CStr(temp_gm.right_top_x) '枠右上座標X
		w_str(22) = CStr(temp_gm.right_top_y) '枠右上座標Y
		w_str(23) = CStr(temp_gm.left_top_x) '枠左上座標X
		w_str(24) = CStr(temp_gm.left_top_y) '枠左上座標Y
		w_str(25) = form_no.w_hem_width.Text '縁取り幅
		w_str(26) = form_no.w_hatch_ang.Text 'ﾊｯﾁﾝｸﾞ角度
		w_str(27) = form_no.w_hatch_width.Text 'ﾊｯﾁﾝｸﾞ幅
		w_str(28) = form_no.w_hatch_space.Text 'ﾊｯﾁﾝｸﾞ間隔
		w_str(29) = form_no.w_hatch_x.Text 'ﾊｯﾁﾝｸﾞ始点X
		w_str(30) = form_no.w_hatch_y.Text 'ﾊｯﾁﾝｸﾞ始点Y
		w_str(31) = form_no.w_base_r.Text '基準Ｒ
		w_str(32) = "'" & form_no.w_old_font_name.Text & "'" '旧ﾌｫﾝﾄ名
		w_str(33) = "'" & form_no.w_old_font_class.Text & "'" '旧ﾌｫﾝﾄ区分
		w_str(34) = "'" & form_no.w_old_name.Text & "'" '旧文字名
		w_str(35) = form_no.w_haiti_pic.Text '配置PIC
		w_str(36) = "'" & "  " & "'" '刻印図面ID(w_gz_id)
		w_str(37) = "'" & "    " & "'" '刻印図面番号(w_gz_no1)
		w_str(38) = "'" & "  " & "'" '刻印図面変番(w_gz_no2)
		w_str(39) = "'" & form_no.w_comment.Text & "'" 'コメント
		w_str(40) = "'" & form_no.w_dep_name.Text & "'" '部署コード
		w_str(41) = "'" & form_no.w_entry_name.Text & "'" '登録者
		
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
		
		w_str(42) = "'" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '登録日



        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1) & ","
        sqlcmd = sqlcmd & " high = " & w_str(8) & ","
        sqlcmd = sqlcmd & " width = " & w_str(9) & ", ang = " & w_str(10) & ","
        sqlcmd = sqlcmd & " moji_high = " & w_str(11) & ", moji_shift = " & w_str(12) & ","
        sqlcmd = sqlcmd & " org_hor = " & w_str(13) & ", org_ver = " & w_str(14) & ","
        sqlcmd = sqlcmd & " org_x = " & w_str(15) & ", org_y = " & w_str(16) & ","
        sqlcmd = sqlcmd & " left_bottom_x = " & w_str(17) & ", left_bottom_y = " & w_str(18) & ","
        sqlcmd = sqlcmd & " right_bottom_x = " & w_str(19) & ", right_bottom_y = " & w_str(20) & ","
        sqlcmd = sqlcmd & " right_top_x = " & w_str(21) & ", right_top_y = " & w_str(22) & ","
        sqlcmd = sqlcmd & " left_top_x = " & w_str(23) & ", left_top_y = " & w_str(24) & ","
        sqlcmd = sqlcmd & " hem_width = " & w_str(25) & ", hatch_ang = " & w_str(26) & ","
        sqlcmd = sqlcmd & " hatch_width = " & w_str(27) & ", hatch_space = " & w_str(28) & ","
        sqlcmd = sqlcmd & " hatch_x = " & w_str(29) & ", hatch_y = " & w_str(30) & ","
        sqlcmd = sqlcmd & " base_r = " & w_str(31) & ", old_font_name = " & w_str(32) & ","
        sqlcmd = sqlcmd & " old_font_class = " & w_str(33) & ", old_name = " & w_str(34) & ","
        sqlcmd = sqlcmd & " haiti_pic = " & w_str(35) & ", gz_id = " & w_str(36) & ","
        sqlcmd = sqlcmd & " gz_no1 = " & w_str(37) & ", gz_no2 = " & w_str(38) & ","
        sqlcmd = sqlcmd & " comment = " & w_str(39) & ", dep_name = " & w_str(40) & ","
        sqlcmd = sqlcmd & " entry_name = " & w_str(41)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ( font_name = " & w_str(3) & " AND"
        sqlcmd = sqlcmd & " font_class1 = " & w_str(4) & " AND"
        sqlcmd = sqlcmd & " font_class2 = " & w_str(5) & " AND"
        sqlcmd = sqlcmd & " name1 = " & w_str(6) & " AND"
        sqlcmd = sqlcmd & " name2 = " & w_str(7) & ")"

        'ｺﾏﾝﾄﾞ実行
        '----- .NET 移行(一旦コメント化) -----
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)

        gm_delete = True
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

        gm_delete = FAIL
    End Function
	

	Function gm_delete_save(ByRef gm_code As String) As Short
		Dim result As Object
		Dim now_time As Object
        '削除フラグをＯＮにします
		Dim w_str(42) As String

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

        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("ﾃﾞｰﾀﾍﾞｰｽにｱｸｾｽ出来ません", MsgBoxStyle.Critical, "SQLｴﾗｰ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        temp_gm.font_name = Mid(gm_code, 1, 6)
		temp_gm.font_class1 = Mid(gm_code, 7, 1)
		temp_gm.font_class2 = Mid(gm_code, 8, 1)
		temp_gm.name1 = Mid(gm_code, 9, 1)
		temp_gm.name2 = Mid(gm_code, 10, 1)
		
		w_str(1) = "1" '削除フラグ(0:OFF 1:ON)
		w_str(2) = "'" & "GM" & "'" 'ＩＤ(GM固定)
        w_str(3) = "'" & temp_gm.font_name & "'" 'ﾌｫﾝﾄ名(KO****)
		w_str(4) = "'" & temp_gm.font_class1 & "'" 'ﾌｫﾝﾄ区分1(A,F,H,B,D,P,N）
		w_str(5) = "'" & temp_gm.font_class2 & "'" 'ﾌｫﾝﾄ区分2(0～9: 自動連番）
		w_str(6) = "'" & temp_gm.name1 & "'" '文字名1
		w_str(7) = "'" & temp_gm.name2 & "'" '文字名2
		w_str(8) = form_no.w_high.Text '高さ
		w_str(9) = form_no.w_width.Text '幅
		w_str(10) = form_no.w_ang.Text '角度
		w_str(11) = form_no.w_moji_high.Text '実高さ
		w_str(12) = form_no.w_moji_shift.Text 'ずれ量
        Select Case form_no.w_org_hor.Text '水平原点位置(現在Cに固定)'20100706コード変更
            Case "Center"
                w_str(13) = "'C'"
            Case "Left end"
                w_str(13) = "'L'"
            Case "Right end"
                w_str(13) = "'R'"
            Case Else
                'Debug.Print "水平原点位置ｴﾗｰ"
                ' -> watanabe edit VerUP(2011)
                ErrMsg = "Horizontal origin position error."
                ErrTtl = "Primitive character delete"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
        Select Case form_no.w_org_ver.Text '垂直原点位置(現在Bに固定)'20100706コード変更
            Case "Center"
                w_str(14) = "'C'"
            Case "Top"
                w_str(14) = "'T'"
            Case "Lower end"
                w_str(14) = "'B'"
            Case Else
                'Debug.Print "垂直原点位置ｴﾗｰ"
                ' -> watanabe edit VerUP(2011)
                ErrMsg = "Vertical origin position error."
                ErrTtl = "Primitive character delete"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
        End Select
		
		w_str(15) = CStr(temp_gm.org_x) '文字原点座標X
		w_str(16) = CStr(temp_gm.org_y) '文字原点座標Y
		w_str(17) = CStr(temp_gm.left_bottom_x) '枠左下座標X
		w_str(18) = CStr(temp_gm.left_bottom_y) '枠左下座標Y
		w_str(19) = CStr(temp_gm.right_bottom_x) '枠右下座標X
		w_str(20) = CStr(temp_gm.right_bottom_y) '枠右下座標Y
		w_str(21) = CStr(temp_gm.right_top_x) '枠右上座標X
		w_str(22) = CStr(temp_gm.right_top_y) '枠右上座標Y
		w_str(23) = CStr(temp_gm.left_top_x) '枠左上座標X
		w_str(24) = CStr(temp_gm.left_top_y) '枠左上座標Y
		w_str(25) = form_no.w_hem_width.Text '縁取り幅
		w_str(26) = form_no.w_hatch_ang.Text 'ﾊｯﾁﾝｸﾞ角度
		w_str(27) = form_no.w_hatch_width.Text 'ﾊｯﾁﾝｸﾞ幅
		w_str(28) = form_no.w_hatch_space.Text 'ﾊｯﾁﾝｸﾞ間隔
		w_str(29) = form_no.w_hatch_x.Text 'ﾊｯﾁﾝｸﾞ始点X
		w_str(30) = form_no.w_hatch_y.Text 'ﾊｯﾁﾝｸﾞ始点Y
		w_str(31) = form_no.w_base_r.Text '基準Ｒ
		w_str(32) = "'" & form_no.w_old_font_name.Text & "'" '旧ﾌｫﾝﾄ名
		w_str(33) = "'" & form_no.w_old_font_class.Text & "'" '旧ﾌｫﾝﾄ区分
		w_str(34) = "'" & form_no.w_old_name.Text & "'" '旧文字名
		w_str(35) = form_no.w_haiti_pic.Text '配置PIC
		w_str(36) = "'" & "  " & "'" '刻印図面ID(w_gz_id)
		w_str(37) = "'" & "    " & "'" '刻印図面番号(w_gz_no1)
		w_str(38) = "'" & "  " & "'" '刻印図面変番(w_gz_no2)
		w_str(39) = "'" & form_no.w_comment.Text & "'" 'コメント
		w_str(40) = "'" & form_no.w_dep_name.Text & "'" '部署コード
		w_str(41) = "'" & form_no.w_entry_name.Text & "'" '登録者
		
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
		
		w_str(42) = "'" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '登録日


        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ","
        sqlcmd = sqlcmd & " font_name = " & w_str(3) & ", font_class1 = " & w_str(4) & ","
        sqlcmd = sqlcmd & " font_class2 = " & w_str(5) & ", name1 = " & w_str(6) & ","
        sqlcmd = sqlcmd & " name2 = " & w_str(7) & ", high = " & w_str(8) & ","
        sqlcmd = sqlcmd & " width = " & w_str(9) & ", ang = " & w_str(10) & ","
        sqlcmd = sqlcmd & " moji_high = " & w_str(11) & ", moji_shift = " & w_str(12) & ","
        sqlcmd = sqlcmd & " org_hor = " & w_str(13) & ", org_ver = " & w_str(14) & ","
        sqlcmd = sqlcmd & " org_x = " & w_str(15) & ", org_y = " & w_str(16) & ","
        sqlcmd = sqlcmd & " left_bottom_x = " & w_str(17) & ", left_bottom_y = " & w_str(18) & ","
        sqlcmd = sqlcmd & " right_bottom_x = " & w_str(19) & ", right_bottom_y = " & w_str(20) & ","
        sqlcmd = sqlcmd & " right_top_x = " & w_str(21) & ", right_top_y = " & w_str(22) & ","
        sqlcmd = sqlcmd & " left_top_x = " & w_str(23) & ", left_top_y = " & w_str(24) & ","
        sqlcmd = sqlcmd & " hem_width = " & w_str(25) & ", hatch_ang = " & w_str(26) & ","
        sqlcmd = sqlcmd & " hatch_width = " & w_str(27) & ", hatch_space = " & w_str(28) & ","
        sqlcmd = sqlcmd & " hatch_x = " & w_str(29) & ", hatch_y = " & w_str(30) & ","
        sqlcmd = sqlcmd & " base_r = " & w_str(31) & ", old_font_name = " & w_str(32) & ","
        sqlcmd = sqlcmd & " old_font_class = " & w_str(33) & ", old_name = " & w_str(34) & ","
        sqlcmd = sqlcmd & " haiti_pic = " & w_str(35) & ", gz_id = " & w_str(36) & ","
        sqlcmd = sqlcmd & " gz_no1 = " & w_str(37) & ", gz_no2 = " & w_str(38) & ","
        sqlcmd = sqlcmd & " comment = " & w_str(39) & ", dep_name = " & w_str(40) & ","
        sqlcmd = sqlcmd & " entry_name = " & w_str(41) & ", entry_date = " & w_str(42)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ( font_name = " & w_str(3) & " )"

        'ｺﾏﾝﾄﾞ実行
        '----- .NET 移行(一旦コメント化) -----
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
        '    ErrTtl = "SQL error"
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)

        'gm_delete = True
        Exit Function
		
error_section: 
		' gm_delete = FAIL

        ' -> watanabe add VerUP(2011)
        If ErrMsg = "" Then
            ErrMsg = Err.Description
            ErrTtl = "System error"
        End If
        MsgBox(ErrMsg, MsgBoxStyle.Critical, ErrTtl)

        On Error Resume Next
        Err.Clear()
        ' <- watanabe add VerUP(2011)

	End Function
End Module