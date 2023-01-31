Option Strict Off
Option Explicit On
Module MJ_GZ
	Function gz_insert() As Short
        Dim j As Integer '20100707 �C��
		Dim w_ret As Object
        Dim i As Integer '20100707 �C��
		Dim now_time As Object
        Dim result As Integer '20100707 �C��
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        'Dim kubun As Short
        ' <- watanabe del VerUP(2011)

        Dim DBTableNameGm As String

        'UPGRADE_ISSUE: �錾�̌^���T�|�[�g����Ă��܂���
        'Dim err_flg(100) As String*3

        ' -> watanabe edit 2013.06.03
        'Dim err_flg(100) As String '20100617�ڐA�ǉ�
        Dim err_flg(200) As String '20100617�ڐA�ǉ�
        ' <- watanabe edit 2013.06.03

        Dim flg As Short

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


        '---------< ����}�� �o�^ �V�K >---------------------------------------------


		DBTableNameGm = DBName & "..gm_kanri"
		
		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
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
        'result = SqlCmd(SqlConn, " flag_delete = 0 AND")
        'result = SqlCmd(SqlConn, " id = 'KO' AND")
        'result = SqlCmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "'")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		Loop 
        '		MsgBox("�}�ʔԍ������ɍ���}�ʂɑ��݂��܂��B" & Chr(13) & "�V�K�ł̓o�^�͏o���܂���", MsgBoxStyle.Critical, "number exist error")
        '		GoTo error_section
        '	End If
        'End If


        '�����L�[�Z�b�g
        key_code = " flag_delete = 0 AND"
        key_code = key_code & " id = 'KO' AND"
        key_code = key_code & " no1 = '" & form_no.w_no1.Text & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT *  FROM " & DBTableName & " WHERE " & key_code

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt > 0 Then
            ErrMsg = "Drawing number exists in the already brand drawing." & Chr(13) & "It is not possible to register a new."
            ErrTtl = "number exist error"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "carved seal drawing registration"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


		w_str(1) = "0" '�폜�t���O
		w_str(2) = "'" & "KO" & "'" '�h�c(KO�Œ�)
		w_str(3) = "'" & form_no.w_no1.Text & "'" '�}�ʔԍ�
		w_str(4) = "'" & form_no.w_no2.Text & "'" '�ϔ�
		w_str(5) = "'" & form_no.w_comment.Text & "'" '�R�����g
		w_str(6) = "'" & form_no.w_dep_name.Text & "'" '�����R�[�h
		w_str(7) = "'" & form_no.w_entry_name.Text & "'" '�o�^��
		
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
		
		w_str(8) = "'" & form_no.w_entry_date.Text & " " & Trim(now_time) & "'" '�o�^��
		
		w_str(9) = form_no.w_gm_num.Text '���n������
		
		'���n�����f�[�^�`�F�b�N(���ɑ��̍���}�ʂɎg�p����Ă���΃G���[)
		flg = 0
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            err_flg(i) = ""
            w_ret = exist_gm_gz(DBTableNameGm, temp_gz.gm_name(i), temp_gz.no1, temp_gz.no2)
            If w_ret = -1 Then
                ' -> watanabe edit VerUP(2011)
                'MsgBox("SQL�װ�ł�", MsgBoxStyle.Critical, "SQL�װ")
                ErrMsg = "SQL error."
                ErrTtl = "SQL error"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
            ElseIf w_ret = 1 Then
                err_flg(i) = "100"
                flg = 1
            ElseIf w_ret = 2 Then
                err_flg(i) = "200"
                flg = 1
            ElseIf w_ret = 3 Then
            Else
            End If
        Next i
		
		'Brand Ver.3 �ύX
		' For i = 1 To 63
		'     w_str(i + 9) = "'" & temp_gz.gm_name(i) & " '"
		' Next i
		
		If flg = 1 Then
			For i = 1 To temp_gz.gm_num Step 2
				w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i), ((i - 1) / 2) + 1, 1)
				If (i + 1) > temp_gz.gm_num Then Exit For
				w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i + 1), ((i - 1) / 2) + 1, 3)
			Next i

            ' -> watanabe add VerUP(2011)
            ErrMsg = "Primitive character that can not be registered is included."
            ErrTtl = "carved seal drawing new registration"
            ' <- watanabe add VerUP(2011)

            GoTo error_section
		End If
		
		'���n�����ɍ���}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "KO", form_no.w_no1.Text, form_no.w_no2.Text)
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "  ", "    ", "  ")
                Next j
                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to add information stamped drawing primitive character code to [" & temp_gz.gm_name(i) & "]"
                ErrTtl = "carved seal drawing new registration"
                ' <- watanabe add VerUP(2011)
                GoTo error_section
            End If
        Next i
		
		
		'����}��̧�قɓo�^
		' MsgBox "����}�ʂɓo�^���܂�", , "�m�F"

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        'w_command = "INSERT INTO " & DBTableName & " VALUES("
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 71
        'For i = 1 To 8
        '	result = SqlCmd(SqlConn, w_str(i) & ",")
        '	w_command = w_command & w_str(i) & ","
        'Next i
        '' result = SqlCmd(SqlConn, w_str(72) & ")")
        '' w_command = w_command & w_str(72) & ")"
        'result = SqlCmd(SqlConn, w_str(9) & ")")
        'w_command = w_command & w_str(9) & ")"
        '
        'result = SqlExec(SqlConn)
        '
        ''����}�ʂ̓o�^�Ɏ��s�������͌��n�����̍���}�ʏ����폜����
        'If result = FAIL Then
        '	MsgBox("����}�ʂ̓o�^�Ɏ��s�����̂Ō��n�����̓o�^�f�[�^���N���A���܂�",  , "�m�F")
        '          For i = 1 To form_no.w_gm_num.Text
        '              w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
        '          Next i
        '	GoTo error_section
        'End If
        'result = SqlResults(SqlConn)


        sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        For i = 1 To 8
            sqlcmd = sqlcmd & w_str(i) & ","
        Next i
        sqlcmd = sqlcmd & w_str(9) & ")"

        '����ގ��s
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)

        '����}�ʂ̓o�^�Ɏ��s�������͌��n�����̍���}�ʏ����폜����
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            For i = 1 To Val(Trim(form_no.w_gm_num.Text))
                w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
            Next i
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()
		
		
		' Brand Ver.3 �ύX
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            init_sql()

            w_str(1) = "'" & "KO" & "'" '�h�c(KO�Œ�)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '�}�ʔԍ�
            w_str(3) = "'" & form_no.w_no2.Text & "'" '�ϔ�
            w_str(4) = i '���n�����ԍ�


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '    result = SqlCmd(SqlConn, w_str(j) & ",")
            'Next j
            'result = SqlCmd(SqlConn, "'" & temp_gz.gm_name(i) & "'")
            'result = SqlCmd(SqlConn, " )")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then
            '    For j = 1 To Val(Trim(form_no.w_gm_num.Text))
            '        w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
            '    Next j
            '    GoTo error_section
            'End If
            'result = SqlResults(SqlConn)


            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            For j = 1 To 4
                sqlcmd = sqlcmd & w_str(j) & ","
            Next j
            sqlcmd = sqlcmd & "'" & temp_gz.gm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                For j = 1 To Val(Trim(form_no.w_gm_num.Text))
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
                Next j
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i
		
		
		gz_insert = True
		
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

        gz_insert = FAIL
    End Function
	
    Function gz_read(ByRef wk_id As String, ByRef wk_no1 As String, ByRef wk_no2 As String) As Short
        Dim w_ret As Object
        Dim gz_code As String '20100707 �C��
        Dim result As Integer '20100707 �C��
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


        '�߸���ԍ�

        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "SELECT entry_name")
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
        '    If SqlNextRow(SqlConn) = REGROW Then
        '        wk_entry_name = CStr(Val(SqlData(SqlConn, 1)))
        '    Else
        '
        '        ' -> watanabe add VerUP(2011)
        '        gz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
        '        ' <- watanabe add VerUP(2011)
        '
        '        MsgBox("�w�肳�ꂽ����}�ʂ�������܂���" & Chr(13) & gz_code, MsgBoxStyle.Critical, "data not found")
        '        gz_read = FAIL
        '        Exit Function
        '    End If
        'Else
        '    MsgBox("SQL �G���[")
        '    gz_read = FAIL
        '    Exit Function
        'End If


        '�����L�[�Z�b�g
        key_code = "flag_delete = 0 AND"
        key_code = key_code & " id = '" & wk_id & "' AND"
        key_code = key_code & " no1 = '" & wk_no1 & "' AND"
        key_code = key_code & " no2 = '" & wk_no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT entry_name FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            gz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
            ErrMsg = "Carved seal drawings specified was not found." & Chr(13) & gz_code
            ErrTtl = "data not found"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "carved seal drawing reading"
            GoTo error_section
        End If

        '����
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            wk_entry_name = CStr(Val(Rs.rdoColumns(0).Value))
        Else
            wk_entry_name = ""
        End If
        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        w_mess = KokuinDir & wk_id & "-" & wk_no1 & "-" & wk_no2
        w_ret = PokeACAD("MDLREAD", w_mess)
        w_ret = RequestACAD("MDLREAD")

        ' -> watanabe add VerUP(2011)
        gz_read = SUCCEED
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

        gz_read = FAIL
        ' <- watanabe add VerUP(2011)

    End Function
	
    Function gz_update() As Short
        Dim j As Integer '20100707 �C��
        Dim w_ret As Object
        Dim i As Integer '20100707 �C��
        Dim result As Integer '20100707 �C��
        Dim now_time As Object
        Dim DBTableNameGm As Object
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe edit VerUP(2011)
        'Dim temp_gz2 As GZ_KANRI '���ƃe�[�u���Q�Ɨp
        Dim temp_gz2 As New GZ_KANRI '���ƃe�[�u���Q�Ɨp
        ' <- watanabe edit VerUP(2011)

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


        '------- ����}�� �o�^ �C�� --------------------


        temp_gz2.Initilize() '20100707 �R�[�h�ǉ�

        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        DBTableNameGm = DBName & "..gm_kanri"

        w_str(1) = "0" '�폜�t���O
        w_str(2) = "'" & "KO" & "'" '�h�c(KO�Œ�)
        w_str(3) = "'" & Trim(form_no.w_no1.Text) & "'" '�}�ʔԍ�
        w_str(4) = "'" & Trim(form_no.w_no2.Text) & "'" '�ϔ�
        w_str(5) = "'" & Trim(form_no.w_comment.Text) & "'" '�R�����g
        w_str(6) = "'" & Trim(form_no.w_dep_name.Text) & "'" '�����R�[�h
        w_str(7) = "'" & Trim(form_no.w_entry_name.Text) & "'" '�o�^��

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

        w_str(8) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '�o�^��

        w_str(9) = Trim(form_no.w_gm_num.Text) '���n������

        ' Brand Ver.3 �ύX
        ' For i = 1 To 63
        '     w_str(i + 9) = "'" & Trim(temp_gz.gm_name(i)) & "'"
        ' Next i

        '�e�[�u������

        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_gz.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_gz.no2 & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '    '    If SqlNextRow(SqlConn) = REGROW Then
        '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        temp_gz.comment = SqlData(SqlConn, 1)
        '        temp_gz.dep_name = SqlData(SqlConn, 2)
        '        temp_gz.entry_name = SqlData(SqlConn, 3)
        '        temp_gz.entry_date = SqlData(SqlConn, 4)
        '    Loop
        '    '    Else
        '    '       MsgBox "SQL�װ"
        '    '       GoTo error_section
        '    '    End If
        'Else
        '    MsgBox("����}�ʂ�����܂���B�C���ł͓o�^�ł��܂���", MsgBoxStyle.Critical, "number is exist")
        '    GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_gz.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_gz.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "There is no Stamp drawing. You can not register the correct."
            ErrTtl = "number is exist"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz.comment = Rs.rdoColumns(0).Value
            Else
                temp_gz.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_gz.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_gz.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_gz.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_gz.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_gz.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_gz.entry_date = ""
            End If

            Rs.Close()

        End If
        ' <- watanabe edit VerUP(2011)


        '' 1997.6.13 yam Mod.........
        '// �e�[�u������(����) --------------------------------------------------------------------
        temp_gz2.no1 = temp_gz.no1
        '----- .NET �ڍs -----
        'temp_gz2.no2 = VB6.Format(CDbl(Val(Trim(temp_gz.no2))) - 1, "00")
        temp_gz2.no2 = (CDbl(Val(Trim(temp_gz.no2))) - 1).ToString("00")


        ' -> watanabe edit VerUP(2011)
        '' Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, gm_num,")
        'result = sqlcmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, gm_num ")
        '
        '' Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''    result = SqlCmd(SqlConn, " gm_name" & Format(i, "000") & ",")
        '' Next i
        '' result = SqlCmd(SqlConn, " gm_name" & Format(63, "000"))
        '
        'result = sqlcmd(SqlConn, " FROM " & DBTableName)
        'result = sqlcmd(SqlConn, " WHERE ( no1 = '" & temp_gz2.no1 & "' AND")
        'result = sqlcmd(SqlConn, " no2 = '" & temp_gz2.no2 & "' )")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        temp_gz2.comment = SqlData(SqlConn, 1)
        '        temp_gz2.dep_name = SqlData(SqlConn, 2)
        '        temp_gz2.entry_name = SqlData(SqlConn, 3)
        '        temp_gz2.entry_date = SqlData(SqlConn, 4)
        '        temp_gz2.gm_num = Val(SqlData(SqlConn, 5))
        '
        '        'Brand Ver.3 �ύX
        '        '       For i = 1 To 63
        '        '          temp_gz2.gm_name(i) = SqlData$(SqlConn, 5 + i)
        '        '       Next i
        '    Loop
        'Else
        '    GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_gz2.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_gz2.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date, gm_num FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "carved seal drawing update registration"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz2.comment = Rs.rdoColumns(0).Value
            Else
                temp_gz2.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_gz2.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_gz2.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_gz2.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_gz2.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_gz2.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_gz2.entry_date = ""
            End If

            If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                temp_gz2.gm_num = Val(Rs.rdoColumns(4).Value)
            Else
                temp_gz2.gm_num = 0
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        For i = 1 To temp_gz2.gm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT gm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ("
            'w_command = w_command & " no1 = '" & temp_gz2.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_gz2.no2 & "' AND"
            'w_command = w_command & " gm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_gz2.gm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = " no1 = '" & temp_gz2.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_gz2.no2 & "' AND"
            key_code = key_code & " gm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT gm_name FROM " & DBTableName2 & " WHERE (" & key_code & " )"

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Or cnt = -1 Then
                Exit For
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz2.gm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_gz2.gm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        end_sql()
        init_sql()

        '// ���e�[�u���̌��n�����̍���}�ʃf�[�^���N���A���� ----------------------------------------
        '���n�����̍���}�ʏ����폜����

        For i = 1 To temp_gz2.gm_num
            w_ret = update_gm_gz(DBTableNameGm, temp_gz2.gm_name(i), "  ", "    ", "  ")
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz2.gm_name(j), "KO", temp_gz2.no1, temp_gz2.no2)
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to change information stamped drawing primitive character code [" & temp_gz2.gm_name(i) & "]"
                ErrTtl = "carved seal drawing update registration"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i

        '// �e�[�u���̌��n�����ɍ���}�ʃf�[�^��o�^���� ----------------------------------------
        '���n�����ɍ���}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "KO", form_no.w_no1.Text, form_no.w_no2.Text)
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "  ", "    ", "  ")
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to change information stamped drawing primitive character code [" & temp_gz.gm_name(i) & "]"
                ErrTtl = "carved seal drawing update registration"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i
        'Yam Mod End


        '�e�[�u���X�V

        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "UPDATE " & DBTableName)
        'result = sqlcmd(SqlConn, " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ",")
        'result = sqlcmd(SqlConn, " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ",")
        'result = sqlcmd(SqlConn, " comment = " & w_str(5) & ", dep_name = " & w_str(6) & ",")
        'result = sqlcmd(SqlConn, " entry_name = " & w_str(7) & ", entry_date = " & w_str(8) & ",")
        ''Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, " gm_num = " & w_str(9) & ",")
        'result = sqlcmd(SqlConn, " gm_num = " & w_str(9))
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''     gname = "gm_name" & Format(i, "000")
        ''     result = SqlCmd(SqlConn, gname & " = " & w_str(9 + i) & ",")
        '' Next i
        '' gname = "gm_name" & Format(63, "000")
        '' result = SqlCmd(SqlConn, gname & " = " & w_str(9 + 63))
        '
        'result = sqlcmd(SqlConn, " From " & DBTableName & "(PAGLOCK)")
        'result = sqlcmd(SqlConn, " WHERE ")
        'result = sqlcmd(SqlConn, " id = 'KO' AND")
        'result = sqlcmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "' AND")
        'result = sqlcmd(SqlConn, " no2 = '" & form_no.w_no2.Text & "'")
        ''Send the command to SQL Server and start execution.
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)


        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ","
        sqlcmd = sqlcmd & " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ","
        sqlcmd = sqlcmd & " comment = " & w_str(5) & ", dep_name = " & w_str(6) & ","
        sqlcmd = sqlcmd & " entry_name = " & w_str(7) & ", entry_date = " & w_str(8) & ","
        sqlcmd = sqlcmd & " gm_num = " & w_str(9)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE "
        sqlcmd = sqlcmd & " id = 'KO' AND"
        sqlcmd = sqlcmd & " no1 = '" & form_no.w_no1.Text & "' AND"
        sqlcmd = sqlcmd & " no2 = '" & form_no.w_no2.Text & "'"

        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        '���f�[�^�폜
        init_sql()


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "DELETE FROM " & DBTableName2 & " WHERE ( ")
        'result = sqlcmd(SqlConn, "id = 'KO' AND ")
        'result = sqlcmd(SqlConn, "no1 = '" & form_no.w_no1.Text & "' AND ")
        'result = sqlcmd(SqlConn, "no2 = '" & form_no.w_no2.Text & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)

        sqlcmd = "DELETE FROM " & DBTableName2 & " WHERE ( "
        sqlcmd = sqlcmd & "id = 'KO' AND "
        sqlcmd = sqlcmd & "no1 = '" & form_no.w_no1.Text & "' AND "
        sqlcmd = sqlcmd & "no2 = '" & form_no.w_no2.Text & "' )"

        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not delete the existing data from the database.(" & DBTableName2 & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        '�V�K�o�^
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            init_sql()

            w_str(1) = "'" & "KO" & "'" '�h�c(KO�Œ�)
            w_str(2) = "'" & Trim(form_no.w_no1.Text) & "'" '�}�ʔԍ�
            w_str(3) = "'" & Trim(form_no.w_no2.Text) & "'" '�ϔ�
            w_str(4) = i '���n�����ԍ�
            w_str(5) = "'" & Trim(temp_gz.gm_name(i)) & "'" '���n�����R�[�h


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'result = sqlcmd(SqlConn, w_str(1) & ", ")
            'result = sqlcmd(SqlConn, w_str(2) & ", ")
            'result = sqlcmd(SqlConn, w_str(3) & ", ")
            'result = sqlcmd(SqlConn, w_str(4) & ", ")
            'result = sqlcmd(SqlConn, w_str(5) & " )")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then GoTo error_section
            'result = SqlResults(SqlConn)

            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            sqlcmd = sqlcmd & w_str(1) & ", "
            sqlcmd = sqlcmd & w_str(2) & ", "
            sqlcmd = sqlcmd & w_str(3) & ", "
            sqlcmd = sqlcmd & w_str(4) & ", "
            sqlcmd = sqlcmd & w_str(5) & " )"

            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        gz_update = True

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

        gz_update = FAIL
    End Function

    Function gz_addnum() As Short
        Dim now_time As Object
        Dim j As Integer '20100707 �C��
        Dim w_ret As Object
        Dim i As Integer '20100707 �C��
        Dim result As Integer '20100707 �C��
        Dim DBTableNameGm As Object
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe edit VerUP(2011)
        'Dim temp_gz2 As GZ_KANRI '���ƃe�[�u���Q�Ɨp
        Dim temp_gz2 As New GZ_KANRI '���ƃe�[�u���Q�Ɨp
        ' <- watanabe edit VerUP(2011)

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


        ' -> watanabe add VerUP(2011)
        temp_gz2.Initilize() '20100707 �R�[�h�ǉ�
        ' <- watanabe add VerUP(2011)


        DBTableNameGm = DBName & "..gm_kanri"

        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        '// �e�[�u������(����) --------------------------------------------------------------------
        '���Ƃ̍���}�ʏ������o���܂�

        temp_gz2.no1 = temp_gz.no1
        '----- .NET �ڍs -----
        'temp_gz2.no2 = VB6.Format(CDbl(Val(Trim(temp_gz.no2))) - 1, "00")
        temp_gz2.no2 = (CDbl(Val(Trim(temp_gz.no2))) - 1).ToString("00")


        ' -> watanabe edit VerUP(2011)
        ''Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, gm_num,")
        'result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, gm_num ")
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''    result = SqlCmd(SqlConn, " gm_name" & Format(i, "000") & ",")
        '' Next i
        '' result = SqlCmd(SqlConn, " gm_name" & Format(63, "000"))
        '
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_gz2.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_gz2.no2 & "' )")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '        temp_gz2.comment = SqlData(SqlConn, 1)
        '        temp_gz2.dep_name = SqlData(SqlConn, 2)
        '        temp_gz2.entry_name = SqlData(SqlConn, 3)
        '        temp_gz2.entry_date = SqlData(SqlConn, 4)
        '        temp_gz2.gm_num = Val(SqlData(SqlConn, 5))
        '
        '        'Brand Ver.3 �ύX
        '        '       For i = 1 To 63
        '        '          temp_gz2.gm_name(i) = SqlData$(SqlConn, 5 + i)
        '        '       Next i
        '    Loop
        'Else
        '    GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_gz2.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_gz2.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date, gm_num FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "carved seal drawing update registration"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz2.comment = Rs.rdoColumns(0).Value
            Else
                temp_gz2.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_gz2.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_gz2.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_gz2.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_gz2.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_gz2.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_gz2.entry_date = ""
            End If

            If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                temp_gz2.gm_num = Val(Rs.rdoColumns(4).Value)
            Else
                temp_gz2.gm_num = 0
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        For i = 1 To temp_gz2.gm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT gm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ("
            'w_command = w_command & " no1 = '" & temp_gz2.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_gz2.no2 & "' AND"
            'w_command = w_command & " gm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_gz2.gm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = " no1 = '" & temp_gz2.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_gz2.no2 & "' AND"
            key_code = key_code & " gm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT gm_name FROM " & DBTableName2 & " WHERE ( " & key_code & " )"

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Or cnt = -1 Then
                Exit For
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz2.gm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_gz2.gm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        end_sql()
        init_sql()

        '// ���e�[�u���̌��n�����̍���}�ʃf�[�^���N���A���� ----------------------------------------
        '���n�����̍���}�ʏ����폜����
        For i = 1 To temp_gz2.gm_num
            w_ret = update_gm_gz(DBTableNameGm, temp_gz2.gm_name(i), "  ", "    ", "  ")
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz2.gm_name(j), "KO", temp_gz2.no1, temp_gz2.no2)
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to change information stamped drawing primitive character code [" & temp_gz2.gm_name(i) & "]"
                ErrTtl = "carved seal drawing change number registration"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i


        '// �e�[�u���̌��n�����ɍ���}�ʃf�[�^��o�^���� ----------------------------------------
        '���n�����ɍ���}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "KO", form_no.w_no1.Text, form_no.w_no2.Text)
            '1�ł����s����Α��̌��n�����̍���}�ʃf�[�^���N���A����
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "  ", "    ", "  ")
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to change information stamped drawing primitive character code [" & temp_gz.gm_name(i) & "]"
                ErrTtl = "carved seal drawing change number registration"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i


        '// ����}�ʂ̓o�^ -------------------------------------------------------------------
        w_str(1) = "0" '�폜�t���O
        w_str(2) = "'" & "KO" & "'" '�h�c(KO�Œ�)
        w_str(3) = "'" & Trim(form_no.w_no1.Text) & "'" '�}�ʔԍ�
        w_str(4) = "'" & Trim(form_no.w_no2.Text) & "'" '�ϔ�
        w_str(5) = "'" & Trim(form_no.w_comment.Text) & "'" '�R�����g
        w_str(6) = "'" & Trim(form_no.w_dep_name.Text) & "'" '�����R�[�h
        w_str(7) = "'" & Trim(form_no.w_entry_name.Text) & "'" '�o�^��

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

        w_str(8) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '�o�^��

        w_str(9) = Trim(form_no.w_gm_num.Text) '���n������

        'Brand Ver.3 �ύX
        ' For i = 1 To 63
        '     w_str(i + 9) = "'" & Trim(temp_gz.gm_name(i)) & "'"
        ' Next i

        '����}��̧�قɓo�^
        ' MsgBox "����}�ʂɓo�^���܂�", , "�m�F"


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        'w_command = "INSERT INTO " & DBTableName & " VALUES("
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 71
        'For i = 1 To 8
        '    result = sqlcmd(SqlConn, w_str(i) & ",")
        '    w_command = w_command & w_str(i) & ","
        'Next i
        '' result = SqlCmd(SqlConn, w_str(72) & ")")
        '' w_command = w_command & w_str(72) & ")"
        'result = sqlcmd(SqlConn, w_str(9) & ")")
        'w_command = w_command & w_str(9) & ")"
        '
        'result = SqlExec(SqlConn)
        ''����}�ʂ̓o�^�Ɏ��s�������͌��n�����̍���}�ʏ����폜����
        'If result = FAIL Then
        '    For i = 1 To Val(Trim(form_no.w_gm_num.Text))
        '        w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
        '    Next i
        '    GoTo error_section
        'End If
        'result = SqlResults(SqlConn)


        sqlcmd = "INSERT INTO " & DBTableName & " VALUES("
        For i = 1 To 8
            sqlcmd = sqlcmd & w_str(i) & ","
        Next i
        sqlcmd = sqlcmd & w_str(9) & ")"

        '����ގ��s
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)

        '�o�^�Ɏ��s�������͐}�ʏ����폜����
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            For i = 1 To Val(Trim(form_no.w_gm_num.Text))
                w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
            Next i
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        ' Brand Ver.3 �ύX
        For i = 1 To Val(Trim(form_no.w_gm_num.Text))
            init_sql()
            w_str(1) = "'" & "KO" & "'" '�h�c(KO�Œ�)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '�}�ʔԍ�
            w_str(3) = "'" & form_no.w_no2.Text & "'" '�ϔ�
            w_str(4) = i '���n�����ԍ�


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '    result = sqlcmd(SqlConn, w_str(j) & ",")
            'Next j
            'result = sqlcmd(SqlConn, "'" & temp_gz.gm_name(i) & "'")
            'result = sqlcmd(SqlConn, " )")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then
            '    For j = 1 To Val(Trim(form_no.w_gm_num.Text))
            '        w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "  ", "    ", "  ")
            '    Next j
            '    GoTo error_section
            'End If
            'result = SqlResults(SqlConn)


            sqlcmd = "INSERT INTO " & DBTableName2 & " VALUES("
            For j = 1 To 4
                sqlcmd = sqlcmd & w_str(j) & ","
            Next j
            sqlcmd = sqlcmd & "'" & temp_gz.gm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            '����ގ��s
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                For j = 1 To Val(Trim(form_no.w_gm_num.Text))
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "  ", "    ", "  ")
                Next j
                ErrMsg = "Can not be registered in the database.(" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i


        gz_addnum = True

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

        gz_addnum = FAIL
    End Function

    Function gz_search(ByRef gz_code1 As String, ByRef gz_code2 As String, ByRef flag As Short) As Short
        Dim i As Object
        Dim w_ret As Integer '20100707 �C��
        Dim result As Integer '20100707 �C��
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
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        errflg = 0
        ' <- watanabe add VerUP(2011)


        'flag 0:�폜�t���O��0�̃f�[�^�̂݌���
        'flag 1:���ׂẴf�[�^������


        If SqlConn = 0 Then
            MsgBox("Can not access the database.", MsgBoxStyle.Critical, "SQL error")
            ' -> watanabe add VerUP(2011)
            errflg = 1
            ' <- watanabe add VerUP(2011)
            GoTo error_section
        End If

        'GZ_KANRI�e�[�u�����Y�����錴�n�����f�[�^�����߂�
        temp_gz.no1 = gz_code1
        temp_gz.no2 = gz_code2


        ' -> watanabe edit VerUP(2011)
        'w_command = "SELECT flag_delete, comment, dep_name, entry_name, entry_date, gm_num"
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 63
        ''    w_command = w_command & ", gm_name" & Format(i, "000")
        '' Next i
        '
        'w_command = w_command & " FROM " & DBTableName
        'If flag = 0 Then
        '    w_command = w_command & " WHERE (flag_delete = 0 AND no1 = '" & temp_gz.no1 & "' AND"
        'Else
        '    w_command = w_command & " WHERE (no1 = '" & temp_gz.no1 & "' AND"
        'End If
        'w_command = w_command & " no2 = '" & temp_gz.no2 & "')"
        '
        'result = SqlCmd(SqlConn, w_command)
        '
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '    '   Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '    If SqlNextRow(SqlConn) = REGROW Then
        '        temp_gz.flag_delete = CByte(SqlData(SqlConn, 1))
        '        temp_gz.comment = SqlData(SqlConn, 2)
        '        temp_gz.dep_name = SqlData(SqlConn, 3)
        '        temp_gz.entry_name = SqlData(SqlConn, 4)
        '        ww = SqlData(SqlConn, 5)
        '        w_ret = SqlDateCrack(SqlConn, df, ww)
        '        temp_gz.entry_date = df.Year_Renamed & df.Month_Renamed & df.Day_Renamed
        '
        '        temp_gz.gm_num = Val(SqlData(SqlConn, 6))
        '
        '        ' Brand Ver.3 �ύX
        '        '     For i = 1 To 63
        '        '       temp_gz.gm_name(i) = SqlData$(SqlConn, 6 + i)
        '        '     Next i
        '    Else
        '        GoTo error_section
        '    End If
        '
        '    '   Loop
        'Else
        '    GoTo error_section
        'End If


        '�����L�[�Z�b�g
        If flag = 0 Then
            key_code = "flag_delete = 0 AND no1 = '" & temp_gz.no1 & "' AND"
        Else
            key_code = "no1 = '" & temp_gz.no1 & "' AND"
        End If
        key_code = key_code & " no2 = '" & temp_gz.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT flag_delete, comment, dep_name, entry_name, entry_date, gm_num FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = 0 Then
            errflg = 1
            GoTo error_section
        ElseIf cnt = -1 Then
            errflg = 1
            GoTo error_section
        End If

        '����
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            temp_gz.flag_delete = CByte(Rs.rdoColumns(0).Value)
        Else
            temp_gz.flag_delete = 0
        End If

        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
            temp_gz.comment = Rs.rdoColumns(1).Value
        Else
            temp_gz.comment = ""
        End If

        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
            temp_gz.dep_name = Rs.rdoColumns(2).Value
        Else
            temp_gz.dep_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
            temp_gz.entry_name = Rs.rdoColumns(3).Value
        Else
            temp_gz.entry_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(4).Value) = False Then
            Dim tmpstr As String
            tmpstr = Rs.rdoColumns(4).Value
            temp_gz.entry_date = Left(tmpstr, 4) & Mid(tmpstr, 6, 2) & Mid(tmpstr, 9, 2)
        Else
            temp_gz.entry_date = ""
        End If

        If IsDBNull(Rs.rdoColumns(5).Value) = False Then
            temp_gz.gm_num = Val(Rs.rdoColumns(5).Value)
        Else
            temp_gz.gm_num = 0
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        For i = 1 To temp_gz.gm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT gm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ( "
            'w_command = w_command & " no1 = '" & temp_gz.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_gz.no2 & "' AND"
            'w_command = w_command & " gm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_gz.gm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = " no1 = '" & temp_gz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_gz.no2 & "' AND"
            key_code = key_code & " gm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT gm_name FROM " & DBTableName2 & " WHERE ( " & key_code & " )"

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Then
                Exit For
            ElseIf cnt = -1 Then
                Exit For
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz.gm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_gz.gm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i


        gz_search = True

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

        gz_search = FAIL
    End Function

    Function temp_gz_set(ByRef hexdata As String) As Short
        Dim i As Object

        Dim aa As String

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)


        '========================================
        '���n�����f�[�^���g�d�w���ϊ����܂�
        '========================================
        'w_ret = HextoSht(Mid$(hexdata, 1, 4), temp_gz.gm_num)

        temp_gz.gm_num = 0

        ' -> watanabe edit 2013.05.29
        'For i = 1 To 63
        For i = 1 To 130
            ' <- watanabe edit 2013.05.29

            temp_gz.gm_name(i) = ""
        Next i

        If open_mode = "NEW" Then
            temp_gz.gm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_gz.gm_num
                temp_gz.gm_name(i) = Mid(hexdata, (i - 1) * 10 + 4, 10)
            Next i
            temp_gz.id = "KO"
            temp_gz.no1 = ""
            temp_gz.no2 = "00"
            temp_gz.comment = ""
            temp_gz.dep_name = ""
            temp_gz.entry_name = ""
            Call true_date(aa)
            temp_gz.entry_date = aa
        ElseIf open_mode = "Revision number" Then
            temp_gz.id = "KO"
            temp_gz.gm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_gz.gm_num
                temp_gz.gm_name(i) = Mid(hexdata, (i - 1) * 10 + 4, 10)
            Next i
            Call true_date(aa)
            temp_gz.entry_date = aa
        ElseIf open_mode = "modify" Then
            temp_gz.id = "KO"
            temp_gz.gm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_gz.gm_num
                temp_gz.gm_name(i) = Mid(hexdata, (i - 1) * 10 + 4, 10)
            Next i
            Call true_date(aa)
            temp_gz.entry_date = aa
        End If

    End Function

    Function zumen_no_set(ByRef hexdata As String) As Short
        Dim t4 As String '20100707 �C��
        Dim t3 As String '20100707 �C��
        Dim t2 As String '20100707 �C��
        Dim t1 As String '20100707 �C��
        Dim nn As Integer
        Dim result As Integer '20100707 �C��

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String '20100707 �C��
        'Dim aa As String
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


        '========================================
        '�}�ʃf�[�^���g�d�w���ϊ����܂�
        '========================================
        'w_ret = HextoSht(Mid$(hexdata, 1, 4), temp_gz.gm_num)

        'MsgBox "�}�ʖ����獏��}�ʃe�[�u�����������܂� " & hexdata & "," & Mid$(hexdata, 4, 4) & "," & Mid$(hexdata, 9, 2)

        If open_mode = "modify" Then
            temp_gz.id = "KO"
            temp_gz.no1 = Mid(hexdata, 4, 4)
            temp_gz.no2 = Mid(hexdata, 9, 2)


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT comment, dep_name, entry_name, entry_date "
            'w_command = w_command & " FROM " & DBTableName
            'w_command = w_command & " WHERE "
            'w_command = w_command & " flag_delete = 0 AND"
            'w_command = w_command & " id = 'KO' AND"
            'w_command = w_command & " no1 = '" & temp_gz.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_gz.no2 & "'"
            '
            ''    MsgBox "w_command = " & w_command
            '
            'result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date ")
            'result = SqlCmd(SqlConn, " FROM " & DBTableName)
            'result = SqlCmd(SqlConn, " WHERE ")
            'result = SqlCmd(SqlConn, " flag_delete = 0 AND")
            'result = SqlCmd(SqlConn, " id = 'KO' AND")
            'result = SqlCmd(SqlConn, " no1 = '" & temp_gz.no1 & "' AND")
            'result = SqlCmd(SqlConn, " no2 = '" & temp_gz.no2 & "'")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then GoTo error_section
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    '        If SqlNextRow(SqlConn) = REGROW Then
            '    Do Until SqlNextRow(SqlConn) = NOMOREROWS
            '        temp_gz.comment = SqlData(SqlConn, 1)
            '        temp_gz.dep_name = SqlData(SqlConn, 2)
            '        temp_gz.entry_name = SqlData(SqlConn, 3)
            '        temp_gz.entry_date = SqlData(SqlConn, 4)
            '    Loop
            '    '        Else
            '    '           GoTo error_section
            '    '        End If
            '    If temp_gz.entry_name = "" Then
            '        MsgBox("����}�ʃf�[�^������܂���" & Chr(13) & "�C�������͏o���܂���", MsgBoxStyle.Critical, "�װ")
            '        GoTo error_section
            '    End If
            'Else
            '    GoTo error_section
            'End If


            '�����L�[�Z�b�g
            key_code = " flag_delete = 0 AND"
            key_code = key_code & " id = 'KO' AND"
            key_code = key_code & " no1 = '" & temp_gz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_gz.no2 & "'"

            '�����R�}���h�쐬
            sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE " & key_code

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = 0 Then
                MsgBox("There is no Stamp drawing data." & Chr(13) & "Can not modify processing.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            ElseIf cnt = -1 Then
                MsgBox("An error occurred on the existing record during the database search.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_gz.comment = Rs.rdoColumns(0).Value
            Else
                temp_gz.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_gz.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_gz.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_gz.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_gz.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_gz.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_gz.entry_date = ""
            End If

            Rs.Close()
            ' <- watanabe add VerUP(2011)


        ElseIf open_mode = "Revision number" Then
            temp_gz.id = "KO"
            temp_gz.no1 = Mid(hexdata, 4, 4)
            temp_gz.no2 = Mid(hexdata, 9, 2)

            '�����L�[�Z�b�g
            key_code = " id = 'KO' AND"
            key_code = key_code & " no1 = '" & temp_gz.no1 & "'"

            '�����R�}���h�쐬
            sqlcmd = "SELECT no2, comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE " & key_code

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = 0 Then
                MsgBox("There is no Stamp drawing data." & Chr(13) & "It is not possible to revision number processing.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            ElseIf cnt = -1 Then
                MsgBox("An error occurred on the existing record during the database search.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            temp_gz.no2 = "-1"

            Do Until Rs.EOF

                If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                    nn = Val(Rs.rdoColumns(0).Value)
                Else
                    nn = -1
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

                If Val(temp_gz.no2) < nn Then
                    '----- .NET �ڍs -----
                    'temp_gz.no2 = VB6.Format(nn, "00")
                    temp_gz.no2 = nn.ToString("00")

                    temp_gz.comment = t1
                    temp_gz.dep_name = t2
                    temp_gz.entry_name = t3
                    temp_gz.entry_date = t4
                End If

                Rs.MoveNext()
            Loop

            If Val(temp_gz.no2) < 0 Then
                MsgBox("There is no Stamp drawing data." & Chr(13) & "It is not possible to revision number processing.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            Else
                '----- .NET �ڍs -----
                'temp_gz.no2 = VB6.Format(Val(temp_gz.no2) + 1, "00")
                temp_gz.no2 = (Val(temp_gz.no2) + 1).ToString("00")
            End If

            Rs.Close()
            ' <- watanabe add VerUP(2011)


        End If

        zumen_no_set = True
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

        zumen_no_set = FAIL
    End Function

    Function gz_delete(ByRef gz_code1 As String, ByRef gz_code2 As String) As Short
        Dim j As Object
        Dim w_ret As Object
        Dim i As Object
        Dim result As Integer '20100707 �C��
        Dim DBTableNameGm As Object

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


        DBTableNameGm = DBName & "..gm_kanri"


        If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
        End If

        w_str(1) = "1" '�폜�t���O
        w_str(2) = "'" & "KO" & "'" '�h�c(KO�Œ�)
        w_str(3) = "'" & gz_code1 & "'" '�}�ʔԍ�(****)
        w_str(4) = "'" & gz_code2 & "'" '�ϔ�(00~99�j
        ' w_str(5) = "'" & form_no.w_comment.Text & "'"                  '�R�����g
        ' w_str(6) = "'" & form_no.w_dep_name.Text & "'"                 '�����R�[�h
        ' w_str(7) = "'" & form_no.w_entry_name.Text & "'"               '�o�^��
        ' w_str(8) = "'" & form_no.w_entry_date.Text & "'"               '�o�^��
        ' w_str(9) = form_no.w_gm_num.Text                               '���n������


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

        '����ގ��s
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not be registered in the database.(" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        'MsgBox "���n�����ɂ̍���}�ʏ����폜", , "debug"
        '���n�����ɍ���}�ʏ����폜����
        For i = 1 To form_no.w_gm_num
            '    MsgBox "���n�����R�[�h[" & temp_gz.gm_name(i) & "]�̍���}�ʏ����폜���܂�", , "�m�F"
            w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(i), "  ", "    ", "  ")
            '1�ł����s����Α��̌��n�����̍���}�ʃf�[�^�����A����
            If w_ret = -1 Then
                '        MsgBox "���n�����R�[�h[" & temp_gz.gm_name(i) & "]�̍���}�ʏ��폜�Ɏ��s���܂���" & Chr$(13) & "���܂œo�^�������͕��A���܂�", vbOK, "�m�F"
                For j = 1 To i
                    '            MsgBox temp_gz.gm_name(j) & "�𕜋A���Ă��܂�"
                    w_ret = update_gm_gz(DBTableNameGm, temp_gz.gm_name(j), "KO", form_no.w_no1.Text, form_no.w_no2.Text)
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to change information stamped drawing primitive character code [" & temp_gz.gm_name(i) & "]"
                ErrTtl = "carved seal drawing delete"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i

        gz_delete = True

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

        gz_delete = FAIL
    End Function
End Module