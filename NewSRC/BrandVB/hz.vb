Option Strict Off
Option Explicit On
Module MJ_HZ
	
	Function hz_insert() As Short
		Dim j As Object
		Dim w_ret As Object
		Dim i As Object
		Dim now_time As Object
        Dim result As Integer '20100707 �C��
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        'Dim kubun As Short
        ' <- watanabe del VerUP(2011)

        Dim DBTableNameHm As String

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


		'Brand Ver.3 �ύX
		' DBTableNameHm = DBName & "..hm_kanri"
		DBTableNameHm = DBName & "..hm_kanri1"
		
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
        '
        'result = SqlCmd(SqlConn, " WHERE ")
        'result = SqlCmd(SqlConn, " id = 'HE' AND")
        'result = SqlCmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "'")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		Loop 
        '		MsgBox("�}�ʔԍ������ɕҏW�����}�ʂɑ��݂��܂��B" & Chr(13) & "�V�K�ł̓o�^�͏o���܂���", MsgBoxStyle.Critical, "number exist error")
        '		GoTo error_section
        '	End If
        'Else
        '	GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = " id = 'HE' AND"
        key_code = key_code & " no1 = '" & form_no.w_no1.Text & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT *  FROM " & DBTableName & " WHERE " & key_code

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt > 0 Then
            ErrMsg = "Drawing number exists in the already editing characters drawing." & Chr(13) & "It is not possible to register a new."
            ErrTtl = "number exist error"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Editing characters drawing registration"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


		w_str(1) = "0" '�폜�t���O
		w_str(2) = "'" & "HE" & "'" '�h�c(HE�Œ�)
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
		
		w_str(8) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '�o�^��
		
		w_str(9) = CStr(Val(form_no.w_hm_num.Text)) '���n������
		
		'�ҏW�����f�[�^�`�F�b�N(���ɑ��̕ҏW�����}�ʂɎg�p����Ă���΃G���[)
		For i = 1 To Val(form_no.w_hm_num.Text)
			w_ret = exist_hm_hz(DBTableNameHm, temp_hz.hm_name(i), form_no.w_no1.Text, form_no.w_no2.Text)
			If w_ret = -1 Then
                ' -> watanabe edit VerUP(2011)
                'MsgBox("SQL�װ�ł�", MsgBoxStyle.Critical, "SQL�װ")
                ErrMsg = "SQL error."
                ErrTtl = "SQL error"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
			ElseIf w_ret = 1 Then 
                ' -> watanabe edit VerUP(2011)
                'MsgBox("�ҏW�����R�[�h[" & temp_hz.hm_name(i) & "]�͊��ɑ��̍���}�ʂŎg�p����Ă��܂��̂œo�^�o���܂���", MsgBoxStyle.Critical, "����}�ʐV�K�o�^�װ")
                ErrMsg = "It is not possible to register and Editing characters code [" & temp_hz.hm_name(i) & "] because it is used in editing characters other drawings already"
                ErrTtl = "Editing characters drawing new registration error"
                ' <- watanabe edit VerUP(2011)
                GoTo error_section
			End If
			' Brand Ver.3 �ύX
			'     w_str(i + 9) = "'" & temp_hz.hm_name(i) & "'"
		Next i
		
		' Brand Ver.3 �ύX
		' For i = Val(form_no.w_hm_num.Text) + 1 To 63
		'     w_str(i + 9) = "'" & Space$(8) & "'"
		' Next i
		
		
		'�ҏW�����ɕҏW�����}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            '    MsgBox "�ҏW�����R�[�h[" & temp_hz.hm_name(i) & "]�ɕҏW�����}�ʏ����L�q���܂�", vbOK, "�m�F"
            w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "HE", form_no.w_no1.Text, form_no.w_no2.Text)
            '1�ł����s����Α��̕ҏW�����̕ҏW�����}�ʃf�[�^���N���A����
            If w_ret = -1 Then
                '        MsgBox "�ҏW�����R�[�h[" & temp_hz.hm_name(i) & "]�̕ҏW�����}�ʏ��ǉ��Ɏ��s���܂���" & Chr$(13) & "���܂œo�^���������N���A���܂�", vbOK, "�m�F"
                For j = 1 To i
                    '            MsgBox j & "���N���A���Ă��܂�"
                    w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(j), "  ", "    ", "  ")
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to add editing characters drawing information editing characters code to [" & temp_hz.hm_name(i) & "]"
                ErrTtl = "Editing characters drawing new registration error"
                ' <- watanabe add VerUP(2011)
                GoTo error_section
            End If
        Next i
		
		
		'�ҏW�����}��̧�قɓo�^

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 71
        'For i = 1 To 8
        '	result = SqlCmd(SqlConn, w_str(i) & ",")
        'Next i
        '' result = SqlCmd(SqlConn, w_str(72) & ")")
        'result = SqlCmd(SqlConn, w_str(9) & ")")
        '
        'result = SqlExec(SqlConn)
        '
        ''�ҏW�����}�ʂ̓o�^�Ɏ��s�������͕ҏW�����̕ҏW�����}�ʏ����폜����
        'If result = FAIL Then
        '          For i = 1 To Val(Trim(form_no.w_hm_num.Text))
        '              w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "  ", "    ", "  ")
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

        '�ҏW�����}�ʂ̓o�^�Ɏ��s�������͕ҏW�����̕ҏW�����}�ʏ����폜����
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            For i = 1 To Val(Trim(form_no.w_hm_num.Text))
                w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "  ", "    ", "  ")
            Next i
            ErrMsg = "Can not be registered in the database (" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        ' Brand Ver.3 �ύX
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            init_sql()

            w_str(1) = "'" & "HE" & "'" '�h�c(HE�Œ�)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '�}�ʔԍ�
            w_str(3) = "'" & form_no.w_no2.Text & "'" '�ϔ�
            w_str(4) = i '�ҏW�����ԍ�


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '    result = sqlcmd(SqlConn, w_str(j) & ",")
            'Next j
            'result = sqlcmd(SqlConn, "'" & temp_hz.hm_name(i) & "'")
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
            sqlcmd = sqlcmd & "'" & temp_hz.hm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            '����ގ��s
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database (" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i


        hz_insert = True

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

        hz_insert = FAIL
    End Function
	
	Function hz_read(ByRef wk_id As String, ByRef wk_no1 As String, ByRef wk_no2 As String) As Short
		Dim w_ret As Object
		Dim hz_code As Object
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
        '	Else
        '
        '              ' -> watanabe add VerUP(2011)
        '              hz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
        '              ' <- watanabe add VerUP(2011)
        '
        '
        '              MsgBox("�w�肳�ꂽ�ҏW�����}�ʂ�������܂���" & Chr(13) & hz_code, MsgBoxStyle.Critical, "data not found")
        '		hz_read = FAIL
        '		Exit Function
        '	End If
        'Else
        '	MsgBox("SQL �G���[")
        '	hz_read = FAIL
        '	Exit Function
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
            hz_code = "(" & wk_id & "-" & wk_no1 & "-" & wk_no2 & ")"
            ErrMsg = "Editing characters drawings specified was not found." & Chr(13) & hz_code
            ErrTtl = "data not found"
            GoTo error_section
        ElseIf cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Editing characters drawing reading"
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


        w_mess = HensyuZumenDir & wk_id & "-" & wk_no1 & "-" & wk_no2
		w_ret = PokeACAD("MDLREAD", w_mess)
		w_ret = RequestACAD("MDLREAD")

        ' -> watanabe add VerUP(2011)
        hz_read = SUCCEED
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

        hz_read = FAIL
        ' <- watanabe add VerUP(2011)

	End Function
	
	Function hz_update() As Short
		Dim now_time As Object
		Dim j As Object
		Dim w_ret As Object
		Dim i As Object
		Dim result As Object
		Dim DBTableNameHm As Object
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe edit VerUP(2011)
        'Dim temp_hz2 As HZ_KANRI '���ƃe�[�u���Q�Ɨp
        Dim temp_hz2 As New HZ_KANRI '���ƃe�[�u���Q�Ɨp
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


        '------------- < �ҏW�����}�� �C�� > --------------------------------------------------


        temp_hz2.Initilize() '20100707 �R�[�h�ǉ�

		' Brand Ver.3 �ύX
		' DBTableNameHm = DBName & "..hm_kanri"
		DBTableNameHm = DBName & "..hm_kanri1"
		
		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If
		

        '�e�[�u������

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date")
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_hz.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_hz.no2 & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		temp_hz.comment = SqlData(SqlConn, 1)
        '		temp_hz.dep_name = SqlData(SqlConn, 2)
        '		temp_hz.entry_name = SqlData(SqlConn, 3)
        '		temp_hz.entry_date = SqlData(SqlConn, 4)
        '	Loop 
        'Else
        '	MsgBox("�ҏW�����}�ʂ�����܂���B", MsgBoxStyle.Critical, "number is exist")
        '	GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_hz.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_hz.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "There is no Editing characters drawing."
            ErrTtl = "number is exist"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz.comment = Rs.rdoColumns(0).Value
            Else
                temp_hz.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_hz.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_hz.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_hz.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_hz.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_hz.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_hz.entry_date = ""
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        'Mody
        '// �e�[�u������(����) --------------------------------------------------------------------
        temp_hz2.no1 = temp_hz.no1

        '----- .NET �ڍs -----
        'temp_hz2.no2 = VB6.Format(CDbl(Val(Trim(temp_hz.no2))) - 1, "00")
        temp_hz2.no2 = (CDbl(Val(Trim(temp_hz.no2))) - 1).ToString("00")


        ' -> watanabe edit VerUP(2011)
        '      'Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, hm_num,")
        'result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, hm_num ")
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''    result = SqlCmd(SqlConn, " hm_name" & Format(i, "000") & ",")
        '' Next i
        '' result = SqlCmd(SqlConn, " hm_name" & Format(63, "000"))
        '
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        '      result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_hz2.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_hz2.no2 & "' )")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		temp_hz2.comment = SqlData(SqlConn, 1)
        '		temp_hz2.dep_name = SqlData(SqlConn, 2)
        '		temp_hz2.entry_name = SqlData(SqlConn, 3)
        '		temp_hz2.entry_date = SqlData(SqlConn, 4)
        '		temp_hz2.hm_num = Val(SqlData(SqlConn, 5))
        '
        '		'Brand Ver.3 �ύX
        '		'       For i = 1 To 63
        '		'          temp_hz2.hm_name(i) = SqlData$(SqlConn, 5 + i)
        '		'       Next i
        '	Loop 
        'Else
        '	GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_hz2.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_hz2.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date, hm_num FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Editing characters drawing update registration"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz2.comment = Rs.rdoColumns(0).Value
            Else
                temp_hz2.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_hz2.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_hz2.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_hz2.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_hz2.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_hz2.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_hz2.entry_date = ""
            End If

            If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                temp_hz2.hm_num = Val(Rs.rdoColumns(4).Value)
            Else
                temp_hz2.hm_num = 0
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        For i = 1 To temp_hz2.hm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT hm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ("
            'w_command = w_command & " no1 = '" & temp_hz2.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_hz2.no2 & "' AND"
            'w_command = w_command & " hm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_hz2.hm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = " no1 = '" & temp_hz2.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_hz2.no2 & "' AND"
            key_code = key_code & " hm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT hm_name FROM " & DBTableName2 & " WHERE (" & key_code & " )"

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Or cnt = -1 Then
                Exit For
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz2.hm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_hz2.hm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        end_sql()
        init_sql()

        '// ���e�[�u���̌��n�����̕ҏW�����}�ʃf�[�^���N���A���� ----------------------------------------
        '�ҏW�����}�ʏ����폜����
        For i = 1 To temp_hz2.hm_num
            w_ret = update_hm_hz(DBTableNameHm, temp_hz2.hm_name(i), "  ", "    ", "  ")
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_hm_hz(DBTableNameHm, temp_hz2.hm_name(j), "KO", temp_hz2.no1, temp_hz2.no2)
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to Editing characters drawing information change Editing characters code [" & temp_hz2.hm_name(i) & "]"
                ErrTtl = "Editing characters drawing update registration error"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i

        '// �e�[�u���̌��n�����ɕҏW�����}�ʃf�[�^��o�^���� ----------------------------------------
        '�ҏW�����}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "HE", form_no.w_no1.Text, form_no.w_no2.Text)
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(j), "  ", "    ", "  ")
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to Editing characters drawing information change Editing characters code [" & temp_hz2.hm_name(i) & "]"
                ErrTtl = "Editing characters drawing update registration error"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i
        'mod end

        w_str(1) = "0" '�폜�t���O
        w_str(2) = "'" & "HE" & "'" '�h�c(HE�Œ�)
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

        w_str(8) = "'" & Trim(form_no.w_entry_date.Text) & " " & Trim(now_time) & "'" '�o�^��

        w_str(9) = Trim(form_no.w_hm_num.Text) '�ҏW������


        'Brand Ver.3 �ύX
        ' For i = 1 To 63
        '     w_str(i + 9) = "'" & Trim(temp_hz.hm_name(i)) & "'"
        ' Next i


        '�e�[�u���X�V

        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "UPDATE " & DBTableName)
        'result = sqlcmd(SqlConn, " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ",")
        'result = sqlcmd(SqlConn, " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ",")
        'result = sqlcmd(SqlConn, " comment = " & w_str(5) & ", dep_name = " & w_str(6) & ",")
        'result = sqlcmd(SqlConn, " entry_name = " & w_str(7) & ", entry_date = " & w_str(8) & ",")
        '
        ''Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, " hm_num = " & w_str(9) & ",")
        'result = sqlcmd(SqlConn, " hm_num = " & w_str(9))
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''     gname = " hm_name" & Format(i, "000")
        ''     result = SqlCmd(SqlConn, gname & " = " & w_str(9 + i) & ",")
        '' Next i
        '' gname = " hm_name" & Format(63, "000")
        '' result = SqlCmd(SqlConn, gname & " = " & w_str(9 + 63))
        '
        'result = sqlcmd(SqlConn, " From " & DBTableName & "(PAGLOCK)")
        'result = sqlcmd(SqlConn, " WHERE ( ")
        'result = sqlcmd(SqlConn, " id = 'HE' AND")
        'result = sqlcmd(SqlConn, " no1 = '" & form_no.w_no1.Text & "' AND")
        'result = sqlcmd(SqlConn, " no2 = '" & form_no.w_no2.Text & "' )")
        '
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)


        sqlcmd = "UPDATE " & DBTableName
        sqlcmd = sqlcmd & " SET flag_delete = " & w_str(1) & ", id = " & w_str(2) & ","
        sqlcmd = sqlcmd & " no1 = " & w_str(3) & ", no2 = " & w_str(4) & ","
        sqlcmd = sqlcmd & " comment = " & w_str(5) & ", dep_name = " & w_str(6) & ","
        sqlcmd = sqlcmd & " entry_name = " & w_str(7) & ", entry_date = " & w_str(8) & ","
        sqlcmd = sqlcmd & " hm_num = " & w_str(9)
        sqlcmd = sqlcmd & " From " & DBTableName & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ( "
        sqlcmd = sqlcmd & " id = 'HE' AND"
        sqlcmd = sqlcmd & " no1 = '" & form_no.w_no1.Text & "' AND"
        sqlcmd = sqlcmd & " no2 = '" & form_no.w_no2.Text & "' )"

        '����ގ��s
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
        'result = sqlcmd(SqlConn, "id = 'HE' AND ")
        'result = sqlcmd(SqlConn, "no1 = '" & form_no.w_no1.Text & "' AND ")
        'result = sqlcmd(SqlConn, "no2 = '" & form_no.w_no2.Text & "' )")
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)

        sqlcmd = "DELETE FROM " & DBTableName2 & " WHERE ( "
        sqlcmd = sqlcmd & "id = 'HE' AND "
        sqlcmd = sqlcmd & "no1 = '" & form_no.w_no1.Text & "' AND "
        sqlcmd = sqlcmd & "no2 = '" & form_no.w_no2.Text & "' )"

        '����ގ��s
        GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        If GL_T_RDO.Con.RowsAffected() = 0 Then
            ErrMsg = "Can not delete the existing data from the database (" & DBTableName2 & ")."
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        '�V�K�o�^
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            init_sql()

            w_str(1) = "'" & "HE" & "'" '�h�c(KO�Œ�)
            w_str(2) = "'" & form_no.w_no1.Text & "'" '�}�ʔԍ�
            w_str(3) = "'" & form_no.w_no2.Text & "'" '�ϔ�
            w_str(4) = i '�ҏW�����ԍ�
            w_str(5) = "'" & Trim(temp_hz.hm_name(i)) & "'" '�ҏW�����R�[�h


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

            '����ގ��s
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database (" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)

            end_sql()
        Next i


        hz_update = True

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

        hz_update = FAIL
    End Function
	
	
	Function hz_addnum() As Short
		Dim now_time As Object
		Dim j As Object
		Dim w_ret As Object
		Dim i As Object
        Dim result As Integer '20100707 �C��
		Dim DBTableNameHm As Object
        Dim w_str(100) As String

        ' -> watanabe del VerUP(2011)
        'Dim w_command As String
        ' <- watanabe del VerUP(2011)

        ' -> watanabe edit VerUP(2011)
        'Dim temp_hz2 As HZ_KANRI '���ƃe�[�u���Q�Ɨp
        Dim temp_hz2 As New HZ_KANRI '���ƃe�[�u���Q�Ɨp
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


        temp_hz2.Initilize() '20100707 �R�[�h�ǉ�
		
		' Brand Ver.3 �ύX
		'  DBTableNameHm = DBName & "..hm_kanri"
		DBTableNameHm = DBName & "..hm_kanri1"
		
		
		If SqlConn = 0 Then
            ' -> watanabe edit VerUP(2011)
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If

        '// �e�[�u������(����) --------------------------------------------------------------------
        temp_hz2.no1 = temp_hz.no1

        '----- .NET �ڍs -----
        'temp_hz2.no2 = VB6.Format(CDbl(Val(Trim(temp_hz.no2))) - 1, "00")
        temp_hz2.no2 = (CDbl(Val(Trim(temp_hz.no2))) - 1).ToString("00")


        ' -> watanabe edit VerUP(2011)
        '      'Brand Ver.3 �ύX
        '' result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, hm_num,")
        'result = SqlCmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date, hm_num ")
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 62
        ''    result = SqlCmd(SqlConn, " hm_name" & Format(i, "000") & ",")
        '' Next i
        '' result = SqlCmd(SqlConn, " hm_name" & Format(63, "000"))
        '
        'result = SqlCmd(SqlConn, " FROM " & DBTableName)
        'result = SqlCmd(SqlConn, " WHERE ( no1 = '" & temp_hz2.no1 & "' AND")
        'result = SqlCmd(SqlConn, " no2 = '" & temp_hz2.no2 & "' )")
        'result = SqlExec(SqlConn)
        'If result = FAIL Then GoTo error_section
        'result = SqlResults(SqlConn)
        'If result = SUCCEED Then
        '	'    If SqlNextRow(SqlConn) = REGROW Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		temp_hz2.comment = SqlData(SqlConn, 1)
        '		temp_hz2.dep_name = SqlData(SqlConn, 2)
        '		temp_hz2.entry_name = SqlData(SqlConn, 3)
        '		temp_hz2.entry_date = SqlData(SqlConn, 4)
        '		temp_hz2.hm_num = Val(SqlData(SqlConn, 5))
        '
        '		'Brand Ver.3 �ύX
        '		'       For i = 1 To 63
        '		'          temp_hz2.hm_name(i) = SqlData$(SqlConn, 5 + i)
        '		'       Next i
        '	Loop 
        'Else
        '	GoTo error_section
        'End If


        '�����L�[�Z�b�g
        key_code = "no1 = '" & temp_hz2.no1 & "' AND"
        key_code = key_code & " no2 = '" & temp_hz2.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date, hm_num FROM " & DBTableName & " WHERE ( " & key_code & " )"

        '�q�b�g���`�F�b�N
        cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        If cnt = -1 Then
            ErrMsg = "An error occurred on the existing record during the database search."
            ErrTtl = "Editing characters drawing update registration"
            GoTo error_section

        ElseIf cnt > 0 Then

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz2.comment = Rs.rdoColumns(0).Value
            Else
                temp_hz2.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_hz2.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_hz2.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_hz2.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_hz2.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_hz2.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_hz2.entry_date = ""
            End If

            If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                temp_hz2.hm_num = Rs.rdoColumns(4).Value
            Else
                temp_hz2.hm_num = 0
            End If

            Rs.Close()
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()


        'Brand Ver.3 �ǉ�
        For i = 1 To temp_hz2.hm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT hm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ("
            'w_command = w_command & " no1 = '" & temp_hz2.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_hz2.no2 & "' AND"
            'w_command = w_command & " hm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_hz2.hm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = "no1 = '" & temp_hz2.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_hz2.no2 & "' AND"
            key_code = key_code & " hm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT hm_name FROM " & DBTableName2 & " WHERE ( " & key_code & " )"

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName2, key_code)
            If cnt = 0 Or cnt = -1 Then
                Exit For
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz2.hm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_hz2.hm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        end_sql()
        init_sql()

        '// ���e�[�u���̌��n�����̕ҏW�����}�ʃf�[�^���N���A���� ----------------------------------------
        '�ҏW�����}�ʏ����폜����
        For i = 1 To temp_hz2.hm_num
            w_ret = update_hm_hz(DBTableNameHm, temp_hz2.hm_name(i), "  ", "    ", "  ")
            '1�ł����s����Α��̐}�ʃf�[�^���N���A����
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_hm_hz(DBTableNameHm, temp_hz2.hm_name(j), "KO", temp_hz2.no1, temp_hz2.no2)
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to Editing characters drawing information change Editing characters code [" & temp_hz2.hm_name(i) & "]"
                ErrTtl = "Editing characters drawing Change number registration error"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i


        '// �e�[�u���̌��n�����ɕҏW�����}�ʃf�[�^��o�^���� ----------------------------------------
        '�ҏW�����}�ʏ���ǉ�����
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "HE", form_no.w_no1.Text, form_no.w_no2.Text)
            '1�ł����s����Α��̕ҏW�����̕ҏW�����}�ʃf�[�^���N���A����
            If w_ret = -1 Then
                For j = 1 To i
                    w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(j), "  ", "    ", "  ")
                Next j

                ' -> watanabe add VerUP(2011)
                ErrMsg = "Failed to Editing characters drawing information change Editing characters code [" & temp_hz2.hm_name(i) & "]"
                ErrTtl = "Editing characters drawing Change number registration error"
                ' <- watanabe add VerUP(2011)

                GoTo error_section
            End If
        Next i


        '// �ҏW�����}�ʂ̓o�^ -------------------------------------------------------------------
        w_str(1) = "0" '�폜�t���O
        w_str(2) = "'" & "HE" & "'" '�h�c(HE�Œ�)
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

        w_str(9) = Trim(form_no.w_hm_num.Text) '�ҏW������

        'Brand Ver.3 �ύX
        ' For i = 1 To 63
        '     w_str(i + 9) = "'" & Trim(temp_hz.hm_name(i)) & "'"
        ' Next i

        '�ҏW�����}��̧�قɓo�^
        ' MsgBox "�ҏW�����}�ʂɓo�^���܂�", , "�m�F"


        ' -> watanabe edit VerUP(2011)
        'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName & " VALUES(")
        'w_command = "INSERT INTO " & DBTableName & " VALUES("
        '
        '' Brand Ver.3 �ύX
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
        '
        '
        ''�o�^�Ɏ��s�������͐}�ʏ����폜����
        'If result = FAIL Then
        '    For i = 1 To Val(Trim(form_no.w_hm_num.Text))
        '        w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "  ", "    ", "  ")
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
            For i = 1 To Val(Trim(form_no.w_hm_num.Text))
                w_ret = update_hm_hz(DBTableNameHm, temp_hz.hm_name(i), "  ", "    ", "  ")
            Next i
            ErrMsg = "Can not be registered in the database (" & DBTableName & ")"
            ErrTtl = "SQL error"
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


        end_sql()

        ' Brand Ver.3 �ύX
        For i = 1 To Val(Trim(form_no.w_hm_num.Text))
            init_sql()

            w_str(1) = "'" & "HE" & "'" '�h�c(HE�Œ�)
            w_str(2) = "'" & Trim(form_no.w_no1.Text) & "'" '�}�ʔԍ�
            w_str(3) = "'" & Trim(form_no.w_no2.Text) & "'" '�ϔ�
            w_str(4) = i '�ҏW�����ԍ�


            ' -> watanabe edit VerUP(2011)
            'result = sqlcmd(SqlConn, "INSERT INTO " & DBTableName2 & " VALUES(")
            'For j = 1 To 4
            '    result = sqlcmd(SqlConn, w_str(j) & ",")
            'Next j
            'result = sqlcmd(SqlConn, "'" & temp_hz.hm_name(i) & "'")
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
            sqlcmd = sqlcmd & "'" & temp_hz.hm_name(i) & "'"
            sqlcmd = sqlcmd & " )"

            '����ގ��s
            GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
            If GL_T_RDO.Con.RowsAffected() = 0 Then
                ErrMsg = "Can not be registered in the database (" & DBTableName2 & ")"
                ErrTtl = "SQL error"
                GoTo error_section
            End If
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i

        hz_addnum = True

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

        hz_addnum = FAIL
    End Function
	
	Function hz_search(ByRef hz_code1 As String, ByRef hz_code2 As String, ByRef flag As Short) As Short
		Dim i As Object
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
		
		'HZ_KANRI�e�[�u�����Y�����錴�n�����f�[�^�����߂�
		temp_hz.no1 = hz_code1
		temp_hz.no2 = hz_code2
		

        ' -> watanabe edit VerUP(2011)
        '      w_command = "SELECT comment, dep_name, entry_name, entry_date, hm_num"
        '
        ''Brand Ver.3 �ύX
        '' For i = 1 To 63
        ''    w_command = w_command & ", hm_name" & Format(i, "000")
        '' Next i
        '
        'w_command = w_command & " FROM " & DBTableName
        'If flag = 0 Then
        '	w_command = w_command & " WHERE (flag_delete = 0 AND no1 = '" & temp_hz.no1 & "' AND"
        'Else
        '	w_command = w_command & " WHERE (no1 = '" & temp_hz.no1 & "' AND"
        'End If
        'w_command = w_command & " no2 = '" & temp_hz.no2 & "')"
        '
        '
        'result = SqlCmd(SqlConn, w_command)
        '
        'result = SqlExec(SqlConn)
        'result = SqlResults(SqlConn)
        '
        'If result = SUCCEED Then
        '	'   Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '	If SqlNextRow(SqlConn) = REGROW Then
        '		temp_hz.comment = SqlData(SqlConn, 1)
        '		temp_hz.dep_name = SqlData(SqlConn, 2)
        '		temp_hz.entry_name = SqlData(SqlConn, 3)
        '		ww = SqlData(SqlConn, 4)
        '		w_ret = SqlDateCrack(SqlConn, df, ww)
        '		temp_hz.entry_date = df.Year_Renamed & df.Month_Renamed & df.Day_Renamed
        '		temp_hz.hm_num = Val(SqlData(SqlConn, 5))
        '
        '		'Brand Ver.3 �ύX
        '		'     For i = 1 To 63
        '		'       temp_hz.hm_name(i) = SqlData$(SqlConn, 5 + i)
        '		'     Next i
        '	Else
        '		GoTo error_section
        '	End If
        '
        '	'   Loop
        'Else
        '	GoTo error_section
        'End If


        '�����L�[�Z�b�g
        If flag = 0 Then
            key_code = "flag_delete = 0 AND no1 = '" & temp_hz.no1 & "' AND"
        Else
            key_code = "no1 = '" & temp_hz.no1 & "' AND"
        End If
        key_code = key_code & " no2 = '" & temp_hz.no2 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT comment, dep_name, entry_name, entry_date, hm_num FROM " & DBTableName & " WHERE (" & key_code & ")"

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
            temp_hz.comment = Rs.rdoColumns(0).Value
        Else
            temp_hz.comment = ""
        End If

        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
            temp_hz.dep_name = Rs.rdoColumns(1).Value
        Else
            temp_hz.dep_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(2).Value) = False Then
            temp_hz.entry_name = Rs.rdoColumns(2).Value
        Else
            temp_hz.entry_name = ""
        End If

        If IsDBNull(Rs.rdoColumns(3).Value) = False Then
            Dim tmpstr As String
            tmpstr = Rs.rdoColumns(3).Value
            temp_hz.entry_date = Left(tmpstr, 4) & Mid(tmpstr, 6, 2) & Mid(tmpstr, 9, 2)
        Else
            temp_hz.entry_date = ""
        End If

        If IsDBNull(Rs.rdoColumns(4).Value) = False Then
            temp_hz.hm_num = Val(Rs.rdoColumns(4).Value)
        Else
            temp_hz.hm_num = 0
        End If

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        end_sql()

        'Brand Ver.3 �ǉ�
        For i = 1 To temp_hz.hm_num
            init_sql()


            ' -> watanabe edit VerUP(2011)
            'w_command = "SELECT hm_name"
            'w_command = w_command & " FROM " & DBTableName2 & " WHERE ( "
            'w_command = w_command & " no1 = '" & temp_hz.no1 & "' AND"
            'w_command = w_command & " no2 = '" & temp_hz.no2 & "' AND"
            'w_command = w_command & " hm_no = " & i & " )"
            'result = sqlcmd(SqlConn, w_command)
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '    If SqlNextRow(SqlConn) = REGROW Then
            '        temp_hz.hm_name(i) = SqlData(SqlConn, 1)
            '    Else
            '        Exit For
            '    End If
            'Else
            '    Exit For
            'End If


            '�����L�[�Z�b�g
            key_code = "no1 = '" & temp_hz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_hz.no2 & "' AND"
            key_code = key_code & " hm_no = " & i

            '�����R�}���h�쐬
            sqlcmd = "SELECT hm_name FROM " & DBTableName2 & " WHERE ( " & key_code & " )"

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
                temp_hz.hm_name(i) = Rs.rdoColumns(0).Value
            Else
                temp_hz.hm_name(i) = ""
            End If

            Rs.Close()
            ' <- watanabe edit VerUP(2011)


            end_sql()
        Next i


        hz_search = True

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

        hz_search = FAIL
    End Function
	
	Function temp_hz_set(ByRef hexdata As String) As Short
		
		Dim aa As String
		Dim i As Short

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)


        '========================================
		'�ҏW�����}�ʃf�[�^���g�d�w���ϊ����܂�
		'========================================
		
		temp_hz.hm_num = 0

        ' -> watanabe edit 2013.05.29
        'For i = 1 To 63
        For i = 1 To 130
            ' <- watanabe edit 2013.05.29

            temp_hz.hm_name(i) = ""
        Next i

        If open_mode = "NEW" Then
            temp_hz.hm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_hz.hm_num
                temp_hz.hm_name(i) = Mid(hexdata, (i - 1) * 8 + 4, 8)
            Next i
            temp_hz.id = "HE"
            temp_hz.no1 = ""
            temp_hz.no2 = "00"
            temp_hz.comment = ""
            temp_hz.dep_name = ""
            temp_hz.entry_name = ""
            Call true_date(aa)
            temp_hz.entry_date = aa
        ElseIf open_mode = "Revision number" Then
            temp_hz.id = "HE"
            temp_hz.hm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_hz.hm_num
                temp_hz.hm_name(i) = Mid(hexdata, (i - 1) * 8 + 4, 8)
            Next i
            Call true_date(aa)
            temp_hz.entry_date = aa
        ElseIf open_mode = "modify" Then
            temp_hz.id = "HE"
            temp_hz.hm_num = Val(Mid(hexdata, 1, 3))
            For i = 1 To temp_hz.hm_num
                temp_hz.hm_name(i) = Mid(hexdata, (i - 1) * 8 + 4, 8)
            Next i
            Call true_date(aa)
            temp_hz.entry_date = aa
        End If

    End Function
	
	
	Function hz_delete(ByRef hz_code1 As String, ByRef hz_code2 As String) As Short
        Dim result As Integer '20100707 �C��
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
            'MsgBox("�ް��ް��ɱ����o���܂���", MsgBoxStyle.Critical, "SQL�װ")
            ErrMsg = "Can not access the database."
            ErrTtl = "SQL error"
            ' <- watanabe edit VerUP(2011)
            GoTo error_section
		End If

        w_str(1) = "1" '�폜�t���O
		w_str(2) = "'" & "HE" & "'" '�h�c(HE�Œ�)
		w_str(3) = "'" & hz_code1 & "'" '�}�ʔԍ�(****)
		w_str(4) = "'" & hz_code2 & "'" '�ϔ�(00~99�j
		' w_str(5) = "'" & form_no.w_comment.Text & "'"                  '�R�����g
		' w_str(6) = "'" & form_no.w_dep_name.Text & "'"                 '�����R�[�h
		' w_str(7) = "'" & form_no.w_entry_name.Text & "'"               '�o�^��
		' w_str(8) = "'" & form_no.w_entry_date.Text & "'"               '�o�^��
		' w_str(9) = form_no.w_hm_num.Text                               '���n������
		

        ' -> watanabe edit VerUP(2011)
        '      result = sqlcmd(SqlConn, "UPDATE " & DBTableName)
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


        hz_delete = True
		
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

        hz_delete = FAIL
    End Function

	Function zumen_no_set_hz(ByRef hexdata As String) As Short
		Dim t4 As String
		Dim t3 As String
		Dim t2 As String
		Dim t1 As String
		Dim result As Object
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim nn As Short

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


        If open_mode = "modify" Then
            temp_hz.id = "HE"
            temp_hz.no1 = Mid(hexdata, 4, 4)
            temp_hz.no2 = Mid(hexdata, 9, 2)


            ' -> watanabe add VerUP(2011)
            '         result = sqlcmd(SqlConn, "SELECT comment, dep_name, entry_name, entry_date ")
            'result = SqlCmd(SqlConn, " FROM " & DBTableName)
            'result = SqlCmd(SqlConn, " WHERE ")
            'result = SqlCmd(SqlConn, " flag_delete = 0 AND")
            'result = SqlCmd(SqlConn, " id = 'HE' AND")
            'result = SqlCmd(SqlConn, " no1 = '" & temp_hz.no1 & "' AND")
            'result = SqlCmd(SqlConn, " no2 = '" & temp_hz.no2 & "'")
            'result = SqlExec(SqlConn)
            'If result = FAIL Then GoTo error_section
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
            '		temp_hz.comment = SqlData(SqlConn, 1)
            '		temp_hz.dep_name = SqlData(SqlConn, 2)
            '		temp_hz.entry_name = SqlData(SqlConn, 3)
            '		temp_hz.entry_date = SqlData(SqlConn, 4)
            '	Loop 
            '	If temp_hz.entry_name = "" Then
            '		MsgBox("�ҏW�����}�ʃf�[�^������܂���" & Chr(13) & "�C�������͏o���܂���", MsgBoxStyle.Critical, "�װ")
            '		GoTo error_section
            '	End If
            'Else
            '	GoTo error_section
            'End If


            '�����L�[�Z�b�g
            key_code = " flag_delete = 0 AND"
            key_code = key_code & " id = 'HE' AND"
            key_code = key_code & " no1 = '" & temp_hz.no1 & "' AND"
            key_code = key_code & " no2 = '" & temp_hz.no2 & "'"

            '�����R�}���h�쐬
            sqlcmd = "SELECT comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE " & key_code

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = 0 Then
                MsgBox("There is no Editing characters drawing data." & Chr(13) & "Can not modify processing.", MsgBoxStyle.Critical, "error")
                errflg = 1
                GoTo error_section
            ElseIf cnt = -1 Then
                MsgBox("An error occurred on the existing record during the database search.", MsgBoxStyle.Critical, "error")
                errflg = 1
                GoTo error_section
            End If

            '����
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            Rs.MoveFirst()

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                temp_hz.comment = Rs.rdoColumns(0).Value
            Else
                temp_hz.comment = ""
            End If

            If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                temp_hz.dep_name = Rs.rdoColumns(1).Value
            Else
                temp_hz.dep_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                temp_hz.entry_name = Rs.rdoColumns(2).Value
            Else
                temp_hz.entry_name = ""
            End If

            If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                temp_hz.entry_date = Rs.rdoColumns(3).Value
            Else
                temp_hz.entry_date = ""
            End If

            Rs.Close()
            ' <- watanabe add VerUP(2011)


        ElseIf open_mode = "Revision number" Then
            temp_hz.id = "HE"
            temp_hz.no1 = Mid(hexdata, 4, 4)
            temp_hz.no2 = Mid(hexdata, 9, 2)

            '�����L�[�Z�b�g
            key_code = " id = 'HE' AND"
            key_code = key_code & " no1 = '" & temp_hz.no1 & "'"

            '�����R�}���h�쐬
            sqlcmd = "SELECT no2, comment, dep_name, entry_name, entry_date FROM " & DBTableName & " WHERE " & key_code

            '�q�b�g���`�F�b�N
            cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
            If cnt = 0 Then
                MsgBox("There is no Editing characters drawing data." & Chr(13) & "It is not possible to revision number processing.", MsgBoxStyle.Critical, "�װ")
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

            temp_hz.no2 = "-1"

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

                If Val(temp_hz.no2) < nn Then
                    '----- .NET �ڍs -----
                    'temp_hz.no2 = VB6.Format(nn, "00")
                    temp_hz.no2 = nn.ToString("00")

                    temp_hz.comment = t1
                    temp_hz.dep_name = t2
                    temp_hz.entry_name = t3
                    temp_hz.entry_date = t4
                End If

                Rs.MoveNext()
            Loop

            If Val(temp_hz.no2) < 0 Then
                MsgBox("There is no Editing characters drawing data." & Chr(13) & "It is not possible to revision number processing.", MsgBoxStyle.Critical, "�װ")
                errflg = 1
                GoTo error_section
            Else
                '----- .NET �ڍs -----
                'temp_hz.no2 = VB6.Format(Val(temp_hz.no2) + 1, "00")
                temp_hz.no2 = (Val(temp_hz.no2) + 1).ToString("00")
            End If

            Rs.Close()
            ' <- watanabe add VerUP(2011)


        End If

        zumen_no_set_hz = True
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

        zumen_no_set_hz = FAIL
    End Function
End Module