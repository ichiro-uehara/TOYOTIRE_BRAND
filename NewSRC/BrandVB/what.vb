Option Strict Off
Option Explicit On

Imports System.Collections.Generic

Module MJ_What

    '�E�C���h�E�Y�̃v���O�������Ăяo�����߂̐錾
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer


    '�g�p���Ă��Ȃ����ߑS�ăR�����g
    'Function what_pic_GM(id As String, font_name As String) As Integer
    'GM_KANRI�e�[�u�����V�K�̔z�uPIC�����߂�
    'ID���t�H���g�����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�Ȕz�uPIC�̒l��1���������̂�V�K�z�uPIC�̒l�Ƃ���
    '1�`63�𒴂���ƃG���[�Ƃ���
    ' Dim pic_no As Integer
    ' Dim w_pic(0 To 256) As Byte
    ' Dim i As Integer
    '
    ' w_pic(0) = 1
    ' For i = 1 To 63
    '   w_pic(i) = 0
    ' Next i
    '
    ' pic_no = -1
    ' result = SqlCmd(SqlConn, "SELECT haiti_pic")
    ' result = SqlCmd(SqlConn, " FROM yama..gm_kanri")
    ' result = SqlCmd(SqlConn, " WHERE (id = '" & id & "' AND")
    ' result = SqlCmd(SqlConn, " font_name = '" & font_name & "')")
    '
    ' result = SqlExec(SqlConn)
    ' result = SqlResults(SqlConn)
    ' If result = SUCCEED Then
    '   Do Until SqlNextRow(SqlConn) = NOMOREROWS
    '     w_pic(Val(SqlData$(SqlConn, 1))) = 1
    '   Loop
    ' Else
    '   'Debug.Print "SqlResults FAIL....."
    '   GoTo error_section
    ' End If
    ' For i = 0 To 62
    '   If (w_pic(i) = 1) Then
    '      pic_no = i + 1
    '   End If
    ' Next i
    '
    ' If pic_no = -1 Then
    '    MsgBox "�z�uPIC���o�^�o���܂���", 64, "Out of Range"
    '    GoTo error_section
    ' End If
    '
    'what_pic_GM = pic_no
    ''MsgBox "�z�uPIC��" & pic_no
    'Exit Function
    '
    'error_section:
    ' what_pic_GM = -1
    '
    'End Function

    Function what_font_class2_GM(ByRef font_name As String, ByRef class1 As String, ByRef name1 As String, ByRef name2 As String) As Short
        Dim L_DAT1 As String
		Dim kubun As Short
		Dim result As Integer

        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(�R�����g��) -----
        'Dim Rs As RDO.rdoResultset

        On Error GoTo error_section
        Err.Clear()

        'GM_KANRI�e�[�u�����V�K�̃t�H���g�敪�����߂�
        '�t�H���g�����t�H���g�敪1��������1���������Q���}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȃt�H���g�敪�̒l��1���������̂�V�K�t�H���g�敪�̒l�Ƃ���
        '�͈́i�O�`�X�j�𒴂���ƃG���[�Ƃ���
        kubun = 0


        '�����L�[�Z�b�g
        key_code = "font_name = '" & font_name & "' AND"
        key_code = key_code & " font_class1 = '" & class1 & "' AND"
        key_code = key_code & " name1 = '" & name1 & "' AND"
        key_code = key_code & " name2 = '" & name2 & "'"

        '----- .NET �ڍs -----
        ''�����R�}���h�쐬
        'sqlcmd = "SELECT font_class2 FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    kubun = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    kubun = 0
        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If kubun <= Val(L_DAT1) Then
        '            kubun = Val(L_DAT1) + 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        '---------------------------------------------------------------

        cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

        If cnt = 0 Then
            kubun = 0
        ElseIf cnt = -1 Then
            GoTo error_section
        Else
            '����
            Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
            Dim dataList As List(Of List(Of String)) = New List(Of List(Of String))
            Dim param As ADO_PARAM_Struct

            param.DataSize = 0
            param.Value = Nothing
            param.Sign = ""

            param.ColumnName = "font_class2"
            param.SqlDbType = SqlDbType.Char
            paramList.Add(param)

            If VBADO_Search(GL_T_ADO, DBTableName, key_code, paramList, dataList) <> 1 Then
                MsgBox("Failed to find table record.")
                GoTo error_section
            End If

            kubun = 0
            For Each recordList As List(Of String) In dataList
                L_DAT1 = recordList(0)
                If L_DAT1 = " " Or L_DAT1 = "" Then
                    L_DAT1 = "0"
                End If

                If kubun <= Val(L_DAT1) Then
                    kubun = Val(L_DAT1) + 1
                End If

            Next

        End If

        '----- .NET �ڍs -----

        If kubun > 9 Then
            MsgBox("Can not register a number of 10 or more." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_font_class2_GM = kubun
        Exit Function

error_section:
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(�R�����g��) -----
        'Rs.Close()

        what_font_class2_GM = -1
    End Function

    Function what_no2_GZ(ByRef gz_code1 As String) As Short
        Dim L_DAT1 As String
        Dim kubun As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        'GZ_KANRI�e�[�u�����V�K�̕ϔԂ����߂�
        '�}�ʔԍ����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȕϔԂ̒l��1���������̂�V�K�ϔԂ̒l�Ƃ���
        '�͈́i�O�`�X�X�j�𒴂���ƃG���[�Ƃ���
        kubun = 0



        '�����L�[�Z�b�g
        key_code = "no1 = '" & gz_code1 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT no2 FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    kubun = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    kubun = 0
        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If kubun <= Val(L_DAT1) Then
        '            kubun = Val(L_DAT1) + 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)


        If kubun > 99 Then
            MsgBox("Can not register a number of 100 or more." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_no2_GZ = kubun
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_no2_GZ = -1
    End Function

    Function what_no2_BZ(ByRef bz_code1 As String) As Short
        Dim L_DAT1 As String
        Dim kubun As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        'BZ_KANRI�e�[�u�����V�K�̕ϔԂ����߂�
        '�}�ʔԍ����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȕϔԂ̒l��1���������̂�V�K�ϔԂ̒l�Ƃ���
        '�͈́i�O�`�X�X�j�𒴂���ƃG���[�Ƃ���
        kubun = 0



        '�����L�[�Z�b�g
        key_code = "no1 = '" & bz_code1 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT no2 FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    kubun = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    kubun = 0
        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If kubun <= Val(L_DAT1) Then
        '            kubun = Val(L_DAT1) + 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)


        If kubun > 99 Then
            MsgBox("Can not register a number of 100 or more." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_no2_BZ = kubun
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_no2_BZ = -1
    End Function

    Function what_no2_HZ(ByRef hz_code1 As String) As Short
        Dim L_DAT1 As String
        Dim kubun As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        'HZ_KANRI�e�[�u�����V�K�̕ϔԂ����߂�
        '�}�ʔԍ����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȕϔԂ̒l��1���������̂�V�K�ϔԂ̒l�Ƃ���
        '�͈́i�O�`�X�X�j�𒴂���ƃG���[�Ƃ���
        kubun = 0


        '�����L�[�Z�b�g
        key_code = "no1 = '" & hz_code1 & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT no2 FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    kubun = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    kubun = 0
        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If kubun <= Val(L_DAT1) Then
        '            kubun = Val(L_DAT1) + 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)


        If kubun > 99 Then
            MsgBox("Can not register a number of 100 or more." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_no2_HZ = kubun
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_no2_HZ = -1
    End Function

    Function what_no_HM(ByRef font_name As String) As Short
        Dim L_DAT1 As String
        Dim kubun As Short
        Dim result As Integer

        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(�R�����g��) -----
        'Dim Rs As RDO.rdoResultset

        On Error GoTo error_section
        Err.Clear()

        'hm_kanri1 �e�[�u�����V�K�̃t�H���g�敪�����߂�
        '�t�H���g�����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȃt�H���g�敪�̒l��1���������̂�V�K�t�H���g�敪�̒l�Ƃ���
        '�͈́i�O�`�X�X�j�𒴂���ƃG���[�Ƃ���
        kubun = 0


        '�����L�[�Z�b�g
        key_code = "font_name = '" & font_name & "'"

        '----- .NET �ڍs -----
        ''�����R�}���h�쐬
        'sqlcmd = "SELECT no FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    kubun = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    kubun = 0
        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If kubun <= Val(L_DAT1) Then
        '            kubun = Val(L_DAT1) + 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        '---------------------------------------------------------------

        cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

        If cnt = 0 Then
            kubun = 0
        ElseIf cnt = -1 Then
            GoTo error_section
        Else
            '����
            Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
            Dim dataList As List(Of List(Of String)) = New List(Of List(Of String))
            Dim param As ADO_PARAM_Struct

            param.DataSize = 0
            param.Value = Nothing
            param.Sign = ""

            param.ColumnName = "no"
            param.SqlDbType = SqlDbType.Char
            paramList.Add(param)

            If VBADO_Search(GL_T_ADO, DBTableName, key_code, paramList, dataList) <> 1 Then
                MsgBox("Failed to find table record.")
                GoTo error_section
            End If

            kubun = 0
            For Each recordList As List(Of String) In dataList
                L_DAT1 = recordList(0)
                If L_DAT1 = " " Or L_DAT1 = "" Then
                    L_DAT1 = "0"
                End If

                If kubun <= Val(L_DAT1) Then
                    kubun = Val(L_DAT1) + 1
                End If
            Next

        End If

        '----- .NET �ڍs -----

        If kubun > 99 Then
            MsgBox("Can not register a number of 100 or more." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_no_HM = kubun
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_no_HM = -1

    End Function

    Function what_name2_GM(ByRef font_name As String, ByRef class1 As String, ByRef name1 As String) As Short
        Dim L_DAT1 As String
        Dim s_name As Short
        Dim kubun As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)

        'GM_KANRI�e�[�u�����V�K�̕�����2�����߂�
        '�t�H���g�����t�H���g�敪1��������1���}�b�`���O���ē���̃f�[�^�̈�ԑ傫�ȃt�H���g�敪�̒l��1���������̂�V�K�t�H���g�敪�̒l�Ƃ���
        '�͈́i1�`�X,A�`Z�j�𒴂���ƃG���[�Ƃ���
        kubun = 0





        '�����L�[�Z�b�g
        key_code = "font_name = '" & font_name & "' AND"
        key_code = key_code & " font_class1 = '" & class1 & "' AND"
        key_code = key_code & " name1 = '" & name1 & "' AND"

        '�����R�}���h�쐬
        sqlcmd = "SELECT font_name2 FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = 0 Then
        '    s_name = CShort("0")

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    s_name = CShort("0")
        '    Do Until Rs.EOF
        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            L_DAT1 = Rs.rdoColumns(0).Value
        '        Else
        '            L_DAT1 = "0"
        '        End If

        '        If s_name <= CDbl(L_DAT1) Then
        '            If InStr("123456789", Left(L_DAT1, 1)) = 0 Then
        '                s_name = CShort(Str(CDbl(L_DAT1) + 1))
        '            ElseIf InStr("9", Left(L_DAT1, 1)) = 0 Then
        '                s_name = CShort("A")
        '            ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXY", Left(L_DAT1, 1)) = 0 Then
        '                's_name = CShort(Chr(Asc(CStr(CDbl(Left(L_DAT1, 1)) + 1))))
        '                s_name = CShort(Val(Chr(Asc(CStr(CDbl(Left(L_DAT1, 1)) + 1))))) '20100617�ڐA�ǉ�
        '            End If
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' -> watanabe edit VerUP(2011)


        If s_name > CDbl("Z") Then
            MsgBox("Value of 1 ~ 9, A ~ Z or more can not be registered to the character names 2." & "It was not possible to auto-numbering." & Chr(13), 64, "Out of Range")
            GoTo error_section
        End If

        what_name2_GM = kubun
        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_name2_GM = -1
    End Function

    Function what_pic_no(ByRef id As String, ByRef font_name As String) As Short

        'HM_KANRI�e�[�u�����V�K�̔z�uPIC�����߂�
        'ID���t�H���g�����}�b�`���O���ē���̃f�[�^�̈�ԑ傫�Ȕz�uPIC�̒l��1���������̂�V�K�z�uPIC�̒l�Ƃ���
        '1�`63�𒴂���ƃG���[�Ƃ���

        Dim pic_no As Short
        Dim w_pic(256) As Byte
        Dim i As Short
        Dim result As Integer
        Dim FileName As String
        '----- .NET �ڍs -----
        'Dim key_value As New VB6.FixedLengthString(255)
        Dim key_value As System.Text.StringBuilder = New System.Text.StringBuilder(256)

        Dim work As Short
        Dim max_pic As Short

        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(�R�����g��) -----
        'Dim Rs As RDO.rdoResultset

        FileName = Environ("ACAD_SET")
        FileName = FileName & "BR_DRAWCONF.ini"

        If id = "GM" Then
            work = GetPrivateProfileString("GENSI", "GM_PIC_SAVE_END", "", key_value, key_value.Capacity - 1, FileName)
        Else
            work = GetPrivateProfileString("HENSYU", "HE_PIC_SAVE_END", "", key_value, key_value.Capacity - 1, FileName)
        End If

        '----- .NET �ڍs -----
        'key_value.Value = Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)
        'max_pic = Val(Left(key_value.Value, InStr(1, key_value.Value, ";", 0) - 1))
        max_pic = Val(Left(key_value.ToString(), InStr(1, key_value.ToString(), ";", 0) - 1))


        w_pic(0) = 1
        For i = 1 To max_pic
            w_pic(i) = 0
        Next i

        pic_no = -1



        '�����L�[�Z�b�g
        key_code = "id = '" & id & "' AND"
        key_code = key_code & " font_name = '" & font_name & "' AND"
        key_code = key_code & " flag_delete = 0 "

        '----- .NET �ڍs -----
        '�����R�}���h�쐬
        'sqlcmd = "SELECT haiti_pic FROM " & DBTableName & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        'cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        'If cnt = -1 Then
        '    GoTo error_section

        'ElseIf cnt > 0 Then
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            Dim aryno As Integer
        '            aryno = Val(Rs.rdoColumns(0).Value)
        '            w_pic(aryno) = 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' -> watanabe edit VerUP(2011)
        '-------------------------------------------------------

        '�e�[�u�����R�[�h���`�F�b�N
        cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)
        If cnt = -1 Then
            GoTo error_section
        ElseIf cnt > 0 Then
            Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
            Dim dataList As List(Of List(Of String)) = New List(Of List(Of String))
            Dim param As ADO_PARAM_Struct
            param.DataSize = 0
            param.Value = Nothing
            param.Sign = ""
            param.ColumnName = "haiti_pic"
            param.SqlDbType = SqlDbType.TinyInt
            paramList.Add(param)

            If VBADO_Search(GL_T_ADO, DBTableName, key_code, paramList, dataList) <> 1 Then
                MsgBox("Failed to find table record.")
                GoTo error_section
            End If

            For Each recordList As List(Of String) In dataList
                Dim aryno As Integer = Val(recordList(0))
                w_pic(aryno) = 1
            Next

        End If

        '----- .NET �ڍs -----

        For i = 0 To max_pic
            If (w_pic(i) = 1) Then
                pic_no = i + 1
            Else
                Exit For
            End If
        Next i

        If pic_no = -1 Or pic_no > max_pic Then
            MsgBox("Can not register the Placement picture.", 64, "Out of Range")
            GoTo error_section
        End If

        what_pic_no = pic_no

        Exit Function

error_section:
        On Error Resume Next
        Err.Clear()

        '----- .NET �ڍs -----
        'Rs.Close()

        what_pic_no = -1
    End Function
	
	Function what_pic_from_hmcode(ByRef hm_code As String) As Short
		
		'HM_KANRI�e�[�u�����R�[�h�Ō������Ĕz�uPIC�����߂�

        Dim w_table_Name As String
		Dim pic_no As Short
		Dim result As Integer
		
        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        'Brand Ver.3 �ύX
        ' w_table_Name = DBName & "..hm_kanri"
        w_table_Name = DBName & "..hm_kanri1"
		
		pic_no = -1





        '�����L�[�Z�b�g
        key_code = " font_name = '" & Mid(hm_code, 1, 6) & "' AND"
        key_code = key_code & " no = '" & Mid(hm_code, 7, 2) & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT haiti_pic FROM " & w_table_Name & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, w_table_Name, key_code)
        'If cnt = 0 Then
        '    pic_no = -1

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            pic_no = Val(Rs.rdoColumns(0).Value)
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)


        what_pic_from_hmcode = pic_no
        Exit Function
		
error_section: 
        MsgBox("There is no  editing characters code drawing corresponding. " & hm_code, MsgBoxStyle.Critical, "editing characters code error")

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_pic_from_hmcode = -1
    End Function
	
    Function what_pic_from_gmcode(ByRef gm_code As String) As Short

        'GM_KANRI�e�[�u�����R�[�h�Ō������Ĕz�uPIC�����߂�

        Dim w_table_Name As String
        Dim pic_no As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET �ڍs(��U�R�����g��) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        w_table_Name = DBName & "..gm_kanri"

        pic_no = -1



        '�����L�[�Z�b�g
        key_code = " font_name = '" & Mid(gm_code, 1, 6) & "' AND"
        key_code = key_code & " font_class1 = '" & Mid(gm_code, 7, 1) & "' AND"
        key_code = key_code & " font_class2 = '" & Mid(gm_code, 8, 1) & "' AND"
        key_code = key_code & " name1 = '" & Mid(gm_code, 9, 1) & "' AND"
        key_code = key_code & " name2 = '" & Mid(gm_code, 10, 1) & "'"

        '�����R�}���h�쐬
        sqlcmd = "SELECT haiti_pic FROM " & w_table_Name & " WHERE (" & key_code & ")"

        '�q�b�g���`�F�b�N
        '----- .NET �ڍs(��U�R�����g��) -----
        'cnt = VBRDO_Count(GL_T_RDO, w_table_Name, key_code)
        'If cnt = 0 Then
        '    pic_no = -1

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '����
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            pic_no = Val(Rs.rdoColumns(0).Value)
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)

        what_pic_from_gmcode = pic_no
        Exit Function

error_section:
        MsgBox("There is no Primitive character code corresponding. " & gm_code, MsgBoxStyle.Critical, "Primitive character code error")

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET �ڍs(��U�R�����g��) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        what_pic_from_gmcode = -1
    End Function
End Module