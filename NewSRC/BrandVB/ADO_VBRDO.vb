Option Strict Off
Option Explicit On
Imports System.Data.SqlClient '2011/7/25 moriya add
Module ADO_VBRDO
    '/***************************************************************************
    '����̧�ق� �ȉ��̼��тŋ��L���Ă��܂��B
    '�ύX����ۂɂ́A�䒍�ӊ肢�܂��B
    '
    '                          Comment�F1998 8/25     by f.yamamoto
    '
    '       �̔Լ���
    '       �b�\
    '       �L���A�h���C�A�E�g
    '       ����ށE�ް���ݸ�            1998 11-12 add by yamamoto
    '       ����ށEӰ������̧��         1998 11-12 add by yamamoto
    '
    '
    '   [ Notes ]
    '       ����̧�ق�ݸٰ�ނ���ɂͤ�ȉ���̧�ق��K�v�ł��
    '           �Emsrdo20.dll (Microsoft Remote Data Object 2.0)
    '           �i�Q�Ɛݒ�Ųݸٰ�ނł��܂���j
    '
    '   [ Contents ]
    '       Declarations
    '       VBRDO_Init         --- RDO�ڑ��p�ݒ�̧�ٓǍ�
    '       VBRDO_OpenEnv      --- RDO���ݒ�
    '       VBRDO_Count        --- DB�w�����ں��ސ�����
    '       VBRDO_Delete       --- DB�w�����ں��ލ폜
    '       VBRDO_Connect      --- RDO�ڑ��i�ȸ��݂��J���j
    '       VBRDO_Discon       --- RDO�ڑ��ؒf�i�ȸ��݂�ؒf����j
    '       VBRDO_RDORegistry  --- RDO�ڑ��pDSN�쐬
    '       VBRDO_CloseEnv     --- RDO�����
    '       VBRDO_T_RDOInit    --- RDO�ڑ��p�\���̏�����

    '***************************************************************************/ ' ���ׂĂ̕ϐ��𖾎��I�ɐ錾����悤�ɂ��܂��B

    '2011/7/25 moriya add start
    Dim dataada As New SqlDataAdapter()
    Dim dataset As New DataSet()
    Dim cnn As New SqlConnection()
    Dim cmnd As New SqlCommand()
    Public Const CUAD As Integer = 0
    Public Const CPMASTER As Integer = 1
    Public Const DESIGN_STANDERD As Integer = 2
    Public Const NUMBER As Integer = 3
    Public Const PROFILE As Integer = 4
    '2011/7/25 moriya add end

    '2012/5/29 moriya udpate start
    ''test moriya start  �S�Ă̋@�\��ADO�ɐ؂�ւ�����Ƃ���"_ADO"���폜���邱��
    'Public Const DEF_MSG_E9000_ADO As String = "�ݒ�t�@�C���ǂݍ��݃G���[�ł��B "
    ''Public Const DEF_MSG_E9001 As String = "DSN�쐬���ɃG���[�������܂����B"
    'Public Const DEF_MSG_E9002_ADO As String = "ADO�ڑ��G���[�ł��B"
    'Public Const DEF_MSG_E9003_ADO As String = "�c�a�o�^�������ɃG���[�������܂����B"
    'Public Const DEF_MSG_E9004_ADO As String = "�c�a���R�[�h�폜�������ɃG���[�������܂����B"

    'test moriya start  �S�Ă̋@�\��ADO�ɐ؂�ւ�����Ƃ���"_ADO"���폜���邱��
    Public Const DEF_MSG_E9000_ADO As String = "It occures error reading setting file."
    Public Const DEF_MSG_E9002_ADO As String = "ADO connection error."
    Public Const DEF_MSG_E9003_ADO As String = "Error occurred saving DB."
    Public Const DEF_MSG_E9004_ADO As String = "Error occurred deleting DB record."
    '2012/5/29 moriya udpate end

    '�ڑ��p�ϐ�
    'Structure T_RDO_Struct
    '    Dim DSN As String 'RDO�ڑ��p�ް������
    '    Dim UID As String 'RDO�ڑ��pհ��ID
    '    Dim PWD As String 'RDO�ڑ��p�߽ܰ��
    '    Dim DBName As String 'RDO�ڑ��p�ް��ް���
    '    Dim Server As String 'RDO�ڑ��p����-
    '    Dim Con As RDO.rdoConnection 'RDO�ڑ��p��޼ު��
    'End Structure

    '2011/8/4 moriya add start
    Structure T_ADO_Struct
        Dim DSN As String 'ADO�ڑ��p�ް������
        Dim UID As String 'ADO�ڑ��pհ��ID
        Dim PWD As String 'ADO�ڑ��p�߽ܰ��
        Dim DBName As String 'ADO�ڑ��p�ް��ް���
        Dim Server As String 'ADO�ڑ��p����-
    End Structure
    '2011/8/4 moriya add end



    'WINAPI  �ݒ�̧�ٓǍ��ݗp�֐��Ŏg�p
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer


    '�T�v  �F�ݒ�̧�ق��RDO�ݒ�l���擾
    '���Ұ��FT_RDO,  I/O,T_RDO_Struct,RDO�ڑ��p�\����
    '      �Fsection,I,  String,      ����ݖ�
    '      �Ffname,  I,  String,      ̧�ٖ�
    '      �Fsw,     I,  Long,        ����
    '      �F�߂�l, O,Long,          ���� OK or NG
    '����  �F�ݒ�̧�ق��RDO�ݒ�l���擾
    '2011/8/4 moriya update start
    '����  �FRDO��ADO�ɕύX
    'Function VBRDO_Init(ByRef T_RDO As T_RDO_Struct, ByRef lpszSection As String, ByRef fname As String, ByRef sw As Integer) As Integer
    Function VBADO_Init(ByRef T_ADO As T_ADO_Struct, ByRef lpszSection As String, ByRef fname As String, ByRef sw As Integer) As Integer
        '2011/8/4 moriya update end

        Dim IRet As Integer
        Dim str_Renamed As New VB6.FixedLengthString(256)
        Dim Cmdstr(4) As String
        Dim iCnt As Integer
        Dim posi As Integer
        Dim wkfname As String
        Dim AcadDir As String

        '2011/8/4 moriya update start
        'VBRDO_Init = 0
        VBADO_Init = 0
        '2011/8/4 moriya update end

        wkfname = fname

        '==================================================
        '           ̧�يi�[�ިڸ�ؐݒ�     Start
        '==================================================
        If sw = 0 Then

            AcadDir = New String(Chr(0), 255)
            IRet = GetEnvironmentVariable("ACAD_SET", AcadDir, Len(AcadDir))
            If IRet = 0 Then
                Exit Function
            End If

            AcadDir = Left(AcadDir, InStr(1, AcadDir, Chr(0), CompareMethod.Binary) - 1)
            If Right(AcadDir, 1) <> "\" Then
                AcadDir = AcadDir & "\"
            End If
            wkfname = Trim(AcadDir) & wkfname

        End If

        '==================================================
        '           ̧�يi�[�ިڸ�ؐݒ�     End
        '==================================================


        '==================================================
        '           �ݒ�l�擾     Start
        '==================================================

        Cmdstr(0) = "DSN" 'DSN
        Cmdstr(1) = "UID" '���[�U�[ID
        Cmdstr(2) = "DBName" '�f�[�^�x�[�X��
        Cmdstr(3) = "PWD" '�p�X���[�h
        Cmdstr(4) = "Server" '�T�[�o�[


        For iCnt = 0 To 4

            '�w�辸��݂̎w�跰ܰ�ނ̒l���擾
            IRet = GetPrivateProfileString(lpszSection, Cmdstr(iCnt), "ERROR", str_Renamed.Value, 256, wkfname)

            If IRet <> 0 Then

                If InStr(1, str_Renamed.Value, "ERROR", CompareMethod.Binary) > 0 Then
                    Exit Function
                End If

                '�к�݂���菜���A��߰��දĂ���
                posi = InStr(1, str_Renamed.Value, ";", CompareMethod.Binary)
                If posi <> 0 Then
                    str_Renamed.Value = Trim(Left(str_Renamed.Value, posi - 1))
                Else
                    str_Renamed.Value = Trim(Mid(str_Renamed.Value, 1, InStr(1, str_Renamed.Value, Chr(0), CompareMethod.Binary) - 1))
                End If

                Select Case iCnt
                    '2011/8/4 moriya update start
                    'Case 0 : T_RDO.DSN = Trim(str_Renamed.Value)
                    'Case 1 : T_RDO.UID = Trim(str_Renamed.Value)
                    'Case 2 : T_RDO.DBName = Trim(str_Renamed.Value)
                    'Case 3 : T_RDO.PWD = Trim(str_Renamed.Value)
                    'Case 4 : T_RDO.Server = Trim(str_Renamed.Value)
                    Case 0 : T_ADO.DSN = Trim(str_Renamed.Value)
                    Case 1 : T_ADO.UID = Trim(str_Renamed.Value)
                    Case 2 : T_ADO.DBName = Trim(str_Renamed.Value)
                    Case 3 : T_ADO.PWD = Trim(str_Renamed.Value)
                    Case 4 : T_ADO.Server = Trim(str_Renamed.Value)

                        '2011/8/4 moriya update end
                End Select

            Else
                Exit Function
            End If

        Next

        '==================================================
        '           �ݒ�l�擾     Start
        '==================================================

        '2011/8/8 moriya update start
        'VBRDO_Init = 1
        VBADO_Init = 1
        '2011/8/8 moriya update end

    End Function

    '2011/8/2 moriya delete start
    '�T�v  �FRDO�ڑ��p���ϐ����
    '���Ұ��FRDOEnv,I/O,rdoEnvironment,RDO�ڑ��p���ϐ�
    '      �F�߂�l, O,Long,            -------
    '����  �FRDO�ڑ��p���ϐ����
    'Function VBRDO_CloseEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer

    '    On Error Resume Next

    '    VBRDO_CloseEnv = 1

    '    RDOEnv.Close()
    '    RDOEnv = Nothing


    'End Function
    '2011/8/2 moriya delete end

    '�T�v  �Fð���ں��ސ�����
    '���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
    '      �FTBName,I,String,      ��������ð��ٖ�
    '      �Fjoken, I,String,      ��������
    '      �F�߂�l, O,Long,       �Y��ں��ސ�
    '����  �F�����ɊY������ں��ނ���
    '2011/8/8 moriya update start
    'Function VBRDO_Count(ByRef T_RDO As T_RDO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
    Function VBADO_Count(ByRef T_ADO As T_ADO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
        '2011/8/8 moriya update end

        '2011/7/25 moriya update start
        'Dim Rs As RDO.rdoResultset
        Dim Rs As New DataTable
        '2011/7/25 moriya update end
        Dim cnt As Integer
        Dim sqlcmd As String

        On Error GoTo ErrHandler

        sqlcmd = "SELECT COUNT (*) FROM " & TBLName & " WHERE " & joken

        '2011/7/25 moriya update start
        'Rs = T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'cnt = Rs.rdoColumns(0).Value
        dataset.Clear()

        ADO_DB_Search(sqlcmd, "Custom1", T_ADO, dataset)

        cnt = dataset.Tables("Custom1").Rows(0).Item(0)
        '2011/7/25 moriya update end

        '2011/7/25 moriya update start
        'VBRDO_Count = cnt
        VBADO_Count = cnt
        '2011/7/25 moriya update end

ExitFunc:
        On Error Resume Next
        'Rs.Close() '2011/7/25 moriya delete
        Exit Function

ErrHandler:
        '2011/7/25 moriya update start
        'VBRDO_Count = -1
        VBADO_Count = -1
        '2011/7/25 moriya update end

        'Dim er As rdoError
        'For Each er In rdoErrors
        'MsgBox er.Description
        'Next
        '   MsgBox Error$(Err), 64
        Resume ExitFunc

    End Function
    '2011/8/8 moriya delete start
    '�T�v  �Fð���ں��ސ�����
    '���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
    '      �FTBName,I,String,      ��������ð��ٖ�
    '      �Fjoken, I,String,      ��������
    '      �F�߂�l, O,Long,       �Y��ں��ސ�
    '����  �F�����ɊY������ں��ނ��� 
    '    Function VBRDO_Count(ByRef T_RDO As T_RDO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer

    '        Dim Rs As RDO.rdoResultset
    '        Dim cnt As Integer
    '        Dim sqlcmd As String

    '        On Error GoTo ErrHandler

    '        sqlcmd = "SELECT COUNT (*) FROM " & TBLName & " WHERE " & joken

    '        '2011/7/25 moriya update start
    '        Rs = T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
    '        cnt = Rs.rdoColumns(0).Value

    '        VBRDO_Count = cnt

    'ExitFunc:
    '        On Error Resume Next
    '        Rs.Close() '2011/7/25 moriya delete
    '        Exit Function

    'ErrHandler:
    '        VBRDO_Count = -1
    '        'Dim er As rdoError
    '        'For Each er In rdoErrors
    '        'MsgBox er.Description
    '        'Next
    '        '   MsgBox Error$(Err), 64
    '        Resume ExitFunc

    '    End Function
    '2011/8/8 moriya delete end

    '�T�v  �FDatabaseں��ލ폜����
    '���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
    '      �FTBName,I,String,      �폜����ð��ٖ�
    '      �Fjoken, I,String,      ����
    '      �F�߂�l, O,Long,       �����n�j or �m�f
    '����  �F�����ɊY������ں��ނ��폜

    '2011/8/3 moriya update start
    'Function VBRDO_Delete(ByRef T_RDO As T_RDO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
    Function VBADO_Delete(ByVal T_ADO As T_ADO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
        '2011/8/3 moriya update end

        Dim sqlcmd As String

        On Error GoTo ErrHandler

        sqlcmd = "DELETE FROM " & TBLName & " WHERE " & joken
        '2011/8/3 moriya update start
        'T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        ADO_DB_Event(T_ADO, sqlcmd)
        '2011/8/3 moriya update end

        '2011/8/3 moriya update start
        'VBRDO_Delete = 1
        VBADO_Delete = 1
        '2011/8/3 moriya update end

        Exit Function

ErrHandler:

        '2011/8/3 moriya update start
        'VBRDO_Delete = 0
        VBADO_Delete = 0
        '2011/8/3 moriya update end

        '    MsgBox Error$(Err) , 64, DEF_TestTitle1


    End Function

    '2011/9/13 moriya delete start
    '�T�v  �FDataBase�ڑ�
    '���Ұ��FRDOEnv,I,rdoEnvironment,RDO�ڑ��p���ϐ�
    '      �FT_RDO, I,T_RDO_Struct,  RDO�ڑ��p�\����
    '      �F�߂�l,O,Long,          ���� OK or NG
    '����  �FDataBase �Ɛڑ�����
    '    Function VBRDO_Connect(ByRef RDOEnv As RDO.rdoEnvironment, ByRef T_RDO As T_RDO_Struct) As Integer

    '        Dim ConStr As String

    '        On Error GoTo ErrHandler

    '        With T_RDO

    '            '    con = "UID =sa;PWD=;Database =brand;"
    '            ConStr = "UID=" & .UID & ";PWD=" & .PWD & ";Database=" & .DBName & ";"
    '            .Con = RDOEnv.OpenConnection(.DSN, RDO.PromptConstants.rdDriverNoPrompt, False, ConStr)

    '        End With

    '        VBRDO_Connect = 1

    '        Exit Function

    'ErrHandler:
    '        VBRDO_Connect = 0
    '        '    wkMsg = "DataBase �ڑ��������ɃG���[�������܂����B"
    '        '    GL_ErrMsg = wkMsg
    '        '    MsgBox Error$(Err) & vbCrLf & wkMsg, 64, "Connect Error"
    '        'Dim er As rdoError
    '        '    For Each er In rdoErrors
    '        '        MsgBox er.Description, er.Number
    '        '    Next er


    '    End Function
    '2011/9/13 moriya delete end

    '2011/8/2 moriya delete start
    ''�T�v  �FDataBase�ڑ��ؒf
    ''���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
    ''      �F�߂�l, O,Long,        -------
    ''����  �FDataBase �Ƃ̐ڑ���ؒf����
    'Function VBRDO_Discon(ByRef T_RDO As T_RDO_Struct) As Integer

    '    On Error Resume Next

    '    T_RDO.Con.Close()
    '    T_RDO.Con = Nothing


    'End Function
    '2011/8/2 moriya delete end

    '2011/9/13 moriya delete start
    '�T�v  �FDSN�쐬
    '���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
    '      �F�߂�l,O.Long,        ���� OK or NG
    '����  �FDSN���쐬
    '    Function VBRDO_DSNRegistry(ByRef T_RDO As T_RDO_Struct) As Integer

    '        Dim odbcAttr As String

    '        On Error GoTo ErrHandler

    '        '    odbcAttr = "Database=brand" & vbCr _
    '        ''            & odbcAttr & "Description=FFTEST" & vbCr _
    '        ''            & odbcAttr & "c:\windows\system\sqlsrv32.dll" & vbCr _
    '        ''            & odbcAttr & "Langage=japanese" & vbCr _
    '        ''            & odbcAttr & "OemToAnsi=NO" & vbCr _
    '        ''            & odbcAttr & "Server=Mother" & vbCr _
    '        ''            & odbcAttr & "UseProcForPrepare=Yes"

    '        odbcAttr = "Database=" & T_RDO.DBName & vbCr & "Description=" & vbCr & "OemToAnsi=No" & vbCr & "Server=" & T_RDO.Server

    '        RDOrdoEngine_definst.rdoRegisterDataSource(T_RDO.DSN, "SQL Server", True, odbcAttr)


    '        VBRDO_DSNRegistry = 1
    '        Exit Function

    'ErrHandler:
    '        VBRDO_DSNRegistry = 0



    '    End Function

    '�T�v  �FRDO�ڑ��p���ϐ��ݒ�
    '���Ұ��FRDOEnv, I/O,rdoEnvironment,RDO�ڑ��p���ϐ�
    '      �F�߂�l,O,Long,             ���� OK or NG
    '����  �FRDO�ڑ��p���ϐ���ݒ肷��
    '    Function VBRDO_OpenEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer

    '        On Error GoTo ErrHandler

    '        'RDO���ϐ�
    '        RDOEnv = RDOrdoEngine_definst.rdoEnvironments(0)

    '        VBRDO_OpenEnv = 1

    '        Exit Function

    'ErrHandler:
    '        VBRDO_OpenEnv = 0


    '    End Function
    '2011/9/13 moriya delete end

    '�T�v  �FT_RDO������
    '���Ұ��FT_RDO,I/O,T_RDO_Struct,RDO�ڑ��p���ϐ�
    '      �F�߂�l,O,Long,          -------
    '����  �FRDO�ڑ��p�\���̂�����������
    '2011/8/8 moriya update start
    'Function VBRDO_T_RDOInit(ByRef T_RDO As T_RDO_Struct) As Integer
    Function VBADO_T_ADOInit(ByRef T_ADO As T_ADO_Struct) As Integer
        '2011/8/8 moriya update end

        '2011/8/8 moriya update start
        'VBRDO_T_RDOInit = 1
        VBADO_T_ADOInit = 1
        '2011/8/8 moriya update end

        '2011/8/8 moriya update start
        'With T_RDO
        '    .DSN = "" 'RDO�ڑ��p�ް������
        '    .UID = "" 'RDO�ڑ��pհ��ID
        '    .PWD = "" 'RDO�ڑ��p�߽ܰ��
        '    .DBName = "" 'RDO�ڑ��p�ް��ް���
        '    .Server = "" 'RDO�ڑ��p����-
        '    .Con = Nothing 'RDO�ڑ��p��޼ު��
        'End With

        With T_ADO
            .DSN = "" 'ADO�ڑ��p�ް������
            .UID = "" 'ADO�ڑ��pհ��ID
            .PWD = "" 'ADO�ڑ��p�߽ܰ��
            .DBName = "" 'ADO�ڑ��p�ް��ް���
            .Server = "" 'ADO�ڑ��p����-
        End With
        '2011/8/8 moriya update start


    End Function

    '2011/7/25 moriya add method
    'ADO��SQL����p����DB����
    Function ADO_DB_Search(ByVal sqlcmd As String, ByVal dtname As String, ByVal T_ADO As T_ADO_Struct, ByRef ds As DataSet) As Integer
        On Error GoTo ErrHandler

        '�R�l�N�V������ݒ�
        'cnn.ConnectionString = "user id=cuad;password=cuad;initial catalog=cpmasterDB;data source=IHDB66;Connect Timeout=30"
        cnn.ConnectionString = "user id=" + T_ADO.UID + ";password=" + T_ADO.PWD + ";initial catalog=" + T_ADO.DBName + ";data source=" + T_ADO.Server + ";Connect Timeout=30"

        'SQL���̐ݒ�
        cmnd.CommandText = sqlcmd

        '�R�l�N�V�����̐ݒ�
        cmnd.Connection = cnn

        '�f�[�^�A�_�v�^�[�ɃR�}���h��ݒ�
        dataada.SelectCommand = cmnd

        '�f�[�^�Z�b�g�Ƀf�[�^�̎��Ԃ��擾����()
        dataada.Fill(ds, dtname)

        '�R�l�N�V���������
        cnn.Close()

        ADO_DB_Search = 0

        Exit Function
ErrHandler:
        ADO_DB_Search = -1
    End Function

    '2011/7/25 moriya add method
    'ADO��SQL����p����DB����i�o�^�A�X�V�A�폜�j
    Function ADO_DB_Event(ByVal ADO_str As T_ADO_Struct, ByVal sqlcmd As String) As Integer
        Dim cn As SqlConnection
        Dim DB_cmd As SqlCommand

        On Error GoTo ErrHandler

        '�R�l�N�V������ݒ�
        cn = New SqlConnection("user id=" + ADO_str.UID + ";password=" + ADO_str.PWD + ";initial catalog=" + ADO_str.DBName + ";data source=" + ADO_str.Server + ";Connect Timeout=30")
        'cn = New SqlConnection("user id=cuad;password=cuad;initial catalog=number;data source=IHDB66;Connect Timeout=30")

        '��ڂ̈�����SQL��������
        DB_cmd = New SqlCommand(sqlcmd, cn)

        cn.Open()

        'SQL���̎��s
        DB_cmd.ExecuteNonQuery()

        cn.Close()

        ADO_DB_Event = 0

        Exit Function

ErrHandler:
        ADO_DB_Event = -1
    End Function

    'ADO��SQL����p����DB����(Windows�F��)
    Function ADO_Win_Search(ByVal sqlcmd As String, ByVal dtname As String, ByVal T_ADO As T_ADO_Struct, ByRef ds As DataSet) As Integer
        On Error GoTo ErrHandler

        '�R�l�N�V������ݒ�
        'cnn.ConnectionString = "user id=cuad;password=cuad;initial catalog=cpmasterDB;data source=IHDB66;Connect Timeout=30"
        cnn.ConnectionString = "initial catalog=" + T_ADO.DBName + ";data source=" + T_ADO.Server + ";Integrated Security=SSPI"

        'SQL���̐ݒ�
        cmnd.CommandText = sqlcmd

        '�R�l�N�V�����̐ݒ�
        cmnd.Connection = cnn

        '�f�[�^�A�_�v�^�[�ɃR�}���h��ݒ�
        dataada.SelectCommand = cmnd

        '�f�[�^�Z�b�g�Ƀf�[�^�̎��Ԃ��擾����()
        dataada.Fill(ds, dtname)

        '�R�l�N�V���������
        cnn.Close()

        ADO_Win_Search = 0

        Exit Function
ErrHandler:
        ADO_Win_Search = -1
    End Function

    'ADO��SQL����p����DB����i�o�^�A�X�V�A�폜�j(Windows�F��)
    Function ADO_Win_DB_Event(ByVal ADO_str As T_ADO_Struct, ByVal sqlcmd As String) As Integer
        Dim cn As SqlConnection
        Dim DB_cmd As SqlCommand

        On Error GoTo ErrHandler

        '�R�l�N�V������ݒ�
        cn = New SqlConnection("initial catalog=" + ADO_str.DBName + ";data source=" + ADO_str.Server + ";Integrated Security=SSPI")

        '��ڂ̈�����SQL��������
        DB_cmd = New SqlCommand(sqlcmd, cn)

        cn.Open()

        'SQL���̎��s
        DB_cmd.ExecuteNonQuery()

        cn.Close()

        ADO_Win_DB_Event = 0

        Exit Function

ErrHandler:
        ADO_Win_DB_Event = -1
    End Function

    '2012/9/7 moriya add start
    '���������ɍH����������Ƃ��̌������o�͏���_CD
    Function VBADO_Count_FctCD(ByRef T_ADO As T_ADO_Struct, ByRef TBLName As String, ByVal Code As String, ByVal MainNo As String, _
                             ByVal Revno As String, ByRef joken As String) As Integer
        Dim Rs As New DataTable
        Dim cnt, i As Integer
        Dim sqlcmd As String

        On Error GoTo ErrHandler

        sqlcmd = "SELECT DISTINCT T_TM_FCT.C_CODE,T_TM_FCT.C_MAINNO,T_TM_FCT.C_REVNO FROM " & TBLName _
               & " INNER JOIN T_TM_FCT ON (" & TBLName & Code & "=T_TM_FCT.C_CODE) AND (" & TBLName _
               & MainNo & "=T_TM_FCT.C_MAINNO) AND (" & TBLName & Revno & "=T_TM_FCT.C_REVNO) WHERE " & joken

        dataset.Clear()

        ADO_DB_Search(sqlcmd, "Custom2", T_ADO, dataset)

        '�����A�ԍ������Ԃ��Ă����ꍇ�̓J�E���g���Ȃ�
        cnt = dataset.Tables("Custom2").Rows.Count

        VBADO_Count_FctCD = cnt

ExitFunc:
        On Error Resume Next
        Exit Function

ErrHandler:
        VBADO_Count_FctCD = -1
        Resume ExitFunc

    End Function

    '���������ɍH����������Ƃ��̌������o�͏���_CD
    Function VBADO_Count_FctBG(ByRef T_ADO As T_ADO_Struct, ByRef TBLName As String, ByVal Code As String, ByVal MainNo As String, _
                             ByVal Revno As String, ByRef joken As String) As Integer
        Dim Rs As New DataTable
        Dim cnt, i As Integer
        Dim sqlcmd As String

        On Error GoTo ErrHandler

        sqlcmd = "SELECT DISTINCT T_BG_FCT.C_CODE,T_BG_FCT.C_MAINNO,T_BG_FCT.C_REVNO FROM " & TBLName _
               & " INNER JOIN T_BG_FCT ON (" & TBLName & "." & Code & "=T_BG_FCT.C_CODE) AND (" & TBLName _
                & "." & MainNo & "=T_BG_FCT.C_MAINNO) AND (" & TBLName & "." & Revno & "=T_BG_FCT.C_REVNO) WHERE " & joken

        dataset.Clear()

        ADO_DB_Search(sqlcmd, "Custom4", T_ADO, dataset)

        '�����A�ԍ������Ԃ��Ă����ꍇ�̓J�E���g���Ȃ�
        cnt = dataset.Tables("Custom4").Rows.Count

        VBADO_Count_FctBG = cnt

ExitFunc:
        On Error Resume Next
        Exit Function

ErrHandler:
        VBADO_Count_FctBG = -1
        Resume ExitFunc

    End Function

    '���������ɍH����������Ƃ��̉�ʏo�͌����`�F�b�N
    Function VBADO_Check_Fct(ByRef T_ADO As T_ADO_Struct, ByVal BGNO As String, ByVal fact() As String) As Integer
        Dim ds As New DataSet()
        Dim dt As New DataTable()
        Dim sqlcmd As String
        Dim i, j As Integer

        On Error GoTo ErrHandler

        'BGNO�̍H������āA�������Ȃ��ꍇ�͓ǂݍ��܂Ȃ��B
        sqlcmd = "SELECT * FROM T_BG_FCT WHERE C_CODE='" & BGNO.Substring(0, 2) & _
                 "' AND C_MAINNO='" & BGNO.Substring(2, 8) & _
                 "' AND C_REVNO='" & BGNO.Substring(10, 4) & "'"
        ds.Clear()
        ADO_DB_Search(sqlcmd, "CUSBGFCT1", T_ADO, ds)
        dt = ds.Tables("CUSBGFCT1")

        i = 0
        While i < dt.Rows.Count
            For j = 0 To fact.Length - 1
                If fact(j) = Trim(dt.Rows(i).Item("FctCode")) Or _
                    fact(j) = "IMH" Or _
                    fact(j) = "ALL" Or _
                    Trim(dt.Rows(i).Item("FctCode")) = "IMH" Or _
                    Trim(dt.Rows(i).Item("FctCode")) = "ALL" Then
                    VBADO_Check_Fct = 0
                    Exit Function
                End If
            Next
            i = i + 1
        End While

        VBADO_Check_Fct = -1

        Exit Function

ExitFunc:
        On Error Resume Next
        Exit Function

ErrHandler:
        VBADO_Check_Fct = -1
        Resume ExitFunc

    End Function
    '2012/9/7 moriya add end
End Module