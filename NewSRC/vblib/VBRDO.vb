Option Strict Off
Option Explicit On
Module VBRDO
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

	'Public Const DEF_MSG_E9000 As String = "�ݒ�t�@�C���ǂݍ��݃G���[�ł��B "
	'Public Const DEF_MSG_E9001 As String = "DSN�쐬���ɃG���[�������܂����B"
	'Public Const DEF_MSG_E9002 As String = "RDO�ڑ��G���[�ł��B"
	'Public Const DEF_MSG_E9003 As String = "�c�a�o�^�������ɃG���[�������܂����B"
	'Public Const DEF_MSG_E9004 As String = "�c�a���R�[�h�폜�������ɃG���[�������܂����B"

	'�ڑ��p�ϐ�
	Structure T_RDO_Struct
		Dim DSN As String 'RDO�ڑ��p�ް������
		Dim UID As String 'RDO�ڑ��pհ��ID
		Dim PWD As String 'RDO�ڑ��p�߽ܰ��
		Dim DBName As String 'RDO�ڑ��p�ް��ް���
		Dim Server As String 'RDO�ڑ��p����-
		Dim Con As RDO.rdoConnection 'RDO�ڑ��p��޼ު��
	End Structure
	
	'WINAPI  �ݒ�̧�ٓǍ��ݗp�֐��Ŏg�p
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32"  Alias "GetEnvironmentVariableA"(ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	
	'�T�v  �F�ݒ�̧�ق��RDO�ݒ�l���擾
	'���Ұ��FT_RDO,  I/O,T_RDO_Struct,RDO�ڑ��p�\����
	'      �Fsection,I,  String,      ����ݖ�
	'      �Ffname,  I,  String,      ̧�ٖ�
	'      �Fsw,     I,  Long,        ����
	'      �F�߂�l, O,Long,          ���� OK or NG
	'����  �F�ݒ�̧�ق��RDO�ݒ�l���擾
	Function VBRDO_Init(ByRef T_RDO As T_RDO_Struct, ByRef lpszSection As String, ByRef fname As String, ByRef sw As Integer) As Integer
		
		Dim IRet As Integer
        Dim str_Renamed As New VB6.FixedLengthString(256)
		Dim Cmdstr(4) As String
		Dim iCnt As Integer
		Dim posi As Integer
		Dim wkfname As String
		Dim AcadDir As String
		
		VBRDO_Init = 0
		
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
					Case 0 : T_RDO.DSN = Trim(str_Renamed.Value)
					Case 1 : T_RDO.UID = Trim(str_Renamed.Value)
					Case 2 : T_RDO.DBName = Trim(str_Renamed.Value)
					Case 3 : T_RDO.PWD = Trim(str_Renamed.Value)
					Case 4 : T_RDO.Server = Trim(str_Renamed.Value)
				End Select
				
			Else
				Exit Function
			End If
			
		Next 
		
		'==================================================
		'           �ݒ�l�擾     Start
		'==================================================
		
		VBRDO_Init = 1
		
		
	End Function
	
	'�T�v  �FRDO�ڑ��p���ϐ����
	'���Ұ��FRDOEnv,I/O,rdoEnvironment,RDO�ڑ��p���ϐ�
	'      �F�߂�l, O,Long,            -------
	'����  �FRDO�ڑ��p���ϐ����
	Function VBRDO_CloseEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer
		
		On Error Resume Next
		
		VBRDO_CloseEnv = 1
		
		RDOEnv.Close()
        RDOEnv = Nothing
		
		
	End Function
	
	'�T�v  �Fð���ں��ސ�����
	'���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
	'      �FTBName,I,String,      ��������ð��ٖ�
	'      �Fjoken, I,String,      ��������
	'      �F�߂�l, O,Long,       �Y��ں��ސ�
	'����  �F�����ɊY������ں��ނ���
	Function VBRDO_Count(ByRef T_RDO As T_RDO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
		
		Dim Rs As RDO.rdoResultset
		Dim cnt As Integer
		Dim sqlcmd As String
		
		On Error GoTo ErrHandler
		
		sqlcmd = "SELECT COUNT (*) FROM " & TBLName & " WHERE " & joken
		
		Rs = T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
		cnt = Rs.rdoColumns(0).Value
		
		VBRDO_Count = cnt
		
ExitFunc: 
		On Error Resume Next
		Rs.Close()
		Exit Function
		
ErrHandler: 
		VBRDO_Count = -1
		'Dim er As rdoError
		'For Each er In rdoErrors
		'MsgBox er.Description
		'Next
		'   MsgBox Error$(Err), 64
		Resume ExitFunc
		
	End Function
	
	'�T�v  �FDatabaseں��ލ폜����
	'���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
	'      �FTBName,I,String,      �폜����ð��ٖ�
	'      �Fjoken, I,String,      ����
	'      �F�߂�l, O,Long,       �����n�j or �m�f
	'����  �F�����ɊY������ں��ނ��폜
	Function VBRDO_Delete(ByRef T_RDO As T_RDO_Struct, ByRef TBLName As String, ByRef joken As String) As Integer
		
		Dim sqlcmd As String
		
		On Error GoTo ErrHandler
		
		sqlcmd = "DELETE FROM " & TBLName & " WHERE " & joken
		T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)

        ' -> watanabe add VerUP(2011)
        If T_RDO.Con.RowsAffected() = 0 Then
            GoTo ErrHandler
        End If
        ' <- watanabe add VerUP(2011)

		VBRDO_Delete = 1
		
		Exit Function
		
ErrHandler: 
		VBRDO_Delete = 0
		'    MsgBox Error$(Err) , 64, DEF_TestTitle1
		
		
	End Function
	
	
	'�T�v  �FDataBase�ڑ�
	'���Ұ��FRDOEnv,I,rdoEnvironment,RDO�ڑ��p���ϐ�
	'      �FT_RDO, I,T_RDO_Struct,  RDO�ڑ��p�\����
	'      �F�߂�l,O,Long,          ���� OK or NG
	'����  �FDataBase �Ɛڑ�����
	Function VBRDO_Connect(ByRef RDOEnv As RDO.rdoEnvironment, ByRef T_RDO As T_RDO_Struct) As Integer
		
		Dim ConStr As String
		
		On Error GoTo ErrHandler
		
		With T_RDO
			
			'    con = "UID =sa;PWD=;Database =brand;"
			ConStr = "UID=" & .UID & ";PWD=" & .PWD & ";Database=" & .DBName & ";"
			.Con = RDOEnv.OpenConnection(.DSN, RDO.PromptConstants.rdDriverNoPrompt, False, ConStr)
			
		End With
		
		VBRDO_Connect = 1
		
		Exit Function
		
ErrHandler: 
		VBRDO_Connect = 0
		'    wkMsg = "DataBase �ڑ��������ɃG���[�������܂����B"
		'    GL_ErrMsg = wkMsg
		'    MsgBox Error$(Err) & vbCrLf & wkMsg, 64, "Connect Error"
		'Dim er As rdoError
		'    For Each er In rdoErrors
		'        MsgBox er.Description, er.Number
		'    Next er
		
		
	End Function
	
	'�T�v  �FDataBase�ڑ��ؒf
	'���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
	'      �F�߂�l, O,Long,        -------
	'����  �FDataBase �Ƃ̐ڑ���ؒf����
	Function VBRDO_Discon(ByRef T_RDO As T_RDO_Struct) As Integer
		
		On Error Resume Next
		
		T_RDO.Con.Close()
        T_RDO.Con = Nothing
		
		
	End Function
	
	
	'�T�v  �FDSN�쐬
	'���Ұ��FT_RDO, I,T_RDO_Struct,RDO�ڑ��p�\����
	'      �F�߂�l,O.Long,        ���� OK or NG
	'����  �FDSN���쐬
	Function VBRDO_DSNRegistry(ByRef T_RDO As T_RDO_Struct) As Integer
		
		Dim odbcAttr As String
		
		On Error GoTo ErrHandler
		
		'    odbcAttr = "Database=brand" & vbCr _
		''            & odbcAttr & "Description=FFTEST" & vbCr _
		''            & odbcAttr & "c:\windows\system\sqlsrv32.dll" & vbCr _
		''            & odbcAttr & "Langage=japanese" & vbCr _
		''            & odbcAttr & "OemToAnsi=NO" & vbCr _
		''            & odbcAttr & "Server=Mother" & vbCr _
		''            & odbcAttr & "UseProcForPrepare=Yes"
		
		odbcAttr = "Database=" & T_RDO.DBName & vbCr & "Description=" & vbCr & "OemToAnsi=No" & vbCr & "Server=" & T_RDO.Server
		
		RDOrdoEngine_definst.rdoRegisterDataSource(T_RDO.DSN, "SQL Server", True, odbcAttr)
		
		
		VBRDO_DSNRegistry = 1
		Exit Function
		
ErrHandler: 
		VBRDO_DSNRegistry = 0
		
		
		
	End Function
	
	'�T�v  �FRDO�ڑ��p���ϐ��ݒ�
	'���Ұ��FRDOEnv, I/O,rdoEnvironment,RDO�ڑ��p���ϐ�
	'      �F�߂�l,O,Long,             ���� OK or NG
	'����  �FRDO�ڑ��p���ϐ���ݒ肷��
	Function VBRDO_OpenEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer
		
		On Error GoTo ErrHandler
		
		'RDO���ϐ�
		RDOEnv = RDOrdoEngine_definst.rdoEnvironments(0)
		
		VBRDO_OpenEnv = 1
		
		Exit Function
		
ErrHandler: 
		VBRDO_OpenEnv = 0
		
		
	End Function
	
	
	
	'�T�v  �FT_RDO������
	'���Ұ��FT_RDO,I/O,T_RDO_Struct,RDO�ڑ��p���ϐ�
	'      �F�߂�l,O,Long,          -------
	'����  �FRDO�ڑ��p�\���̂�����������
	Function VBRDO_T_RDOInit(ByRef T_RDO As T_RDO_Struct) As Integer
		
		
		VBRDO_T_RDOInit = 1
		
		With T_RDO
			.DSN = "" 'RDO�ڑ��p�ް������
			.UID = "" 'RDO�ڑ��pհ��ID
			.PWD = "" 'RDO�ڑ��p�߽ܰ��
			.DBName = "" 'RDO�ڑ��p�ް��ް���
			.Server = "" 'RDO�ڑ��p����-
            .Con = Nothing 'RDO�ڑ��p��޼ު��
		End With
		
		
	End Function
End Module