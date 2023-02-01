Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.Collections.Generic


Module VBADO
	'/***************************************************************************
	'
	'	ADO.NET �֐����W���[��
	'                          Comment�F2022 12/23     by K.Taira
	'
	'
	'   [ Contents ]
	'       Declarations
	'       VBADO_Init         --- ADO�ڑ��p�ݒ�̧�ٓǍ�
	'       VBADO_Count        --- DB�w�����ں��ސ�����
	'       VBADO_Delete       --- DB�w�����ں��ލ폜
	'       VBADO_Connect      --- ADO�ڑ��i�ȸ��݂��J���j
	'       VBADO_Discon       --- ADO�ڑ��ؒf�i�ȸ��݂�ؒf����j
	'       VBADO_T_ADOInit    --- ADO�ڑ��p�\���̏�����

	'***************************************************************************/

	Public Const DEF_MSG_E9000 As String = "�ݒ�t�@�C���ǂݍ��݃G���[�ł��B "
	Public Const DEF_MSG_E9001 As String = "DSN�쐬���ɃG���[�������܂����B"
	Public Const DEF_MSG_E9002 As String = "ADO�ڑ��G���[�ł��B"
	Public Const DEF_MSG_E9003 As String = "�c�a�o�^�������ɃG���[�������܂����B"
	Public Const DEF_MSG_E9004 As String = "�c�a���R�[�h�폜�������ɃG���[�������܂����B"

	'�ڑ��p�ϐ�
	Structure T_ADO_Struct
		Dim DSN As String 'ADO�ڑ��p�f�[�^�\�[�X��
		Dim UID As String 'ADO�ڑ��p���[�UID
		Dim PWD As String 'ADO�ڑ��p�p�X���[�h
		Dim DBName As String 'ADO�ڑ��p�f�[�^�x�[�X��
		Dim Server As String 'ADO�ڑ��p�T�[�o
		Dim Con As SqlConnection 'ADO�ڑ��I�u�W�F�N�g
	End Structure

	'ADO�p�����[�^�\����
	Structure ADO_PARAM_Struct
		Dim ColumnName As String    '��
		Dim SqlDbType As SqlDbType  'DB�f�[�^�^
		Dim DataSize As Integer     '�f�[�^�T�C�Y(�f�[�^�^�� VarChar���ɒ�`����)
		Dim Value As Object         '�l(�l��ݒ肷��ꍇ�ɒ�`����)
		Dim Sign As String          '��r���Z�q('=', '>'���̒�`������ꍇ��WHERE��Ɏg�p����)
	End Structure

	'WINAPI  �ݒ�̧�ٓǍ��ݗp�֐��Ŏg�p
	Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer


	'�T�v    �ݒ�t�@�C�����ADO�ݒ�l���擾
	'�p�����[�^�FT_ADO		I/O		ADO�ڑ��p�\����
	'          �Fsection	I		�Z�N�V������
	'          �Ffname		I		�t�@�C����
	'          �Fsw			I		�X�C�b�` (0:���ϐ����ݒ�t�@�C���̊i�[�t�H���_�p�X���擾����)
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����    �ݒ�t�@�C�����ADO�ݒ�l���擾
	Function VBADO_Init(ByRef T_ADO As T_ADO_Struct, ByVal section As String, ByVal fname As String, ByVal sw As Integer) As Integer

		Try
			Dim IRet As Integer
			Dim str_Renamed As System.Text.StringBuilder = New System.Text.StringBuilder(256)
			Dim Cmdstr(4) As String
			Dim iCnt As Integer
			Dim posi As Integer
			Dim wkfname As String
			Dim AcadDir As String

			VBADO_Init = 0

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
				wkfname = Trim(AcadDir) & fname

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
				IRet = GetPrivateProfileString(section, Cmdstr(iCnt), "ERROR", str_Renamed, str_Renamed.Capacity - 1, wkfname)

				If IRet <> 0 Then
					Dim strWork As String = str_Renamed.ToString()

					If InStr(1, strWork, "ERROR", CompareMethod.Binary) > 0 Then
						Exit Function
					End If

					'�к�݂���菜���A��߰��දĂ���
					posi = InStr(1, strWork, ";", CompareMethod.Binary)
					If posi <> 0 Then
						strWork = Trim(Left(strWork, posi - 1))
					Else
						strWork = Trim(Mid(strWork, 1, InStr(1, strWork, Chr(0), CompareMethod.Binary) - 1))
					End If

					Select Case iCnt
						Case 0 : T_ADO.DSN = Trim(strWork)
						Case 1 : T_ADO.UID = Trim(strWork)
						Case 2 : T_ADO.DBName = Trim(strWork)
						Case 3 : T_ADO.PWD = Trim(strWork)
						Case 4 : T_ADO.Server = Trim(strWork)
					End Select

				Else
					Exit Function
				End If

			Next

			'==================================================
			'           �ݒ�l�擾     Start
			'==================================================

			VBADO_Init = 1

		Catch ex As Exception

			VBADO_Init = 0

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'�T�v  �e�[�u�����R�[�h���`�F�b�N
	'�p�����[�^�FT_ADO	I/O		ADO�ڑ��p�\����
	'          �FTBName	I		�`�F�b�N����e�[�u����
	'          �Fjoken	I		��������
	'          �F�߂�l			�Y�����R�[�h��(NG:-1)
	'����  �����ɊY�����郌�R�[�h���J�E���g
	Function VBADO_Count(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal joken As String) As Integer

		Try
			Dim cnt As Integer = 0

			Using cmd As New SqlClient.SqlCommand()
				cmd.Connection = T_ADO.Con
				cmd.CommandType = System.Data.CommandType.Text
				'SQL�R�}���h�ݒ�
				cmd.CommandText = "SELECT COUNT (*) FROM " & TBLName & " WHERE " & joken

				'���o�������R�[�h�����擾
				cnt = cmd.ExecuteScalar()
			End Using

			VBADO_Count = cnt

		Catch ex As Exception

			VBADO_Count = -1

			'DataBase�ؒf
			VBADO_Discon(T_ADO)

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'�T�v  Database���R�[�h��������
	'�p�����[�^�FT_ADO		I/O		ADO�ڑ��p�\����
	'          �FTBName		I		�폜����e�[�u����
	'          �Fjoken		I		��������
	'          �FparamList	I		�p�����[�^���X�g(ADO�p�����[�^�\���̃��X�g �񖼂�DB�f�[�^�^�̂ݎg�p)
	'          �FdataList	I		�����f�[�^���X�g(1������:���R�[�h�A2������:��)
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����  �����ɊY�����郌�R�[�h���폜
	Function VBADO_Search(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal joken As String, ByVal paramList As List(Of ADO_PARAM_Struct), ByRef dataList As List(Of List(Of String))) As Integer

		Try
			dataList = New List(Of List(Of String))

			Using cmd As New SqlClient.SqlCommand()

				cmd.Connection = T_ADO.Con
				cmd.CommandType = System.Data.CommandType.Text

				cmd.CommandText = "SELECT * " & "FROM " & TBLName & " " & "WHERE " & joken

				Dim dread As SqlClient.SqlDataReader = cmd.ExecuteReader()

				Do While dread.Read()

					Dim list As List(Of String) = New List(Of String)

					For Each param As ADO_PARAM_Struct In paramList

						'Database�e�[�u����P�ʂł̃f�[�^�Ǎ�
						Dim data As String = VBADO_Read(dread, param.ColumnName, param.SqlDbType)
						list.Add(data)
					Next

					dataList.Add(list)
				Loop

				dread.Close()

			End Using

			VBADO_Search = 1

		Catch ex As Exception

			VBADO_Search = 0

			'DataBase�ؒf
			VBADO_Discon(T_ADO)

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'�T�v  Database�e�[�u����P�ʂł̃f�[�^�Ǎ�
	'�p�����[�^�Fdread		I/O		SQL Data Reader
	'          �FcolumnName	I		�e�[�u���̗�
	'          �FsqlDbType	I		DB�f�[�^�^
	'          �F�߂�l				�Ǎ��񂾃f�[�^(String�^)
	'����  DB�e�[�u������f�[�^�^�ɂ��f�[�^�Ǎ����s��
	Function VBADO_Read(ByRef dread As SqlClient.SqlDataReader, ByVal columnName As String, ByVal sqlDbType As SqlDbType) As String

		Dim result As String = ""

		If sqlDbType = SqlDbType.Char Or sqlDbType = SqlDbType.VarChar Then
			result = dread.GetString(dread.GetOrdinal(columnName))
		ElseIf sqlDbType = SqlDbType.TinyInt Then
			result = dread.GetByte(dread.GetOrdinal(columnName)).ToString()
		ElseIf sqlDbType = SqlDbType.SmallInt Then
			result = dread.GetInt16(dread.GetOrdinal(columnName)).ToString()
		ElseIf sqlDbType = SqlDbType.Int Then
			result = dread.GetInt32(dread.GetOrdinal(columnName)).ToString()
		ElseIf sqlDbType = SqlDbType.BigInt Then
			result = dread.GetInt64(dread.GetOrdinal(columnName)).ToString()
		ElseIf sqlDbType = SqlDbType.Float Then
			result = dread.GetDouble(dread.GetOrdinal(columnName)).ToString()
		ElseIf sqlDbType = SqlDbType.SmallDateTime Then
			result = dread.GetDateTime(dread.GetOrdinal(columnName)).ToString("yyyyMMdd")
		Else
			result = ""
		End If

		VBADO_Read = result

	End Function

	'�T�v  Database���R�[�h�ǉ�����
	'�p�����[�^�FT_ADO			I/O		ADO�ڑ��p�\����
	'          �FTBName			I		�ǉ�����e�[�u����
	'          �FparamList		I		�p�����[�^���X�g
	'          �FisDisconnect	I		�G���[����DB�ؒf���邩�ۂ�
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����  �p�����[�^�̏����Ń��R�[�h��ǉ�
	Function VBADO_Insert(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal paramList As List(Of ADO_PARAM_Struct), Optional isDisconnect As Boolean = False) As Integer

		Try

			Using cmd As New SqlClient.SqlCommand()

				cmd.Connection = T_ADO.Con

				cmd.CommandText = "INSERT INTO " & TBLName & " ("

				For i As Integer = 0 To paramList.Count - 1
					If i > 0 Then
						cmd.CommandText += ", "
					End If

					cmd.CommandText += paramList(i).ColumnName
				Next

				cmd.CommandText += ") VALUES("

				For i As Integer = 0 To paramList.Count - 1
					If i > 0 Then
						cmd.CommandText += ", "
					End If

					cmd.CommandText += ("@" & paramList(i).ColumnName)
				Next

				cmd.CommandText += ")"

				'SQL�R�}���h�p�����[�^�쐬
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'�N�G���[���s
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Insert = 1

		Catch ex As Exception

			VBADO_Insert = 0

			If isDisconnect = True Then
				'DataBase�ؒf
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'�T�v  Database���R�[�h�X�V����
	'�p�����[�^�FT_ADO			I/O		ADO�ڑ��p�\����
	'          �FTBName			I		�X�V����e�[�u����
	'          �FparamList		I		�p�����[�^���X�g
	'          �FisDisconnect	I		�G���[����DB�ؒf���邩�ۂ�
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����  �p�����[�^�̏����Ń��R�[�h���X�V
	Function VBADO_Update(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal paramList As List(Of ADO_PARAM_Struct), Optional isDisconnect As Boolean = False) As Integer

		Try

			Using cmd As New SqlClient.SqlCommand()

				cmd.Connection = T_ADO.Con

				cmd.CommandText = "UPDATE " & TBLName & " SET "

				For i As Integer = 0 To paramList.Count - 1
					If i > 0 Then
						cmd.CommandText += ", "
					End If

					cmd.CommandText += (paramList(i).ColumnName & " = @" & paramList(i).ColumnName)
				Next

				cmd.CommandText += " WHERE "

				Dim sgnCnt As Integer = 0

				For i As Integer = 0 To paramList.Count - 1
					If paramList(i).Sign = "" Then
						Continue For
					End If

					If sgnCnt > 0 Then
						cmd.CommandText += " AND "
					End If
					sgnCnt += 1

					cmd.CommandText += (paramList(i).ColumnName & " " & paramList(i).Sign & " @" & paramList(i).ColumnName)
				Next

				'SQL�R�}���h�p�����[�^�쐬
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'�N�G���[���s
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Update = 1

		Catch ex As Exception

			VBADO_Update = 0

			If isDisconnect = True Then
				'DataBase�ؒf
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'�T�v  Database���R�[�h�폜����
	'�p�����[�^�FT_ADO			I/O		ADO�ڑ��p�\����
	'          �FTBName			I		�폜����e�[�u����
	'          �FparamList		I		�p�����[�^���X�g
	'          �FisDisconnect	I		�G���[����DB�ؒf���邩�ۂ�
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����  �����ɊY�����郌�R�[�h���폜
	Function VBADO_Delete(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal paramList As List(Of ADO_PARAM_Struct), Optional isDisconnect As Boolean = False) As Integer

		Try
			Using cmd As New SqlClient.SqlCommand()

				'SQL�R�}���h�ݒ�
				cmd.Connection = T_ADO.Con

				cmd.CommandText = "DELETE FROM " & TBLName & " WHERE "

				For i As Integer = 0 To paramList.Count - 1

					If i > 0 Then
						cmd.CommandText += " AND "
					End If

					If paramList(i).Sign.Length > 0 Then
						cmd.CommandText = cmd.CommandText & paramList(i).ColumnName & " " & paramList(i).Sign & " @" & paramList(i).ColumnName
					Else
						cmd.CommandText = cmd.CommandText & paramList(i).ColumnName & " = @" & paramList(i).ColumnName
					End If
				Next

				'SQL�R�}���h�p�����[�^�쐬
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'�N�G���[���s
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Delete = 1

		Catch ex As Exception

			VBADO_Delete = 0

			If isDisconnect = True Then
				'DataBase�ؒf
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function


	'�T�v  �FDataBase�ڑ�
	'���Ұ��FT_ADO		I	T_ADO_Struct	ADO�ڑ��p�\����
	'      �F�߂�l			Integer			�������� (1:OK / 0:NG)
	'����  �FDataBase �Ɛڑ�����
	Function VBADO_Connect(ByRef T_ADO As T_ADO_Struct) As Integer

		Try
			Dim ConStr As String

			With T_ADO

				.Con = New SqlConnection()

				ConStr = "UID=" & .UID & ";PWD=" & .PWD & ";Database=" & .DBName & ";"
				.Con.ConnectionString = " Data Source = " & .Server &
						";Initial Catalog = " & .DBName &
						";User ID = " & .UID &
						";Password =" & .PWD

				.Con.Open()

			End With

			VBADO_Connect = 1

		Catch ex As Exception

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

			VBADO_Connect = 0

		End Try

	End Function

	'�T�v  �FDataBase�ؒf
	'���Ұ��FT_ADO		I	T_ADO_Struct	ADO�ڑ��p�\����
	'      �F�߂�l			Integer			�������� (1:OK / 0:NG)
	'����  �FDataBase �Ƃ̐ڑ���ؒf����
	Function VBADO_Discon(ByRef T_ADO As T_ADO_Struct) As Integer

		Try
			T_ADO.Con.Close()
			T_ADO.Con.Dispose()

			VBADO_Discon = 1

		Catch ex As Exception

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

			VBADO_Discon = 0

		End Try

	End Function


	'�T�v  �FT_ADO������
	'���Ұ��FT_ADO	I/O			ADO�ڑ��p���ϐ�
	'      �F�߂�l	Integer		-------
	'����  �FADO�ڑ��p�\���̂�����������
	Function VBADO_T_ADOInit(ByRef T_ADO As T_ADO_Struct) As Integer

		VBADO_T_ADOInit = 1

		With T_ADO
			.DSN = "" 'ADO�ڑ��p�ް������
			.UID = "" 'ADO�ڑ��pհ��ID
			.PWD = "" 'ADO�ڑ��p�߽ܰ��
			.DBName = "" 'ADO�ڑ��p�ް��ް���
			.Server = "" 'ADO�ڑ��p����-
			.Con = Nothing 'ADO�ڑ��p��޼ު��
		End With

	End Function


	'�T�v  SQL�R�}���h�p�����[�^�쐬
	'�p�����[�^�Fcmd		I/O		SqlCommand
	'          �Fparam		I		ADO�p�����[�^
	'          �F�߂�l				�������� (1:OK / 0:NG)
	'����  �p�����[�^���X�g���SQL�R�}���h���쐬
	Function MakeCommandParam(ByRef cmd As SqlClient.SqlCommand, ByVal param As ADO_PARAM_Struct) As Integer

		Try
			Dim paramName = "@" & param.ColumnName

			If param.SqlDbType = SqlDbType.VarChar Then
				cmd.Parameters.Add(paramName, param.SqlDbType, param.DataSize).Value = CType(param.Value, String)
			ElseIf param.SqlDbType = SqlDbType.Char Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = CType(param.Value, String)
			ElseIf param.SqlDbType = SqlDbType.Float Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = CType(param.Value, Double)
			ElseIf param.SqlDbType = SqlDbType.SmallDateTime Or param.SqlDbType = SqlDbType.DateTime Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = DateTime.Parse(param.Value.ToString())
			ElseIf param.SqlDbType = SqlDbType.Int Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = CType(param.Value, Int32)
			ElseIf param.SqlDbType = SqlDbType.SmallInt Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = CType(param.Value, Int16)
			ElseIf param.SqlDbType = SqlDbType.TinyInt Then
				cmd.Parameters.Add(paramName, param.SqlDbType).Value = CType(param.Value, Byte)
			Else
				MakeCommandParam = 0
				MessageBox.Show("SqlDbType = " & param.SqlDbType.GetTypeCode().ToString(), "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)
			End If

			MakeCommandParam = 1

		Catch ex As Exception

			MakeCommandParam = 0

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

End Module