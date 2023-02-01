Option Strict Off
Option Explicit On
Module VBLIB
	'/*********************************************************************************************
	'����̧�ق� �u�a�v���O�����̋��Ļ�قł��B
	'�ύX����ۂɂ́A�䒍�ӊ肢�܂��B
	'
	'                       Comment�F1998 8/25  by f.yamamoto
	'
	'   [ Contents ]
	'       Declarations
	'       VBLIB_GetProfileini  ---  �w�肳�ꂽ������̧�ٓ��̎w�肳�ꂽ����݂��當������擾
	'                                 by yamamoto
	'       VBLIB_GetProfileini2 ---  �w�肳�ꂽ������̧�ٓ��̎w�肳�ꂽ����݂��緰����S�Ď擾
	'                                 by yamamoto  1998 12 Add
	'
	'
	'*********************************************************************************************/
	
	'WINAPI  �ݒ�̧�ٓǍ��ݗp�֐��Ŏg�p ( VBLIB_GetProfileIni)
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32"  Alias "GetEnvironmentVariableA"(ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	'�w�肳�ꂽ������̧�ٓ��̎w�肳�ꂽ�Z�N�V�������當������擾���܂��
	'WinAPI��GetprivateProfileString���J�X�^�}�C�Y�������́B
	'�T�v  �F������擾
	'���Ұ��FlpszSection,     I,String, ����ݖ���
	'      �FlpszKey,         I,String, ������
	'      �FlpszDefault,     I,String, ��̫�ĕ�����
	'      �FlpszReturnBuffer,I,String, �]�����ޯ̧
	'      �FcchReturnBuffer, I,Long,   �]�����ޯ̧�̻���
	'      �FlpszFile,        I,String, ������̧�ٖ���
	'      �Fsw,              I,Long,   ACAD���ϐ����Q�Ƃ��邩(sw = 0)�AWin�޲ڸ�؂��Q�Ƃ��邩
	'                                   ���̑��F���͈��������̂܂܎g�p
	'      �F�߂�l,          I,Long,    ���� �n�j or �m�f
	'����  �F�w��̧�ق��`�b�`�c���ϐ��ݒ�f�B���N�g���A�܂��͂v�h�m�f�B���N�g�����Ō������A
	'      �F�w��Z�N�V�����́A�w��L�[���[�h�̒l���擾����
	Function VBLib_GetProfileIni(ByRef lpszSection As String, ByRef lpszKey As String, ByRef lpszDefault As String, ByRef lpszReturnBuffer As String, ByRef cchReturnBuffer As Integer, ByRef lpszFile As String, ByRef sw As Integer) As Integer
		
		Dim posi As Integer
		Dim IRet As Integer
		Dim fname As String
		Dim AcadDir As String
		Dim SrhChar As String
		Dim wkReturnbuffer As New VB6.FixedLengthString(256)
		
		VBLib_GetProfileIni = 0
		
		'̧�ٖ����Ƃ��Ă���
		fname = lpszFile
		'���ݒl��Ҕ�
		'wkReturnBuffer = lpszReturnBuffer
		wkReturnbuffer.Value = New String(Chr(0), 255)
		
		If sw = 0 Then
			
			'        'ACAD���ϐ����Q��
			'        AcadDir = Environ("ACAD_SET")
			'        '���ϐ����ݒ肳��Ă��Ȃ����
			'        If Len(Trim$(AcadDir)) = 0 Then
			'            Exit Function
			'        End If
			AcadDir = New String(Chr(0), 255)
			IRet = GetEnvironmentVariable("ACAD_SET", AcadDir, Len(AcadDir))
			If IRet = 0 Then
				Exit Function
			End If
			
			AcadDir = Left(AcadDir, InStr(1, AcadDir, Chr(0), CompareMethod.Binary) - 1)
			If Right(AcadDir, 1) <> "\" Then
				AcadDir = AcadDir & "\"
			End If
			fname = Trim(AcadDir) & fname
			
		End If
		
		'�w�辸��݂̎w�跰ܰ�ނ̒l���擾
		IRet = GetPrivateProfileString(lpszSection, lpszKey, lpszDefault, wkReturnbuffer.Value, cchReturnBuffer, fname)
		If IRet = 0 Then
			Exit Function
			
		Else
			'�к�݂���菜���A��߰��දĂ���
			SrhChar = ";"
			posi = InStr(1, wkReturnbuffer.Value, SrhChar, CompareMethod.Binary)
			If posi <> 0 Then
				lpszReturnBuffer = Trim(Left(wkReturnbuffer.Value, posi - 1))
			Else
				lpszReturnBuffer = Trim(Mid(wkReturnbuffer.Value, 1, InStr(1, wkReturnbuffer.Value, Chr(0), CompareMethod.Binary) - 1))
			End If
			
		End If
		
		VBLib_GetProfileIni = 1
		
		
	End Function
	
	
	'�w�肳�ꂽ�������t�@�C�����̎w�肳�ꂽ�Z�N�V�������當������擾���܂��
	'WinAPI��GetprivateProfileString���J�X�^�}�C�Y�������́B
	'�T�v  �F������擾
	'���Ұ��FlpszSection,     I,String, ����ݖ���
	'      �FlpszKey,         I,String, ������
	'      �FlpszDefault,     I,String, ��̫�ĕ�����
	'      �FlpszReturnBuffer,I,String, �]�����ޯ̧
	'      �FcchReturnBuffer, I,Long,   �]�����ޯ̧�̻���
	'      �FlpszFile,        I,String, ������̧�ٖ���
	'      �Fsw,              I,Long,   ACAD���ϐ����Q�Ƃ��邩(sw = 0)�AWin�޲ڸ�؂��Q�Ƃ��邩
	'                                   ���̑��F���͈��������̂܂܎g�p
	'      �F�߂�l,          I,Long,    ���� �n�j or �m�f
	'����  �F�w��t�@�C�����`�b�`�c���ϐ��ݒ�f�B���N�g���A�܂��͂v�h�m�f�B���N�g�����Ō������A
	'      �F�w��Z�N�V�����́A�w��L�[���[�h�̒l���擾����
	Function VBLib_GetProfileIni2(ByRef lpszSection As String, ByRef lpszKey As String, ByRef lpszDefault As String, ByRef lpszReturnBuffer As String, ByRef cchReturnBuffer As Integer, ByRef lpszFile As String, ByRef sw As Integer) As Integer
		
		Dim posi As Integer
		Dim posi2 As Integer
		Dim IRet As Integer
		Dim fname As String
		Dim AcadDir As String
		Dim SrhChar As String
		Dim wkReturnbuffer As New VB6.FixedLengthString(256)
		
		VBLib_GetProfileIni2 = 0
		
		'̧�ٖ����Ƃ��Ă���
		fname = lpszFile
		'���ݒl��Ҕ�
		'wkReturnBuffer = lpszReturnBuffer
		wkReturnbuffer.Value = New String(Chr(0), 255)
		
		
		If sw = 0 Then
			
			'ACAD���ϐ����Q��
			AcadDir = Environ("ACAD_SET")
			'���ϐ����ݒ肳��Ă��Ȃ����
			If Len(Trim(AcadDir)) = 0 Then
				Exit Function
			End If
			AcadDir = New String(Chr(0), 255)
			IRet = GetEnvironmentVariable("ACAD_SET", AcadDir, Len(AcadDir))
			If IRet = 0 Then
				Exit Function
			End If
			
			AcadDir = Left(AcadDir, InStr(1, AcadDir, Chr(0), CompareMethod.Binary) - 1)
			If Right(AcadDir, 1) <> "\" Then
				AcadDir = AcadDir & "\"
			End If
			fname = Trim(AcadDir) & fname
			
		End If
		
		'�w�辸��݂̎w�跰ܰ�ނ̒l���擾
		IRet = GetPrivateProfileString(lpszSection, lpszKey, lpszDefault, wkReturnbuffer.Value, cchReturnBuffer, fname)
		If IRet = 0 Then
			Exit Function
			
		Else
			'�к�݂���菜���A��߰��දĂ���
			SrhChar = ";"
			posi = InStr(1, wkReturnbuffer.Value, SrhChar, CompareMethod.Binary)
			If posi <> 0 Then
				lpszReturnBuffer = Trim(Left(wkReturnbuffer.Value, posi - 1))
			Else
				
				'NULL�����̌���
				posi = 1
				Do 
					'����NULL�����̊J�n�ʒu�擾
					'�����̂�NULL������؂�œ����Ă����
					posi = InStr(posi, wkReturnbuffer.Value, Chr(0), CompareMethod.Binary)
					posi2 = InStr(posi + 1, wkReturnbuffer.Value, Chr(0), CompareMethod.Binary)
					If posi2 - posi <= 1 Then
						posi = posi - 1
						Exit Do
					End If
					posi = posi + 1
					
				Loop 
				
				'�������߰����
				lpszReturnBuffer = Left(wkReturnbuffer.Value, posi)
			End If
			
		End If
		
		VBLib_GetProfileIni2 = 1
		
		
	End Function
End Module