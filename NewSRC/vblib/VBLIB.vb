Option Strict Off
Option Explicit On
Module VBLIB
	'/*********************************************************************************************
	'このﾌｧｲﾙは ＶＢプログラムの共有ﾌｧｲﾙです。
	'変更する際には、御注意願います。
	'
	'                       Comment：1998 8/25  by f.yamamoto
	'
	'   [ Contents ]
	'       Declarations
	'       VBLIB_GetProfileini  ---  指定された初期化ﾌｧｲﾙ内の指定されたｾｸｼｮﾝから文字列を取得
	'                                 by yamamoto
	'       VBLIB_GetProfileini2 ---  指定された初期化ﾌｧｲﾙ内の指定されたｾｸｼｮﾝからｷｰ名を全て取得
	'                                 by yamamoto  1998 12 Add
	'
	'
	'*********************************************************************************************/
	
	'WINAPI  設定ﾌｧｲﾙ読込み用関数で使用 ( VBLIB_GetProfileIni)
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32"  Alias "GetEnvironmentVariableA"(ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	'指定された初期化ﾌｧｲﾙ内の指定されたセクションから文字列を取得します｡
	'WinAPIのGetprivateProfileStringをカスタマイズしたもの。
	'概要  ：文字列取得
	'ﾊﾟﾗﾒｰﾀ：lpszSection,     I,String, ｾｸｼｮﾝ名称
	'      ：lpszKey,         I,String, ｷｰ名称
	'      ：lpszDefault,     I,String, ﾃﾞﾌｫﾙﾄ文字列
	'      ：lpszReturnBuffer,I,String, 転送先ﾊﾞｯﾌｧ
	'      ：cchReturnBuffer, I,Long,   転送先ﾊﾞｯﾌｧのｻｲｽﾞ
	'      ：lpszFile,        I,String, 初期化ﾌｧｲﾙ名称
	'      ：sw,              I,Long,   ACAD環境変数を参照するか(sw = 0)、Winﾃﾞｲﾚｸﾄﾘを参照するか
	'                                   その他：入力引数をそのまま使用
	'      ：戻り値,          I,Long,    処理 ＯＫ or ＮＧ
	'説明  ：指定ﾌｧｲﾙをＡＣＡＤ環境変数設定ディレクトリ、またはＷＩＮディレクトリ下で検索し、
	'      ：指定セクションの、指定キーワードの値を取得する
	Function VBLib_GetProfileIni(ByRef lpszSection As String, ByRef lpszKey As String, ByRef lpszDefault As String, ByRef lpszReturnBuffer As String, ByRef cchReturnBuffer As Integer, ByRef lpszFile As String, ByRef sw As Integer) As Integer
		
		Dim posi As Integer
		Dim IRet As Integer
		Dim fname As String
		Dim AcadDir As String
		Dim SrhChar As String
		Dim wkReturnbuffer As New VB6.FixedLengthString(256)
		
		VBLib_GetProfileIni = 0
		
		'ﾌｧｲﾙ名をとっておく
		fname = lpszFile
		'ﾘﾀｰﾝ値を待避
		'wkReturnBuffer = lpszReturnBuffer
		wkReturnbuffer.Value = New String(Chr(0), 255)
		
		If sw = 0 Then
			
			'        'ACAD環境変数を参照
			'        AcadDir = Environ("ACAD_SET")
			'        '環境変数が設定されていなければ
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
		
		'指定ｾｸｼｮﾝの指定ｷｰﾜｰﾄﾞの値を取得
		IRet = GetPrivateProfileString(lpszSection, lpszKey, lpszDefault, wkReturnbuffer.Value, cchReturnBuffer, fname)
		If IRet = 0 Then
			Exit Function
			
		Else
			'ｾﾐｺﾛﾝを取り除き、ｽﾍﾟｰｽもｶｯﾄする
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
	
	
	'指定された初期化ファイル内の指定されたセクションから文字列を取得します｡
	'WinAPIのGetprivateProfileStringをカスタマイズしたもの。
	'概要  ：文字列取得
	'ﾊﾟﾗﾒｰﾀ：lpszSection,     I,String, ｾｸｼｮﾝ名称
	'      ：lpszKey,         I,String, ｷｰ名称
	'      ：lpszDefault,     I,String, ﾃﾞﾌｫﾙﾄ文字列
	'      ：lpszReturnBuffer,I,String, 転送先ﾊﾞｯﾌｧ
	'      ：cchReturnBuffer, I,Long,   転送先ﾊﾞｯﾌｧのｻｲｽﾞ
	'      ：lpszFile,        I,String, 初期化ﾌｧｲﾙ名称
	'      ：sw,              I,Long,   ACAD環境変数を参照するか(sw = 0)、Winﾃﾞｲﾚｸﾄﾘを参照するか
	'                                   その他：入力引数をそのまま使用
	'      ：戻り値,          I,Long,    処理 ＯＫ or ＮＧ
	'説明  ：指定ファイルをＡＣＡＤ環境変数設定ディレクトリ、またはＷＩＮディレクトリ下で検索し、
	'      ：指定セクションの、指定キーワードの値を取得する
	Function VBLib_GetProfileIni2(ByRef lpszSection As String, ByRef lpszKey As String, ByRef lpszDefault As String, ByRef lpszReturnBuffer As String, ByRef cchReturnBuffer As Integer, ByRef lpszFile As String, ByRef sw As Integer) As Integer
		
		Dim posi As Integer
		Dim posi2 As Integer
		Dim IRet As Integer
		Dim fname As String
		Dim AcadDir As String
		Dim SrhChar As String
		Dim wkReturnbuffer As New VB6.FixedLengthString(256)
		
		VBLib_GetProfileIni2 = 0
		
		'ﾌｧｲﾙ名をとっておく
		fname = lpszFile
		'ﾘﾀｰﾝ値を待避
		'wkReturnBuffer = lpszReturnBuffer
		wkReturnbuffer.Value = New String(Chr(0), 255)
		
		
		If sw = 0 Then
			
			'ACAD環境変数を参照
			AcadDir = Environ("ACAD_SET")
			'環境変数が設定されていなければ
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
		
		'指定ｾｸｼｮﾝの指定ｷｰﾜｰﾄﾞの値を取得
		IRet = GetPrivateProfileString(lpszSection, lpszKey, lpszDefault, wkReturnbuffer.Value, cchReturnBuffer, fname)
		If IRet = 0 Then
			Exit Function
			
		Else
			'ｾﾐｺﾛﾝを取り除き、ｽﾍﾟｰｽもｶｯﾄする
			SrhChar = ";"
			posi = InStr(1, wkReturnbuffer.Value, SrhChar, CompareMethod.Binary)
			If posi <> 0 Then
				lpszReturnBuffer = Trim(Left(wkReturnbuffer.Value, posi - 1))
			Else
				
				'NULL文字の検索
				posi = 1
				Do 
					'後ろのNULL文字の開始位置取得
					'ｷｰ名称はNULL文字区切りで入っている為
					posi = InStr(posi, wkReturnbuffer.Value, Chr(0), CompareMethod.Binary)
					posi2 = InStr(posi + 1, wkReturnbuffer.Value, Chr(0), CompareMethod.Binary)
					If posi2 - posi <= 1 Then
						posi = posi - 1
						Exit Do
					End If
					posi = posi + 1
					
				Loop 
				
				'ｷｰ名をｺﾋﾟｰする
				lpszReturnBuffer = Left(wkReturnbuffer.Value, posi)
			End If
			
		End If
		
		VBLib_GetProfileIni2 = 1
		
		
	End Function
End Module