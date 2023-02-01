Option Strict Off
Option Explicit On
Module VBRDO
	'/***************************************************************************
	'このﾌｧｲﾙは 以下のｼｽﾃﾑで共有しています。
	'変更する際には、御注意願います。
	'
	'                          Comment：1998 8/25     by f.yamamoto
	'
	'       採番ｼｽﾃﾑ
	'       Ｃ表
	'       キュアドレイアウト
	'       ｷｭｱﾄﾞ・ﾋﾞｰﾄﾞﾘﾝｸﾞ            1998 11-12 add by yamamoto
	'       ｷｭｱﾄﾞ・ﾓｰﾙﾄﾞﾌﾟﾛﾌｧｲﾙ         1998 11-12 add by yamamoto
	'
	'
	'   [ Notes ]
	'       このﾌｧｲﾙをｲﾝｸﾙｰﾄﾞするには､以下のﾌｧｲﾙが必要です｡
	'           ・msrdo20.dll (Microsoft Remote Data Object 2.0)
	'           （参照設定でｲﾝｸﾙｰﾄﾞできます｡）
	'
	'   [ Contents ]
	'       Declarations
	'       VBRDO_Init         --- RDO接続用設定ﾌｧｲﾙ読込
	'       VBRDO_OpenEnv      --- RDO環境設定
	'       VBRDO_Count        --- DB指定条件ﾚｺｰﾄﾞ数ｶｳﾝﾄ
	'       VBRDO_Delete       --- DB指定条件ﾚｺｰﾄﾞ削除
	'       VBRDO_Connect      --- RDO接続（ｺﾈｸｼｮﾝを開く）
	'       VBRDO_Discon       --- RDO接続切断（ｺﾈｸｼｮﾝを切断する）
	'       VBRDO_RDORegistry  --- RDO接続用DSN作成
	'       VBRDO_CloseEnv     --- RDO環境解放
	'       VBRDO_T_RDOInit    --- RDO接続用構造体初期化

	'***************************************************************************/ ' すべての変数を明示的に宣言するようにします。

	'Public Const DEF_MSG_E9000 As String = "設定ファイル読み込みエラーです。 "
	'Public Const DEF_MSG_E9001 As String = "DSN作成中にエラーが生じました。"
	'Public Const DEF_MSG_E9002 As String = "RDO接続エラーです。"
	'Public Const DEF_MSG_E9003 As String = "ＤＢ登録処理中にエラーが生じました。"
	'Public Const DEF_MSG_E9004 As String = "ＤＢレコード削除処理中にエラーが生じました。"

	'接続用変数
	Structure T_RDO_Struct
		Dim DSN As String 'RDO接続用ﾃﾞｰﾀｿｰｽ名
		Dim UID As String 'RDO接続用ﾕｰｻﾞID
		Dim PWD As String 'RDO接続用ﾊﾟｽﾜｰﾄﾞ
		Dim DBName As String 'RDO接続用ﾃﾞｰﾀﾍﾞｰｽ名
		Dim Server As String 'RDO接続用ｻｰﾊﾞ-
		Dim Con As RDO.rdoConnection 'RDO接続用ｵﾌﾞｼﾞｪｸﾄ
	End Structure
	
	'WINAPI  設定ﾌｧｲﾙ読込み用関数で使用
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32"  Alias "GetEnvironmentVariableA"(ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	
	'概要  ：設定ﾌｧｲﾙよりRDO設定値を取得
	'ﾊﾟﾗﾒｰﾀ：T_RDO,  I/O,T_RDO_Struct,RDO接続用構造体
	'      ：section,I,  String,      ｾｸｼｮﾝ名
	'      ：fname,  I,  String,      ﾌｧｲﾙ名
	'      ：sw,     I,  Long,        ｽｲｯﾁ
	'      ：戻り値, O,Long,          処理 OK or NG
	'説明  ：設定ﾌｧｲﾙよりRDO設定値を取得
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
		'           ﾌｧｲﾙ格納ﾃﾞｨﾚｸﾄﾘ設定     Start
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
		'           ﾌｧｲﾙ格納ﾃﾞｨﾚｸﾄﾘ設定     End
		'==================================================
		
		
		'==================================================
		'           設定値取得     Start
		'==================================================
		
		Cmdstr(0) = "DSN" 'DSN
		Cmdstr(1) = "UID" 'ユーザーID
		Cmdstr(2) = "DBName" 'データベース名
		Cmdstr(3) = "PWD" 'パスワード
		Cmdstr(4) = "Server" 'サーバー
		
		
		For iCnt = 0 To 4
			
			'指定ｾｸｼｮﾝの指定ｷｰﾜｰﾄﾞの値を取得
			IRet = GetPrivateProfileString(lpszSection, Cmdstr(iCnt), "ERROR", str_Renamed.Value, 256, wkfname)
			
			If IRet <> 0 Then
				
				If InStr(1, str_Renamed.Value, "ERROR", CompareMethod.Binary) > 0 Then
					Exit Function
				End If
				
				'ｾﾐｺﾛﾝを取り除き、ｽﾍﾟｰｽもｶｯﾄする
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
		'           設定値取得     Start
		'==================================================
		
		VBRDO_Init = 1
		
		
	End Function
	
	'概要  ：RDO接続用環境変数解放
	'ﾊﾟﾗﾒｰﾀ：RDOEnv,I/O,rdoEnvironment,RDO接続用環境変数
	'      ：戻り値, O,Long,            -------
	'説明  ：RDO接続用環境変数解放
	Function VBRDO_CloseEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer
		
		On Error Resume Next
		
		VBRDO_CloseEnv = 1
		
		RDOEnv.Close()
        RDOEnv = Nothing
		
		
	End Function
	
	'概要  ：ﾃｰﾌﾞﾙﾚｺｰﾄﾞ数ﾁｪｯｸ
	'ﾊﾟﾗﾒｰﾀ：T_RDO, I,T_RDO_Struct,RDO接続用構造体
	'      ：TBName,I,String,      ﾁｪｯｸするﾃｰﾌﾞﾙ名
	'      ：joken, I,String,      検索条件
	'      ：戻り値, O,Long,       該当ﾚｺｰﾄﾞ数
	'説明  ：条件に該当するﾚｺｰﾄﾞをｶｳﾝﾄ
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
	
	'概要  ：Databaseﾚｺｰﾄﾞ削除処理
	'ﾊﾟﾗﾒｰﾀ：T_RDO, I,T_RDO_Struct,RDO接続用構造体
	'      ：TBName,I,String,      削除するﾃｰﾌﾞﾙ名
	'      ：joken, I,String,      条件
	'      ：戻り値, O,Long,       処理ＯＫ or ＮＧ
	'説明  ：条件に該当するﾚｺｰﾄﾞを削除
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
	
	
	'概要  ：DataBase接続
	'ﾊﾟﾗﾒｰﾀ：RDOEnv,I,rdoEnvironment,RDO接続用環境変数
	'      ：T_RDO, I,T_RDO_Struct,  RDO接続用構造体
	'      ：戻り値,O,Long,          処理 OK or NG
	'説明  ：DataBase と接続する
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
		'    wkMsg = "DataBase 接続処理中にエラーが生じました。"
		'    GL_ErrMsg = wkMsg
		'    MsgBox Error$(Err) & vbCrLf & wkMsg, 64, "Connect Error"
		'Dim er As rdoError
		'    For Each er In rdoErrors
		'        MsgBox er.Description, er.Number
		'    Next er
		
		
	End Function
	
	'概要  ：DataBase接続切断
	'ﾊﾟﾗﾒｰﾀ：T_RDO, I,T_RDO_Struct,RDO接続用構造体
	'      ：戻り値, O,Long,        -------
	'説明  ：DataBase との接続を切断する
	Function VBRDO_Discon(ByRef T_RDO As T_RDO_Struct) As Integer
		
		On Error Resume Next
		
		T_RDO.Con.Close()
        T_RDO.Con = Nothing
		
		
	End Function
	
	
	'概要  ：DSN作成
	'ﾊﾟﾗﾒｰﾀ：T_RDO, I,T_RDO_Struct,RDO接続用構造体
	'      ：戻り値,O.Long,        処理 OK or NG
	'説明  ：DSNを作成
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
	
	'概要  ：RDO接続用環境変数設定
	'ﾊﾟﾗﾒｰﾀ：RDOEnv, I/O,rdoEnvironment,RDO接続用環境変数
	'      ：戻り値,O,Long,             処理 OK or NG
	'説明  ：RDO接続用環境変数を設定する
	Function VBRDO_OpenEnv(ByRef RDOEnv As RDO.rdoEnvironment) As Integer
		
		On Error GoTo ErrHandler
		
		'RDO環境変数
		RDOEnv = RDOrdoEngine_definst.rdoEnvironments(0)
		
		VBRDO_OpenEnv = 1
		
		Exit Function
		
ErrHandler: 
		VBRDO_OpenEnv = 0
		
		
	End Function
	
	
	
	'概要  ：T_RDO初期化
	'ﾊﾟﾗﾒｰﾀ：T_RDO,I/O,T_RDO_Struct,RDO接続用環境変数
	'      ：戻り値,O,Long,          -------
	'説明  ：RDO接続用構造体を初期化する
	Function VBRDO_T_RDOInit(ByRef T_RDO As T_RDO_Struct) As Integer
		
		
		VBRDO_T_RDOInit = 1
		
		With T_RDO
			.DSN = "" 'RDO接続用ﾃﾞｰﾀｿｰｽ名
			.UID = "" 'RDO接続用ﾕｰｻﾞID
			.PWD = "" 'RDO接続用ﾊﾟｽﾜｰﾄﾞ
			.DBName = "" 'RDO接続用ﾃﾞｰﾀﾍﾞｰｽ名
			.Server = "" 'RDO接続用ｻｰﾊﾞ-
            .Con = Nothing 'RDO接続用ｵﾌﾞｼﾞｪｸﾄ
		End With
		
		
	End Function
End Module