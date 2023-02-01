Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports System.Collections.Generic


Module VBADO
	'/***************************************************************************
	'
	'	ADO.NET 関数モジュール
	'                          Comment：2022 12/23     by K.Taira
	'
	'
	'   [ Contents ]
	'       Declarations
	'       VBADO_Init         --- ADO接続用設定ﾌｧｲﾙ読込
	'       VBADO_Count        --- DB指定条件ﾚｺｰﾄﾞ数ｶｳﾝﾄ
	'       VBADO_Delete       --- DB指定条件ﾚｺｰﾄﾞ削除
	'       VBADO_Connect      --- ADO接続（ｺﾈｸｼｮﾝを開く）
	'       VBADO_Discon       --- ADO接続切断（ｺﾈｸｼｮﾝを切断する）
	'       VBADO_T_ADOInit    --- ADO接続用構造体初期化

	'***************************************************************************/

	Public Const DEF_MSG_E9000 As String = "設定ファイル読み込みエラーです。 "
	Public Const DEF_MSG_E9001 As String = "DSN作成中にエラーが生じました。"
	Public Const DEF_MSG_E9002 As String = "ADO接続エラーです。"
	Public Const DEF_MSG_E9003 As String = "ＤＢ登録処理中にエラーが生じました。"
	Public Const DEF_MSG_E9004 As String = "ＤＢレコード削除処理中にエラーが生じました。"

	'接続用変数
	Structure T_ADO_Struct
		Dim DSN As String 'ADO接続用データソース名
		Dim UID As String 'ADO接続用ユーザID
		Dim PWD As String 'ADO接続用パスワード
		Dim DBName As String 'ADO接続用データベース名
		Dim Server As String 'ADO接続用サーバ
		Dim Con As SqlConnection 'ADO接続オブジェクト
	End Structure

	'ADOパラメータ構造体
	Structure ADO_PARAM_Struct
		Dim ColumnName As String    '列名
		Dim SqlDbType As SqlDbType  'DBデータ型
		Dim DataSize As Integer     'データサイズ(データ型が VarChar時に定義する)
		Dim Value As Object         '値(値を設定する場合に定義する)
		Dim Sign As String          '比較演算子('=', '>'等の定義がある場合はWHERE句に使用する)
	End Structure

	'WINAPI  設定ﾌｧｲﾙ読込み用関数で使用
	Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer


	'概要    設定ファイルよりADO設定値を取得
	'パラメータ：T_ADO		I/O		ADO接続用構造体
	'          ：section	I		セクション名
	'          ：fname		I		ファイル名
	'          ：sw			I		スイッチ (0:環境変数より設定ファイルの格納フォルダパスを取得する)
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明    設定ファイルよりADO設定値を取得
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
				wkfname = Trim(AcadDir) & fname

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
				IRet = GetPrivateProfileString(section, Cmdstr(iCnt), "ERROR", str_Renamed, str_Renamed.Capacity - 1, wkfname)

				If IRet <> 0 Then
					Dim strWork As String = str_Renamed.ToString()

					If InStr(1, strWork, "ERROR", CompareMethod.Binary) > 0 Then
						Exit Function
					End If

					'ｾﾐｺﾛﾝを取り除き、ｽﾍﾟｰｽもｶｯﾄする
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
			'           設定値取得     Start
			'==================================================

			VBADO_Init = 1

		Catch ex As Exception

			VBADO_Init = 0

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'概要  テーブルレコード数チェック
	'パラメータ：T_ADO	I/O		ADO接続用構造体
	'          ：TBName	I		チェックするテーブル名
	'          ：joken	I		検索条件
	'          ：戻り値			該当レコード数(NG:-1)
	'説明  条件に該当するレコードをカウント
	Function VBADO_Count(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal joken As String) As Integer

		Try
			Dim cnt As Integer = 0

			Using cmd As New SqlClient.SqlCommand()
				cmd.Connection = T_ADO.Con
				cmd.CommandType = System.Data.CommandType.Text
				'SQLコマンド設定
				cmd.CommandText = "SELECT COUNT (*) FROM " & TBLName & " WHERE " & joken

				'抽出したレコード件数取得
				cnt = cmd.ExecuteScalar()
			End Using

			VBADO_Count = cnt

		Catch ex As Exception

			VBADO_Count = -1

			'DataBase切断
			VBADO_Discon(T_ADO)

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'概要  Databaseレコード検索処理
	'パラメータ：T_ADO		I/O		ADO接続用構造体
	'          ：TBName		I		削除するテーブル名
	'          ：joken		I		検索条件
	'          ：paramList	I		パラメータリスト(ADOパラメータ構造体リスト 列名とDBデータ型のみ使用)
	'          ：dataList	I		検索データリスト(1次元目:レコード、2次元目:列)
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明  条件に該当するレコードを削除
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

						'Databaseテーブル列単位でのデータ読込
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

			'DataBase切断
			VBADO_Discon(T_ADO)

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'概要  Databaseテーブル列単位でのデータ読込
	'パラメータ：dread		I/O		SQL Data Reader
	'          ：columnName	I		テーブルの列名
	'          ：sqlDbType	I		DBデータ型
	'          ：戻り値				読込んだデータ(String型)
	'説明  DBテーブルからデータ型によるデータ読込を行う
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

	'概要  Databaseレコード追加処理
	'パラメータ：T_ADO			I/O		ADO接続用構造体
	'          ：TBName			I		追加するテーブル名
	'          ：paramList		I		パラメータリスト
	'          ：isDisconnect	I		エラー時にDB切断するか否か
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明  パラメータの条件でレコードを追加
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

				'SQLコマンドパラメータ作成
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'クエリー実行
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Insert = 1

		Catch ex As Exception

			VBADO_Insert = 0

			If isDisconnect = True Then
				'DataBase切断
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'概要  Databaseレコード更新処理
	'パラメータ：T_ADO			I/O		ADO接続用構造体
	'          ：TBName			I		更新するテーブル名
	'          ：paramList		I		パラメータリスト
	'          ：isDisconnect	I		エラー時にDB切断するか否か
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明  パラメータの条件でレコードを更新
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

				'SQLコマンドパラメータ作成
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'クエリー実行
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Update = 1

		Catch ex As Exception

			VBADO_Update = 0

			If isDisconnect = True Then
				'DataBase切断
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

	'概要  Databaseレコード削除処理
	'パラメータ：T_ADO			I/O		ADO接続用構造体
	'          ：TBName			I		削除するテーブル名
	'          ：paramList		I		パラメータリスト
	'          ：isDisconnect	I		エラー時にDB切断するか否か
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明  条件に該当するレコードを削除
	Function VBADO_Delete(ByRef T_ADO As T_ADO_Struct, ByVal TBLName As String, ByVal paramList As List(Of ADO_PARAM_Struct), Optional isDisconnect As Boolean = False) As Integer

		Try
			Using cmd As New SqlClient.SqlCommand()

				'SQLコマンド設定
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

				'SQLコマンドパラメータ作成
				For Each param As ADO_PARAM_Struct In paramList
					MakeCommandParam(cmd, param)
				Next

				'クエリー実行
				cmd.ExecuteNonQuery()

			End Using

			VBADO_Delete = 1

		Catch ex As Exception

			VBADO_Delete = 0

			If isDisconnect = True Then
				'DataBase切断
				VBADO_Discon(T_ADO)
			End If

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function


	'概要  ：DataBase接続
	'ﾊﾟﾗﾒｰﾀ：T_ADO		I	T_ADO_Struct	ADO接続用構造体
	'      ：戻り値			Integer			処理結果 (1:OK / 0:NG)
	'説明  ：DataBase と接続する
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

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

			VBADO_Connect = 0

		End Try

	End Function

	'概要  ：DataBase切断
	'ﾊﾟﾗﾒｰﾀ：T_ADO		I	T_ADO_Struct	ADO接続用構造体
	'      ：戻り値			Integer			処理結果 (1:OK / 0:NG)
	'説明  ：DataBase との接続を切断する
	Function VBADO_Discon(ByRef T_ADO As T_ADO_Struct) As Integer

		Try
			T_ADO.Con.Close()
			T_ADO.Con.Dispose()

			VBADO_Discon = 1

		Catch ex As Exception

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

			VBADO_Discon = 0

		End Try

	End Function


	'概要  ：T_ADO初期化
	'ﾊﾟﾗﾒｰﾀ：T_ADO	I/O			ADO接続用環境変数
	'      ：戻り値	Integer		-------
	'説明  ：ADO接続用構造体を初期化する
	Function VBADO_T_ADOInit(ByRef T_ADO As T_ADO_Struct) As Integer

		VBADO_T_ADOInit = 1

		With T_ADO
			.DSN = "" 'ADO接続用ﾃﾞｰﾀｿｰｽ名
			.UID = "" 'ADO接続用ﾕｰｻﾞID
			.PWD = "" 'ADO接続用ﾊﾟｽﾜｰﾄﾞ
			.DBName = "" 'ADO接続用ﾃﾞｰﾀﾍﾞｰｽ名
			.Server = "" 'ADO接続用ｻｰﾊﾞ-
			.Con = Nothing 'ADO接続用ｵﾌﾞｼﾞｪｸﾄ
		End With

	End Function


	'概要  SQLコマンドパラメータ作成
	'パラメータ：cmd		I/O		SqlCommand
	'          ：param		I		ADOパラメータ
	'          ：戻り値				処理結果 (1:OK / 0:NG)
	'説明  パラメータリストよりSQLコマンドを作成
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
				MessageBox.Show("SqlDbType = " & param.SqlDbType.GetTypeCode().ToString(), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
			End If

			MakeCommandParam = 1

		Catch ex As Exception

			MakeCommandParam = 0

			MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Function

End Module