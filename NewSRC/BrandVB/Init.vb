Option Strict Off
Option Explicit On
Module MJ_Init

    ' -> watanabe add VerUP(2011)
    'ウインドウズのプログラムを呼び出すための宣言
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    ' <- watanabe add VerUP(2011)


	'----------------< 設定ファイルの読込み >----------------------------
	Function config_read(ByRef FileName As String) As Short
		Dim w_str As String
		Dim key_word As String
		Dim key_value As String
		Dim i As Object
		Dim j As Short
		Dim c As String
		Dim fn As Short
		
		On Error GoTo error_section
		
		fn = FreeFile
		FileOpen(fn, FileName, OpenMode.Input)
		
		Do While Not EOF(fn)
			w_str = LineInput(fn)
			'MsgBox w_str
			If Mid(Trim(w_str), 1, 1) = "'" Then GoTo LOOP_CONT 'コメント行
			If Mid(Trim(w_str), 1, 1) = "[" Then GoTo LOOP_CONT 'ブロック名
            If IsDBNull(Mid(Trim(w_str), 1, 1)) Then GoTo LOOP_CONT 'NULL行
			'キーワード検索
			key_word = ""
			key_value = ""
			For i = 1 To Len(w_str)
				c = Mid(w_str, i, 1)
				Select Case c
					Case " "
					Case "="
						Exit For
					Case Else
						key_word = key_word & c
				End Select
			Next i
			'キーワード検索
			For j = i + 1 To Len(w_str)
				c = Mid(w_str, j, 1)
				If c <> """" Then
					key_value = key_value & c
				End If
			Next j
			key_word = Trim(key_word)
			key_value = Trim(key_value)
			
			Select Case key_word
				Case ""
				Case "timeOutSecond"
					timeOutSecond = CShort(key_value)
				Case "TIFFDir"
					TIFFDir = key_value
				Case "TmpTIFFName"
					TmpTIFFName = key_value
				Case "DBServer"
					DBServer = key_value
				Case "DBLoginID"
					DBLoginID = key_value
				Case "DBpasswd"
					DBpasswd = key_value
				Case "DBexample"
					DBexample = key_value
				Case "DBName"
					DBName = key_value
				Case "ACADTransAppName"
					ACADTransAppName = key_value
				Case "ACADTransTopic"
					ACADTransTopic = key_value
				Case "ACADTransItem"
					ACADTransItem = key_value
				Case Else
                    MsgBox("Keyword [" & key_word & "] is invalid.", MsgBoxStyle.Critical, "Read error")
					GoTo error_section
			End Select
LOOP_CONT: 
		Loop 
		
		FileClose(1)
		
		config_read = True
		Exit Function
		
error_section: 
		
		config_read = False
		Exit Function
		
	End Function
	
	'----------------< 設定ファイルの読込み >----------------------------
	Function set_read(ByRef FileName As String) As Short
		Dim key_value As New VB6.FixedLengthString(255)
		Dim work As Short
		
        'ディレクトリ取得
        'UPGRADE_WARNING: オブジェクトの既定プロパティを解決できませんでした。
		work = GetPrivateProfileString("DIRECTORY", "gensi_model", "", key_value.Value, Len(key_value.Value), FileName)
        GensiDir = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "hensyu_model", "", key_value.Value, Len(key_value.Value), FileName)
		HensyuDir = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "gensi_zumen", "", key_value.Value, Len(key_value.Value), FileName)
		KokuinDir = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "hensyu_zumen", "", key_value.Value, Len(key_value.Value), FileName)
		HensyuZumenDir = Trim(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		work = GetPrivateProfileString("DIRECTORY", "brand_zumen", "", key_value.Value, Len(key_value.Value), FileName)
		BrandDir = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "view_tiff1", "", key_value.Value, Len(key_value.Value), FileName)
		TIFFDirGM = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "view_tiff2", "", key_value.Value, Len(key_value.Value), FileName)
		TIFFDirHM = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DIRECTORY", "Help_File", "", key_value.Value, Len(key_value.Value), FileName)
		HelpFileName = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		'BrandVB システム定数取得
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBtimeout", "", key_value.Value, Len(key_value.Value), FileName)
		timeOutSecond = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBTIFFDir", "", key_value.Value, Len(key_value.Value), FileName)
		TMPTIFFDir = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBTmpTIFFName", "", key_value.Value, Len(key_value.Value), FileName)
		TmpTIFFName = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBACADAppName", "", key_value.Value, Len(key_value.Value), FileName)
		ACADTransAppName = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBACADTopic", "", key_value.Value, Len(key_value.Value), FileName)
		ACADTransTopic = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("VBSYSTEM", "BrandVBACADItem", "", key_value.Value, Len(key_value.Value), FileName)
		ACADTransItem = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		'データベース定数取得
		work = GetPrivateProfileString("DB", "DBServer", "", key_value.Value, Len(key_value.Value), FileName)
		DBServer = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DB", "DBLoginID", "", key_value.Value, Len(key_value.Value), FileName)
		DBLoginID = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DB", "DBpasswd", "", key_value.Value, Len(key_value.Value), FileName)
		DBpasswd = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DB", "DBexample", "", key_value.Value, Len(key_value.Value), FileName)
		DBexample = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DB", "DBName", "", key_value.Value, Len(key_value.Value), FileName)
		DBName = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("DB", "STANDARD_DBName", "", key_value.Value, Len(key_value.Value), FileName)
		STANDARD_DBName = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		'テンプレート 設定ファイル名取得
		'テンプレート タイプ １
		'work = GetPrivateProfileString("TMPFILENAME", "Size1", "", key_value, Len(key_value), FileName)
		'Tmp_Size1_ini = Trim(Left(key_value, InStr(key_value, Chr$(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Load1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Load1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Pattern1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Pattern1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Serial1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Serial1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Mold_no1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Mold_no1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "E_no1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_E_no1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Utqg1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Utqg1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Maxload1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Maxload1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Ply1", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Ply1_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Ply2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Ply2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "etc", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_ETC_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
        '2014/12/15 moriya add start
        work = GetPrivateProfileString("TMPFILENAME", "MARK", "", key_value.Value, Len(key_value.Value), FileName)
        Tmp_MARK_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
        '2014/12/15 moriya add end

		'テンプレート タイプ ２
		work = GetPrivateProfileString("TMPFILENAME", "Size2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Size2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Load2S", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Load2S_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Load2D", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Load2D_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Pattern2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Pattern2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Lt2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Lt2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Pr2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Pr2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Psi2", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Psi2_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		
		' -> watanabe add 2007.03
		
		'テンプレート タイプ ３
		work = GetPrivateProfileString("TMPFILENAME", "Utqg3", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Utqg3_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Maxload3", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Maxload3_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Ply1_3", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Ply1_3_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "Ply2_3", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Ply2_3_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		work = GetPrivateProfileString("TMPFILENAME", "etc3", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_ETC3_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		' <- watanabe add 2007.03
		
		
		'テンプレート プレート
		work = GetPrivateProfileString("TMPFILENAME", "Plate", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_Plate_ini = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		'テンプレート タイプ２ 作成用共通データ
		work = GetPrivateProfileString("TMP2DATA", "DummyHM", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp2_Dummy_HM = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
		
		AskNum = 0
		
		set_read = True
		Exit Function
		
error_section: 
		
		set_read = False
		Exit Function
		
	End Function
	
	'----------------< 設定ファイルの読込み２ テンプレート フォントデータ >----------------------------
	Function set_read2(ByRef FileName As String, ByRef temp_name As String) As Short
		Dim key_value As New VB6.FixedLengthString(255)
		Dim key_word As String
		Dim work As Short
		Dim i As Short
		
		If temp_name = "serial1" Or temp_name = "serial2" Then
			work = GetPrivateProfileString("DATA", "TmpSerialWidth", "", key_value.Value, Len(key_value.Value), FileName)
			TmpSerialWidth = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
			work = GetPrivateProfileString("DATA", "TmpSerialMove", "", key_value.Value, Len(key_value.Value), FileName)
			TmpSerialMove = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		End If
		
		work = GetPrivateProfileString("FONT", "allcnt", "", key_value.Value, Len(key_value.Value), FileName)
		Tmp_font_cnt = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		
		For i = 1 To Tmp_font_cnt + 1
			key_word = "num" & i
			
			work = GetPrivateProfileString("FONT", key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			If Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)) = "ERROR" Then
				Exit For
			End If
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			Tmp_font_word(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_font_size(i) = Val(Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1)))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_font_block(i) = Val(Trim(key_value.Value))
		Next i
		
		set_read2 = True
		Exit Function
		
error_section: 
		
		set_read2 = False
		Exit Function
		
	End Function
	
	'----------------< 設定ファイルの読込み３ テンプレート タイプ１ >----------------------------
	Function set_read3(ByRef FileName As String, ByRef temp_name As String, ByRef block_no As Short) As Short
		Dim key_value As New VB6.FixedLengthString(255)
		Dim key_block As String
		Dim key_word As String
		Dim gm_word As String
		Dim gm_code As String
		Dim work As Short
		Dim i As Short
		
		key_block = "TYPE" & block_no
		
		work = GetPrivateProfileString(key_block, "replace", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			ReplaceMode = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		End If
		
		For i = 1 To MaxSelNum
			Tmp_hm_word(i) = ""
			Tmp_hm_code(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "hmtype" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) = "ERROR" Then
				Exit For
			End If
			Tmp_hm_word(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_code(i) = Trim(key_value.Value)
		Next i
		
		For i = 0 To 26
			GensiALPH(i) = ""
		Next i
		
		For i = 0 To 26
			GensiALPHS(i) = ""
		Next i
		
		For i = 0 To 10
			GensiNUM(i) = ""
		Next i
		
		For i = 0 To 128
			GensiKIGO(i) = ""
		Next i
		
		'Brand Ver4.0追加
		For i = 0 To 26
			GensiALPHS(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "gmcode" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) <> "ERROR" Then
				gm_word = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
				key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
				gm_code = Trim(key_value.Value)
				If Len(Trim(gm_word)) = 1 And Asc("A") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("Z") Then
					GensiALPH(Asc(Trim(gm_word)) - Asc("A")) = gm_code
					'Brand Ver4.0 追加
				ElseIf Len(Trim(gm_word)) = 1 And Asc("a") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("z") Then 
					GensiALPHS(Asc(Trim(gm_word)) - Asc("a")) = gm_code
				ElseIf Len(Trim(gm_word)) = 1 And Asc("0") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("9") Then 
					GensiNUM(Asc(Trim(gm_word)) - Asc("0")) = Trim(gm_code)
					'Brand Ver4.0 変更
				ElseIf Len(Trim(gm_word)) = 1 Then 
					GensiKIGO(Asc(Trim(gm_word))) = Trim(gm_code)
				End If
			End If
		Next i
		
		set_read3 = True
		Exit Function
		
error_section: 
		
		set_read3 = False
		Exit Function
		
	End Function
	
	'----------------< 設定ファイルの読込み４ テンプレート タイプ２ >----------------------------
	Function set_read4(ByRef FileName As String, ByRef temp_name As String, ByRef block_no As Short) As Short
		Dim key_value As New VB6.FixedLengthString(255)
		Dim key_block As String
		Dim key_word As String
		Dim gm_word As String
		Dim gm_code As String
		Dim work As Short
		Dim i As Short
		
		key_block = "TYPE" & block_no
		
		work = GetPrivateProfileString(key_block, "replace", "", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			ReplaceMode = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		End If
		
		'    For i = 1 To MaxSelNum
		'        Tmp_rule_word(i) = ""
		'        Tmp_rule_type(i) = ""
		'        Tmp_rule_x(i) = 0
		'        Tmp_rule_y(i) = 0
		'    Next i
		
		'    For i = 1 To MaxSelNum
		'        key_word = "rule" & i
		'        work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value, Len(key_value), FileName)
		'        key_value = Trim(Left(key_value, InStr(key_value, Chr$(0)) - 1))
		'        If Trim(key_value) = "ERROR" Then
		'            Exit For
		'        End If
		'        Tmp_rule_word(i) = Trim(Left(key_value, InStr(key_value, "|") - 1))
		'        key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		'        If Trim(Tmp_rule_word(i)) <> "all" Then
		'            Tmp_rule_type(i) = Trim(Left(key_value, InStr(key_value, "|") - 1))
		'        Else
		'            Tmp_rule_type(i) = "all"
		'        End If
		'        key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		'        Tmp_rule_x(i) = Val(Trim(Left(key_value, InStr(key_value, "|") - 1)))
		'        Tmp_rule_y(i) = Val(Trim(Mid(key_value, InStr(key_value, "|") + 1)))
		'    Next i
		
		
		For i = 0 To 26
			GensiALPH(i) = ""
		Next i
		
		For i = 0 To 26
			GensiALPHS(i) = ""
		Next i
		
		For i = 0 To 10
			GensiNUM(i) = ""
		Next i
		
		For i = 0 To 128
			GensiKIGO(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "gmcode" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) <> "ERROR" Then
				gm_word = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
				gm_code = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
				If Len(Trim(gm_word)) = 1 And Asc("A") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("Z") Then
					GensiALPH(Asc(Trim(gm_word)) - Asc("A")) = Trim(gm_code)
					'Brand Ver4.0 追加
				ElseIf Len(Trim(gm_word)) = 1 And Asc("a") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("z") Then 
					GensiALPHS(Asc(Trim(gm_word)) - Asc("a")) = gm_code
				ElseIf Len(Trim(gm_word)) = 1 And Asc("0") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("9") Then 
					GensiNUM(Asc(Trim(gm_word)) - Asc("0")) = Trim(gm_code)
					'Brand Ver4.0 変更
				ElseIf Len(Trim(gm_word)) = 1 Then 
					GensiKIGO(Asc(Trim(gm_word))) = Trim(gm_code)
				End If
			End If
		Next i
		
		set_read4 = True
		Exit Function
		
error_section: 
		
		set_read4 = False
		Exit Function
		
	End Function
	
	'----------------< 設定ファイルの読込み５ テンプレート タイプ１－２>----------------------------
	Function set_read5(ByRef FileName As String, ByRef temp_name As String, ByRef block_no As Short) As Short
		
		Dim key_value As New VB6.FixedLengthString(255)
		Dim key_block As String
		Dim key_word As String
		Dim gm_word As String
		Dim gm_code As String
		Dim work As Short
		Dim i As Short
		
		key_block = "TYPE" & block_no
		
		work = GetPrivateProfileString(key_block, "replace", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			ReplaceMode = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		End If
		
		For i = 1 To MaxSelNum
			Tmp_prcs_code(i) = ""
			Tmp_hm_word(i) = ""
			Tmp_hm_code(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "hmtype" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) = "ERROR" Then
				Exit For
			End If
			Tmp_prcs_code(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_word(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_code(i) = Trim(key_value.Value)
		Next i
		
		For i = 0 To 26
			GensiALPH(i) = ""
		Next i
		
		For i = 0 To 26
			GensiALPHS(i) = ""
		Next i
		
		For i = 0 To 10
			GensiNUM(i) = ""
		Next i
		
		For i = 0 To 128
			GensiKIGO(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "gmcode" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) <> "ERROR" Then
				gm_word = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
				key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
				gm_code = Trim(key_value.Value)
				If Len(Trim(gm_word)) = 1 And Asc("A") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("Z") Then
					GensiALPH(Asc(Trim(gm_word)) - Asc("A")) = gm_code
					'Brand Ver4.0 追加
				ElseIf Len(Trim(gm_word)) = 1 And Asc("a") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("z") Then 
					GensiALPHS(Asc(Trim(gm_word)) - Asc("a")) = gm_code
				ElseIf Len(Trim(gm_word)) = 1 And Asc("0") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("9") Then 
					GensiNUM(Asc(Trim(gm_word)) - Asc("0")) = Trim(gm_code)
					'Brand Ver4.0 変更
				ElseIf Len(Trim(gm_word)) = 1 Then 
					GensiKIGO(Asc(Trim(gm_word))) = Trim(gm_code)
				End If
			End If
		Next i
		
		' -> watanabe add 2007.03
		work = GetPrivateProfileString("EDIT", "plate_w", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			Tmp_plate_w = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		Else
			Tmp_plate_w = 0#
		End If
		
		work = GetPrivateProfileString("EDIT", "plate_h", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			Tmp_plate_h = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		Else
			Tmp_plate_h = 0#
		End If
		
		work = GetPrivateProfileString("EDIT", "plate_r", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			Tmp_plate_r = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		Else
			Tmp_plate_r = 0#
		End If
		
		work = GetPrivateProfileString("EDIT", "plate_n", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			Tmp_plate_n = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		Else
			Tmp_plate_n = 0#
		End If
		' <- watanabe add 2007.03
		
		set_read5 = True
		Exit Function
		
error_section: 
		
		set_read5 = False
		Exit Function
		
	End Function
	
	
	' -> watanabe add 2007.03
	
	'----------------< 設定ファイルの読込み６ テンプレート タイプ３>----------------------------
	Function set_read6(ByRef FileName As String, ByRef temp_name As String, ByRef block_no As Short) As Short
		Dim key_value As New VB6.FixedLengthString(255)
		Dim key_block As String
		Dim key_word As String
		Dim gm_word As String
		Dim gm_code As String
		Dim work As Short
		Dim i As Short
		
		key_block = "TYPE" & block_no
		
		work = GetPrivateProfileString(key_block, "replace", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			ReplaceMode = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		End If
		
		For i = 1 To MaxSelNum
			Tmp_prcs_code(i) = ""
			Tmp_hm_word(i) = ""
			Tmp_hm_code(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "hmtype" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) = "ERROR" Then
				Exit For
			End If
			
			Tmp_prcs_code(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_word(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_code(i) = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
			key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
			Tmp_hm_group(i) = Trim(key_value.Value)
		Next i
		
		For i = 0 To 26
			GensiALPH(i) = ""
		Next i
		
		For i = 0 To 26
			GensiALPHS(i) = ""
		Next i
		
		For i = 0 To 10
			GensiNUM(i) = ""
		Next i
		
		For i = 0 To 128
			GensiKIGO(i) = ""
		Next i
		
		For i = 1 To MaxSelNum
			key_word = "gmcode" & i
			work = GetPrivateProfileString(key_block, key_word, "ERROR", key_value.Value, Len(key_value.Value), FileName)
			key_value.Value = Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1))
			If Trim(key_value.Value) <> "ERROR" Then
				gm_word = Trim(Left(key_value.Value, InStr(key_value.Value, "|") - 1))
				key_value.Value = Trim(Mid(key_value.Value, InStr(key_value.Value, "|") + 1))
				gm_code = Trim(key_value.Value)
				If Len(Trim(gm_word)) = 1 And Asc("A") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("Z") Then
					GensiALPH(Asc(Trim(gm_word)) - Asc("A")) = gm_code
				ElseIf Len(Trim(gm_word)) = 1 And Asc("a") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("z") Then 
					GensiALPHS(Asc(Trim(gm_word)) - Asc("a")) = gm_code
				ElseIf Len(Trim(gm_word)) = 1 And Asc("0") <= Asc(Trim(gm_word)) And Asc(Trim(gm_word)) <= Asc("9") Then 
					GensiNUM(Asc(Trim(gm_word)) - Asc("0")) = Trim(gm_code)
				ElseIf Len(Trim(gm_word)) = 1 Then 
					GensiKIGO(Asc(Trim(gm_word))) = Trim(gm_code)
				End If
			End If
		Next i
		
		' -> watanabe add 2007.06
		work = GetPrivateProfileString(key_block, "TMPETC3", "ERROR", key_value.Value, Len(key_value.Value), FileName)
		If Trim(key_value.Value) <> "ERROR" Then
			Tmp_brd_no = Val(Trim(Left(key_value.Value, InStr(key_value.Value, Chr(0)) - 1)))
		Else
			Tmp_brd_no = 0
		End If
		' <- watanabe add 2007.06
		
		set_read6 = True
		Exit Function
		
error_section: 
		set_read6 = False
	End Function
	
	' <- watanabe add 2007.03
	
	
	'----- 12/10 1997 yamamoto start -----
	'ﾃﾞﾊﾞｯｸﾞ用
	Function output_command_line(ByRef commnad_line As String) As Short
		
		Dim fname As String
		Dim fno As Object
		
		fname = "c:\acad11\exe\output_command_line.txt"
		
		fno = FreeFile
		
		On Error GoTo error_section
		FileOpen(fno, fname, OpenMode.Append)
		On Error Resume Next
		
		PrintLine(fno, Now, "    [", commnad_line, "]")
		
		FileClose(fno)
		
		Exit Function
		
error_section: 
		
        MsgBox("file open error [" & fname & "]")
		
	End Function
	'----- 12/10 1997 yamamoto end -------
End Module