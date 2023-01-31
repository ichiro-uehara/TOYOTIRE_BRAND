Option Strict Off
Option Explicit On
Module MJ_env_get
	
	Function env_get() As Short
		
		Dim key_word(15) As String
		Dim w_str As String
		Dim i As Short
		
		key_word(1) = "BrandVBtimeout"
		key_word(2) = "BrandVBTIFFDir"
		key_word(3) = "BrandVBTmpTIFFName"
		key_word(4) = "BrandVBACADAppName"
		key_word(5) = "BrandVBACADTopic"
		key_word(6) = "BrandVBACADItem"
		
		'ﾃﾞｰﾀﾍﾞｰｽ変数の設定
		DBServer = "YAMAOKA"
		DBLoginID = "sa"
		DBpasswd = ""
		DBexample = ""
		'ブランド管理DB名称
		DBName = "brand"
		'サイズ・規格DB名称
		STANDARD_DBName = "standard"
		
		For i = 1 To 6
			w_str = Environ(key_word(i))
			If Len(w_str) = 0 Then
                MsgBox("Environment variable [" & key_word(i) & "] is not set.")
				GoTo error_section
			Else
				Select Case key_word(i)
					Case "BrandVBtimeout"
						timeOutSecond = Val(w_str)
					Case "BrandVBTIFFDir"
						TMPTIFFDir = Trim(w_str)
					Case "BrandVBTmpTIFFName"
						TmpTIFFName = Trim(w_str)
					Case "BrandVBACADAppName"
						ACADTransAppName = Trim(w_str)
					Case "BrandVBACADTopic"
						ACADTransTopic = Trim(w_str)
					Case "BrandVBACADItem"
						ACADTransItem = Trim(w_str)
				End Select
			End If
		Next i
		
		env_get = True
		Exit Function
		
error_section: 
		env_get = False
	End Function
End Module