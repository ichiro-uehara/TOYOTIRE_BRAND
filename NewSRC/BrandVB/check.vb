Option Strict Off
Option Explicit On
Module MJ_Check
	Function charnum0_check(ByRef num As String) As Short
		
		'/NULL    = 1/
		'/num < 1 = 1/
		'/ERROR   = 2/
		'/OK      = 0/
		
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			charnum0_check = 1
		Else
			If Len(num) < 1 Then
				charnum0_check = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(num, lp, 1)) = 0 Then
						charnum0_check = 2
						Exit Function
					End If
					
				Next lp
				charnum0_check = 0
			End If
		End If
		
	End Function
	'1999/05/11 yamamoto.f  英数＋小数点ﾁｪｯｸ追加 Start
	Function charnum0_1_check(ByRef num As String) As Short
		
		'/NULL    = 1/
		'/num < 1 = 1/
		'/ERROR   = 2/
		'/OK      = 0/
		
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			charnum0_1_check = 1
		Else
			If Len(num) < 1 Then
				charnum0_1_check = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ.", Mid(num, lp, 1)) = 0 Then
						charnum0_1_check = 2
						Exit Function
					End If
					
				Next lp
				charnum0_1_check = 0
			End If
		End If
		
	End Function
	'1999/05/11 yamamoto.f  英数＋小数点ﾁｪｯｸ追加 End
	
	Function charnum1_check(ByRef num As String) As Short
		
		'/NULL    = 1/
		'/num < 1 = 1/
		'/ERROR   = 2/
		'/OK      = 0/
		
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			charnum1_check = 1
		Else
			If Len(num) < 1 Then
				charnum1_check = 1
			Else
				For lp = 1 To Len(num)
					
					'1998.12.15 watanabe ”．”追加
					'            If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(num, lp, 1)) = 0 Then
					If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.", Mid(num, lp, 1)) = 0 Then
						charnum1_check = 2
						Exit Function
					End If
					
				Next lp
				charnum1_check = 0
			End If
		End If
		
	End Function
	Function charnum2_check(ByRef num As String) As Short
		
		'/NULL    = 1/
		'/num < 1 = 1/
		'/ERROR   = 2/
		'/OK      = 0/
		
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			charnum2_check = 1
		Else
			If Len(num) < 1 Then
				charnum2_check = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.", Mid(num, lp, 1)) = 0 Then
						charnum2_check = 2
						Exit Function
					End If
					
				Next lp
				charnum2_check = 0
			End If
		End If
		
	End Function
	
	Function num_check(ByRef num As String) As Short
		
		'/文字列が数字であるか判定する/
		'/    戻り値                /
		'/      ０: 数字            /
		'/      １: NULL           /
		'/      ２: エラー         /
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			num_check = 1
		Else
			If Len(num) < 1 Then
				num_check = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("0123456789", Mid(num, lp, 1)) = 0 Then
						num_check = 2
						Exit Function
					End If
					
				Next lp
				num_check = 0
			End If
		End If
		
	End Function
	
	Function num_check_1(ByRef num As String) As Short
		
		'/文字列が数字であるか判定する/ 小数点も追加 2002.08.27 Kawaguchi
		'/    戻り値                /
		'/      ０: 数字+小数点          /
		'/      １: NULL           /
		'/      ２: エラー         /
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			num_check_1 = 1
		Else
			If Len(num) < 1 Then
				num_check_1 = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("0123456789.", Mid(num, lp, 1)) = 0 Then
						num_check_1 = 2
						Exit Function
					End If
					
				Next lp
				num_check_1 = 0
			End If
		End If
		
	End Function
	
	Function char_check(ByRef num As String) As Short
		
		'/NULL    = 1/
		'/num < 1 = 1/
		'/ERROR   = 2/
		'/OK      = 0/
		
		Dim lp As Short
		
		If num Is System.DBNull.Value.ToString Then
			char_check = 1
		Else
			If Len(num) < 1 Then
				char_check = 1
			Else
				For lp = 1 To Len(num)
					
					If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(num, lp, 1)) = 0 Then
						char_check = 2
						Exit Function
					End If
					
				Next lp
				char_check = 0
			End If
		End If
		
	End Function
	
	Function check_0(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short, ByRef f As System.Windows.Forms.Control) As Short
		'
		'  英数のチェック（英字は大文字のみ）
		
		'  flg = 0 指定長さと同じかチェック
		'
		'  check_0 : 0 = ok
		'            1 = length = 0
		'            2 = ng
		'
		Dim irt As Short
		
		irt = len_check(num, nagasa, flg)
		If irt > 0 Then
			check_0 = 2
		ElseIf irt = 0 Then 
			check_0 = 1
		Else
			check_0 = charnum0_check(num)
		End If
		
		If (check_0 = 2) Then
            MsgBox("Input data are incorrect.", 64)
			f.Focus()
		End If
		
		
	End Function
	
	'1999/05/11 yamamoto.f 追加 英数＋小数点可 チェック関数 Start
	Function check_0_1(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short, ByRef f As System.Windows.Forms.Control) As Short
		'
		'  英数のチェック（英字は大文字のみ）
		
		'  flg = 0 指定長さと同じかチェック
		'
		'  check_0 : 0 = ok
		'            1 = length = 0
		'            2 = ng
		'
		Dim irt As Short
		
		irt = len_check(num, nagasa, flg)
		If irt > 0 Then
			check_0_1 = 2
		ElseIf irt = 0 Then 
			check_0_1 = 1
		Else
			check_0_1 = charnum0_1_check(num)
		End If
		
		If (check_0_1 = 2) Then
            MsgBox("Input data are incorrect.", 64)
			f.Focus()
		End If
		
		
	End Function
	'1999/05/11 yamamoto.f 追加 英数＋小数点可 チェック関数 End
	
	Function check_1(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short, ByRef f As System.Windows.Forms.Control) As Short
		'
		'  英数のチェック（英字は大文字のみ）
		'
		'  flg = 0 指定長さと同じかチェック
		'
		'  check_1 : 0 = ok
		'            1 = length = 0
		'            2 = ng
		'
		Dim irt As Short
		
		irt = len_check(num, nagasa, flg)
		If irt > 0 Then
			check_1 = 2
		ElseIf irt = 0 Then 
			check_1 = 1
		Else
			check_1 = charnum1_check(num)
		End If
		
		If check_1 = 2 Then
            MsgBox("Input data are incorrect.", 64)
			f.Focus()
		End If
		
	End Function
	
	Function check_2(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short, ByRef f As System.Windows.Forms.Control) As Short
		'
		'  英数のチェック
		'
		'  flg = 0 指定長さと同じかチェック
		'
		'  check_2 : 0 = ok
		'            1 = length = 0
		'            2 = ng
		'
		Dim irt As Short
		
		irt = len_check(num, nagasa, flg)
		If irt > 0 Then
			check_2 = 2
		ElseIf irt = 0 Then 
			check_2 = 1
		Else
			check_2 = charnum2_check(num)
		End If
		
		If check_2 = 2 Then
            MsgBox("Input data are incorrect.", 64)
			f.Focus()
		End If
		
	End Function
	
	Function len_check2(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short) As Short
		'
		'  unicode 対応（２バイトコード）
		'
		'  flg = 0 指定長さと同じかチェック
		'
		'  len_check2:-1 チェックがＯＫ
		'            :以外は文字の長さをセット
		'
		If num Is System.DBNull.Value.ToString Then
			len_check2 = 0
		Else
			'      len_check2 = LenB(StrConv(num, vbUnicode))
            'len_check2 = LenB(StrConv(num, vbFromUnicode))
            len_check2 = System.Text.Encoding.GetEncoding(932).GetByteCount(num) '20100616移植追加
		End If
		
		If flg = 0 Then
			If len_check2 = nagasa Then
				len_check2 = -1
			End If
		Else
			If len_check2 <= nagasa Then
				len_check2 = -1
			End If
		End If
		
	End Function
	Function float_check2(ByRef num As String, ByRef nagasa As Short, ByRef s_ika As Short) As Short
		'
		'  unicode 対応（２バイトコード）
		'
		'
		'  float_check2:-1 チェックがＯＫ
		'            :以外は文字の長さをセット
		'
		Dim lp As Short
		Dim s_point As Short
		Dim s_num As Short
		
		If num Is System.DBNull.Value.ToString Then
			float_check2 = 0
		Else
            'float_check2 = LenB(StrConv(num, vbFromUnicode))
            float_check2 = System.Text.Encoding.GetEncoding(932).GetByteCount(num) '20100616移植追加
		End If
		
		If float_check2 <= nagasa Then
			
			s_point = Len(num) + 1
			s_num = 0
			For lp = 1 To Len(num)
				
				If InStr("0123456789.", Mid(num, lp, 1)) = 0 Then
					Exit Function
				ElseIf InStr("0123456789.", Mid(num, lp, 1)) = 11 Then 
					If s_num <> 0 Then Exit Function
					s_num = s_num + 1
					s_point = lp
					
				End If
				
			Next lp
			' 小数点以上は
			If s_point - 1 > nagasa - s_ika - 1 Then Exit Function
			
			' 小数点以下は
			If Len(num) - s_point > s_ika Then Exit Function
			
			float_check2 = -1
			
			
		End If
		
	End Function
	
	Function len_check(ByRef num As String, ByRef nagasa As Short, ByRef flg As Short) As Short
		'
		'  flg = 0 指定長さと同じかチェック
		'
		'  len_check:-1 チェックがＯＫ
		'           :以外は文字の長さをセット
		'
		If num Is System.DBNull.Value.ToString Then
			len_check = 0
		Else
			len_check = Len(num)
		End If
		
		If flg = 0 Then
			If len_check = nagasa Then
				len_check = -1
			End If
		Else
			If len_check <= nagasa Then
				len_check = -1
			End If
		End If
		
	End Function
	
	Function apos_check(ByRef check_text As String) As String
		
		Dim i As Short
		Dim moji_no As Integer
		
		For i = 0 To Len(Trim(check_text))
			
			moji_no = InStr(check_text, "'")
			If moji_no > 0 Then
				Mid(check_text, moji_no, 1) = "#"
			Else
				Exit For
			End If
			
		Next i
		
		apos_check = check_text
		
	End Function
End Module