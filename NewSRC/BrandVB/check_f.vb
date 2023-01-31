Option Strict Off
Option Explicit On
Module MJ_CheckForm
	
	Function check_F_GMSAVE() As Short
		Dim irt2 As Object
		Dim irt1 As Object
		'
		'check_form_no = 0: OK
		'
		Dim irt As Short
		Dim f As System.Windows.Forms.Control
		'***********************************************
		'12/5.1997 yamamoto start
		'  新規に no_data_sectonを作りました
		'  ﾌｫｰﾑのﾛｽﾄﾌｫｰｶｽﾌﾟﾛｼｰｼﾞｬを持ってきて修正
		
        If Trim(form_no.w_font_name.Text) = "" Then GoTo no_data_section
		'   If Trim(form_no.w_font_class2.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_name2.Text) = "" Then GoTo no_data_section

        If Trim(form_no.w_dep_name.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_entry_name.Text) = "" Then GoTo no_data_section
        If Trim(form_no.w_entry_date.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_high.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_width.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_ang.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_moji_high.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_moji_shift.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hem_width.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hatch_ang.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hatch_width.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hatch_space.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hatch_x.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_hatch_y.Text) = "" Then GoTo no_data_section
        If Trim(form_no.w_org_hor.Text) = "" Then GoTo no_data_section
        If Trim(form_no.w_org_ver.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_base_r.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_font_class1.Text) = "" Then GoTo no_data_section
		If Trim(form_no.w_name1.Text) = "" Then GoTo no_data_section
		
		'/フォント名/
        f = form_no.w_font_name
		irt = check_0(form_no.w_font_name.Text, 6, 0, f)
        If irt <> 2 And Left(form_no.w_font_name.Text, 2) <> "KO" Then
            MsgBox("Font name is 6 characters starting with KO.", 64)
            f.Focus()
            GoTo error_section
        ElseIf irt = 2 Then
            GoTo error_section
        End If
		
		'/文字名2/
		f = form_no.w_name2
		irt = check_0(form_no.w_name2.Text, 1, 0, f)
		If irt = 0 Then
            irt1 = len_check(form_no.w_name1.Text, 255, 1)
            If irt1 <> -1 Then
                MsgBox("Character name 1 has not been set.", 64)
                f.Focus()
                GoTo error_section
            End If
			
            If Mid(form_no.w_name1.Text, 1, 1) = "A" Then
                irt2 = char_check(form_no.w_name2.Text)
            ElseIf Mid(form_no.w_name1.Text, 1, 1) = "B" Then
                irt2 = num_check(form_no.w_name2.Text)
            Else
                irt2 = 0
            End If
			
			If irt2 = 2 Then
                MsgBox("Appropriate value has not been set.", 64)
				f.Focus()
				GoTo error_section
			End If
			'99/8/17 kitamura Add
		Else
			GoTo error_section
		End If
		
		
		'/コメント/
		f = form_no.w_comment
		
        irt = len_check2(form_no.w_comment.Text, 255, 1)
		If irt > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
			f.Focus()
			GoTo error_section
		End If
		
		'/部署/
        f = form_no.w_dep_name
        irt = check_0(form_no.w_dep_name.Text, 2, 1, f)
        If irt = 2 Then GoTo error_section

        '/登録日/
        f = form_no.w_entry_name
        irt = check_1(form_no.w_entry_name.Text, 4, 1, f)
        If irt = 0 Then
            irt = num_check(form_no.w_entry_name.Text)
            If irt = 2 Then
                MsgBox("Input data (Record Date) is incorrect.", 64)
                f.Focus()
            End If
        ElseIf irt = 2 Then
            GoTo error_section
		End If

        '/基準R/
        f = form_no.w_base_r
        irt = check_2(form_no.w_base_r.Text, 9, 1, f)
		If irt = 0 Then
            irt2 = float_check2(form_no.w_base_r.Text, 9, 4)
            If irt2 > 0 Then
                MsgBox("Input data (Base R) is incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
		ElseIf irt = 2 Then 
			GoTo error_section
		End If
		
		'/縁取り幅/
        f = form_no.w_hem_width
        irt = check_2(form_no.w_hem_width.Text, 5, 1, f)
		If irt = 0 Then
            irt2 = float_check2(form_no.w_hem_width.Text, 5, 2)
            If irt2 > 0 Then
                MsgBox("Input data (Border width) is incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
            If Left(Trim(form_no.w_font_class1.Text), 1) = "F" Or Left(Trim(form_no.w_font_class1.Text), 1) = "B" Then
                If Val(form_no.w_hem_width.Text) = 0 Then
                    MsgBox("Input data (Border width) is incorrect.", 64)
                    f.Focus()
                    GoTo error_section
                End If
            End If
		ElseIf irt = 2 Then 
			GoTo error_section
		End If
		
		
		'/旧フォント区分名/
        f = form_no.w_old_font_class
        irt = check_0(form_no.w_old_font_class.Text, 2, 1, f)
		If irt = 2 Then GoTo error_section
		
		'/旧フォント名/
        f = form_no.w_old_font_name
        irt = check_0(form_no.w_old_font_name.Text, 6, 0, f)
		If irt = 2 Then GoTo error_section
		
		'/旧文字/
        f = form_no.w_old_name
        irt = check_0(form_no.w_old_name.Text, 2, 1, f)
		If irt = 2 Then GoTo error_section
        irt = check_0(form_no.w_font_name.Text, 6, 0, form_no.w_font_name)
		If irt = 2 Then GoTo error_section
		
		check_F_GMSAVE = 0
		Exit Function
		
no_data_section: 
        MsgBox("There are undefined data.", 64)

error_section:
        check_F_GMSAVE = 1
		
    End Function

	Function check_F_TMP_ENO() As Short
		Dim i As Object
		'
		'check_F_TMP_ENO = 0: OK
		'
		'エラーとなった項目へのセットフォーカスも同時に行う
        'Dim ss As String '20100616移植削除
		Dim slen As Short
		Dim w_ret As Short
        Dim wstr As String

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)

		'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If
		
		
		'20000126 追加
		For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                wstr = Tmp_prcs_code(i)
                Exit For
            End If
		Next i
		
		
		' -> watanabe Add 2007.03
        'If form_no.chk_shonin.Value = 1 Then
        If form_no.chk_shonin.CheckState = 1 Then
            If Trim(form_no.w_shonin.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_shonin.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            '承認番号のチェック
            '   長さチェック
            slen = Len(Trim(form_no.w_shonin.Text))
            '20000126 修正
            '   If form_no.w_type.Text = "E4" Then
            '       If slen < 6 Or slen > 7 Then
            '           MsgBox "入力データが不適切です。" & Chr$(13) & "承認番号は６～７桁です。", 64
            '           form_no.w_shonin.SetFocus
            '           GoTo error_section
            '       End If
            '   ElseIf form_no.w_type.Text = "E5" Then
            '       If slen <> 5 Then
            '           MsgBox "入力データが不適切です。" & Chr$(13) & "承認番号は５桁です。", 64
            '           form_no.w_shonin.SetFocus
            '           GoTo error_section
            '       End If
            '   End If
            If Trim(wstr) = "ENO5" Then
                If slen <> 5 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "The approval number is 5 characters.", 64)
                    form_no.w_shonin.Focus() 'SetFocus()
                    GoTo error_section
                End If
            ElseIf Trim(wstr) = "ENO6" Then
                If slen <> 6 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "The approval number is 6 characters.", 64)
                    form_no.w_shonin.Focus() 'SetFocus()
                    GoTo error_section
                End If
            ElseIf Trim(wstr) = "ENO7" Then
                If slen <> 7 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "The approval number is 7 characters.", 64)
                    form_no.w_shonin.Focus() 'SetFocus()
                    GoTo error_section
                End If
            ElseIf Trim(wstr) = "ENO8" Then
                If slen <> 8 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "The approval number is 8 characters.", 64)
                    form_no.w_shonin.Focus() 'SetFocus()
                    GoTo error_section
                End If
            Else
                MsgBox("Contents of the configuration file is incorrect." & Chr(13) & "Type", 64)
                form_no.w_type.Focus() 'SetFocus()
                GoTo error_section
            End If

            '   データチェック
            w_ret = num_check(form_no.w_shonin.Text)
            If w_ret <> 0 Then
                MsgBox("Input data are incorrect." & Chr(13) & "Approval number is numeric only.", 64)
                form_no.w_shonin.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03


        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        'If form_no.ImgThumbnail1.Picture = 0 Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        '2011/12/08 uriu add start
        'Ｓ番号のチェック
        If form_no.w_s.Enabled = True Then
            If form_no.chk_s.CheckState = 1 Then
                If Trim(form_no.w_s.Text) <> "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                    form_no.w_s.Focus() 'SetFocus()
                    GoTo error_section
                End If
            Else
                w_ret = num_check(form_no.w_s.Text)
                If w_ret <> 0 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "S number is numerical value only.", 64)
                    form_no.w_s.Focus() 'SetFocus()
                    GoTo error_section
                End If
            End If
        End If

        'Ｒ番号のチェック
        If form_no.w_r.Enabled = True Then
            If form_no.chk_r.CheckState = 1 Then
                If Trim(form_no.w_r.Text) <> "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                    form_no.w_r.Focus() 'SetFocus()
                    GoTo error_section
                End If
            Else
                w_ret = num_check(form_no.w_r.Text)
                If w_ret <> 0 Then
                    MsgBox("Input data are incorrect." & Chr(13) & "R number is numerical value only.", 64)
                    form_no.w_r.Focus() 'SetFocus()
                    GoTo error_section
                End If
            End If
        End If
        '2011/12/08 uriu add end

        check_F_TMP_ENO = 0
        Exit Function

error_section:
        check_F_TMP_ENO = 1

    End Function

    '***** 12/9 1997 yamamoto start *****
    '新規に追加
    Function check_F_TMP_PTNCODE() As Short
        'check_F_TMP_PTNCODE = 0: OK
        'エラーとなった項目へのセットフォーカスも同時に行う
        'Dim ss As String'20100616移植削除
        Dim slen As Short
        'Dim w_ret As Short'20100616移植削除
        Dim i As Short
        Dim w_str As String
        Dim w_msg As String

        '/pattern codeのチェック/
        ' 長さチェック
        slen = Len(form_no.w_ptncode.Text)
        If slen > 6 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Pattern code is 6 characters.", 64)
            form_no.w_ptncode.Focus() 'SetFocus()
            GoTo error_section
        End If
        ' データチェック
        w_str = Left(form_no.w_ptncode.Text, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", w_str, 0) = 0 Then
            w_msg = "Input data are incorrect." & Chr(13)
            w_msg = w_msg & "Pattern code first character only English letter (capital letter) or numerical value."
            MsgBox(w_msg, 64)
            GoTo error_section
        End If

        w_str = Mid(form_no.w_ptncode.Text, 2)
        For i = 1 To Len(w_str)
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+-", Mid(w_str, i, 1), 0) = 0 Then
                w_msg = "Input data are incorrect." & Chr(13)
                w_msg = w_msg & "Pattern cord only as for the English letter( capital letter), the numerical value or the sign(+ or -)."
                MsgBox(w_msg, 64)
                GoTo error_section
            End If
        Next i


        '/タイプのチェック/
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        '/ピクチャのチェック/
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        'If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then '20100629コード変更
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PTNCODE = 0
        Exit Function

error_section:
        check_F_TMP_PTNCODE = 1

    End Function
    Function check_F_TMP_SERIARU() As Short
        '
        'check_F_TMP_SERIARU = 0: OK
        '
        'エラーとなった項目へのセットフォーカスも同時に行う

        If w_size_chk(0) <> 0 Then
            GoTo error_section
        End If

        If Trim(form_no.w_syurui.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Tire type", 64)
            form_no.w_syurui.Focus() 'SetFocus()
            GoTo error_section
        End If
        If Trim(form_no.w_plant.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plant", 64)
            form_no.w_plant.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_SERIARU = 0
        Exit Function

error_section:
        check_F_TMP_SERIARU = 1

    End Function

    'パラメータ
    'syori_flg  :1 = SQL検索時
    '            2 = 置換読込時
    '
    'check_form_no = 0: OK
    '
    'エラーとなった項目へのセットフォーカスも同時に行う

    Function check_F_TMP_MAXLOAD(ByRef syori_flg As Short) As Short
        Dim slen As Object
        Dim w_ret As Object
        'Dim irt As Short'20100616移植削除
        Dim check_flg As Short
        Dim ret As Short
        Dim wstr As String

        Select Case syori_flg
            Case 1
                If Trim(form_no.w_syurui.Text) = "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "Tire type", 64)
                    form_no.w_syurui.Focus() 'SetFocus()
                    GoTo error_section
                End If
                If Trim(form_no.w_type.Text) = "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
                    form_no.w_type.Focus() 'SetFocus()
                    GoTo error_section
                End If
                If w_size_chk(0) <> 0 Then
                    GoTo error_section
                End If

            Case 2
                check_flg = 0

                ' -> watanabe Add 2007.03
                If form_no.chk_max_load_kg.CheckState = 1 Then
                    If Trim(form_no.w_max_load_kg.Text) <> "" Then
                        MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                        form_no.w_max_load_kg.Focus() 'SetFocus()
                        GoTo error_section
                    End If

                Else
                    ' <- watanabe Add 2007.03

                    '// MAXLOAD KG 入力チェック
                    If form_no.w_max_load_kg.Enabled = True Then
                        w_ret = num_check(form_no.w_max_load_kg.Text)
                        If w_ret <> 0 Then
                            MsgBox("Input data are incorrect." & Chr(13) & "MAXLOAD KG", 64)
                            form_no.w_max_load_kg.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        slen = Len(form_no.w_max_load_kg.Text)
                        If slen <> 3 And slen <> 4 Then
                            MsgBox("max_load_kg is 3-4 characters.", 64)
                            form_no.w_max_load_kg.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_load_kg.Text) < Val(form_no.w_kikaku_max_load_kg.Text) Then
                            MsgBox("Please enter a value greater than the standard value." & Chr(13) & "MAXLOAD KG", 64)
                            form_no.w_max_load_kg.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_load_kg.Text) <> Val(form_no.w_kikaku_max_load_kg.Text) Then
                            check_flg = 1
                        End If
                    End If

                    ' -> watanabe Add 2007.03
                End If

                If form_no.chk_max_load_lbs.CheckState = 1 Then
                    If Trim(form_no.w_max_load_lbs.Text) <> "" Then
                        MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                        form_no.w_max_load_lbs.Focus() 'SetFocus()
                        GoTo error_section
                    End If

                Else
                    ' <- watanabe Add 2007.03

                    '// MAXLOAD LBS 入力チェック
                    If form_no.w_max_load_lbs.Enabled = True Then
                        w_ret = num_check(form_no.w_max_load_lbs.Text)
                        If w_ret <> 0 Then
                            MsgBox("Input data are incorrect." & Chr(13) & "MAXLOAD LBS", 64)
                            form_no.w_max_load_lbs.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        slen = Len(form_no.w_max_load_lbs.Text)
                        If slen <> 3 And slen <> 4 Then
                            MsgBox("max_load_lbs is 3-4characters.", 64)
                            form_no.w_max_load_lbs.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_load_lbs.Text) < Val(form_no.w_kikaku_max_load_lbs.Text) Then
                            MsgBox("Please enter a value greater than the standard value." & Chr(13) & "MAXLOAD LBS", 64)
                            form_no.w_max_load_lbs.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_load_lbs.Text) <> Val(form_no.w_kikaku_max_load_lbs.Text) Then
                            If check_flg = 0 Then
                                check_flg = 2
                            End If
                        End If
                    End If

                    ' -> watanabe Add 2007.03
                End If

                If form_no.chk_max_press_kpa.CheckState = 1 Then
                    If Trim(form_no.w_max_press_kpa.Text) <> "" Then
                        MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                        form_no.w_max_press_kpa.Focus() 'SetFocus()
                        GoTo error_section
                    End If

                Else
                    ' <- watanabe Add 2007.03

                    '// MAXPRESS KPA 入力チェック
                    If form_no.w_max_press_kpa.Enabled = True Then
                        w_ret = num_check(form_no.w_max_press_kpa.Text)
                        If w_ret <> 0 Then
                            MsgBox("Input data are incorrect." & Chr(13) & "MAXPRESS KPA", 64)
                            form_no.w_max_press_kpa.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        slen = Len(form_no.w_max_press_kpa.Text)
                        If slen <> 3 Then
                            MsgBox("max_press_kpa is 3 characters.", 64)
                            form_no.w_max_press_kpa.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_press_kpa.Text) < Val(form_no.w_kikaku_max_press_kpa.Text) Then
                            MsgBox("Please enter a value greater than the standard value." & Chr(13) & "MAXPRESS KPA", 64)
                            form_no.w_max_press_kpa.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_press_kpa.Text) <> Val(form_no.w_kikaku_max_press_kpa.Text) And Val(form_no.w_max_press_kpa.Text) <> 300 Then
                            If check_flg = 0 Then
                                check_flg = 3
                            End If
                        End If
                    End If

                    ' -> watanabe Add 2007.03
                End If

                If form_no.chk_max_press_psi.CheckState = 1 Then
                    If Trim(form_no.w_max_press_psi.Text) <> "" Then
                        MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                        form_no.w_max_press_psi.Focus() 'SetFocus()
                        GoTo error_section
                    End If

                Else
                    ' <- watanabe Add 2007.03

                    '// MAXPRESS PSI 入力チェック
                    If form_no.w_max_press_psi.Enabled = True Then
                        w_ret = num_check(form_no.w_max_press_psi.Text)
                        If w_ret <> 0 Then
                            MsgBox("Input data are incorrect." & Chr(13) & "MAXPRESS PSI", 64)
                            form_no.w_max_press_psi.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        slen = Len(form_no.w_max_press_psi.Text)
                        If slen <> 2 Then
                            MsgBox("max_press_psi is 2 characters.", 64)
                            form_no.w_max_press_psi.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_press_psi.Text) < Val(form_no.w_kikaku_max_press_psi.Text) Then
                            MsgBox("Please enter a value greater than the standard value." & Chr(13) & "MAXPRESS PSI", 64)
                            form_no.w_max_press_psi.Focus() 'SetFocus()
                            GoTo error_section
                        End If
                        If Val(form_no.w_max_press_psi.Text) <> Val(form_no.w_kikaku_max_press_psi.Text) And Val(form_no.w_max_press_psi.Text) <> 44 Then
                            If check_flg = 0 Then
                                check_flg = 4
                            End If
                        End If
                    End If

                    ' -> watanabe Add 2007.03
                End If
                ' <- watanabe Add 2007.03

                'タイプのチェック
                If Trim(form_no.w_type.Text) = "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
                    form_no.w_type.Focus() 'SetFocus()
                    GoTo error_section
                End If

                'ピクチャのチェック
                'Brand Ver.5 TIFF->BMP 変更 start
                '       If form_no.ImgThumbnail1.Image = "" Then
                If form_no.ImgThumbnail1.Image Is Nothing Then
                    'Brand Ver.5 TIFF->BMP 変更 end
                    MsgBox("Picture is not specified.", 64)
                    form_no.w_type.Focus() 'SetFocus()
                    GoTo error_section
                End If

                If check_flg <> 0 Then

                    ' -> watanabe add VerUP(2011)
                    wstr = ""
                    ' <- watanabe add VerUP(2011)

                    If check_flg = 1 Then
                        wstr = "MAXLOAD KG"
                    ElseIf check_flg = 2 Then
                        wstr = "MAXLOAD LBS"
                    ElseIf check_flg = 3 Then
                        wstr = "MAXPRESS KPA"
                    ElseIf check_flg = 4 Then
                        wstr = "MAXPRESS PSI"
                    End If

                    ret = MsgBox("Design value of the " & wstr & " has been changed. Do you want to do the processing?", MsgBoxStyle.YesNo)

                    If ret = MsgBoxResult.No Then
                        If check_flg = 1 Then
                            form_no.w_max_load_kg.Focus() 'SetFocus()
                        ElseIf check_flg = 2 Then
                            form_no.w_max_load_lbs.Focus() 'SetFocus()
                        ElseIf check_flg = 3 Then
                            form_no.w_max_press_kpa.Focus() 'SetFocus()
                        ElseIf check_flg = 4 Then
                            form_no.w_max_press_psi.Focus() 'SetFocus()
                        End If
                        GoTo error_section
                    End If
                End If

        End Select

        check_F_TMP_MAXLOAD = 0

        Exit Function

error_section:
        check_F_TMP_MAXLOAD = 1

    End Function

    '
    'check_F_TMP_PLATE = 0: OK
    '
    'エラーとなった項目へのセットフォーカスも同時に行う

    Function check_F_TMP_PLATE() As Short

        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe add 2007.03
        If Trim(form_no.w_plate_w.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate width", 64)
            form_no.w_plate_w.Focus() 'SetFocus()
            GoTo error_section
        End If
        If IsNumeric(Trim(form_no.w_plate_w.Text)) = False Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate width", 64)
            form_no.w_plate_w.Focus() 'SetFocus()
            GoTo error_section
        End If

        If Trim(form_no.w_plate_h.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate height", 64)
            form_no.w_plate_h.Focus() 'SetFocus()
            GoTo error_section
        End If
        If IsNumeric(Trim(form_no.w_plate_h.Text)) = False Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate height", 64)
            form_no.w_plate_h.Focus() 'SetFocus()
            GoTo error_section
        End If

        If Trim(form_no.w_plate_r.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate corner R", 64)
            form_no.w_plate_r.Focus() 'SetFocus()
            GoTo error_section
        End If
        If IsNumeric(Trim(form_no.w_plate_r.Text)) = False Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate corner R", 64)
            form_no.w_plate_r.Focus() 'SetFocus()
            GoTo error_section
        End If

        If Trim(form_no.w_plate_n.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate screw position", 64)
            form_no.w_plate_n.Focus() 'SetFocus()
            GoTo error_section
        End If
        If IsNumeric(Trim(form_no.w_plate_n.Text)) = False Then
            MsgBox("Input data are incorrect." & Chr(13) & "Plate screw position", 64)
            form_no.w_plate_n.Focus() 'SetFocus()
            GoTo error_section
        End If
        ' <- watanabe add 2007.03


        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PLATE = 0
        Exit Function

error_section:
        check_F_TMP_PLATE = 1

    End Function

    '
    'check_form_no = 0: OK
    '
    'エラーとなった項目へのセットフォーカスも同時に行う

    Function check_F_TMP_PLY() As Short
        Dim i As Object
        Dim w_w_n As Short
        Dim flg As Short
        Dim wstr As String

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)

        'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type:Null", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If


        '20000124 追加
        flg = 0
        If form_no.w_sidewall.Text >= "2" Then
            flg = 1
        End If

        '20000124 追加
        For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                If flg = 0 And Mid(Tmp_prcs_code(i), 5, 1) <> "S" Then
                    wstr = Tmp_prcs_code(i)
                    Exit For
                ElseIf flg = 1 And Mid(Tmp_prcs_code(i), 5, 1) = "S" Then
                    wstr = Tmp_prcs_code(i)
                    Exit For
                End If
            End If
        Next i

        w_w_n = 0

        '20000124 修正
        '   If Trim$(form_no.w_type.Text) = "POLYESTER+STEEL" Then w_w_n = 2
        '   If Trim$(form_no.w_type.Text) = "POLYESTER+STEEL+NYLON" Then w_w_n = 3
        '   If Trim$(form_no.w_type.Text) = "RAYON+STEEL" Then w_w_n = 2
        '   If Trim$(form_no.w_type.Text) = "RAYON+STEEL+NYLON" Then w_w_n = 3
        '   If Trim$(form_no.w_type.Text) = "NYLON" Then w_w_n = 1
        If (Trim(wstr) = "PLY1") Or (Trim(wstr) = "PLY1S") Then w_w_n = 1
        If (Trim(wstr) = "PLY2") Or (Trim(wstr) = "PLY2S") Then w_w_n = 2
        If (Trim(wstr) = "PLY3") Or (Trim(wstr) = "PLY3S") Then w_w_n = 3

        If w_w_n = 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type:" & Trim(wstr), 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe Add 2007.03 -> del
        '    If form_no.chk_tread1.Value = 0 Or form_no.chk_tread2.Value = 0 Or form_no.chk_tread3.Value = 0 Then
        ' <- watanabe Add 2007.03 -> del

        'TREADのチェック
        If Trim(form_no.w_tread.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "TREAD", 64)
            form_no.w_tread1.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe Add 2007.03 -> del
        '    End If
        '
        '    If form_no.chk_tread1.Value = 1 Then
        '        If Trim(form_no.w_tread1.Text) <> "" Then
        '            MsgBox "入力データが不適切です。" & Chr$(13) & "加工保留チェックがある時は、空欄にしてください。", 64
        '            form_no.w_tread1.SetFocus
        '            GoTo error_section
        '        End If
        '
        '    Else
        ' <- watanabe Add 2007.03 -> del

        'TREAD1のチェック
        If (form_no.w_tread1.Text <> "") And ((form_no.w_tread1.Text < "1") Or (form_no.w_tread1.Text > "9")) Then
            MsgBox("Input data are incorrect." & Chr(13) & "TREAD1 is numerical value only.", 64)
            form_no.w_tread1.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe Add 2007.03 -> del
        '    End If
        ' <- watanabe Add 2007.03 -> del

        If w_w_n >= 2 Then

            ' -> watanabe Add 2007.03 -> del
            '        If form_no.chk_tread2.Value = 1 Then
            '            If Trim(form_no.w_tread2.Text) <> "" Then
            '                MsgBox "入力データが不適切です。" & Chr$(13) & "加工保留チェックがある時は、空欄にしてください。", 64
            '                form_no.w_tread2.SetFocus
            '                GoTo error_section
            '            End If
            '
            '        Else
            ' <- watanabe Add 2007.03 -> del

            'TREAD2のチェック
            If (form_no.w_tread2.Text <> "") And ((form_no.w_tread2.Text < "1") Or (form_no.w_tread2.Text > "9")) Then
                MsgBox("Input data are incorrect." & Chr(13) & "TREAD2 is numerical value only.", 64)
                form_no.w_tread2.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03 -> del
            '        End If
            ' <- watanabe Add 2007.03 -> del

        End If

        If w_w_n >= 3 Then

            ' -> watanabe Add 2007.03 -> del
            '        If form_no.chk_tread3.Value = 1 Then
            '            If Trim(form_no.w_tread3.Text) <> "" Then
            '                MsgBox "入力データが不適切です。" & Chr$(13) & "加工保留チェックがある時は、空欄にしてください。", 64
            '                form_no.w_tread3.SetFocus
            '                GoTo error_section
            '            End If
            '
            '        Else
            ' <- watanabe Add 2007.03 -> del

            'TREAD3のチェック
            If (form_no.w_tread3.Text <> "") And ((form_no.w_tread3.Text < "1") Or (form_no.w_tread3.Text > "9")) Then
                MsgBox("Input data are incorrect." & Chr(13) & "TREAD3 is numerical value only.", 64)
                form_no.w_tread3.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03 -> del
            '        End If
            ' <- watanabe Add 2007.03 -> del

        End If


        '**************************************
        '12/5 1997 yamamoto start
        '   ・ﾌｫｰﾑよりﾛｽﾄﾌｫｰｶｽｲﾍﾞﾝﾄを流用、修正
        '*************************************
        '/TREAD,TREAD1,TREAD2,TREAD3/
        If (Val(form_no.w_tread1.Text) + Val(form_no.w_tread2.Text) + Val(form_no.w_tread3.Text)) = 0 Then
            form_no.w_tread.Text = ""
        End If

        If (Val(form_no.w_tread1.Text) + Val(form_no.w_tread2.Text) + Val(form_no.w_tread3.Text)) > 9 Then
            '        If Trim$(w_tread.Text) <> "" Then
            MsgBox("TREAD will be more than 10.", 64)
            form_no.w_tread.Text = ""
            form_no.w_tread1.Focus() 'SetFocus()
            '        End If
            GoTo error_section
        End If

        '----- .NET 移行 -----
        'form_no.w_tread.Text = VB6.Format(Val(form_no.w_tread1.Text) + Val(form_no.w_tread2.Text) + Val(form_no.w_tread3.Text), "#")
        form_no.w_tread.Text = (Val(form_no.w_tread1.Text) + Val(form_no.w_tread2.Text) + Val(form_no.w_tread3.Text)).ToString("#")


        '******************************
        '12/5 1997 yamamoto end
        '******************************

        ' -> watanabe Add 2007.03 -> del
        '    If form_no.chk_sidewall.Value = 1 Then
        '        If Trim(form_no.w_sidewall.Text) <> "" Then
        '            MsgBox "入力データが不適切です。" & Chr$(13) & "加工保留チェックがある時は、空欄にしてください。", 64
        '            form_no.w_sidewall.SetFocus
        '            GoTo error_section
        '        End If
        '
        '    Else
        ' <- watanabe Add 2007.03 -> del

        'SIDEWALLのチェック
        If Trim(form_no.w_sidewall.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL", 64)
            form_no.w_sidewall.Focus() 'SetFocus()
            GoTo error_section
        End If
        If (form_no.w_sidewall.Text < "1") Or (form_no.w_sidewall.Text > "9") Then
            MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL only numbers.", 64)
            form_no.w_sidewall.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe Add 2007.03 -> del
        '    End If
        ' <- watanabe Add 2007.03 -> del


        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PLY = 0
        Exit Function

error_section:
        check_F_TMP_PLY = 1

    End Function

    '
    'check_form_no = 0: OK
    '
    'エラーとなった項目へのセットフォーカスも同時に行う
    Function check_F_TMP_PLY2() As Short
        Dim i As Object
        Dim w_w_n As Short
        Dim sw_sw As Short

        'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        w_w_n = 0
        sw_sw = 0

        For i = 1 To MaxSelNum
            If (form_no.w_type.Text = Tmp_hm_word(i)) Then
                If (Tmp_prcs_code(i) = "PLY11") Then
                    w_w_n = 1
                    sw_sw = 1
                ElseIf (Tmp_prcs_code(i) = "PLY12") Then
                    w_w_n = 1
                    sw_sw = 2
                ElseIf (Tmp_prcs_code(i) = "PLY21") Then
                    w_w_n = 2
                    sw_sw = 1
                ElseIf (Tmp_prcs_code(i) = "PLY22") Then
                    w_w_n = 2
                    sw_sw = 2
                ElseIf (Tmp_prcs_code(i) = "PLY31") Then
                    w_w_n = 3
                    sw_sw = 1
                ElseIf (Tmp_prcs_code(i) = "PLY32") Then
                    w_w_n = 3
                    sw_sw = 2
                End If
            End If
        Next i

        If w_w_n = 0 Or sw_sw = 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If


        ' -> watanabe Add 2007.03
        If form_no.chk_tread1.CheckState = 1 Then
            If Trim(form_no.w_tread1.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_tread1.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            'TREAD1のチェック
            If (form_no.w_tread1.Text <> "") And ((form_no.w_tread1.Text < "1") Or (form_no.w_tread1.Text > "9")) Then
                MsgBox("Input data are incorrect." & Chr(13) & "TREAD1 is numerical value only.", 64)
                form_no.w_tread1.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03

        If w_w_n >= 2 Then

            ' -> watanabe Add 2007.03
            If form_no.chk_tread2.CheckState = 1 Then
                If Trim(form_no.w_tread2.Text) <> "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                    form_no.w_tread2.Focus() 'SetFocus()
                    GoTo error_section
                End If

            Else
                ' <- watanabe Add 2007.03

                'TREAD2のチェック
                If (form_no.w_tread2.Text <> "") And ((form_no.w_tread2.Text < "1") Or (form_no.w_tread2.Text > "9")) Then
                    MsgBox("Input data are incorrect." & Chr(13) & "TREAD2 is numerical value only.", 64)
                    form_no.w_tread2.Focus() 'SetFocus()
                    GoTo error_section
                End If

                ' -> watanabe Add 2007.03
            End If
            ' <- watanabe Add 2007.03

        End If

        If w_w_n >= 3 Then

            ' -> watanabe Add 2007.03
            If form_no.chk_tread3.CheckState = 1 Then
                If Trim(form_no.w_tread3.Text) <> "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                    form_no.w_tread3.Focus() 'SetFocus()
                    GoTo error_section
                End If

            Else
                ' <- watanabe Add 2007.03

                'TREAD3のチェック
                If (form_no.w_tread3.Text <> "") And ((form_no.w_tread3.Text < "1") Or (form_no.w_tread3.Text > "9")) Then
                    MsgBox("Input data are incorrect." & Chr(13) & "TREAD3 is numerical value only.", 64)
                    form_no.w_tread3.Focus() 'SetFocus()
                    GoTo error_section
                End If

                ' -> watanabe Add 2007.03
            End If
            ' <- watanabe Add 2007.03

        End If


        ' -> watanabe Add 2007.03
        If form_no.chk_sidewall1.CheckState = 1 Then
            If Trim(form_no.w_sidewall1.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_sidewall1.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            'SIDEWALLのチェック
            If Trim(form_no.w_sidewall1.Text) = "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL1", 64)
                form_no.w_sidewall1.Focus() 'SetFocus()
                GoTo error_section
            End If
            If (form_no.w_sidewall1.Text < "1") Or (form_no.w_sidewall1.Text > "9") Then
                MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL1 only numbers.", 64)
                form_no.w_sidewall1.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03

        If sw_sw = 2 Then

            ' -> watanabe Add 2007.03
            If form_no.chk_sidewall2.CheckState = 1 Then
                If Trim(form_no.w_sidewall2.Text) <> "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                    form_no.w_sidewall2.Focus() 'SetFocus()
                    GoTo error_section
                End If

            Else
                ' <- watanabe Add 2007.03

                If Trim(form_no.w_sidewall2.Text) = "" Then
                    MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL2", 64)
                    form_no.w_sidewall2.Focus() 'SetFocus()
                    GoTo error_section
                End If
                If (form_no.w_sidewall2.Text < "1") Or (form_no.w_sidewall2.Text > "9") Then
                    MsgBox("Input data are incorrect." & Chr(13) & "SIDEWALL2 only numbers.", 64)
                    form_no.w_sidewall2.Focus() 'SetFocus()
                    GoTo error_section
                End If

                ' -> watanabe Add 2007.03
            End If
            ' <- watanabe Add 2007.03

        End If


        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PLY2 = 0
        Exit Function

error_section:
        check_F_TMP_PLY2 = 1

    End Function

    Function check_F_TMP_ETC() As Short
        Dim i As Object
        '
        'check_form_no = 0: OK
        '
        'エラーとなった項目へのセットフォーカスも同時に行う
        Dim w_w_n As Short

        'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        w_w_n = 0

        For i = 1 To MaxSelNum
            If (form_no.w_type.Text = Tmp_hm_word(i)) Then
                If (Tmp_prcs_code(i) = "ETC1") Then
                    w_w_n = 1
                ElseIf (Tmp_prcs_code(i) = "ETC2") Then
                    w_w_n = 2
                ElseIf (Tmp_prcs_code(i) = "ETC3") Then
                    w_w_n = 3
                ElseIf (Tmp_prcs_code(i) = "ETC4") Then
                    w_w_n = 4
                ElseIf (Tmp_prcs_code(i) = "ETC5") Then
                    w_w_n = 5
                ElseIf (Tmp_prcs_code(i) = "ETC6") Then
                    w_w_n = 6
                ElseIf (Tmp_prcs_code(i) = "ETC7") Then
                    w_w_n = 7
                ElseIf (Tmp_prcs_code(i) = "ETC8") Then
                    w_w_n = 8
                ElseIf (Tmp_prcs_code(i) = "ETC9") Then
                    w_w_n = 9
                ElseIf (Tmp_prcs_code(i) = "ETC10") Then
                    w_w_n = 10
                End If
            End If
        Next i

        If w_w_n = 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        For i = 1 To w_w_n
            If Trim(form_no.w_etc(i).Text) = "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "Input value " & i, 64)
                form_no.w_etc(i).Focus() 'SetFocus()
                GoTo error_section
            End If
        Next i

        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_ETC = 0
        Exit Function

error_section:
        check_F_TMP_ETC = 1

    End Function


    Function check_F_TMP_MOLD() As Short
        Dim i As Object
        '
        'check_F_TMP_MOLD = 0: OK
        '
        'エラーとなった項目へのセットフォーカスも同時に行う
        Dim ss As String
        Dim s1 As New VB6.FixedLengthString(1)
        Dim slen As Short
        Dim lp As Short
        Dim wstr As String

        ' -> watanabe add VerUP(2011)
        wstr = ""
        ' <- watanabe add VerUP(2011)


        '区分のチェック

        '20000124 追加
        For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                wstr = Tmp_prcs_code(i)
                Exit For
            End If
        Next i

        ss = Trim(form_no.w_kubun.Text)
        slen = Len(ss)
        '20000124 修正
        '   If form_no.w_type.Text = "区分(1桁)＋番号" Then
        If wstr = "MNO2" Then
            If slen <> 1 Then
                MsgBox("Input data are incorrect." & Chr(13) & "Category (1 digit)", 64)
                form_no.w_kubun.Focus() 'SetFocus()
                GoTo error_section
            Else
                If (Asc(ss) < Asc("A")) Or (Asc(ss) > Asc("Z")) Then
                    MsgBox("Input data is only English letter." & Chr(13) & "Category (1 digit)", 64)
                    form_no.w_kubun.Focus() 'SetFocus()
                    GoTo error_section
                End If
            End If
            '20000124 修正
            '   ElseIf form_no.w_type.Text = "区分(2桁)＋番号" Then
        ElseIf wstr = "MNO3" Then
            If slen <> 2 Then
                MsgBox("Input data are incorrect." & Chr(13) & "Category (2 digit)", 64)
                form_no.w_kubun.Focus() 'SetFocus()
                GoTo error_section
            Else
                For i = 1 To slen
                    s1.Value = Mid(ss, i, 1)
                    If (Asc(s1.Value) < Asc("A")) Or (Asc(s1.Value) > Asc("Z")) Then
                        MsgBox("Input data is only English letter.", 64)
                        form_no.w_kubun.Focus() 'SetFocus()
                        GoTo error_section
                    End If
                Next i

            End If
        End If
        '97.04.25 comment start .........................
        '   If form_no.w_type.Text <> "番号のみ" Then
        '      For lp = 1 To slen
        '         If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid$(ss, lp, 1)) = 0 Then
        '            MsgBox "入力データが不適切です。" & Chr$(13) & "モールド番号－区分", 64
        '            form_no.w_kubun.SetFocus
        '            GoTo error_section
        '         End If
        '      Next lp
        '   End If
        '97.04.25 comment ended .........................

        '番号のチェック
        ss = Trim(form_no.w_no.Text)
        slen = Len(ss)
        If ss = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Mold number - number", 64)
            form_no.w_no.Focus() 'SetFocus()
            GoTo error_section
        End If
        If (slen > 4 Or slen < 1) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Mold number - number", 64)
            form_no.w_no.Focus() 'SetFocus()
            GoTo error_section
        End If
        For lp = 1 To slen
            '      2005.05.09 O.Kawaguchi 変更 A 追加
            '      2010.12.21 T.Uriu 変更 M 追加
            '      2011.04.14 T.Uriu 変更 Z 追加
            '      2012.08.31 T.Uriu 変更 B 追加
            '      If InStr("0123456789T", Mid$(ss, lp, 1)) = 0 Then
            If InStr("0123456789ABTMZ", Mid(ss, lp, 1)) = 0 Then
                MsgBox("Input data are incorrect." & Chr(13) & "Mold number - number", 64)
                form_no.w_no.Focus() 'SetFocus()
                GoTo error_section
            End If
        Next lp

        'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_MOLD = 0
        Exit Function

error_section:
        check_F_TMP_MOLD = 1

    End Function

    Function check_F_HMSAVE() As Short
        Dim irt2 As Object
        Dim i As Object
        Dim irt As Object

        'check_form_no = 0: OK

        '******************************************
        '12/5 1997 yamamoto start
        '   ・no_data_sectoinを付加
        '   ・ﾌｫｰﾑのﾛｽﾄﾌｫｰｶｽﾌﾟﾛｼｰｼﾞｬを持ってきて修正
        '******************************************
        Dim f As System.Windows.Forms.Control

        If form_no.w_font_name.Text = "" Then GoTo no_data_section

        If open_mode <> "NEW" Then
            If form_no.w_no.Text = "" Then GoTo no_data_section
        End If

        If form_no.w_spell.Text = "" Then GoTo no_data_section
        'アポストロフィ検索
        irt = InStr(form_no.w_spell.Text, "'")

        If irt <> 0 Then
            MsgBox("Can not register ' (apostrophe) to spell.", 64)
            GoTo error_section
        End If
        '97.04.23 n.matsumi update 1 line
        '   If form_no.w_comment.Text = "" Then GoTo no_data_section

        If form_no.w_dep_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_date.Text = "" Then GoTo no_data_section
        If form_no.w_width.Text = "" Then GoTo no_data_section
        If form_no.w_high.Text = "" Then GoTo no_data_section
        If form_no.w_ang.Text = "" Then GoTo no_data_section
        If form_no.w_gm_num.Text = "" Then GoTo no_data_section

        '/フォント名/
        f = form_no.w_font_name
        irt = check_0(form_no.w_font_name.Text, 6, 0, f)
        If irt <> 2 And Left(form_no.w_font_name.Text, 2) <> "HE" Then
            MsgBox("Font name is 6 characters starting with HE.", 64)
            f.Focus()
            GoTo error_section
        ElseIf irt = 2 Then
            GoTo error_section
        End If

        '正式部品チェック(Brand CAD System Ver.3 UP )
        If IsNumeric(Mid(Trim(form_no.w_font_name.Text), 3, 4)) = True Then
            For i = 1 To temp_hm.gm_num
                If IsNumeric(Mid(Trim(temp_hm.gm_name(i)), 3, 4)) = False Then
                    MsgBox("Can not be regular registration for individual parts are included.", 64)
                    f.Focus()
                    GoTo error_section
                End If
            Next i
        End If


        '/スペル/
        f = form_no.w_spell
        irt = len_check2(form_no.w_spell.Text, 255, 1)
        If irt > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
            f.Focus()
            GoTo error_section
        End If
        'アポストロフィ検索
        irt = InStr(form_no.w_spell.Text, "'")
        If irt <> 0 Then
            MsgBox("Can not register ' (apostrophe).", 64)
            f.Focus()
            GoTo error_section
        End If

        '/コメント/
        f = form_no.w_comment
        irt = len_check2(form_no.w_comment.Text, 255, 1)
        If irt > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
            f.Focus()
            GoTo error_section
        End If

        '/部署/
        f = form_no.w_dep_name
        ' 1997.04.23 n.matsumi update start ....................
        '   irt = check_1(w_dep_name.Text, 2, 1, f)
        irt = check_0(form_no.w_dep_name.Text, 2, 1, f)
        ' 1997.04.23 n.matsumi update ended ....................
        If irt = 2 Then GoTo error_section

        '/登録者/
        f = form_no.w_entry_name
        irt = check_1(form_no.w_entry_name.Text, 4, 1, f)
        ' 1997.04.23 n.matsumi update start ....................
        If irt = 0 Then
            irt = num_check(form_no.w_entry_name.Text)
            If irt = 2 Then
                MsgBox("Input data are incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
        ElseIf irt = 2 Then
            GoTo error_section
        End If
        ' 1997.04.23 n.matsumi update ended ....................

        '/基準高さ/
        f = form_no.w_high
        irt = check_2(form_no.w_high.Text, 9, 1, f)
        If irt = 0 Then
            irt2 = float_check2(form_no.w_high.Text, 9, 4)
            If irt2 > 0 Then
                MsgBox("Input data are incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
        ElseIf irt = 2 Then
            GoTo error_section
        End If

        '/基準角度/
        f = form_no.w_ang
        irt = check_2(form_no.w_ang.Text, 9, 1, f)
        If irt = 0 Then
            irt2 = float_check2(form_no.w_ang.Text, 9, 4)
            If irt2 > 0 Then
                MsgBox("Input data are incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
        ElseIf irt = 2 Then
            GoTo error_section
        End If

        check_F_HMSAVE = 0
        Exit Function

no_data_section:
        MsgBox("There is an error in the input data.", 64)

        '******************************************
        '12/5 1997 yamamoto end
        '******************************************

error_section:
        check_F_HMSAVE = 1


    End Function

    Function check_F_GZSAVE() As Short
        Dim irt As Object
        '
        'check_form_no = 0: OK

        '******************************************
        '12/5 yamamoto start
        '   ・no_data_sectionを追加
        '   ・ﾌｫｰﾑのﾛｽﾄﾌｫｰｶｽﾌﾟﾛｼｰｼﾞｬを持ってきて修正
        Dim f As System.Windows.Forms.Control

        If form_no.w_id.Text = "" Then GoTo no_data_section
        If form_no.w_no1.Text = "" Then GoTo no_data_section
        If form_no.w_no2.Text = "" Then GoTo no_data_section

        '97.04.23 n.matsumi update 1 line
        '   If form_no.w_comment.Text = "" Then GoTo no_data_section

        If form_no.w_dep_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_date.Text = "" Then GoTo no_data_section
        If form_no.w_gm_num.Text = "" Then GoTo no_data_section

        '/番号/
        f = form_no.w_no1
        form_no.w_no1.Text = Trim(form_no.w_no1.Text)
        irt = check_0(form_no.w_no1.Text, 4, 0, f)
        If irt <> 0 Then
            MsgBox("Code is invalid.", 64, "Input error")
            f.Focus()
            GoTo error_section
        End If

        '/変番/
        f = form_no.w_no2
        form_no.w_no2.Text = Trim(form_no.w_no2.Text)
        irt = check_0(form_no.w_no2.Text, 2, 0, f)
        If irt <> 0 Then
            MsgBox("Code is invalid.", 64, "Input error")
            f.Focus()
            GoTo error_section
        Else
            ' 1997.04.23 n.matsumi update start ....................
            If irt = 0 Then
                irt = num_check(form_no.w_no2.Text)
                If irt = 2 Then
                    MsgBox("Code is invalid.", 64, "Input error")
                    f.Focus()
                    GoTo error_section
                End If
            End If
            ' 1997.04.23 n.matsumi update ended ....................

        End If

        '/コメント/
        f = form_no.w_comment

        irt = len_check2(form_no.w_comment.Text, 255, 1)
        If irt > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
            f.Focus()
            GoTo error_section
        End If

        '/部署/
        form_no.w_dep_name.Text = Trim(form_no.w_dep_name.Text)
        f = form_no.w_dep_name
        ' 1997.04.23 n.matsumi update start ....................
        '   irt = check_1(w_dep_name.Text, 2, 1, f)
        irt = check_0(form_no.w_dep_name.Text, 2, 1, f)
        If irt = 2 Then GoTo error_section
        ' 1997.04.23 n.matsumi update ended ....................

        '/登録者/
        f = form_no.w_entry_name
        irt = check_1(form_no.w_entry_name.Text, 4, 1, f)
        ' 1997.04.23 n.matsumi update start ....................
        If irt = 0 Then
            irt = num_check(form_no.w_entry_name.Text)
            If irt = 2 Then
                MsgBox("Input data are incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
        ElseIf irt = 2 Then
            GoTo error_section
        End If
        ' 1997.04.23 n.matsumi update ended ....................

        check_F_GZSAVE = 0
        Exit Function

no_data_section:
        MsgBox("There are undefined data.", 64)
        '----- 12/5 yamamoto end -------
error_section:
        check_F_GZSAVE = 1


    End Function

    Function check_F_HZSAVE() As Short
        Dim irt As Object
        '
        'check_form_no = 0: OK
        '
        '*********************************************
        '12/5 yamamoto start
        '    ・no_data_sectionを追加
        '    ・ﾌｫｰﾑからﾛｽﾄﾌｫｰｶｽﾌﾟﾛｼｰｼﾞｬを流用､修正
        Dim f As System.Windows.Forms.Control

        If form_no.w_id.Text = "" Then GoTo no_data_section
        If form_no.w_no1.Text = "" Then GoTo no_data_section
        If form_no.w_no2.Text = "" Then GoTo no_data_section
        '97.04.23 n.matsumi update 1 line
        '   If form_no.w_comment.Text = "" Then GoTo no_data_section

        If form_no.w_dep_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_name.Text = "" Then GoTo no_data_section
        If form_no.w_entry_date.Text = "" Then GoTo no_data_section
        If form_no.w_hm_num.Text = "" Then GoTo no_data_section

        '/番号/
        f = form_no.w_no1
        irt = check_0(form_no.w_no1.Text, 4, 0, f)
        If irt = 2 Then GoTo error_section
        '/変番/
        f = form_no.w_no2
        irt = check_0(form_no.w_no2.Text, 2, 0, f)
        If irt = 2 Then GoTo error_section
        ' 1997.04.23 n.matsumi update start ....................
        If irt = 0 Then
            irt = num_check(form_no.w_no2.Text)
            If irt = 2 Then
                MsgBox("Code is invalid.", 64, "Input error")
                f.Focus()
                GoTo error_section
            End If
        End If
        ' 1997.04.23 n.matsumi update ended ....................

        '/コメント/
        f = form_no.w_comment
        irt = len_check2(form_no.w_comment.Text, 255, 1)
        If irt > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
            f.Focus()
            GoTo error_section
        End If

        '/部署/
        f = form_no.w_dep_name
        ' 1997.04.23 n.matsumi update start ....................
        '   irt = check_1(w_dep_name.Text, 2, 1, f)
        irt = check_0(form_no.w_dep_name.Text, 2, 1, f)
        If irt = 2 Then GoTo error_section
        ' 1997.04.23 n.matsumi update ended ....................

        '/登録者/
        f = form_no.w_entry_name
        irt = check_1(form_no.w_entry_name.Text, 4, 1, f)
        ' 1997.04.23 n.matsumi update start ....................
        If irt = 0 Then
            irt = num_check(form_no.w_entry_name.Text)
            If irt = 2 Then
                MsgBox("Input data are incorrect.", 64)
                f.Focus()
                GoTo error_section
            End If
        ElseIf irt = 2 Then
            GoTo error_section
        End If
        ' 1997.04.23 n.matsumi update ended ....................

        check_F_HZSAVE = 0
        Exit Function

no_data_section:
        MsgBox("There are undefined data.", 64)

        '12/5 1997 yamamoto end
        '**************************
error_section:
        check_F_HZSAVE = 1

    End Function

    Function check_F_BZSAVE() As Short
        '
        'check_F_BZSAVE = 0: OK
        '

        'エラーとなった項目へのセットフォーカスも同時に行う
        Dim w_ret As Short
        Dim i As Short

        '図面番号
        form_no.w_no1.Text = Trim(form_no.w_no1.Text)
        If form_no.w_no1.Text = "" Then
            MsgBox("Please enter the drawing number.")
            form_no.w_no1.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' -> watanabe edit 2007.06
        '' -> watanabe edit 2007.03
        ''   w_ret = check_0(form_no.w_no1.Text, 4, 0, form_no.w_no1)
        '   w_ret = check_0(form_no.w_no1.Text, 5, 0, form_no.w_no1)
        '' <- watanabe edit 2007.03
        form_no.w_no1.Text = Trim(form_no.w_no1.Text)
        If Len(form_no.w_no1.Text) = 4 Then
            w_ret = check_0(form_no.w_no1.Text, 4, 0, form_no.w_no1)
        Else
            w_ret = check_0(form_no.w_no1.Text, 5, 0, form_no.w_no1)
        End If
        ' <- watanabe edit 2007.06

        If w_ret <> 0 Then GoTo error_section

        '業務管理番号
        form_no.w_kanri_no.Text = Trim(form_no.w_kanri_no.Text)
        w_ret = check_1(form_no.w_kanri_no.Text, 8, 1, form_no.w_kanri_no)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_kanri_no.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Control number is only numerical value.", 64)
            form_no.w_kanri_no.Focus() 'SetFocus()
            GoTo error_section
        End If

        'コメント
        form_no.w_comment.Text = Trim(form_no.w_comment.Text)
        form_no.w_comment.Text = Trim(form_no.w_comment.Text)
        w_ret = len_check2(form_no.w_comment.Text, 255, 1)
        If w_ret > 0 Then
            MsgBox("Number of characters exceeds the limit.", 64)
            form_no.w_comment.Focus() 'SetFocus()
            GoTo error_section
        End If

        '部署
        form_no.w_dep_name.Text = Trim(form_no.w_dep_name.Text)
        If form_no.w_dep_name.Text = "" Then
            MsgBox("Please enter a department.")
            form_no.w_dep_name.Focus() 'SetFocus()
            GoTo error_section
        End If
        '97.04.24 update n.matsumi start .............................................
        w_ret = check_0(form_no.w_dep_name.Text, 2, 1, form_no.w_dep_name)
        If w_ret = 2 Then GoTo error_section
        '97.04.24 update n.matsumi ended .............................................

        '97.04.24 update n.matsumi start .............................................
        '   If Mid$(form_no.w_dep_name.Text, 1, 1) < "A" Or Mid$(form_no.w_dep_name.Text, 1, 1) > "Z" Then
        '      MsgBox "入力が不正です", vbCritical
        '      form_no.w_dep_name.SetFocus
        '      GoTo error_section
        '   End If
        '   If Mid$(form_no.w_dep_name.Text, 2, 1) < "A" Or Mid$(form_no.w_dep_name.Text, 2, 1) > "Z" Then
        '      MsgBox "入力が不正です", vbCritical
        '      form_no.w_dep_name.SetFocus
        '      GoTo error_section
        '   End If
        '97.04.24 update n.matsumi ended .............................................

        '登録者
        form_no.w_entry_name.Text = Trim(form_no.w_entry_name.Text)
        If form_no.w_entry_name.Text = "" Then
            MsgBox("Please enter the registrant.")
            form_no.w_entry_name.Focus() 'SetFocus()
            GoTo error_section
        End If
        w_ret = check_1(form_no.w_entry_name.Text, 4, 1, form_no.w_entry_name)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_entry_name.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Registrant only numerical value.", 64)
            form_no.w_entry_name.Focus() 'SetFocus()
            GoTo error_section
        End If

        'タイヤ種類
        If form_no.w_syurui.Text = "" Then
            MsgBox("Please enter the tire type.")
            form_no.w_syurui.Focus() 'SetFocus()
            GoTo error_section
        End If

        'パターン
        form_no.w_pattern.Text = Trim(form_no.w_pattern.Text)
        If form_no.w_pattern.Text = "" Then
            MsgBox("Please enter the pattern.")
            form_no.w_pattern.Focus() 'SetFocus()
            GoTo error_section
        End If
        If Len(form_no.w_pattern.Text) > 6 Then
            MsgBox("The input of the pattern is to six characters.")
            GoTo error_section
        End If
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", Left(form_no.w_pattern.Text, 1), 0) = 0 Then
            MsgBox("Input data are incorrect.")
            form_no.w_pattern.Focus() 'SetFocus()
            GoTo error_section
        End If
        For i = 2 To Len(Trim(form_no.w_pattern.Text))
            If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ+-", Mid(Trim(form_no.w_pattern.Text), i, 1), 0) = 0 Then
                MsgBox("Input data are incorrect.")
                form_no.w_pattern.Focus() 'SetFocus()
                GoTo error_section
            End If
        Next i

        'パターン種別
        If form_no.w_syubetu.Text = "" Then
            MsgBox("Please enter the pattern type.")
            form_no.w_syubetu.Focus() 'SetFocus()
            GoTo error_section
        End If

        'サイズ
        form_no.w_size.Text = Trim(form_no.w_size.Text)
        If form_no.w_size.Text = "" Then
            MsgBox("Please enter the size.")
            form_no.w_size1.Focus() 'SetFocus()
            GoTo error_section
        End If
        'サイズ１
        'ピリオドに対応 2000.07.26 by Kawaguchi
        '   w_ret = check_0(form_no.w_size1.Text, 5, 1, form_no.w_size1)
        w_ret = check_0_1(form_no.w_size1.Text, 5, 1, form_no.w_size1)
        If w_ret = 2 Then GoTo error_section
        'サイズ２
        If Len(form_no.w_size2.Text) > 1 Then
            MsgBox("Joint: Input is invalid.", MsgBoxStyle.Critical)
            form_no.w_size2.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_size2.Text <> "" Then
            If (Left(form_no.w_size2.Text, 1) < "A" Or Left(form_no.w_size2.Text, 1) > "Z") And (Left(form_no.w_size2.Text, 1) <> "/") Then
                MsgBox("Section width: input is invalid.", MsgBoxStyle.Critical)
                form_no.w_size2.Focus() 'SetFocus()
                GoTo error_section
            End If
        End If
        'サイズ３
        If form_no.w_size3.Text = "" Then
            MsgBox("Please enter the section width.")
            form_no.w_size3.Focus() 'SetFocus()
            GoTo error_section
        End If
        'ピリオドに対応 2000.07.26 by Kawaguchi
        '   w_ret = check_0(form_no.w_size3.Text, 5, 1, form_no.w_size3)
        w_ret = check_0_1(form_no.w_size3.Text, 5, 1, form_no.w_size3)
        If w_ret = 2 Then GoTo error_section
        'サイズ４
        If Len(form_no.w_size4.Text) > 1 Then
            MsgBox("Input error", MsgBoxStyle.Critical)
            form_no.w_size4.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_size4.Text <> "" Then
            If Left(form_no.w_size4.Text, 1) < "A" Or Left(form_no.w_size4.Text, 1) > "Z" Then
                MsgBox("Speed​​: input is invalid.", MsgBoxStyle.Critical)
                form_no.w_size4.Focus() 'SetFocus()
                GoTo error_section
            End If
        End If
        'サイズ５
        If Len(form_no.w_size5.Text) > 1 Then
            MsgBox("Input error", MsgBoxStyle.Critical)
            form_no.w_size5.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_size5.Text <> "" Then
            If (Left(form_no.w_size5.Text, 1) <> "R") And (Left(form_no.w_size5.Text, 1) <> "D") And (Left(form_no.w_size5.Text, 1) <> "-") Then
                MsgBox("Structure: Input is wrong.", MsgBoxStyle.Critical)
                form_no.w_size5.Focus() 'SetFocus()
                GoTo error_section
            End If
        End If
        'サイズ６
        If form_no.w_size6.Text = "" Then
            MsgBox("Please enter the rim diameter.")
            form_no.w_size6.Focus() 'SetFocus()
            GoTo error_section
        End If
        'w_ret = check_0(form_no.w_size6.Text, 4, 1, form_no.w_size6)
        w_ret = check_0_1(form_no.w_size6.Text, 4, 1, form_no.w_size6) '2002.08.27 Kawaguchi 小数点も可にする
        If w_ret = 2 Then GoTo error_section
        'w_ret = num_check(form_no.w_size6.Text)
        w_ret = num_check_1(form_no.w_size6.Text) '2002.08.27 Kawaguchi 小数点も可にする
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Rim diameter of only numbers.", 64)
            form_no.w_size6.Focus() 'SetFocus()
            GoTo error_section
        End If
        'サイズ７
        w_ret = check_0(form_no.w_size7.Text, 2, 1, form_no.w_size7)
        If w_ret = 2 Then GoTo error_section
        'サイズ８
        w_ret = check_0(form_no.w_size8.Text, 2, 1, form_no.w_size8)
        If w_ret = 2 Then GoTo error_section

        '工場
        If form_no.w_plant.Text = "" Then
            MsgBox("Please enter the plant.")
            form_no.w_plant.Focus() 'SetFocus()
            GoTo error_section
        End If

        '規格
        '  If form_no.w_kikaku1.Text = "" Then GoTo error_section
        '  If form_no.w_kikaku2.Text = "" Then GoTo error_section
        '  If form_no.w_kikaku3.Text = "" Then GoTo error_section
        '  If form_no.w_kikaku4.Text = "" Then GoTo error_section
        '  If form_no.w_kikaku5.Text = "" Then GoTo error_section
        '  If form_no.w_kikaku6.Text = "" Then GoTo error_section

        If form_no.w_tos_moyou.Text = "" Then
            MsgBox("Please select the TOS corresponding pattern.")
            form_no.w_tos_moyou.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_side_moyou.Text = "" Then
            MsgBox("Please choose a side uneven pattern.")
            form_no.w_side_moyou.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_side_kenti.Text = "" Then
            MsgBox("Please choose a side Indentation detected.")
            form_no.w_side_kenti.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_peak_mark.Text = "" Then
            MsgBox("Please select a peak mark.")
            form_no.w_peak_mark.Focus() 'SetFocus()
            GoTo error_section
        End If
        If form_no.w_nasiji.Text = "" Then
            MsgBox("Please select a satin processing.")
            form_no.w_nasiji.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_BZSAVE = 0
        Exit Function

error_section:
        check_F_BZSAVE = 1


    End Function

    Function check_F_TMP_PSI2() As Short
        Dim w_ret As Object

        If form_no.w_psi1.Text = "" Then
            MsgBox("Please enter the Durable air pressure.")
            form_no.w_psi1.Focus() 'SetFocus()
            GoTo error_section
        End If
        ' 1998/10/19 修正
        '    w_ret = check_0(form_no.w_psi1.Text, 2, 1, form_no.w_psi1)
        w_ret = check_0(form_no.w_psi1.Text, 3, 1, form_no.w_psi1)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_psi1.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Durable air pressure only numerical value.", 64)
            form_no.w_psi1.Focus() 'SetFocus()
            GoTo error_section
        End If

        ' 1998/10/22 追加
        If (Val(form_no.w_psi1.Text) Mod 5) <> 0 Then
            MsgBox("Please input PSI with a multiple of 5.", 64)
            form_no.w_psi1.Focus() 'SetFocus()
            GoTo error_section
        End If



        check_F_TMP_PSI2 = 0
        Exit Function

error_section:
        check_F_TMP_PSI2 = 1

    End Function

    Function check_F_TMP_PR2() As Short
        Dim w_ret As Object

        If form_no.w_ply1.Text = "" Then
            MsgBox("Please enter the ply.")
            form_no.w_ply1.Focus() 'SetFocus()
            GoTo error_section
        End If
        w_ret = check_0(form_no.w_ply1.Text, 2, 1, form_no.w_ply1)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_ply1.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "The ply only as for the numerical value.", 64)
            form_no.w_ply1.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PR2 = 0
        Exit Function

error_section:
        check_F_TMP_PR2 = 1

    End Function

    Function check_F_TMP_KAJUU2D() As Short
        Dim w_ret As Object

        If form_no.w_load_index1.Text = "" Then
            MsgBox("Please enter the load index 1.")
            form_no.w_load_index1.Focus() 'SetFocus()
            GoTo error_section
        End If
        w_ret = check_0(form_no.w_load_index1.Text, 3, 1, form_no.w_load_index1)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_load_index1.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Load index 1 is only numerical value.", 64)
            form_no.w_load_index1.Focus() 'SetFocus()
            GoTo error_section
        End If

        If form_no.w_load_index2.Text = "" Then
            MsgBox("Please enter the load index 2.")
            form_no.w_load_index2.Focus() 'SetFocus()
            GoTo error_section
        End If
        w_ret = check_0(form_no.w_load_index2.Text, 3, 1, form_no.w_load_index2)
        If w_ret = 2 Then GoTo error_section
        w_ret = num_check(form_no.w_load_index2.Text)
        If w_ret <> 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Load index 2 is only numerical value.", 64)
            form_no.w_load_index2.Focus() 'SetFocus()
            GoTo error_section
        End If

        If Trim(form_no.w_sokudo.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Speed symbol", 64)
            form_no.w_sokudo.Focus() 'SetFocus()
            GoTo error_section
        End If
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(Trim(form_no.w_sokudo.Text), 1)) = 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Speed symbol", 64)
            form_no.w_sokudo.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_KAJUU2D = 0
        Exit Function

error_section:
        check_F_TMP_KAJUU2D = 1

    End Function

    Function check_F_TMP_KAJUU() As Short
        '
        'check_F_TMP_KAJUU = 0: OK
        '
        'エラーとなった項目へのセットフォーカスも同時に行う

        If w_size_chk(0) <> 0 Then
            GoTo error_section
        End If

        If Trim(form_no.w_syurui.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Tire type", 64)
            form_no.w_syurui.Focus() 'SetFocus()
            GoTo error_section
        End If
        If Trim(form_no.w_kikaku.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Standard degree", 64)
            form_no.w_kikaku.Focus() 'SetFocus()
            GoTo error_section
        End If
        If Trim(form_no.w_sokudo.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Speed symbol", 64)
            form_no.w_sokudo.Focus() 'SetFocus()
            GoTo error_section
        End If
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Left(Trim(form_no.w_sokudo.Text), 1)) = 0 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Speed symbol", 64)
            form_no.w_sokudo.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_KAJUU = 0
        Exit Function

error_section:
        check_F_TMP_KAJUU = 1

    End Function

    '
    'check_F_TMP_UTQG = 0: OK
    '
    'エラーとなった項目へのセットフォーカスも同時に行う
    Function check_F_TMP_UTQG() As Short
        Dim ss As String
        Dim slen As Short
        Dim w_ret As Short


        ' -> watanabe Add 2007.03
        If form_no.chk_treadwear.CheckState = 1 Then
            If Trim(form_no.w_treadwear.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_treadwear.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            '/ TREADWEARのチェック
            slen = Len(form_no.w_treadwear.Text)
            '2000/7/31 修正 TREADWEAR２桁対応 by Kawaguchi
            '    If slen <> 3 Then
            '        MsgBox "TREADWEARは３桁です。", 64
            If slen < 2 Or slen > 3 Then
                MsgBox("TREADWEAR is 2-3 characters.", 64)
                form_no.w_treadwear.Focus() 'SetFocus()
                GoTo error_section
            End If
            w_ret = num_check(form_no.w_treadwear.Text)
            If w_ret <> 0 Then
                MsgBox("Input data are incorrect." & Chr(13) & "TREADWEAR is numerical value only.", 64)
                form_no.w_treadwear.Focus() 'SetFocus()
                GoTo error_section
            End If
            If (Val(form_no.w_treadwear.Text) Mod 20) <> 0 Then
                MsgBox("Please input ＴＲＥＡＤＷＥＡＲ with a multiple of 20.", 64)
                form_no.w_treadwear.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03



        ' -> watanabe Add 2007.03
        If form_no.chk_traction.CheckState = 1 Then
            If Trim(form_no.w_traction.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_traction.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            '/ TRACTIONのチェック
            slen = Len(form_no.w_traction.Text)
            If slen <> 1 And slen <> 2 Then
                MsgBox("TRACTION is 1-2 characters.", 64)
                form_no.w_traction.Focus() 'SetFocus()
                GoTo error_section
            End If
            ss = form_no.w_traction.Text
            If ss <> "A" And ss <> "B" And ss <> "C" And ss <> "AA" Then
                MsgBox("Input data are incorrect." & Chr(13) & "TRACTION is only A or B or C or AA.", 64)
                form_no.w_traction.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03


        ' -> watanabe Add 2007.03
        If form_no.chk_temperature.CheckState = 1 Then
            If Trim(form_no.w_temperature.Text) <> "" Then
                MsgBox("Input data are incorrect." & Chr(13) & "If there is a check to ""pending processing"", a blanks, please.", 64)
                form_no.w_temperature.Focus() 'SetFocus()
                GoTo error_section
            End If

        Else
            ' <- watanabe Add 2007.03

            '/ TEMPERRATUREのチェック
            slen = Len(form_no.w_temperature.Text)
            If slen <> 1 Then
                MsgBox("TEMPERRATURE is 1 digit.", 64)
                form_no.w_temperature.Focus() 'SetFocus()
                GoTo error_section
            End If
            ss = form_no.w_temperature.Text
            If ss <> "A" And ss <> "B" And ss <> "C" Then
                MsgBox("Input data are incorrect." & Chr(13) & "TEMPERRATURE is only A or B or C.", 64)
                form_no.w_temperature.Focus() 'SetFocus()
                GoTo error_section
            End If

            ' -> watanabe Add 2007.03
        End If
        ' <- watanabe Add 2007.03


        'タイプのチェック
        If Trim(form_no.w_type.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If


        'ピクチャのチェック
        'Brand Ver.5 TIFF->BMP 変更 start
        '   If form_no.ImgThumbnail1.Image = "" Then
        If form_no.ImgThumbnail1.Image Is Nothing Then
            'Brand Ver.5 TIFF->BMP 変更 end
            MsgBox("Picture is not specified.", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_UTQG = 0
        Exit Function

error_section:
        check_F_TMP_UTQG = 1

    End Function

    Sub Combo_Sousa(ByRef wk_combo As System.Windows.Forms.ComboBox, ByRef wk_str As Short)
        Dim list_cnt As Short
        Dim now_index As Short
        Dim wk_index As Short

        now_index = wk_combo.SelectedIndex
        For list_cnt = 1 To wk_combo.Items.Count
            wk_index = now_index + list_cnt
            If wk_index >= wk_combo.Items.Count Then
                wk_index = wk_index - wk_combo.Items.Count
            End If
            If Mid(VB6.GetItemString(wk_combo, wk_index), 1, 1) = Chr(wk_str) Then
                wk_combo.SelectedIndex = wk_index
                Exit Sub
            End If
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(wk_str)) <> 0 Then
                If Mid(VB6.GetItemString(wk_combo, wk_index), 1, 1) = Chr(wk_str + 32) Then
                    wk_combo.SelectedIndex = wk_index
                    Exit Sub
                End If
            End If
            If InStr("abcdefghijklmnopqrstuvwxyz", Chr(wk_str)) <> 0 Then
                If Mid(VB6.GetItemString(wk_combo, wk_index), 1, 1) = Chr(wk_str - 32) Then
                    wk_combo.SelectedIndex = wk_index
                    Exit Sub
                End If
            End If
        Next list_cnt

    End Sub


    Function w_size_chk(ByRef syori_flg As Short) As Short
        'パラメータ
        'syori_flg  :0 = すべてのチェックを行う
        '            1 = レングスチェックは行わない
        'w_size_chk :0 = ok
        '            1 = length_err
        '            2 = ng
        '
        'エラーとなった項目へのセットフォーカスも同時に行う
        Dim wk_ctrl As System.Windows.Forms.Control
        Dim ret As Short

        wk_ctrl = form_no.w_size1
        ret = charnum1_check(Trim(wk_ctrl.Text))
        '   If (ret = 2) Or (syori_flg = 0 And ret = 1) Then
        If (ret = 2) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Size - overall Diameter", 64)
            wk_ctrl.Focus()
            w_size_chk = ret
            Exit Function
        End If
        wk_ctrl = form_no.w_size2
        '   If (syori_flg = 0) And (Len(Trim$(wk_ctrl.Text)) <> 1) Then
        If (syori_flg = 0) And (Len(Trim(wk_ctrl.Text)) > 1) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Size -  /", 64)
            wk_ctrl.Focus()
            w_size_chk = 1
            Exit Function
        ElseIf Len(Trim(wk_ctrl.Text)) <> 0 Then
            If (Left(Trim(wk_ctrl.Text), 1) < "A" Or Left(Trim(wk_ctrl.Text), 1) > "Z") And Left(Trim(wk_ctrl.Text), 1) <> "/" Then
                MsgBox("Input data are incorrect." & Chr(13) & "Size -  /", 64)
                wk_ctrl.Focus()
                w_size_chk = 2
                Exit Function
            End If
        End If
        wk_ctrl = form_no.w_size3
        ret = charnum1_check(Trim(wk_ctrl.Text))
        If (ret = 2) Or (syori_flg = 0 And ret = 1) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Size - section width", 64)
            wk_ctrl.Focus()
            w_size_chk = ret
            Exit Function
        End If
        wk_ctrl = form_no.w_size5
        If (syori_flg = 0) And (Len(Trim(wk_ctrl.Text)) <> 1) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Size - structure", 64)
            wk_ctrl.Focus()
            w_size_chk = 1
            Exit Function
        ElseIf Len(Trim(wk_ctrl.Text)) <> 0 Then
            If (Left(Trim(wk_ctrl.Text), 1) <> "R") And (Left(Trim(wk_ctrl.Text), 1) <> "D") And (Left(Trim(wk_ctrl.Text), 1) <> "-") Then
                MsgBox("Input data is R · D · -." & Chr(13) & "Size - structure", 64)
                wk_ctrl.Focus()
                w_size_chk = 2
                Exit Function
            End If
        End If
        wk_ctrl = form_no.w_size6
        ret = charnum1_check(Trim(wk_ctrl.Text))
        If (ret = 2) Or (syori_flg = 0 And ret = 1) Then
            MsgBox("Input data are incorrect." & Chr(13) & "Size - rim diameter", 64)
            wk_ctrl.Focus()
            w_size_chk = ret
            Exit Function
        End If
        w_size_chk = 0

    End Function

    '***** 12/5.1997 yamamoto start *****
    '   ・新規に この関数を作成
    '   ・ﾁｪｯｸはﾌｫｰﾑからﾛｽﾄﾌｫｰｶｽ（一部ﾁｪﾝｼﾞ）ﾌﾟﾛｼｰｼﾞｬを流用、修正
    Function check_F_ZSEARCH_BRAND() As Short
        '
        'check_form_no = 0: OK
        '
        Dim irt As Short
        Dim f As System.Windows.Forms.Control

        form_no.w_pattern.Text = Trim(form_no.w_pattern.Text)
        form_no.w_size1.Text = Trim(form_no.w_size1.Text)
        form_no.w_size2.Text = Trim(form_no.w_size2.Text)
        form_no.w_size3.Text = Trim(form_no.w_size3.Text)
        form_no.w_size4.Text = Trim(form_no.w_size4.Text)
        form_no.w_size5.Text = Trim(form_no.w_size5.Text)
        form_no.w_size6.Text = Trim(form_no.w_size6.Text)
        form_no.w_kanri_no.Text = Trim(form_no.w_kanri_no.Text)
        form_no.w_entry_name.Text = Trim(form_no.w_entry_name.Text)
        form_no.w_entry_date_0.Text = Trim(form_no.w_entry_date_0.Text)
        form_no.w_entry_date_1.Text = Trim(form_no.w_entry_date_1.Text)

        '/パターン/
        f = form_no.w_pattern
        irt = check_1(form_no.w_pattern.Text, 6, 1, f)
        If irt = 2 Then GoTo err_section

        '/サイズ1/
        f = form_no.w_size1
        irt = check_1(form_no.w_size1.Text, 5, 1, f)
        If irt = 2 Then GoTo err_section

        '/サイズ3/
        f = form_no.w_size3
        irt = check_1(form_no.w_size3.Text, 5, 1, f)
        If irt = 2 Then GoTo err_section

        '/サイズ4/
        f = form_no.w_size4
        irt = check_1(form_no.w_size4.Text, 1, 1, f)
        If irt = 2 Then GoTo err_section

        '/サイズ5/
        f = form_no.w_size5
        irt = check_1(form_no.w_size5.Text, 1, 1, f)
        If irt = 2 Then GoTo err_section

        '/サイズ6/
        f = form_no.w_size6
        irt = check_1(form_no.w_size6.Text, 4, 1, f)
        If irt = 2 Then GoTo err_section

        '/業務管理番号/
        f = form_no.w_kanri_no
        irt = check_1(form_no.w_kanri_no.Text, 8, 1, f)
        If irt = 2 Then GoTo err_section

        '/登録者*change-events*/
        f = form_no.w_entry_name
        irt = check_1(form_no.w_entry_name.Text, 4, 1, f)
        If irt = 2 Then GoTo err_section

        check_F_ZSEARCH_BRAND = 0
        Exit Function

err_section:
        check_F_ZSEARCH_BRAND = 1
    End Function
    '***** 12/5.1997 yamamoto end ********

    Function check_F_TMP_PTNCODE2() As Short
        'Dim ss As String'20100616移植削除
        Dim slen As Short
        'Dim w_ret As Short'20100616移植削除
        Dim i As Short
        Dim w_str As String
        Dim w_msg As String

        '/pattern codeのチェック/
        ' 長さチェック
        slen = Len(form_no.w_ptncode.Text)
        If slen > 6 Or slen < 1 Then
            MsgBox("Input data are incorrect." & Chr(13) & "Pattern code is 6 characters.", 64)
            form_no.w_ptncode.Focus() 'SetFocus()
            GoTo error_section
        End If
        ' データチェック
        w_str = Left(form_no.w_ptncode.Text, 1)
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", w_str, 0) = 0 Then
            w_msg = "Input data are incorrect." & Chr(13)
            w_msg = w_msg & "Pattern code first character only English letter (capital letter) or numerical value."
            MsgBox(w_msg, 64)
            GoTo error_section
        End If

        w_str = Mid(form_no.w_ptncode.Text, 2)
        For i = 1 To Len(w_str)
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+-", Mid(w_str, i, 1), 0) = 0 Then
                w_msg = "Input data are incorrect." & Chr(13)
                w_msg = w_msg & "Pattern cord only as for the English letter( capital letter), the numerical value or the sign(+ or -)."
                MsgBox(w_msg, 64)
                GoTo error_section
            End If
        Next i


        '/タイプのチェック/
        If Trim(form_no.w_font.Text) = "" Then
            MsgBox("Input data are incorrect." & Chr(13) & "Type", 64)
            form_no.w_type.Focus() 'SetFocus()
            GoTo error_section
        End If

        check_F_TMP_PTNCODE2 = 0
        Exit Function

error_section:
        check_F_TMP_PTNCODE2 = 1

    End Function
End Module