Option Strict Off
Option Explicit On
Module Utl

    Function exist_gm_gz(ByRef Db_name_gm As Object, ByRef w_code As Object, ByRef w_no1 As String, ByRef w_no2 As String) As Short
        Dim result As Integer
        Dim a1 As String
        Dim a2 As String

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET 移行(一旦コメント化) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


        '原始文字が既に刻印図面に登録されているかチェックする
        '     戻り値 0:登録なし
        '            1:登録済み
        '            2:原始文字無し
        '            3:指定した刻印図面が登録されている
        '            -1:SQLｴﾗｰ

        exist_gm_gz = 0





        '検索キーセット
        key_code = " flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        key_code = key_code & " font_class1 = '" & Mid(w_code, 7, 1) & "' AND"
        key_code = key_code & " font_class2 = '" & Mid(w_code, 8, 1) & "' AND"
        key_code = key_code & " name1 = '" & Mid(w_code, 9, 1) & "' AND"
        key_code = key_code & " name2 = '" & Mid(w_code, 10, 1) & "'"

        '検索コマンド作成
        sqlcmd = "SELECT *  FROM " & Db_name_gm & " WHERE " & key_code

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, Db_name_gm, key_code)
        'If cnt = 0 Then
        '    exist_gm_gz = 2 ' 原始文字が見つからない
        '    Exit Function
        'ElseIf cnt = -1 Then
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)



        '検索キーセット
        key_code = " flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        key_code = key_code & " font_class1 = '" & Mid(w_code, 7, 1) & "' AND"
        key_code = key_code & " font_class2 = '" & Mid(w_code, 8, 1) & "' AND"
        key_code = key_code & " name1 = '" & Mid(w_code, 9, 1) & "' AND"
        key_code = key_code & " name2 = '" & Mid(w_code, 10, 1) & "' AND"
        key_code = key_code & " gz_id = 'KO'"

        '検索コマンド作成
        sqlcmd = "SELECT gz_no1, gz_no2  FROM " & Db_name_gm & " WHERE " & key_code

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, Db_name_gm, key_code)
        'If cnt = 0 Then
        '    exist_gm_gz = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '検索
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            a1 = Rs.rdoColumns(0).Value
        '        Else
        '            a1 = ""
        '        End If

        '        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
        '            a2 = Rs.rdoColumns(1).Value
        '        Else
        '            a2 = ""
        '        End If

        '        If w_no1 = a1 And w_no2 = a2 Then
        '            exist_gm_gz = 3
        '        Else
        '            exist_gm_gz = 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)

        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET 移行(一旦コメント化) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        exist_gm_gz = -1
    End Function

    Function exist_hm_hz(ByRef Db_name_hm As Object, ByRef w_code As Object, ByRef w_no1 As String, ByRef w_no2 As String) As Short
        Dim result As Integer
        Dim a1 As String
        Dim a2 As String

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        '----- .NET 移行(一旦コメント化) -----
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)




        '検索キーセット
        key_code = " flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        key_code = key_code & " no = '" & Mid(w_code, 7, 2) & "' "

        '検索コマンド作成
        sqlcmd = "SELECT *  FROM " & Db_name_hm & " WHERE ( " & key_code & ")"

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, Db_name_hm, key_code)
        'If cnt = 0 Then
        '    exist_hm_hz = 2 ' 編集文字が見つからない
        '    Exit Function
        'ElseIf cnt = -1 Then
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)





        '検索キーセット
        key_code = " flag_delete = 0 AND"
        key_code = key_code & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        key_code = key_code & " no = '" & Mid(w_code, 7, 2) & "' AND"
        key_code = key_code & " hz_id = 'HE'"

        '検索コマンド作成
        sqlcmd = "SELECT hz_no1, hz_no2  FROM " & Db_name_hm & " WHERE " & key_code

        'ヒット数チェック
        '----- .NET 移行(一旦コメント化) -----
        'cnt = VBRDO_Count(GL_T_RDO, Db_name_hm, key_code)
        'If cnt = 0 Then
        '    exist_hm_hz = 0

        'ElseIf cnt = -1 Then
        '    GoTo error_section

        'Else
        '    '検索
        '    Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '    Rs.MoveFirst()

        '    Do Until Rs.EOF

        '        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '            a1 = Rs.rdoColumns(0).Value
        '        Else
        '            a1 = ""
        '        End If

        '        If IsDBNull(Rs.rdoColumns(1).Value) = False Then
        '            a2 = Rs.rdoColumns(1).Value
        '        Else
        '            a2 = ""
        '        End If

        '        If w_no1 = a1 And w_no2 = a2 Then
        '            exist_hm_hz = 3
        '        Else
        '            exist_hm_hz = 1
        '        End If

        '        Rs.MoveNext()
        '    Loop

        '    Rs.Close()
        'End If
        ' <- watanabe edit VerUP(2011)

        Exit Function

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        '----- .NET 移行(一旦コメント化) -----
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

        exist_hm_hz = -1
    End Function
	
    Function update_gm_gz(ByRef Db_name_gm As Object, ByRef w_code As Object, ByRef w_gz_id As Object, ByRef w_gz_no1 As Object, ByRef w_gz_no2 As Object) As Short
        Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


        '原始文字に刻印図面情報を追加する
        '     戻り値  0:登録成功
        '            -1:登録失敗



        sqlcmd = "UPDATE " & Db_name_gm
        sqlcmd = sqlcmd & " SET gz_id = '" & w_gz_id & "',"
        sqlcmd = sqlcmd & " gz_no1 = '" & w_gz_no1 & "',"
        sqlcmd = sqlcmd & " gz_no2 = '" & w_gz_no2 & "'"
        sqlcmd = sqlcmd & " From " & Db_name_gm & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ("
        sqlcmd = sqlcmd & " flag_delete = 0 AND"
        sqlcmd = sqlcmd & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        sqlcmd = sqlcmd & " font_class1 = '" & Mid(w_code, 7, 1) & "' AND"
        sqlcmd = sqlcmd & " font_class2 = '" & Mid(w_code, 8, 1) & "' AND"
        sqlcmd = sqlcmd & " name1 = '" & Mid(w_code, 9, 1) & "' AND"
        sqlcmd = sqlcmd & " name2 = '" & Mid(w_code, 10, 1) & "')"

        'ｺﾏﾝﾄﾞ実行
        '----- .NET 移行(一旦コメント化) -----
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)

        update_gm_gz = 0
        Exit Function

error_section:
        update_gm_gz = -1
    End Function
	
	Function update_hm_hz(ByRef Db_name_hm As Object, ByRef w_code As Object, ByRef w_hz_id As Object, ByRef w_hz_no1 As Object, ByRef w_hz_no2 As Object) As Short
		Dim result As Integer

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


        '編集文字に編集文字図面情報を追加する
        '     戻り値  0:登録成功
        '            -1:登録失敗



        sqlcmd = "UPDATE " & Db_name_hm
        sqlcmd = sqlcmd & " SET  hz_id = '" & w_hz_id & "',"
        sqlcmd = sqlcmd & " hz_no1 = '" & w_hz_no1 & "',"
        sqlcmd = sqlcmd & " hz_no2 = '" & w_hz_no2 & "'"
        sqlcmd = sqlcmd & " From " & Db_name_hm & "(PAGLOCK)"
        sqlcmd = sqlcmd & " WHERE ("
        sqlcmd = sqlcmd & " flag_delete = 0 AND"
        sqlcmd = sqlcmd & " font_name = '" & Mid(w_code, 1, 6) & "' AND"
        sqlcmd = sqlcmd & " no = '" & Mid(w_code, 7, 2) & "' )"

        'ｺﾏﾝﾄﾞ実行
        '----- .NET 移行(一旦コメント化) -----
        'GL_T_RDO.Con.Execute(sqlcmd, RDO.OptionConstants.rdExecDirect)
        'If GL_T_RDO.Con.RowsAffected() = 0 Then
        '    GoTo error_section
        'End If
        ' <- watanabe edit VerUP(2011)

        update_hm_hz = 0
		Exit Function
		
error_section: 
		update_hm_hz = -1
    End Function
	

    'Function Get_Grid_Data(ByRef Sgrid As System.Windows.Forms.Control, ByRef Sdata As String, ByRef Srow As Short, ByRef Scol As Short) As Short
    Function Get_Grid_Data(ByRef Sgrid As Object, ByRef Sdata As String, ByRef Srow As Short, ByRef Scol As Short) As Short '20100616移植追加

        'Dim w_col As Object
        'Dim w_row As Short
        Dim w_col As Integer '20100616移植追加
        Dim w_row As Integer

        '状態の待避
        w_col = Sgrid.Col
        w_row = Sgrid.Row

        Sgrid.Row = Srow
        Sgrid.Col = Scol
        Sdata = Sgrid.Text

        '状態の復帰
        Sgrid.Col = w_col
        Sgrid.Row = w_row

    End Function
    'Function Set_Grid_Data(ByRef Sgrid As System.Windows.Forms.Control, ByRef Sdata As String, ByRef Srow As Short, ByRef Scol As Short) As Short
    Function Set_Grid_Data(ByRef Sgrid As Object, ByRef Sdata As String, ByRef Srow As Short, ByRef Scol As Short) As Short '20100616移植追加

        'Dim w_col As Object
        'Dim w_row As Short
        Dim w_col As Integer '20100616移植追加
        Dim w_row As Integer

        '状態の待避
        w_col = Sgrid.Col
        w_row = Sgrid.Row

        Sgrid.Row = Srow
        Sgrid.Col = Scol
        Sgrid.Text = Sdata

        '状態の復帰
        Sgrid.Col = w_col
        Sgrid.Row = w_row


    End Function

    '概要  ：ロックセット
    'ﾊﾟﾗﾒｰﾀ：rock_level,I,Integer,ロックレベル（ 0 = 通常  1 = 検索中）
    '説明  ：編集文字検索画面のロックセット
    '----- 1/27 1998 update by yamamoto -----
    Sub co_rockset_F_HMSEARCH(ByRef rock_level As Short)

        'Dim _form_no As F_HMSEARCH '20100615移植追加
        '_form_no = form_no

        If rock_level = 0 Then
            form_no.cmd_Cancel.Enabled = False
            '_form_no.cmd_Cancel.Enabled = False '20100615移植追加

            form_no.cmd_Search.Enabled = True
            form_no.cmd_CadRead.Enabled = True
            form_no.cmd_Clear.Enabled = True
            form_no.cmd_End.Enabled = True
            form_no.cmd_Help.Enabled = True
            form_no.w_font_name.Enabled = True
            form_no.w_no.Enabled = True
            form_no.w_spell.Enabled = True
            form_no.w_hikaku.Enabled = True
            form_no.w_high.Enabled = True
            form_no.w_entry_name.Enabled = True
            form_no.w_entry_date_0.Enabled = True
            form_no.w_entry_date_1.Enabled = True
            form_no.cmd_AllRead.Enabled = True
            form_no.cmd_ReadClear.Enabled = True

        Else
            form_no.cmd_Cancel.Enabled = True

            form_no.cmd_Search.Enabled = False
            form_no.cmd_CadRead.Enabled = False
            form_no.cmd_Clear.Enabled = False
            form_no.cmd_End.Enabled = False
            form_no.cmd_Help.Enabled = False
            form_no.w_font_name.Enabled = False
            form_no.w_no.Enabled = False
            form_no.w_spell.Enabled = False
            form_no.w_hikaku.Enabled = False
            form_no.w_high.Enabled = False
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_date_0.Enabled = False
            form_no.w_entry_date_1.Enabled = False
            form_no.cmd_AllRead.Enabled = False
            form_no.cmd_ReadClear.Enabled = False

        End If


    End Sub

    '概要  ：ロックセット
    'ﾊﾟﾗﾒｰﾀ：rock_level,I,Integer,ロックレベル（ 0 = 通常  1 = 検索中）
    '説明  ：編集文字検索画面のロックセット
    '----- 1/27 1998 update by yamamoto -----
    Sub co_rockset_F_HMSEARCH2(ByRef rock_level As Short)

        If rock_level = 0 Then
            form_no.cmd_Cancel.Enabled = False

            form_no.cmd_Search.Enabled = True
            form_no.cmd_CadRead.Enabled = True
            form_no.cmd_Clear.Enabled = True
            form_no.cmd_End.Enabled = True
            form_no.cmd_Help.Enabled = True
            form_no.w_gm_code.Enabled = True
            form_no.w_font_name.Enabled = True
            form_no.cmd_AllRead.Enabled = True
            form_no.cmd_ReadClear.Enabled = True

        Else
            form_no.cmd_Cancel.Enabled = True

            form_no.cmd_Search.Enabled = False
            form_no.cmd_CadRead.Enabled = False
            form_no.cmd_Clear.Enabled = False
            form_no.cmd_End.Enabled = False
            form_no.cmd_Help.Enabled = False
            form_no.w_gm_code.Enabled = True
            form_no.w_font_name.Enabled = False
            form_no.cmd_AllRead.Enabled = False
            form_no.cmd_ReadClear.Enabled = False

        End If


    End Sub

    '概要  ：ロックセット
    'ﾊﾟﾗﾒｰﾀ：rock_level,I,Integer,ロックレベル（ 0 = 通常  1 = 検索中）
    '説明  ：原始文字検索画面のロックセット
    '----- 1/27 1998 update by yamamoto -----
    Sub co_rockset_F_GMSEARCH(ByRef rock_level As Short)

        If rock_level = 0 Then
            form_no.cmd_Cancel.Enabled = False

            form_no.cmd_Search.Enabled = True
            form_no.cmd_CadRead.Enabled = True
            form_no.cmd_Clear.Enabled = True
            form_no.cmd_End.Enabled = True
            form_no.cmd_Help.Enabled = True
            form_no.w_font_name.Enabled = True
            form_no.w_font_class1.Enabled = True
            form_no.w_font_class2.Enabled = True
            form_no.w_name1.Enabled = True
            form_no.w_name2.Enabled = True
            form_no.w_hikaku.Enabled = True
            form_no.w_high.Enabled = True
            form_no.w_entry_name.Enabled = True
            form_no.w_entry_date_0.Enabled = True
            form_no.w_entry_date_1.Enabled = True
            form_no.cmd_AllRead.Enabled = True
            form_no.cmd_ReadClear.Enabled = True

        Else
            form_no.cmd_Cancel.Enabled = True

            form_no.cmd_Search.Enabled = False
            form_no.cmd_CadRead.Enabled = False
            form_no.cmd_Clear.Enabled = False
            form_no.cmd_End.Enabled = False
            form_no.cmd_Help.Enabled = False
            form_no.w_font_name.Enabled = False
            form_no.w_font_class1.Enabled = False
            form_no.w_font_class2.Enabled = False
            form_no.w_name1.Enabled = False
            form_no.w_name2.Enabled = False
            form_no.w_hikaku.Enabled = False
            form_no.w_high.Enabled = False
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_date_0.Enabled = False
            form_no.w_entry_date_1.Enabled = False
            form_no.cmd_AllRead.Enabled = False
            form_no.cmd_ReadClear.Enabled = False

        End If
    End Sub

    '概要  ：ロックセット
    'ﾊﾟﾗﾒｰﾀ：now_posi,I,Integer,  フォーム ( 0 = 番号検索　1 = ブランド検索　2 = 要素検索 ）
    '      ：rock_level,I,Integer,ロックレベル（ 0 = 通常  1 = 検索中）
    '説明  ：図面検索画面のロックセット
    '----- 1/27 1998 update by yamamoto -----
    'Sub co_rockset_F_ZSEARCH(ByRef now_posi As Short, ByRef rock_level As Short)
    Function co_rockset_F_ZSEARCH(ByRef now_posi As Short, ByRef rock_level As Short) As Short '20100616移植追加

        If rock_level = 0 Then
            form_no.cmd_Cancel.Enabled = False

            form_no.cmd_Search.Enabled = True
            form_no.cmd_ZumenRead.Enabled = True
            form_no.cmd_Clear.Enabled = True
            form_no.cmd_End.Enabled = True
            form_no.cmd_Help.Enabled = True

            Select Case now_posi
                Case 0
                    form_no.w_id.Enabled = True
                    form_no.w_no1.Enabled = True
                    form_no.w_no2.Enabled = True
                Case 1
                    form_no.w_pattern.Enabled = True
                    form_no.w_size1.Enabled = True
                    form_no.w_size2.Enabled = True
                    form_no.w_size3.Enabled = True
                    form_no.w_size4.Enabled = True
                    form_no.w_size5.Enabled = True
                    form_no.w_size6.Enabled = True
                    form_no.w_kanri_no.Enabled = True
                    form_no.w_entry_name.Enabled = True
                    form_no.w_entry_date_0.Enabled = True
                    form_no.w_entry_date_1.Enabled = True
                Case 2
                    form_no.w_mojicd.Enabled = True
                    form_no.w_taisho.Enabled = True
            End Select

        Else
            form_no.cmd_Cancel.Enabled = True

            form_no.cmd_Search.Enabled = False
            form_no.cmd_ZumenRead.Enabled = False
            form_no.cmd_Clear.Enabled = False
            form_no.cmd_End.Enabled = False
            form_no.cmd_Help.Enabled = False

            Select Case now_posi
                Case 0
                    form_no.w_id.Enabled = False
                    form_no.w_no1.Enabled = False
                    form_no.w_no2.Enabled = False
                Case 1
                    form_no.w_pattern.Enabled = False
                    form_no.w_size1.Enabled = False
                    form_no.w_size2.Enabled = False
                    form_no.w_size3.Enabled = False
                    form_no.w_size4.Enabled = False
                    form_no.w_size5.Enabled = False
                    form_no.w_size6.Enabled = False
                    form_no.w_kanri_no.Enabled = False
                    form_no.w_entry_name.Enabled = False
                    form_no.w_entry_date_0.Enabled = False
                    form_no.w_entry_date_1.Enabled = False
                Case 2
                    form_no.w_mojicd.Enabled = False
                    form_no.w_taisho.Enabled = False
            End Select

        End If

    End Function

    Function ConvTwipToPixel(ByVal form As Form, ByVal twip As Integer) As Integer


        Dim g As Graphics = form.CreateGraphics()
        Dim value As Integer = CInt((CSng(twip) * g.DpiX) / 1440.0F)

        ConvTwipToPixel = value

    End Function

End Module