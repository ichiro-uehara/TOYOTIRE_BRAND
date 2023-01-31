Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Collections.Generic

Friend Class F_GMSEARCH
    Inherits System.Windows.Forms.Form

    Private Const MarkRead As String = "◆"
    Private Const MarkNotRead As String = "◇"
    Private Const MarkDisp As String = "●"
    Private Const MarkNotDisp As String = "○"

    '概要：ボタンクリック処理
    '説明：キャンセルフラグを立てる
    '----- 1/27 1998 by yamamoto -------
    Private Sub cmd_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Cancel.Click

        GL_cancel_flg = 1

    End Sub

    '概要：ボタンクリック処理
    '説明：追加項目：ロックセット、検索中、キャンセルが選択されると処理を中止する（ グリッドはクリア ）
    '----- 1/27 1998 by yamamoto change control_name
    Private Sub cmd_Search_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_search.Click
        Dim lp As Object
        Dim j As Object
        Dim w_ret As Object
        Dim num As String
        Dim result As Integer

        Dim search_word(20) As String
        Dim L_DAT(20) As String
        Dim i_cnt As Object
        Dim i As Short
        Dim w_str(10) As String
        'Dim TiffFile, ZumenName As Object
        Dim TiffFile As String
        Dim w_file As String
        Dim srh_cnt As Integer
        Dim index_row As String

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer

        '----- .NET 移行 -----
        'Dim Rs As RDO.rdoResultset

        ' <- watanabe add VerUP(2011)

        On Error Resume Next
        Err.Clear()

        GL_cancel_flg = 0
        srh_cnt = 0

        init_sql()

        i_cnt = 1
        search_word(i_cnt) = " flag_delete = 0 "
        i_cnt = i_cnt + 1

        If w_font_name.Text <> "" Then
            search_word(i_cnt) = " font_name LIKE '" & w_font_name.Text & "'"
            i_cnt = i_cnt + 1
        End If

        If w_font_class1.Text <> "" Then
            search_word(i_cnt) = " font_class1 LIKE '" & VB.Left(w_font_class1.Text, 1) & "'"
            i_cnt = i_cnt + 1
        End If

        If w_font_class2.Text <> "" Then
            search_word(i_cnt) = " font_class2 LIKE '" & w_font_class2.Text & "'"
            i_cnt = i_cnt + 1
        End If

        If w_name1.Text <> "" Then
            search_word(i_cnt) = " name1 LIKE '" & VB.Left(w_name1.Text, 1) & "'"
            i_cnt = i_cnt + 1
        End If
        If w_name2.Text <> "" Then
            search_word(i_cnt) = " name2 LIKE '" & w_name2.Text & "'"
            i_cnt = i_cnt + 1
        End If
        If w_high.Text <> "" Then
            search_word(i_cnt) = " high " & w_hikaku.Text & w_high.Text
            i_cnt = i_cnt + 1
        End If
        If w_entry_name.Text <> "" Then
            search_word(i_cnt) = " entry_name LIKE '" & w_entry_name.Text & "'"
            i_cnt = i_cnt + 1
        End If
        If w_entry_date_0.Text <> "" Then
            search_word(i_cnt) = " entry_date >= '" & w_entry_date_0.Text & " 00:00" & "'"
            i_cnt = i_cnt + 1
        End If
        If w_entry_date_1.Text <> "" Then
            search_word(i_cnt) = " entry_date <= '" & w_entry_date_1.Text & " 23:59" & "'"
            i_cnt = i_cnt + 1
        End If

        'Brand Cad System Ver.3 UP
        If w_old_font.Text <> "" Then
            search_word(i_cnt) = " old_font_name LIKE '" & w_old_font.Text & "'"
            i_cnt = i_cnt + 1
        End If

        '----- .NET 移行 -----
        ''検索コマンド作成
        'sqlcmd = "SELECT COUNT(*) FROM " & DBTableName
        'If i_cnt > 1 Then
        '    sqlcmd = sqlcmd & " WHERE "
        '    For i = 1 To i_cnt - 1
        '        sqlcmd = sqlcmd & search_word(i) & " AND "
        '    Next i
        '    sqlcmd = sqlcmd & search_word(i_cnt - 1)
        'End If

        '-----------------------------------------------------

        '検索条件作成
        Dim joken As String = ""
        If i_cnt > 2 Then
            For i = 1 To i_cnt - 2
                joken = joken & search_word(i) & " AND "
            Next i
        End If
        joken = joken & search_word(i_cnt - 1)

        '----- .NET 移行 -----

        '検索

        '----- .NET 移行 -----
        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        'Rs.MoveFirst()

        'w_total.Text = "-1"
        'Do Until Rs.EOF

        '    If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '        num = CStr(Val(Rs.rdoColumns(0).Value))
        '    Else
        '        num = "-1"
        '    End If
        '    w_total.Text = num

        '    Rs.MoveNext()
        'Loop

        'Rs.Close()

        '----------------------------------------------------------------------

        'テーブルレコード数チェック
        Dim count As Integer = VBADO_Count(GL_T_ADO, DBTableName, joken)
        w_total.Text = count.ToString()

        '----- .NET 移行 -----

        If w_total.Text = "-1" Then
            MsgBox("Failed to find.")
            Exit Sub
        ElseIf w_total.Text = "0" Then
            MsgBox("There is no Primitive character corresponding.")
            Exit Sub
        End If

        If CLng(w_total.Text) > AskNum Then
            w_ret = MsgBox("There is " & w_total.Text & " data. Would you like to view?", MsgBoxStyle.YesNo, "Confirmation")
            If w_ret = MsgBoxResult.No Then
                end_sql()
                MsgBox("Canceled the search.", , "Cancel")
                'Brand Ver.5 TIFF->BMP 変更 start
                '         ImgThumbnail1.Image = ""
                ImgThumbnail1.Image = Nothing
                'Brand Ver.5 TIFF->BMP 変更 end
                w_total.Text = ""
                Exit Sub
            End If
        End If

        'ｸﾞﾘｯﾄﾞに検索内容表示
        co_rockset_F_GMSEARCH((1))

        '----- .NET 移行 -----
        'MSFlexGrid1.Redraw = False
        'If CDbl(w_total.Text) > 0 Then
        '    MSFlexGrid1.Rows = CDbl(w_total.Text) + 1
        'Else
        '    MSFlexGrid1.Rows = 2
        '    For i = 0 To MSFlexGrid1.Cols - 1
        '        w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
        '    Next i
        'End If

        'MSFlexGrid1.set_RowHeight(-1, 300)
        'MSFlexGrid1.set_RowHeight(0, 400)
        'index_row = "; NO "


        '検索コマンド作成
        '        sqlcmd = "SELECT font_name, font_class1, font_class2, name1, name2, "
        '        sqlcmd = sqlcmd & " high, width, ang, moji_high, moji_shift, hem_width, "
        '        sqlcmd = sqlcmd & " hatch_ang, hatch_width, hatch_space, base_r, old_font_name, "
        '        sqlcmd = sqlcmd & " entry_name, entry_date FROM " & DBTableName
        '        If i_cnt > 1 Then
        '            sqlcmd = sqlcmd & " WHERE "
        '            For i = 1 To i_cnt - 2
        '                sqlcmd = sqlcmd & search_word(i) & " AND"
        '            Next i
        '            sqlcmd = sqlcmd & search_word(i_cnt - 1)
        '        End If
        '        sqlcmd = sqlcmd & " ORDER BY font_name, name1, name2"

        '        '検索
        '        '----- .NET 移行 -----
        '        'Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '        'Rs.MoveFirst()

        '        i = 0
        '        Do Until Rs.EOF

        '            System.Windows.Forms.Application.DoEvents()
        '            If GL_cancel_flg = 1 Then GoTo cancel_end_section

        '            i = i + 1
        '            MSFlexGrid1.Row = i

        '            For j = 1 To 21
        '                MSFlexGrid1.Col = j
        '                Select Case j
        '                    Case 1
        '                        MSFlexGrid1.Text = ""
        '                    Case 2
        '                        MSFlexGrid1.Text = "◇"
        '                    Case 3
        '                        MSFlexGrid1.Text = "○"
        '                    Case 21
        '                        If IsDBNull(Rs.rdoColumns(17).Value) = False Then
        '                            MSFlexGrid1.Text = VB6.Format(Rs.rdoColumns(17).Value, "yyyymmdd")

        '                        Else
        '                            MSFlexGrid1.Text = ""
        '                        End If
        '                    Case Else
        '                        If IsDBNull(Rs.rdoColumns(j - 4).Value) = False Then
        '                            L_DAT(j - 3) = Rs.rdoColumns(j - 4).Value
        '                        Else
        '                            L_DAT(j - 3) = ""
        '                        End If
        '                        MSFlexGrid1.Text = L_DAT(j - 3)
        '                End Select
        '            Next j

        '            srh_cnt = srh_cnt + 1
        '            w_total.Text = CStr(srh_cnt)

        '            index_row = index_row & "|" & VB6.Format(srh_cnt)

        '            Rs.MoveNext()
        '        Loop

        '        Rs.Close()

        '        grid_num = i

        '        If i > 0 Then
        '            w_show_no.Text = CStr(1)
        '            w_ret = Set_Grid_Data(MSFlexGrid1, "●", 1, 3)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str(1), CShort(w_show_no.Text), 4)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str(2), CShort(w_show_no.Text), 5)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str(3), CShort(w_show_no.Text), 6)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str(4), CShort(w_show_no.Text), 7)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str(5), CShort(w_show_no.Text), 8)

        '            TiffFile = TIFFDir & Trim(w_str(1)) & Trim(w_str(2)) & Trim(w_str(3)) & Trim(w_str(4)) & Trim(w_str(5)) & ".bmp"

        '            'BMPﾌｧｲﾙ表示
        '            w_file = Dir(TiffFile)
        '            If w_file <> "" Then
        '                ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
        '                ImgThumbnail1.Width = 457 '500 '20100701コード変更
        '                ImgThumbnail1.Height = 193 '200 '20100701コード変更
        '            Else
        '                MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
        '            End If
        '            MSFlexGrid1.Enabled = True
        '        Else
        '            MsgBox("There is no Primitive character corresponding.")
        '            MSFlexGrid1.Enabled = False
        '        End If
        '        ' <- watanabe edit VerUP(2011)


        '        end_sql()

        '        '----- 12/11 1997 yamamoto start -----
        '        MSFlexGrid1.FormatString = index_row
        '        MSFlexGrid1.set_FixedAlignment(0, 4)
        '        MSFlexGrid1.Redraw = True
        '        co_rockset_F_GMSEARCH((0))

        '        '----- 12/11 1997 yamamoto end -------
        '        Exit Sub

        'cancel_end_section:

        '        ' -> watanabe add VerUP(2011)
        '        On Error Resume Next
        '        Err.Clear()
        '        Rs.Close()
        '        ' <- watanabe add VerUP(2011)

        '        end_sql()
        '        w_total.Text = ""
        '        'Brand Ver.5 TIFF->BMP 変更 start
        '        '    ImgThumbnail1.Image = ""
        '        ImgThumbnail1.Image = Nothing
        '        'Brand Ver.5 TIFF->BMP 変更 end
        '        MSFlexGrid1.Rows = 2
        '        ' 1998.9.28 修正
        '        '    MSFlexGrid1.Cols = 21
        '        MSFlexGrid1.Cols = 22
        '        For lp = 0 To MSFlexGrid1.Cols - 1
        '            MSFlexGrid1.Row = 1
        '            MSFlexGrid1.Col = lp
        '            MSFlexGrid1.Text = ""
        '        Next lp
        '        MSFlexGrid1.Redraw = True
        '        co_rockset_F_GMSEARCH((0))
        '        MsgBox("Search has been canceled.", 64, "Cancel")

        '----------------------------------------------------------------------

        DataGridViewList.Rows.Clear()

        '検索パラメータ作成
        Dim paramList As List(Of ADO_PARAM_Struct) = New List(Of ADO_PARAM_Struct)
        Dim param As ADO_PARAM_Struct

        param.DataSize = 0
        param.Value = Nothing
        param.Sign = ""

        param.ColumnName = "font_name"
        param.SqlDbType = SqlDbType.Char
        paramList.Add(param)
        param.ColumnName = "font_class1"
        paramList.Add(param)
        param.ColumnName = "font_class2"
        paramList.Add(param)
        param.ColumnName = "name1"
        paramList.Add(param)
        param.ColumnName = "name2"
        paramList.Add(param)
        param.ColumnName = "high"
        param.SqlDbType = SqlDbType.Float
        paramList.Add(param)
        param.ColumnName = "width"
        paramList.Add(param)
        param.ColumnName = "ang"
        paramList.Add(param)
        param.ColumnName = "moji_high"
        paramList.Add(param)
        param.ColumnName = "moji_shift"
        paramList.Add(param)
        param.ColumnName = "hem_width"
        paramList.Add(param)
        param.ColumnName = "hatch_ang"
        paramList.Add(param)
        param.ColumnName = "hatch_width"
        paramList.Add(param)
        param.ColumnName = "hatch_space"
        paramList.Add(param)
        param.ColumnName = "base_r"
        paramList.Add(param)
        param.ColumnName = "old_font_name"
        param.SqlDbType = SqlDbType.Char
        paramList.Add(param)
        param.ColumnName = "entry_name"
        paramList.Add(param)
        param.ColumnName = "entry_date"
        param.SqlDbType = SqlDbType.SmallDateTime
        paramList.Add(param)
        param.ColumnName = "haiti_pic"
        param.SqlDbType = SqlDbType.TinyInt
        paramList.Add(param)


        'Databaseレコード検索処理
        Dim dataList As List(Of List(Of String)) = New List(Of List(Of String))
        If VBADO_Search(GL_T_ADO, DBTableName, joken, paramList, dataList) <> 1 Then
            MsgBox("Failed to find.")
            Exit Sub
        End If

        Dim displayMark As String = MarkDisp
        Dim rowCnt As Integer = 1
        For Each recordList As List(Of String) In dataList
            DataGridViewList.Rows.Add("",
                                      MarkNotRead,
                                      displayMark,
                                      recordList(0),
                                      recordList(1),
                                      recordList(2),
                                      recordList(3),
                                      recordList(4),
                                      recordList(5),
                                      recordList(6),
                                      recordList(7),
                                      recordList(8),
                                      recordList(9),
                                      recordList(10),
                                      recordList(11),
                                      recordList(12),
                                      recordList(13),
                                      recordList(14),
                                      recordList(15),
                                      recordList(16),
                                      recordList(17),
                                      recordList(18))

            rowCnt += 1
            If rowCnt = 2 Then
                displayMark = MarkNotDisp
            End If
        Next

        For rowCnt = 0 To DataGridViewList.Rows.Count - 1
            'DataGridViewList.Rows(rowCnt).HeaderCell.Value = CStr(rowCnt + 1)
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.BackColor = SystemColors.Control
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.ForeColor = SystemColors.WindowText
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.SelectionBackColor = SystemColors.Control
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.SelectionForeColor = SystemColors.WindowText
            'DataGridViewList.Rows(rowCnt).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        '原始文字のイメージ(bmp)を表示
        TiffFile = TIFFDir & dataList(0)(0) & dataList(0)(1) &
                             dataList(0)(2) & dataList(0)(3) &
                             dataList(0)(4) & ".bmp"

        'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            ImgThumbnail1.Width = 457
            ImgThumbnail1.Height = 193
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
        End If

        co_rockset_F_GMSEARCH((0))

        '----- .NET 移行 -----

    End Sub

    Private Sub cmd_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Clear.Click

        Call Clear_F_GMSEARCH()

    End Sub

    Private Sub cmd_End_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_End.Click
        'DB Disconnect
        end_sql()
        form_no.Close()
        End
    End Sub

    Private Sub cmd_Help_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Help.Click

        On Error Resume Next
        Err.Clear()
        Dim oCommonDialog As Object
        oCommonDialog = CreateObject("MSComDlg.CommonDialog")

        If Err.Number = 0 Then
            With oCommonDialog
                .HelpCommand = cdlHelpContext
                .HelpFile = "c:\VBhelp\BRAND.HLP"
                .HelpContext = 302
                .ShowHelp()
            End With
        End If
    End Sub

    Private Sub cmd_AllRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_AllRead.Click
        Dim w_ret As Object

        Dim w_col As Short
        Dim w_row As Short
        Dim w_err As String
        Dim w_select As Short
        Dim w1 As String
        Dim w2 As String
        Dim i As Short

        On Error Resume Next
        Err.Clear()


        '----- .NET 移行 -----
        '' -> watanabe add VerUP(2011)
        'w1 = ""
        'w2 = ""
        '' <- watanabe add VerUP(2011)

        'MSFlexGrid1.Col = 2

        'w_select = 0
        'i = 0
        'For i = 1 To grid_num
        '    MSFlexGrid1.Row = i

        '    w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
        '    w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)


        '    If w1 = "" And w2 = "◆" Then
        '        w_select = w_select + 1
        '    Else
        '        w_select = w_select + 1

        '        If w_select > FreePicNum Then
        '            MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures =" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
        '            Exit Sub
        '        Else
        '            MSFlexGrid1.Text = "◆"
        '            MSFlexGrid1.Text = "◆"
        '        End If
        '    End If

        'Next i

        '---------------------------------------------------

        With DataGridViewList

            w_select = 0

            For i = 0 To .Rows.Count - 1
                If GetCellData(i, 0) = "" Then
                    If GetCellData(i, 1) = MarkRead Then
                        w_select += 1
                    Else
                        w_select += 1

                        If w_select > FreePicNum Then
                            MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures =" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
                            Exit Sub
                        Else
                            SetCellData(i, 1, MarkRead)
                        End If
                    End If
                End If
            Next

        End With

        '----- .NET 移行 -----

    End Sub

    Private Sub cmd_CadRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_CadRead.Click
        Dim result As Object
        Dim pic_no As Integer
        Dim w_ret As Object

        Dim w_str As String
        Dim ss(6) As String
        Dim w_mess As String
        Dim ZumenName As String
        'Dim font_name As String'20100616移植削除
        Dim error_no As String
        Dim i As Short
        Dim time_start As Date
        Dim time_now As Date
        Dim w_err As String
        Dim err_flg As Short

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer

        '----- .NET 移行 -----
        'Dim Rs As RDO.rdoResultset

        ' <- watanabe add VerUP(2011)

        ' -> watanabe add VerUP(2011)
        On Error GoTo error_section
        Err.Clear()
        ' <- watanabe add VerUP(2011)


        ' -> watanabe add VerUP(2011)
        w_str = ""
        w_err = ""
        ' <- watanabe add VerUP(2011)


        err_flg = 0
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor '----- 12/11 1997 yamamoto add -----

        init_sql()

        '----- .NET 移行 -----
        '        For i = 1 To MSFlexGrid1.Rows - 1
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 2)
        '            w_ret = Get_Grid_Data(MSFlexGrid1, w_err, i, 1)
        '            '      If w_str = "◆" Then
        '            If w_str = "◆" And w_err = "" Then
        '                If FreePicNum < 1 Then
        '                    MsgBox("There are no free pictures." & Chr(13) & "Failed to read CAD.", MsgBoxStyle.Critical, "CAD reading error")
        '                    Exit For
        '                End If
        '                ZumenName = "GM-"
        '                w_ret = Get_Grid_Data(MSFlexGrid1, ss(1), i, 4)
        '                w_ret = Get_Grid_Data(MSFlexGrid1, ss(2), i, 5)
        '                w_ret = Get_Grid_Data(MSFlexGrid1, ss(3), i, 6)
        '                w_ret = Get_Grid_Data(MSFlexGrid1, ss(4), i, 7)
        '                w_ret = Get_Grid_Data(MSFlexGrid1, ss(5), i, 8)
        '                ZumenName = ZumenName & Trim(ss(1))

        '                pic_no = -1


        '                '検索キーセット
        '                key_code = " font_name = '" & Trim(ss(1)) & "' AND"
        '                key_code = key_code & " font_class1 = '" & Trim(ss(2)) & "' AND"
        '                key_code = key_code & " font_class2 = '" & Trim(ss(3)) & "' AND"
        '                key_code = key_code & " name1 = '" & Trim(ss(4)) & "' AND"
        '                key_code = key_code & " name2 = '" & Trim(ss(5)) & "'"

        '                '検索コマンド作成
        '                sqlcmd = "SELECT haiti_pic FROM " & DBTableName & " WHERE (" & key_code & ")"

        '                'ヒット数チェック
        '                cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
        '                If cnt = 0 Or cnt = -1 Then
        '                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
        '                    w_ret = Set_Grid_Data(MSFlexGrid1, "999", i, 1)
        '                    GoTo LOOP_EXIT
        '                End If

        '                '検索
        '                Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '                Rs.MoveFirst()

        '                'ﾋﾟｸﾁｬ番号
        '                If IsDBNull(Rs.rdoColumns(0).Value) = False Then
        '                    pic_no = Val(Rs.rdoColumns(0).Value)
        '                Else
        '                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
        '                    w_ret = Set_Grid_Data(MSFlexGrid1, "999", i, 1)
        '                    Rs.Close()
        '                    GoTo LOOP_EXIT
        '                End If

        '                Rs.Close()
        '                ' <- watanabe edit VerUP(2011)


        '                '----- .NET 移行 -----
        '                'w_mess = VB6.Format(pic_no, "000") & GensiDir & ZumenName
        '                w_mess = pic_no.ToString("000") & GensiDir & ZumenName

        '                w_ret = PokeACAD("ACADREAD", w_mess)
        '                w_ret = RequestACAD("ACADREAD")

        '                time_start = Now
        '                Do
        '                    time_now = Now
        '                    If Trim(form_main.Text2.Text) = "" Then
        '                        If System.DateTime.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.DateTime.FromOADate(timeOutSecond) Then
        '                            MsgBox("Time-out error", 64, "ERROR")
        '                            w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
        '                            w_ret = RequestACAD("ERROR")
        '                            GoTo LOOP_EXIT
        '                        End If
        '                    ElseIf VB.Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then
        '                        w_ret = Set_Grid_Data(MSFlexGrid1, "0", i, 1)
        '                        form_main.Text2.Text = ""
        '                        FreePicNum = FreePicNum - 1
        '                        GoTo LOOP_EXIT
        '                    ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
        '                        err_flg = 1
        '                        error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
        '                        w_ret = Set_Grid_Data(MSFlexGrid1, error_no, i, 1)
        '                        form_main.Text2.Text = ""
        '                        GoTo LOOP_EXIT
        '                    Else
        '                        MsgBox("ﾘﾀｰReturn code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
        '                        w_ret = Set_Grid_Data(MSFlexGrid1, "?", i, 1)
        '                        form_main.Text2.Text = ""
        '                        GoTo LOOP_EXIT
        '                    End If
        '                Loop
        'LOOP_EXIT:
        '            End If

        '        Next i

        '-----------------------------------------------------------------------------


        For i = 0 To DataGridViewList.Rows.Count - 1

            Dim mark As String = GetCellData(i, 1)
            Dim err As String = GetCellData(i, 0)
            ss(1) = GetCellData(i, 3)
            ss(2) = GetCellData(i, 4)
            ss(3) = GetCellData(i, 5)
            ss(4) = GetCellData(i, 6)
            ss(5) = GetCellData(i, 7)

            If mark = MarkRead And err = "" Then

                If FreePicNum < 1 Then
                    MsgBox("There are no free pictures." & Chr(13) & "Failed to read CAD.", MsgBoxStyle.Critical, "CAD reading error")
                    Exit For
                End If

                ZumenName = "GM-" & Trim(ss(1))

                pic_no = -1


                '検索キーセット
                key_code = " font_name = '" & Trim(ss(1)) & "' AND"
                key_code = key_code & " font_class1 = '" & Trim(ss(2)) & "' AND"
                key_code = key_code & " font_class2 = '" & Trim(ss(3)) & "' AND"
                key_code = key_code & " name1 = '" & Trim(ss(4)) & "' AND"
                key_code = key_code & " name2 = '" & Trim(ss(5)) & "'"

                'テーブルレコード数チェック
                cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

                If cnt = 0 Or cnt = -1 Then
                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
                    SetCellData(i, 0, "999")
                    Continue For
                End If

                'ピクチャー番号
                pic_no = Val(GetCellData(i, 21))
                w_mess = pic_no.ToString("000") & GensiDir & ZumenName

                'ピクチャー番号と図面パスを通知
                w_ret = PokeACAD("ACADREAD", w_mess)

                '[ACADREAD]リクエスト
                w_ret = RequestACAD("ACADREAD")

                time_start = Now
                Do
                    time_now = Now
                    If Trim(form_main.Text2.Text) = "" Then
                        If System.DateTime.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.DateTime.FromOADate(timeOutSecond) Then
                            MsgBox("Time-out error", 64, "ERROR")
                            w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                            w_ret = RequestACAD("ERROR")
                            Exit Do
                        End If
                    ElseIf VB.Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then
                        SetCellData(i, 0, "0")
                        form_main.Text2.Text = ""
                        FreePicNum = FreePicNum - 1
                        Exit Do
                    ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                        err_flg = 1
                        error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                        SetCellData(i, 0, error_no)
                        form_main.Text2.Text = ""
                        Exit Do
                    Else
                        MsgBox("Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
                        SetCellData(i, 0, "?")
                        form_main.Text2.Text = ""
                        Exit Do
                    End If
                Loop

            End If

        Next

        '----- .NET 移行 -----

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default '----- 12/11 1997 yamamoto add -----

        If err_flg = 0 Then
            MsgBox("CAD reading completion.")
        Else
            MsgBox("There was error reading. CAD reading completion.")
        End If

        end_sql()
        Exit Sub

error_section:
        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        '----- .NET 移行 -----
        'Rs.Close()
        '' <- watanabe add VerUP(2011)

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default '----- 12/11 1997 yamamoto add -----
        MsgBox("Failed to read CAD.", 64)
    End Sub

    Private Sub cmd_ReadClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_ReadClear.Click
        Dim w_ret As Object

        Dim w_col As Short
        Dim w_row As Short
        Dim w_err As String
        Dim w1 As String
        Dim w2 As String
        Dim i As Short

        On Error Resume Next
        Err.Clear()

        '----- .NET 移行 -----

        '' -> watanabe add VerUP(2011)
        'w1 = ""
        'w2 = ""
        '' <- watanabe add VerUP(2011)

        'MSFlexGrid1.Col = 2

        'For i = 1 To grid_num
        '    MSFlexGrid1.Row = i
        '    w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
        '    w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
        '    '              If w1 = "" Then
        '    MSFlexGrid1.Text = "◇"
        '    MSFlexGrid1.Text = "◇"

        '    '              End If
        'Next i

        '----------------------------------------------------

        For i = 0 To DataGridViewList.Rows.Count - 1
            SetCellData(i, 1, MarkNotRead)
        Next

        '----- .NET 移行 -----

    End Sub

    '概要：フォームロード
    '説明：追加項目：セット
    Private Sub F_GMSEARCH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        form_no = Me

        '----- .NET移行 (StartPositionプロパティをCenterScreenで対応) -----
        'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
        'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 4) ' フォームを画面の縦方向にセンタリングします。

        'フォント区分１
        w_font_class1.Items.Clear()
        w_font_class1.Items.Add("A:Solid")
        w_font_class1.Items.Add("F:Hemming letter")
        w_font_class1.Items.Add("H:Hutchings letter")
        w_font_class1.Items.Add("B:Edge & Hutchings")
        w_font_class1.Items.Add("D:Dummy letter")
        w_font_class1.Items.Add("N:Screw")
        w_font_class1.Items.Add("P:Plate")

        '文字名１
        w_name1.Items.Clear()
        w_name1.Items.Add("A:an alphabetic character")
        w_name1.Items.Add("B:Number")
        w_name1.Items.Add("C:Hiragana letter")
        w_name1.Items.Add("D:Katakana letter")
        w_name1.Items.Add("E:kanji letter")
        w_name1.Items.Add("F:Etc")

        Call Clear_F_GMSEARCH()

        '----- .NET移行 (MSFlexGrid⇒DataGridViewで対応) -----
        Dim index_col() As String = {"error", "Read", "Display", "Font" & Chr(13) & "name", "Category" & Chr(13) & "1",
                                     "Category" & Chr(13) & "2", "name" & Chr(13) & "1", "name" & Chr(13) & "2", "Height", "Width",
                                     "Angle", "Character" & Chr(13) & "height", "Shift" & Chr(13) & "length", "Border" & Chr(13) & "width", "Hatching" & Chr(13) & "angle",
                                     "Hatching" & Chr(13) & "width", "Hatching" & Chr(13) & "interval", "Base R", "Old font name", "Registrant", "Record date"}

        With DataGridViewList

            .ColumnCount = 22
            .TopLeftHeaderCell.Value = "NO"

            For i As Integer = 0 To (.ColumnCount - 2)
                .Columns(i).HeaderCell.Value = index_col(i)
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(i).DefaultCellStyle.SelectionBackColor = SystemColors.Window
                .Columns(i).DefaultCellStyle.SelectionForeColor = SystemColors.WindowText
            Next

            'ピクチャー番号の列は非表示にする
            .Columns(21).Visible = False

            .Columns(0).Width = 30
            .Columns(1).Width = 20
            .Columns(2).Width = 20
            .Columns(3).Width = 70
            .Columns(4).Width = 20
            .Columns(5).Width = 20
            .Columns(6).Width = 20
            .Columns(7).Width = 20
            .Columns(8).Width = 50
            .Columns(9).Width = 50
            .Columns(10).Width = 40
            .Columns(11).Width = 50
            .Columns(12).Width = 50
            .Columns(13).Width = 50
            .Columns(14).Width = 50
            .Columns(15).Width = 50
            .Columns(16).Width = 50
            .Columns(17).Width = 50
            .Columns(18).Width = 70
            .Columns(19).Width = 40
            .Columns(20).Width = 70

        End With

        'MSFlexGrid1.Redraw = False

        'MSFlexGrid1.Rows = 2
        'MSFlexGrid1.Cols = 22

        '' 行高さの設定
        'MSFlexGrid1.set_RowHeight(-1, 300) '----- 12/11 1997 yamamoto change 600→400 -----
        'MSFlexGrid1.set_RowHeight(0, 400)


        'MSFlexGrid1.FormatString = index_col

        '' 列幅の設定
        'MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 3)
        'MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 3)
        'MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 7)
        'MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(6, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(7, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(8, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 2)
        'MSFlexGrid1.set_ColWidth(9, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(10, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(11, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 4)
        'MSFlexGrid1.set_ColWidth(12, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(13, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(14, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(15, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(16, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(17, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)
        'MSFlexGrid1.set_ColWidth(18, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 5)

        'MSFlexGrid1.set_ColWidth(19, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 7)
        'MSFlexGrid1.set_ColWidth(20, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 4)
        'MSFlexGrid1.set_ColWidth(21, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 100) / 60 * 7)

        'MSFlexGrid1.Redraw = True

        '高さｺﾝﾎﾞﾎﾞｯｸｽ
        w_hikaku.Items.Clear()
        w_hikaku.Items.Add("=")
        w_hikaku.Items.Add("<")
        w_hikaku.Items.Add(">")
        w_hikaku.Items.Add("<=")
        w_hikaku.Items.Add(">=")

        '----- .NET移行  -----
        'w_hikaku.Text = VB6.GetItemString(w_hikaku, 0)
        w_hikaku.SelectedIndex = 0

        co_rockset_F_GMSEARCH((0))

    End Sub

    '----- .NET移行 (DataGridViewのイベントに変更)-----
#If False Then
    Private Sub MSFlexGrid1_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent) Handles MSFlexGrid1.MouseDownEvent
        Dim TiffFile As Object
        Dim w_ret As Object

        Dim ZumenName As String
        Dim w_str(5) As String
        Dim w_file As String
        Dim w_col As Short
        Dim w_row As Short
        Dim w_err As String
        Dim w_select As Short
        Dim w1 As String
        Dim w2 As String
        Dim i As Short

        On Error Resume Next
        Err.Clear()


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w1 = ""
        w2 = ""
        ' <- watanabe add VerUP(2011)


        w_col = Val(CStr(MSFlexGrid1.Col))
        w_row = Val(CStr(MSFlexGrid1.Row))

        MSFlexGrid1.Redraw = False

        '// w_col=2(読み込み),w_col=3(表示)
        If w_col = 2 Then
            w_ret = Get_Grid_Data(MSFlexGrid1, w_err, w_row, 1)
            If w_err = "" Then
                If MSFlexGrid1.Text = "◆" Then
                    MSFlexGrid1.Text = "◇"
                Else
                    w_select = 0
                    For i = 1 To MSFlexGrid1.Rows - 1
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
                        If w1 = "" And w2 = "◆" Then
                            w_select = w_select + 1
                        End If
                    Next i
                    If w_select >= FreePicNum Then
                        MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures =" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
                    Else
                        MSFlexGrid1.Text = "◆"
                    End If
                End If
            End If

        End If
        If w_col = 3 Then
            If MSFlexGrid1.Text = "●" Then
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str(1), w_row, 4)
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str(2), w_row, 5)
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str(3), w_row, 6)
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str(4), w_row, 7)
                w_ret = Get_Grid_Data(MSFlexGrid1, w_str(5), w_row, 8)
            Else
                If w_row <> CDbl(w_show_no.Text) Then
                    w_ret = Set_Grid_Data(MSFlexGrid1, "○", CShort(w_show_no.Text), w_col)
                    w_ret = Set_Grid_Data(MSFlexGrid1, "●", w_row, w_col)
                    w_show_no.Text = CStr(w_row)
                    '文字コード
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(1), w_row, 4)
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(2), w_row, 5)
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(3), w_row, 6)
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(4), w_row, 7)
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(5), w_row, 8)
                End If
            End If

            'Brand Ver.5 TIFF->BMP 変更 start
            '       TiffFile = TIFFDir & Trim$(w_str(1)) & Trim$(w_str(2)) & Trim$(w_str(3)) & Trim$(w_str(4)) & Trim$(w_str(5)) & ".tif"
            '       'MsgBox "tifffile=" & TiffFile
            '
            '       'Tiffﾌｧｲﾙ表示
            '       w_file = Dir(TiffFile)
            '       If w_file <> "" Then
            '           ImgThumbnail1.Image = TiffFile
            '           ImgThumbnail1.ThumbWidth = 500
            '           ImgThumbnail1.ThumbHeight = 200
            '       Else
            '           MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
            '       End If
            TiffFile = TIFFDir & Trim(w_str(1)) & Trim(w_str(2)) & Trim(w_str(3)) & Trim(w_str(4)) & Trim(w_str(5)) & ".bmp"
            'BMPﾌｧｲﾙ表示
            w_file = Dir(TiffFile)
            If w_file <> "" Then
                ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
                'ImgThumbnail1.ScaleWidth = 500
                ImgThumbnail1.Width = 457 '20100701コード変更
                'ImgThumbnail1.ScaleHeight = 200
                ImgThumbnail1.Height = 193 '20100701コード変更
            Else
                MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            End If
            'Brand Ver.5 TIFF->BMP 変更 end

        End If

        MSFlexGrid1.Redraw = True

    End Sub
#End If

    '----- .NET移行  -----
    'DataGridViewList CellMouseDownイベント
    Private Sub DataGridViewList_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridViewList.CellMouseDown

        Try

            Dim colIndex As Integer = e.ColumnIndex
            Dim rowIndex As Integer = e.RowIndex

            If rowIndex < 0 Then
                Exit Sub
            End If

            With DataGridViewList

                If colIndex = 1 Then
                    '「Read」カラム(読込み)

                    Dim mark As String = GetCellData(rowIndex, colIndex)
                    Dim err As String = GetCellData(rowIndex, colIndex - 1)

                    If err = "" Then
                        If mark = MarkRead Then
                            SetCellData(rowIndex, colIndex, MarkNotRead)
                        Else
                            Dim selCnt As Integer = 0

                            For i As Integer = 0 To .Rows.Count - 1
                                If GetCellData(i, colIndex - 1) = "" And GetCellData(i, colIndex) = MarkRead Then
                                    selCnt += 1
                                End If
                            Next

                            If selCnt >= FreePicNum Then
                                MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures =" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
                            Else
                                SetCellData(rowIndex, colIndex, MarkRead)
                            End If

                        End If
                    End If

                ElseIf colIndex = 2 Then
                    '「Display」カラム(表示)

                    Dim mark As String = GetCellData(rowIndex, colIndex)

                    If mark = MarkNotDisp Then
                        For i As Integer = 0 To .Rows.Count - 1
                            If .Rows(i).Cells(colIndex).Value.ToString() = MarkDisp Then
                                .Rows(i).Cells(colIndex).Value = MarkNotDisp
                                Exit For
                            End If
                        Next

                        SetCellData(rowIndex, colIndex, MarkDisp)

                        '原始文字のイメージ(bmp)を表示
                        Dim TiffFile As String = TIFFDir & .Rows(rowIndex).Cells(colIndex + 1).Value.ToString() &
                                                           .Rows(rowIndex).Cells(colIndex + 2).Value.ToString() &
                                                           .Rows(rowIndex).Cells(colIndex + 3).Value.ToString() &
                                                           .Rows(rowIndex).Cells(colIndex + 4).Value.ToString() &
                                                           .Rows(rowIndex).Cells(colIndex + 5).Value.ToString() & ".bmp"
                        Dim w_file As String = Dir(TiffFile)
                        If w_file <> "" Then
                            ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
                            ImgThumbnail1.Width = 457
                            ImgThumbnail1.Height = 193
                        Else
                            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
                        End If

                    End If

                End If

            End With

        Catch ex As Exception

            MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    '概要  DataGridViewList セルの値取得
    'パラメータ：rowIndex	I		行Index
    '          ：colIndex	I		列Index
    '          ：戻り値				セルの値
    Private Function GetCellData(ByVal rowIndex As Integer, ByVal colIndex As Integer) As String

        Try

            GetCellData = DataGridViewList.Rows(rowIndex).Cells(colIndex).Value.ToString()

        Catch ex As Exception

            GetCellData = ""
            MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Function

    '概要  DataGridViewList セルの値設定
    'パラメータ：rowIndex	I		行Index
    '          ：colIndex	I		列Index
    '          ：data		I		設定データ
    '          ：戻り値				処理結果 (1:OK / 0:NG)
    Private Function SetCellData(ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal data As String) As Integer

        Try

            DataGridViewList.Rows(rowIndex).Cells(colIndex).Value = data
            SetCellData = 1

        Catch ex As Exception

            SetCellData = 0
            MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Function

    'UPGRADE_WARNING: イベント w_font_class1.SelectedIndexChanged は、フォームが初期化されたときに発生します。
    Private Sub w_font_class1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_class1.SelectedIndexChanged

        If w_font_class1.Text = VB6.GetItemString(w_font_class1, 4) Or w_font_class1.Text = VB6.GetItemString(w_font_class1, 5) Or w_font_class1.Text = VB6.GetItemString(w_font_class1, 6) Then
            w_name1.Text = VB6.GetItemString(w_name1, 5)
            w_name1.Enabled = False
            w_name2.Text = ""
        Else
            w_name1.Enabled = True
        End If

    End Sub

    Private Sub w_font_class1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_font_class1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then GoTo EventExitSub
        Call Combo_Sousa(w_font_class1, KeyAscii)
        KeyAscii = 0

EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub w_font_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_name.Leave

        form_no.w_font_name.Text = UCase(Trim(form_no.w_font_name.Text))

    End Sub

    Private Sub w_hikaku_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_hikaku.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then GoTo EventExitSub
        Call Combo_Sousa(w_hikaku, KeyAscii)
        KeyAscii = 0

EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'UPGRADE_WARNING: イベント w_name1.SelectedIndexChanged は、フォームが初期化されたときに発生します。
    Private Sub w_name1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name1.SelectedIndexChanged

        'UPGRADE_WARNING: Null/IsNull() の使用が見つかりました。
        If w_name1.Text Is System.DBNull.Value Or w_name1.Text <> dummy_text Then
            w_name2.Text = ""
        End If

    End Sub

    Private Sub w_name1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name1.Enter

        dummy_text = w_name1.Text

    End Sub

    Private Sub w_name1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_name1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then GoTo EventExitSub
        Call Combo_Sousa(w_name1, KeyAscii)
        KeyAscii = 0

EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub w_name2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_name2.Leave

        form_no.w_name2.Text = UCase(Trim(form_no.w_name2.Text))

    End Sub

    Private Sub w_old_font_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_old_font.Leave

        form_no.w_old_font.Text = UCase(Trim(form_no.w_old_font.Text))

    End Sub

    '----- .NET移行  -----
    'DataGridViewList CellPaintingイベント
    '行番号を描画する
    Private Sub DataGridViewList_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridViewList.CellPainting

        Try

            If e.ColumnIndex < 0 And e.RowIndex >= 0 Then
                'セルを描画する
                e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

                '行番号を描画する範囲を決定する
                Dim indexRect As Rectangle = e.CellBounds

                indexRect.Inflate(-2, -2)
                '行番号を描画する
                TextRenderer.DrawText(e.Graphics,
                    (e.RowIndex + 1).ToString(),
                    e.CellStyle.Font,
                    indexRect,
                    e.CellStyle.ForeColor,
                    TextFormatFlags.Right Or TextFormatFlags.VerticalCenter)
                '描画が完了したことを知らせる
                e.Handled = True
            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class