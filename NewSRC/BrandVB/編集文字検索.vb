Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Collections.Generic

Friend Class F_HMSEARCH
    Inherits System.Windows.Forms.Form

    Private Const MarkRead As String = "◆"
    Private Const MarkNotRead As String = "◇"
    Private Const MarkDisp As String = "●"
    Private Const MarkNotDisp As String = "○"

    '概要：ボタンクリック処理
    '説明：キャンセルフラグを立てる
    '------- 1/23 1997 by yamamoto -------
    Private Sub cmd_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Cancel.Click

        GL_cancel_flg = 1

    End Sub

    '概要：ボタンクリック処理
    '説明：追加項目：ロックセット、キャンセルが選択されると、検索を中止（ グリッドはクリア ）

    Private Sub cmd_Search_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Search.Click
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
        Dim TiffFile, ZumenName As Object
        Dim w_file As String
        Dim srh_cnt As Integer
        Dim index_row As String

        ' -> watanabe add VerUP(2011)
        Dim key_code As String
        Dim sqlcmd As String
        Dim cnt As Integer
        'Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)


        GL_cancel_flg = 0
        srh_cnt = 0

        init_sql()

        On Error Resume Next
        Err.Clear()

        i_cnt = 1
        search_word(i_cnt) = " flag_delete = 0 "
        i_cnt = i_cnt + 1

        If w_font_name.Text <> "" Then
            search_word(i_cnt) = " font_name LIKE '" & w_font_name.Text & "'"
            i_cnt = i_cnt + 1
        End If

        If w_no.Text <> "" Then
            search_word(i_cnt) = " no LIKE '" & w_no.Text & "'"
            i_cnt = i_cnt + 1
        End If

        If w_spell.Text <> "" Then
            search_word(i_cnt) = " spell LIKE '" & w_spell.Text & "'"
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

        '----- .NET 移行 -----

        'ﾃﾞｰﾀﾍﾞｰｽ該当件数を表示

        '検索コマンド作成
        'sqlcmd = "SELECT COUNT(*) FROM " & DBTableName
        'If i_cnt > 1 Then
        '    sqlcmd = sqlcmd & " WHERE "
        '    For i = 1 To i_cnt - 2
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

        '-----------------------------------------------------

        'テーブルレコード数チェック
        Dim count As Integer = VBADO_Count(GL_T_ADO, DBTableName, joken)
        w_total.Text = count.ToString()

        '----- .NET 移行 -----

        If w_total.Text = "-1" Then
            MsgBox("Failed to find.")
            Exit Sub
        ElseIf w_total.Text = "0" Then
            MsgBox("There is no  editing characters drawing corresponding.")
            Exit Sub
        End If


        If CLng(w_total.Text) > AskNum Then
            w_ret = MsgBox("There is " & w_total.Text & " data. Would you like to view?", MsgBoxStyle.YesNo, "Confirmation")
            If w_ret = MsgBoxResult.No Then
                end_sql()
                MsgBox("Canceled the search.", , "Cancel")
                w_total.Text = ""
                'Brand Ver.5 TIFF->BMP 変更 start
                '            ImgThumbnail1.Image = ""
                ImgThumbnail1.Image = Nothing
                'Brand Ver.5 TIFF->BMP 変更 end
                Exit Sub
            End If
        End If

        co_rockset_F_HMSEARCH((1))

        '----- .NET 移行 -----
        '        MSFlexGrid1.Redraw = False

        '        'ｸﾞﾘｯﾄﾞに検索内容表示
        '        If CDbl(w_total.Text) > 0 Then
        '            MSFlexGrid1.Rows = CDbl(w_total.Text) + 1
        '            '   MSFlexGrid1.Cols = 21
        '        Else
        '            MSFlexGrid1.Rows = 2
        '            For i = 1 To MSFlexGrid1.Cols - 1
        '                w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
        '            Next i
        '            '   MSFlexGrid1.Cols = 21
        '        End If

        '        MSFlexGrid1.set_RowHeight(-1, 300)
        '        MSFlexGrid1.set_RowHeight(0, 400)
        '        index_row = "; NO "


        '        '検索コマンド作成
        '        sqlcmd = "SELECT font_name, no, spell, gm_num, width, high, ang, "
        '        sqlcmd = sqlcmd & " entry_name, entry_date FROM " & DBTableName
        '        If i_cnt > 1 Then
        '            sqlcmd = sqlcmd & " WHERE "
        '            For i = 1 To i_cnt - 2
        '                sqlcmd = sqlcmd & search_word(i) & " AND"
        '            Next i
        '            sqlcmd = sqlcmd & search_word(i_cnt - 1)
        '        End If
        '        sqlcmd = sqlcmd & " ORDER BY font_name, no "

        '        '検索
        '        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        '        Rs.MoveFirst()

        '        i = 0
        '        Do Until Rs.EOF

        '            System.Windows.Forms.Application.DoEvents()
        '            If GL_cancel_flg = 1 Then GoTo cancel_end_section

        '            i = i + 1
        '            MSFlexGrid1.Row = i

        '            For j = 1 To 12
        '                MSFlexGrid1.Col = j
        '                Select Case j
        '                    Case 1
        '                        MSFlexGrid1.Text = ""
        '                    Case 2
        '                        MSFlexGrid1.Text = "◇"
        '                    Case 3
        '                        MSFlexGrid1.Text = "○"
        '                    Case 12
        '                        If IsDBNull(Rs.rdoColumns(8).Value) = False Then
        '                            MSFlexGrid1.Text = VB6.Format(Rs.rdoColumns(8).Value, "yyyymmdd")
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

        '            TiffFile = TIFFDir & Trim(w_str(1)) & Trim(w_str(2)) & ".bmp"

        '            'BMPﾌｧｲﾙ表示
        '            w_file = Dir(TiffFile)
        '            If w_file <> "" Then
        '                ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
        '                ImgThumbnail1.Width = 457 '500 '20100701コード変更
        '                ImgThumbnail1.Height = 193 '200 '20100701コード変更
        '            Else
        '                MsgBox("BMP file can not be found." & TiffFile, MsgBoxStyle.Critical)
        '            End If

        '            MSFlexGrid1.Enabled = True
        '        Else
        '            MsgBox("There is no  editing characters drawing corresponding.")
        '            MSFlexGrid1.Enabled = False
        '        End If
        '        ' <- watanabe edit VerUP(2011)


        '        end_sql()

        '        MSFlexGrid1.FormatString = index_row
        '        MSFlexGrid1.set_FixedAlignment(0, 4)
        '        MSFlexGrid1.Redraw = True
        '        co_rockset_F_HMSEARCH((0))
        '        Exit Sub

        'cancel_end_section:

        '        ' -> watanabe add VerUP(2011)
        '        On Error Resume Next
        '        Err.Clear()
        '        Rs.Close()
        '        ' <- watanabe add VerUP(2011)

        '        end_sql()
        '        MSFlexGrid1.Rows = 2
        '        MSFlexGrid1.Cols = 13
        '        For lp = 0 To MSFlexGrid1.Cols - 1
        '            MSFlexGrid1.Row = 1
        '            MSFlexGrid1.Col = lp
        '            MSFlexGrid1.Text = ""
        '        Next lp
        '        MSFlexGrid1.Redraw = True
        '        w_total.Text = ""
        '        'Brand Ver.5 TIFF->BMP 変更 start
        '        '    ImgThumbnail1.Image = ""
        '        ImgThumbnail1.Image = Nothing
        '        'Brand Ver.5 TIFF->BMP 変更 end
        '        co_rockset_F_HMSEARCH((0))
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
        param.ColumnName = "no"
        paramList.Add(param)
        param.ColumnName = "spell"
        paramList.Add(param)
        param.ColumnName = "gm_num"
        param.SqlDbType = SqlDbType.SmallInt
        paramList.Add(param)
        param.ColumnName = "width"
        param.SqlDbType = SqlDbType.Float
        paramList.Add(param)
        param.ColumnName = "high"
        paramList.Add(param)
        param.ColumnName = "ang"
        paramList.Add(param)
        param.ColumnName = "entry_name"
        param.SqlDbType = SqlDbType.Char
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
                                      recordList(9))

            rowCnt += 1
            If rowCnt = 2 Then
                displayMark = MarkNotDisp
            End If
        Next

        For rowCnt = 0 To DataGridViewList.Rows.Count - 1
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.BackColor = SystemColors.Control
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.ForeColor = SystemColors.WindowText
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.SelectionBackColor = SystemColors.Control
            DataGridViewList.Rows(rowCnt).HeaderCell.Style.SelectionForeColor = SystemColors.WindowText
        Next

        '原始文字のイメージ(bmp)を表示
        TiffFile = TIFFDir & dataList(0)(0) & dataList(0)(1) & ".bmp"

        'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            ImgThumbnail1.Width = 457
            ImgThumbnail1.Height = 193
        Else
            MsgBox("BMP file can not be found." & TiffFile, MsgBoxStyle.Critical)
        End If

        co_rockset_F_HMSEARCH((0))

        '----- .NET 移行 -----

    End Sub

    Private Sub cmd_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Clear.Click

        Call Clear_F_HMSEARCH()

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
                .HelpContext = 402
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
        ' -> watanabe add VerUP(2011)
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
        '            MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures=" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
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
#If False Then
        For i = 1 To MSFlexGrid1.Rows - 1
            w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 2)
            w_ret = Get_Grid_Data(MSFlexGrid1, w_err, i, 1)
            '      If w_str = "◆" Then
            If w_str = "◆" And w_err = "" Then
                If FreePicNum < 1 Then
                    MsgBox("There are no free pictures." & Chr(13) & "Failed to read CAD.", MsgBoxStyle.Critical, "CAD reading error")
                    Exit For
                End If

                ZumenName = "HM-"
                w_ret = Get_Grid_Data(MSFlexGrid1, ss(1), i, 4)
                w_ret = Get_Grid_Data(MSFlexGrid1, ss(2), i, 5)
                ZumenName = ZumenName & Trim(ss(1))

                pic_no = -1


                '検索キーセット
                key_code = " font_name = '" & Trim(ss(1)) & "' AND"
                key_code = key_code & " no = '" & Trim(ss(2)) & "'"

                '検索コマンド作成
                sqlcmd = "SELECT haiti_pic FROM " & DBTableName & " WHERE (" & key_code & ")"

                'ヒット数チェック
                cnt = VBRDO_Count(GL_T_RDO, DBTableName, key_code)
                If cnt = 0 Or cnt = -1 Then
                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
                    w_ret = Set_Grid_Data(MSFlexGrid1, "999", i, 1)
                    GoTo LOOP_EXIT
                End If

                '検索
                Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
                Rs.MoveFirst()

                'ﾋﾟｸﾁｬ番号
                If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                    pic_no = Val(Rs.rdoColumns(0).Value)
                Else
                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
                    w_ret = Set_Grid_Data(MSFlexGrid1, "999", i, 1)
                    Rs.Close()
                    GoTo LOOP_EXIT
                End If

                Rs.Close()
                ' <- watanabe edit VerUP(2011)


                '----- .NET 移行 -----
                'w_mess = VB6.Format(pic_no, "000") & HensyuDir & ZumenName
                w_mess = pic_no.ToString("000") & HensyuDir & ZumenName

                w_ret = PokeACAD("ACADREAD", w_mess)
                w_ret = RequestACAD("ACADREAD")

                time_start = Now
                Do
                    time_now = Now
                    If Trim(form_main.Text2.Text) = "" Then
                        If System.DateTime.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.DateTime.FromOADate(timeOutSecond) Then
                            MsgBox("Time-out error.", 64, "ERROR")
                            w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                            w_ret = RequestACAD("ERROR")
                            GoTo LOOP_EXIT
                        End If
                    ElseIf VB.Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then
                        w_ret = Set_Grid_Data(MSFlexGrid1, "0", i, 1)
                        form_main.Text2.Text = ""
                        FreePicNum = FreePicNum - 1
                        GoTo LOOP_EXIT
                    ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                        err_flg = 1
                        error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                        w_ret = Set_Grid_Data(MSFlexGrid1, error_no, i, 1)
                        form_main.Text2.Text = ""
                        GoTo LOOP_EXIT
                    Else
                        MsgBox("Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
                        w_ret = Set_Grid_Data(MSFlexGrid1, "?", i, 1)
                        form_main.Text2.Text = ""
                        GoTo LOOP_EXIT
                    End If
                Loop
LOOP_EXIT:
            End If

        Next i
#End If
        '-----------------------------------------------------------------------------

        For i = 0 To DataGridViewList.Rows.Count - 1

            Dim mark As String = GetCellData(i, 1)
            Dim err As String = GetCellData(i, 0)
            ss(1) = GetCellData(i, 3)
            ss(2) = GetCellData(i, 4)

            If mark = MarkRead And err = "" Then

                If FreePicNum < 1 Then
                    MsgBox("There are no free pictures." & Chr(13) & "Failed to read CAD.", MsgBoxStyle.Critical, "CAD reading error")
                    Exit For
                End If

                ZumenName = "HM-" & Trim(ss(1))

                pic_no = -1


                '検索キーセット
                key_code = " font_name = '" & Trim(ss(1)) & "' AND"
                key_code = key_code & " no = '" & Trim(ss(2)) & "'"

                'テーブルレコード数チェック
                cnt = VBADO_Count(GL_T_ADO, DBTableName, key_code)

                If cnt = 0 Or cnt = -1 Then
                    MsgBox("Failed to read the CAD data of the " & i & " row", 64, "SQL error")
                    SetCellData(i, 0, "999")
                    Continue For
                End If

                'ピクチャー番号
                pic_no = Val(GetCellData(i, 12))
                w_mess = pic_no.ToString("000") & HensyuDir & ZumenName

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
        Err.Clear()
        'Rs.Close()
        ' <- watanabe add VerUP(2011)

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

        ' -> watanabe add VerUP(2011)
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
    '説明：追加項目：ロックセット
    Private Sub F_HMSEARCH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'Dim index_col As String

        form_no = Me
        temp_hm.Initilize() '20100702追加コード

        'test
        'DBTableName = "yama..hm_kanri"

        '----- .NET移行 (StartPositionプロパティをCenterScreenで対応) -----
        'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
        'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 4) ' フォームを画面の縦方向にセンタリングします。

        Call Clear_F_HMSEARCH()

        '----- .NET移行 (MSFlexGrid⇒DataGridViewで対応) -----
        Dim index_col() As String = {"error", "Read", "Display", "Font" & Chr(13) & "name", "Category", "Spell", "Number of" & Chr(13) & "primitive character", "Width",
                                     "Base" & Chr(13) & "height", "Base" & Chr(13) & "angle", "Registrant", "Record date"}

        With DataGridViewList

            .ColumnCount = 13
            .TopLeftHeaderCell.Value = "NO"

            For i As Integer = 0 To (.ColumnCount - 2)
                .Columns(i).HeaderCell.Value = index_col(i)
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(i).DefaultCellStyle.SelectionBackColor = SystemColors.Window
                .Columns(i).DefaultCellStyle.SelectionForeColor = SystemColors.WindowText
            Next

            'ピクチャー番号の列は非表示にする
            .Columns(12).Visible = False

            .Columns(0).Width = 30
            .Columns(1).Width = 20
            .Columns(2).Width = 20
            .Columns(3).Width = 70
            .Columns(4).Width = 30
            .Columns(5).Width = 70
            .Columns(6).Width = 70
            .Columns(7).Width = 60
            .Columns(8).Width = 60
            .Columns(9).Width = 60
            .Columns(10).Width = 60
            .Columns(11).Width = 70

        End With

        'MSFlexGrid1.Redraw = False
        'MSFlexGrid1.Rows = 2
        'MSFlexGrid1.Cols = 13

        '' 行高さの設定
        'MSFlexGrid1.set_RowHeight(-1, 300)
        'MSFlexGrid1.set_RowHeight(0, 400)

        'index_col = "^NO|^error|^Read|^Display|^Font" & Chr(13) & "name|^Category|^Spell|^Number of" & Chr(13) & "primitive character" & "|^Width|^Base" & Chr(13) & "height|^Base" & Chr(13) & "angle|^Registrant|^Record date"

        'MSFlexGrid1.FormatString = index_col

        '' 列幅の設定
        'MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 2)
        'MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 2)
        'MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 2)
        'MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 2)
        'MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 5)
        'MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 2)
        'MSFlexGrid1.set_ColWidth(6, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 4)
        'MSFlexGrid1.set_ColWidth(7, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 4)
        'MSFlexGrid1.set_ColWidth(8, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 4)
        'MSFlexGrid1.set_ColWidth(9, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 4)
        'MSFlexGrid1.set_ColWidth(10, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 4)
        'MSFlexGrid1.set_ColWidth(11, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 3)
        'MSFlexGrid1.set_ColWidth(12, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 44 * 5)

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

        co_rockset_F_HMSEARCH((0))

    End Sub

    '----- .NET移行 (ToDo:DataGridViewのイベントに変更) -----
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
        Dim w1 As String
        Dim w2 As String
        Dim w_select As Short
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
                        w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 2)
                        w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 1)
                        If w1 = "" And w2 = "◆" Then
                            w_select = w_select + 1
                        End If
                    Next i
                    If w_select >= FreePicNum Then
                        MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures=" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
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
            Else
                If w_row <> CDbl(w_show_no.Text) Then
                    w_ret = Set_Grid_Data(MSFlexGrid1, "○", CShort(w_show_no.Text), w_col)
                    w_ret = Set_Grid_Data(MSFlexGrid1, "●", w_row, w_col)
                    w_show_no.Text = CStr(w_row)
                    '文字コード
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(1), w_row, 4)
                    w_ret = Get_Grid_Data(MSFlexGrid1, w_str(2), w_row, 5)
                End If
            End If

            'Brand Ver.5 TIFF->BMP 変更 start
            '       TiffFile = TIFFDir & Trim$(w_str(1)) & Trim$(w_str(2)) & ".tif"
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
            TiffFile = TIFFDir & Trim(w_str(1)) & Trim(w_str(2)) & ".bmp"
            'BMPﾌｧｲﾙ表示
            w_file = Dir(TiffFile)
            If w_file <> "" Then
                ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
                ImgThumbnail1.Width = 457 '500 '20100701コード変更
                ImgThumbnail1.Height = 193 '200 '20100701コード変更
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
                                                           .Rows(rowIndex).Cells(colIndex + 2).Value.ToString() & ".bmp"
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

End Class