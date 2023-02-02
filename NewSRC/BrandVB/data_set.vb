Option Strict Off
Option Explicit On
Module MJ_DataSet
    'Sub dataset_bz_combo(ByRef wk_ctrl As System.Windows.Forms.Control, ByRef wk_index As Short)
    Sub dataset_bz_combo(ByRef wk_ctrl As Object, ByRef wk_index As Short) '20100616移植追加()
        Select Case wk_index
            Case 0, 1
                'wk_ctrl.ListIndex = wk_index
                wk_ctrl.Text = wk_ctrl.GetItemText(wk_ctrl.Items(wk_index)) '20100624コード変更
            Case Else
        End Select
    End Sub

    Sub dataset_bz_int(ByRef wk_ctrl As System.Windows.Forms.Control, ByRef wk_flag As Short)
        'wk_flag = 0:コントロールwk_ctrlに「無し」と表示
        '          1:コントロールwk_ctrlに「有り」と表示
        Select Case wk_flag
            Case 0
                wk_ctrl.Text = "Null"
            Case 1
                wk_ctrl.Text = "YES"
        End Select

    End Sub

    Sub dataset_F_BZDELE()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        'MsgBox "ブランド図面内容確認画面にデータをセットします"

        form_no.w_id.Text = temp_bz.id
        form_no.w_no1.Text = temp_bz.no1
        form_no.w_no2.Text = temp_bz.no2
        form_no.w_kanri_no.Text = temp_bz.kanri_no '20100705コード変更　.Textの追加
        form_no.w_comment.Text = temp_bz.comment
        form_no.w_dep_name.Text = temp_bz.dep_name
        form_no.w_entry_name.Text = temp_bz.entry_name
        form_no.w_entry_date.Text = temp_bz.entry_date
        form_no.w_syurui.Text = temp_bz.syurui
        form_no.w_pattern.Text = temp_bz.pattern
        form_no.w_syubetu.Text = temp_bz.syubetu
        form_no.w_size.Text = temp_bz.Size
        form_no.w_size1.Text = temp_bz.size1
        form_no.w_size2.Text = temp_bz.size2
        form_no.w_size3.Text = temp_bz.size3
        form_no.w_size4.Text = temp_bz.size4
        form_no.w_size5.Text = temp_bz.size5
        form_no.w_size6.Text = temp_bz.size6
        form_no.w_size7.Text = temp_bz.size7
        '97.04.24 update n.matsumi start .......................................
        ' form_no.w_size8.Text = temp_bz.size6
        form_no.w_size8.Text = temp_bz.size8
        '97.04.24 update n.matsumi ended .......................................
        form_no.w_size_code.Text = temp_bz.size_code
        Call dataset_plant(form_no.w_plant, temp_bz.plant, "dele")
        form_no.w_plant_code.Text = temp_bz.plant_code
        form_no.w_kikaku.Text = temp_bz.kikaku
        Call dataset_kikaku(form_no.w_kikaku1, Mid(temp_bz.kikaku, 1, 1))
        Call dataset_kikaku(form_no.w_kikaku2, Mid(temp_bz.kikaku, 2, 1))
        Call dataset_kikaku(form_no.w_kikaku3, Mid(temp_bz.kikaku, 3, 1))
        Call dataset_kikaku(form_no.w_kikaku4, Mid(temp_bz.kikaku, 4, 1))
        Call dataset_kikaku(form_no.w_kikaku5, Mid(temp_bz.kikaku, 5, 1))
        Call dataset_kikaku(form_no.w_kikaku6, Mid(temp_bz.kikaku, 6, 1))
        Call dataset_bz_int(form_no.w_tos_moyou, temp_bz.tos_moyou)
        Call dataset_bz_int(form_no.w_side_moyou, temp_bz.side_moyou)
        Call dataset_bz_int(form_no.w_side_kenti, temp_bz.side_kenti)
        Call dataset_bz_int(form_no.w_peak_mark, temp_bz.peak_mark)
        Call dataset_bz_int(form_no.w_nasiji, temp_bz.nasiji)
        form_no.w_hm_num.Text = Str(temp_bz.hm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_bz.hm_num - 1) / 5) + 2

        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)

        Next i
        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_bz.hm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub


    Sub dataset_F_BZLOOK()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        'MsgBox "ブランド図面内容確認画面にデータをセットします"

        form_no.w_flag_delete.Text = CStr(temp_bz.flag_delete)
        form_no.w_id.Text = temp_bz.id
        form_no.w_no1.Text = temp_bz.no1
        form_no.w_no2.Text = temp_bz.no2
        form_no.w_kanri_no.Text = temp_bz.kanri_no
        form_no.w_comment.Text = temp_bz.comment
        form_no.w_dep_name.Text = temp_bz.dep_name
        form_no.w_entry_name.Text = temp_bz.entry_name
        form_no.w_entry_date.Text = temp_bz.entry_date
        form_no.w_syurui.Text = temp_bz.syurui
        form_no.w_pattern.Text = temp_bz.pattern
        form_no.w_syubetu.Text = temp_bz.syubetu
        form_no.w_size.Text = temp_bz.Size
        form_no.w_size1.Text = temp_bz.size1
        form_no.w_size2.Text = temp_bz.size2
        form_no.w_size3.Text = temp_bz.size3
        form_no.w_size4.Text = temp_bz.size4
        form_no.w_size5.Text = temp_bz.size5
        form_no.w_size6.Text = temp_bz.size6
        form_no.w_size7.Text = temp_bz.size7
        form_no.w_size8.Text = temp_bz.size8
        form_no.w_size_code.Text = temp_bz.size_code
        Call dataset_plant(form_no.w_plant, temp_bz.plant, "look")
        form_no.w_plant_code.Text = temp_bz.plant_code
        form_no.w_kikaku.Text = temp_bz.kikaku
        Call dataset_kikaku(form_no.w_kikaku1, Mid(temp_bz.kikaku, 1, 1))
        Call dataset_kikaku(form_no.w_kikaku2, Mid(temp_bz.kikaku, 2, 1))
        Call dataset_kikaku(form_no.w_kikaku3, Mid(temp_bz.kikaku, 3, 1))
        Call dataset_kikaku(form_no.w_kikaku4, Mid(temp_bz.kikaku, 4, 1))
        Call dataset_kikaku(form_no.w_kikaku5, Mid(temp_bz.kikaku, 5, 1))
        Call dataset_kikaku(form_no.w_kikaku6, Mid(temp_bz.kikaku, 6, 1))
        Call dataset_bz_int(form_no.w_tos_moyou, temp_bz.tos_moyou)
        Call dataset_bz_int(form_no.w_side_moyou, temp_bz.side_moyou)
        Call dataset_bz_int(form_no.w_side_kenti, temp_bz.side_kenti)
        Call dataset_bz_int(form_no.w_peak_mark, temp_bz.peak_mark)
        Call dataset_nasiji(form_no.w_nasiji, temp_bz.nasiji)
        form_no.w_hm_num.Text = temp_bz.hm_num

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_bz.hm_num - 1) / 5) + 2

        ' -> watanabe add VerUP(2011)
        For i = 0 To form_no.MSFlexGrid1.Rows - 1
            form_no.MSFlexGrid1.set_RowHeight(i, 400)
        Next i

        form_no.MSFlexGrid1.set_ColWidth(0, 1000)
        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.set_ColWidth(i, 1900)
        Next i
        ' <- watanabe add VerUP(2011)

        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_bz.hm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub

    Sub dataset_F_GMSAVE()

        'MsgBox "原始文字登録画面にデータをセットします"
        '20100705コード変更'.Textを追加
        form_no.w_font_name.Text = Trim(Right(temp_gm.font_name, 6))
        form_no.w_font_class1.Text = Trim(temp_gm.font_class1)
        form_no.w_font_class2.Text = Trim(temp_gm.font_class2)
        form_no.w_name1.Text = Trim(temp_gm.name1)
        form_no.w_name2.Text = Trim(temp_gm.name2)

        '----- .NET 移行 -----
        'form_no.w_high.Text = VB6.Format(temp_gm.high, "#####0.0000")
        'form_no.w_width.Text = VB6.Format(temp_gm.width, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_gm.ang, "#####0.0000")
        'form_no.w_moji_high.Text = VB6.Format(temp_gm.moji_high, "#####0.0000")
        'form_no.w_moji_shift.Text = VB6.Format(temp_gm.moji_shift, "#####0.0000")

        form_no.w_high.Text = temp_gm.high.ToString("#####0.0000")
        form_no.w_width.Text = temp_gm.width.ToString("#####0.0000")
        form_no.w_ang.Text = temp_gm.ang.ToString("#####0.0000")
        form_no.w_moji_high.Text = temp_gm.moji_high.ToString("#####0.0000")
        form_no.w_moji_shift.Text = temp_gm.moji_shift.ToString("#####0.0000")

        Select Case temp_gm.org_hor '現在Cに固定
            Case "C"
                form_no.w_org_hor.Text = "Center" '.Textを追加
            Case "L"
                form_no.w_org_hor.Text = "Left end"
            Case "R"
                form_no.w_org_hor.Text = "Right end"
            Case Else
                MsgBox("Horizontal origin position error." & temp_gm.org_hor)
                form_no.w_org_hor.Text = ""
        End Select
        Select Case temp_gm.org_ver '現在Bに固定
            Case "C"
                form_no.w_org_ver.Text = "Center" '.Textを追加
            Case "T"
                form_no.w_org_ver.Text = "Top"
            Case "B"
                form_no.w_org_ver.Text = "Lower end"
            Case Else
                MsgBox("Vertical origin position error." & temp_gm.org_ver)
                form_no.w_org_ver.Text = "" '.Textを追加
        End Select


        'フォント区分1 ｢縁取り｣ OR ｢縁＆ハッチング｣ チェック  (Brand CAD System Ver.3 UP)
        If Left(Trim(form_no.w_font_class1.Text), 1) <> "F" And Left(Trim(form_no.w_font_class1.Text), 1) <> "B" Then
            '----- .NET 移行 -----
            'form_no.w_hem_width.Text = VB6.Format(temp_gm.hem_width, "#0.00")
            form_no.w_hem_width.Text = temp_gm.hem_width.ToString("#0.00")

            '縁取り幅ロック
            form_no.w_hem_width.Enabled = False
            form_no.w_hem_width.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
        Else
            form_no.w_hem_width.Text = Trim(CStr(temp_gm.hem_width))
        End If


        '20100705コード変更 以下　.Textを追加
        '----- .NET 移行 -----
        'form_no.w_hatch_ang.Text = VB6.Format(temp_gm.hatch_ang, "#####0.0000")
        'form_no.w_hatch_width.Text = VB6.Format(temp_gm.hatch_width, "#####0.0000")
        'form_no.w_hatch_space.Text = VB6.Format(temp_gm.hatch_space, "#####0.0000")
        'form_no.w_hatch_x.Text = VB6.Format(temp_gm.hatch_x, "#####0.0000")
        'form_no.w_hatch_y.Text = VB6.Format(temp_gm.hatch_y, "#####0.0000")
        'form_no.w_base_r.Text = VB6.Format(temp_gm.base_r, "#####0.0000")

        form_no.w_hatch_ang.Text = temp_gm.hatch_ang.ToString("#####0.0000")
        form_no.w_hatch_width.Text = temp_gm.hatch_width.ToString("#####0.0000")
        form_no.w_hatch_space.Text = temp_gm.hatch_space.ToString("#####0.0000")
        form_no.w_hatch_x.Text = temp_gm.hatch_x.ToString("#####0.0000")
        form_no.w_hatch_y.Text = temp_gm.hatch_y.ToString("#####0.0000")
        form_no.w_base_r.Text = temp_gm.base_r.ToString("#####0.0000")

        form_no.w_old_font_name.Text = Trim(temp_gm.old_font_name)
        form_no.w_old_font_class.Text = Trim(temp_gm.old_font_class)
        form_no.w_old_name.Text = Trim(temp_gm.old_name)
        form_no.w_comment.Text = Trim(temp_gm.comment)
        form_no.w_dep_name.Text = Trim(Right(temp_gm.dep_name, 4))
        form_no.w_entry_name.Text = Trim(temp_gm.entry_name)
        form_no.w_entry_date.Text = Trim(temp_gm.entry_date)

        form_no.w_haiti_pic.Text = Trim(CStr(temp_gm.haiti_pic))

    End Sub

    'Sub dataset_CADREAD()
    '
    ' Dim i As Integer
    ' Dim ir As Integer
    ' Dim ic As Integer
    ' Dim w_ret As Integer
    '
    ' form_no.w_comment = Trim(temp_gz.comment)
    ' form_no.w_dep_name = Trim(temp_gz.dep_name)
    ' form_no.w_entry_name = Trim(temp_gz.entry_name)
    ' form_no.w_entry_date = Trim(temp_gz.entry_date)
    '
    ' '原始文字登録表のセット
    ' '列と行の総数を設定します。
    ' form_no.Grid1.Cols = 6
    ' form_no.Grid1.Rows = Int((temp_hm.gm_num - 1) / (form_no.Grid1.Cols - 1)) + 2
    '
    ' 'MsgBox "Grid1.Cols,Rows=" & form_no.Grid1.Cols & "," & form_no.Grid1.Rows
    '
    ' ' 列幅の設定
    ' form_no.Grid1.ColWidth(0) = (form_no.Grid1.width - 150) / 21 * 1
    ' form_no.Grid1.ColWidth(1) = (form_no.Grid1.width - 150) / 21 * 4
    ' form_no.Grid1.ColWidth(2) = (form_no.Grid1.width - 150) / 21 * 4
    ' form_no.Grid1.ColWidth(3) = (form_no.Grid1.width - 150) / 21 * 4
    ' form_no.Grid1.ColWidth(4) = (form_no.Grid1.width - 150) / 21 * 4
    ' form_no.Grid1.ColWidth(5) = (form_no.Grid1.width - 150) / 21 * 4
    ' For i = 1 To form_no.Grid1.Cols - 1
    '  w_ret = Set_Grid_Data(form_no.Grid1, Str$(i), 0, i)
    '  form_no.Grid1.FixedAlignment(i) = 2
    ' Next i
    ' For i = 1 To form_no.Grid1.Rows - 1
    '  w_ret = Set_Grid_Data(form_no.Grid1, Str$(i), i, 0)
    ' Next i
    '
    ' i = 0
    ' ir = 1
    ' ic = 1
    ' For i = 1 To temp_gz.gm_num
    '   If ic >= form_no.Grid1.Cols Then
    '     ic = 1
    '     ir = ir + 1
    '   End If
    '   w_ret = Set_Grid_Data(form_no.Grid1, temp_gz.gm_name(i), ir, ic)
    '   ic = ic + 1
    ' Next i
    '
    '
    'End Sub


    Sub dataset_F_HMSAVE()

        'Dim ir As Short
        'Dim ic As Short
        Dim i As Short
        'Dim w_ret As Short

        'MsgBox "編集文字登録画面にデータをセットします"
        '20100705コード変更'.Textを追加
        form_no.w_font_name.Text = Trim(Right(temp_hm.font_name, 6))
        form_no.w_no.Text = Trim(temp_hm.no) '.Textを追加
        form_no.w_spell.Text = Trim(temp_hm.spell) '.Textを追加

        '----- .NET 移行 -----
        'form_no.w_width.Text = VB6.Format(temp_hm.width, "#####0.0000")
        'form_no.w_high.Text = VB6.Format(temp_hm.high, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_hm.ang, "#####0.0000")
        '-----------------------------------------------------------------
        form_no.w_width.Text = temp_hm.width.ToString("#####0.0000")
        form_no.w_high.Text = temp_hm.high.ToString("#####0.0000")
        form_no.w_ang.Text = temp_hm.ang.ToString("#####0.0000")
        '----- .NET 移行 -----

        form_no.w_comment.Text = Trim(temp_hm.comment)
        form_no.w_dep_name.Text = Trim(Right(temp_hm.dep_name, 4))
        form_no.w_entry_name.Text = Trim(temp_hm.entry_name)
        form_no.w_entry_date.Text = Trim(temp_hm.entry_date)
        form_no.w_gm_num.Text = Trim(CStr(temp_hm.gm_num))
        form_no.w_haiti_pic.Text = Trim(CStr(temp_hm.haiti_pic))

        '原始文字登録表のセット

        '----- .NET 移行 (MSFlexGrid ⇒ DataGridViewに変更)-----

        ''列と行の総数を設定します。
        'form_no.MSFlexGrid1.Cols = 6
        'form_no.MSFlexGrid1.Rows = Int((temp_hm.gm_num - 1) / (form_no.MSFlexGrid1.Cols - 1)) + 2

        ''MsgBox "MSFlexGrid1.Cols,Rows=" & form_no.MSFlexGrid1.Cols & "," & form_no.MSFlexGrid1.Rows

        '' 列幅の設定
        ''20100705コード変更
        'form_no.MSFlexGrid1.set_ColWidth(0, (((form_no.MSFlexGrid1.width - 150) / 21 * 1) * 15))
        'form_no.MSFlexGrid1.set_ColWidth(1, (((form_no.MSFlexGrid1.width - 150) / 21 * 4) * 15))
        'form_no.MSFlexGrid1.set_ColWidth(2, (((form_no.MSFlexGrid1.width - 150) / 21 * 4) * 15))
        'form_no.MSFlexGrid1.set_ColWidth(3, (((form_no.MSFlexGrid1.width - 150) / 21 * 4) * 15))
        'form_no.MSFlexGrid1.set_ColWidth(4, (((form_no.MSFlexGrid1.width - 150) / 21 * 4) * 15))
        'form_no.MSFlexGrid1.set_ColWidth(5, (((form_no.MSFlexGrid1.width - 150) / 21 * 4) * 15))

        'For i = 1 To form_no.MSFlexGrid1.Cols - 1
        '    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
        '    'form_no.MSFlexGrid1.FixedAlignment(i) = 2
        '    form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        'Next i
        'For i = 1 To form_no.MSFlexGrid1.Rows - 1
        '    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        'Next i

        'i = 0
        'ir = 1
        'ic = 1
        'For i = 1 To temp_hm.gm_num
        '    If ic >= form_no.MSFlexGrid1.Cols Then
        '        ic = 1
        '        ir = ir + 1
        '    End If
        '    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hm.gm_name(i), ir, ic)
        '    ic = ic + 1
        'Next i

        '-----------------------------------------------------------------------------

        With form_no.DataGridViewList

            .ColumnCount = 5
            .TopLeftHeaderCell.Value = "NO"

            For i = 0 To (.ColumnCount - 1)
                .Columns(i).HeaderCell.Value = (i + 1).ToString()
                .Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                .Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .Columns(i).DefaultCellStyle.SelectionBackColor = SystemColors.Window
                .Columns(i).DefaultCellStyle.SelectionForeColor = SystemColors.WindowText
                .Columns(i).Width = 110
            Next

            Dim tmp(4) As String
            For i = 0 To 4
                tmp(i) = ""
            Next
            Dim cnt As Integer = 0

            For i = 1 To temp_hm.gm_num
                tmp(cnt) = temp_hm.gm_name(i)
                cnt += 1

                If cnt >= 5 Then
                    .Rows.Add(tmp(0), tmp(1), tmp(2), tmp(3), tmp(4))

                    For cnt = 0 To 4
                        tmp(cnt) = ""
                    Next cnt
                    cnt = 0
                End If
            Next i

            If cnt > 0 Then
                .Rows.Add(tmp(0), tmp(1), tmp(2), tmp(3), tmp(4))
            End If

            For i = 0 To .Rows.Count - 1
                .Rows(i).HeaderCell.Style.BackColor = SystemColors.Control
                .Rows(i).HeaderCell.Style.ForeColor = SystemColors.WindowText
                .Rows(i).HeaderCell.Style.SelectionBackColor = SystemColors.Control
                .Rows(i).HeaderCell.Style.SelectionForeColor = SystemColors.WindowText
            Next

        End With

        '----- .NET 移行 -----

    End Sub


    Sub dataset_F_GZSAVE()

        'Dim ir As Short'20100616移植削除
        'Dim ic As Short'20100616移植削除
        Dim i As Short
        Dim w_ret As Short
        Dim flg As Short

        ' -> watanabe edit 2013.06.03
        'Dim err_flg(100) As String
        Dim err_flg(200) As String
        ' <- watanabe edit 2013.06.03

        Dim w_str(100) As String
        Dim DBTableNameGm As Object

        'MsgBox "刻印図面登録画面にデータをセットします"

        DBTableNameGm = DBName & "..gm_kanri"

        '20100705コード追加　.Textを追加
        form_no.w_id.Text = Trim(temp_gz.id)
        form_no.w_no1.Text = Trim(temp_gz.no1)
        form_no.w_no2.Text = Trim(temp_gz.no2)
        form_no.w_comment.Text = Trim(temp_gz.comment)
        form_no.w_dep_name.Text = Trim(temp_gz.dep_name)
        form_no.w_entry_name.Text = Trim(temp_gz.entry_name)
        form_no.w_entry_date.Text = Trim(temp_gz.entry_date)
        form_no.w_gm_num.Text = Str(temp_gz.gm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Redraw = False '----- 12/12 1997 yamamoto add -----

        form_no.MSFlexGrid1.Cols = 5
        form_no.MSFlexGrid1.Rows = Int((temp_gz.gm_num - 1) / 2) + 2


        ' -> watanabe add VerUP(2011)   フォームロード時だとグリッドの行が更新されない
        ' 行高さの設定
        For i = 0 To form_no.MSFlexGrid1.Rows - 1
            form_no.MSFlexGrid1.set_RowHeight(i, 300)
        Next i

        ' 列幅の設定
        form_no.MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 1)
        form_no.MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 2)
        form_no.MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 6)
        form_no.MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 2)
        form_no.MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 6)
        For i = 0 To 4
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i

        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "NO", 0, 0)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "error", 0, 1)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "Primitive character code", 0, 2)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "error", 0, 3)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "Primitive character code", 0, 4)
        ' <- watanabe add VerUP(2011)

        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        For i = 1 To temp_gz.gm_num Step 2
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_gz.gm_name(i), Int((i - 1) / 2) + 1, 2)
            If (i + 1) > temp_gz.gm_num Then Exit For
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_gz.gm_name(i + 1), Int((i - 1) / 2) + 1, 4)
        Next i

        form_no.MSFlexGrid1.Redraw = True '----- 12/12 1997 yamamoto add -----


        '原始文字データチェック(既に他の刻印図面に使用されていればエラー)
        flg = 0
        If open_mode <> "Revision number" Then
            For i = 1 To Val(Trim(form_no.w_gm_num.Text))
                err_flg(i) = ""
                w_ret = exist_gm_gz(DBTableNameGm, temp_gz.gm_name(i), temp_gz.no1, temp_gz.no2)
                If w_ret = -1 Then
                    '            MsgBox "SQLｴﾗｰです", vbCritical, "SQLｴﾗｰ"  'yamamoto
                    GoTo error_section
                ElseIf w_ret = 1 Then
                    '            MsgBox "原始文字コード[" & temp_gz.gm_name(i) & "]は既に他の刻印図面で使用されていますので登録出来ません", vbCritical, "刻印図面新規登録ｴﾗｰ"  'yamamoto
                    err_flg(i) = "100"
                    flg = 1
                ElseIf w_ret = 2 Then
                    err_flg(i) = "200"
                    flg = 1
                ElseIf w_ret = 3 Then
                End If
            Next i
            For i = 1 To 63
                w_str(i + 9) = "'" & temp_gz.gm_name(i) & " '"
            Next i
            If flg = 1 Then

                form_no.MSFlexGrid1.Redraw = False '----- 12/12 1997 yamamoto add -----

                For i = 1 To temp_gz.gm_num Step 2
                    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i), ((i - 1) / 2) + 1, 1)
                    If (i + 1) > temp_gz.gm_num Then Exit For
                    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i + 1), ((i - 1) / 2) + 1, 3)
                Next i
                form_no.MSFlexGrid1.Redraw = True '----- 12/12 1997 yamamoto add -----

                GoTo error_section
            End If
        End If

        Exit Sub

error_section:

        If open_mode <> "modify" Then '----- 12/12 1997 yamamoto add -----
            '画面ロック
            form_no.Command1.Enabled = False
            form_no.Command2.Enabled = False
            form_no.Command4.Enabled = False
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_comment.Enabled = False
            form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_dep_name.Enabled = False
            form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If

    End Sub

    Sub dataset_F_GZDELE()
        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        ' -> watanabe edit 2013.06.03
        'Dim err_flg(100) As String
        Dim err_flg(200) As String
        ' <- watanabe edit 2013.06.03

        Dim w_str(100) As Object

        'MsgBox "刻印図面削除画面にデータをセットします"

        '20100705コード追加'.Textを追加
        form_no.w_id.Text = temp_gz.id
        form_no.w_no1.Text = temp_gz.no1
        form_no.w_no2.Text = temp_gz.no2
        form_no.w_comment.Text = temp_gz.comment
        form_no.w_dep_name.Text = temp_gz.dep_name
        form_no.w_entry_name.Text = temp_gz.entry_name
        form_no.w_entry_date.Text = temp_gz.entry_date
        form_no.w_gm_num.Text = CStr(temp_gz.gm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_gz.gm_num - 1) / 5) + 2

        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_gz.gm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_gz.gm_name(i), ir, ic)
            ic = ic + 1
        Next i

        Exit Sub

    End Sub


    Sub dataset_F_BZSAVE()

        '--------------<ブランド登録画面にデータをセットします >--------------

        'Dim ir As Short '20100616移植削除
        'Dim ic As Short '20100616移植削除
        Dim i As Short
        Dim kkno As Short
        Dim kk As New VB6.FixedLengthString(1)
        Dim w_ret As Short

        'MsgBox "dataset_F_BZSAVE:no1=[" & temp_bz.no1 & "],no2=[" & temp_bz.no2 & "]"

        'form_no.w_id = temp_bz.id
        '20100705コード変更'.Textを追加
        form_no.w_no1.Text = Trim(temp_bz.no1)
        form_no.w_no2.Text = Trim(temp_bz.no2)
        form_no.w_kanri_no.Text = Trim(temp_bz.kanri_no)
        form_no.w_comment.Text = Trim(temp_bz.comment)
        form_no.w_dep_name.Text = Trim(temp_bz.dep_name)
        form_no.w_entry_name.Text = Trim(temp_bz.entry_name)
        form_no.w_entry_date.Text = Trim(temp_bz.entry_date)
        form_no.w_syurui.Text = Trim(temp_bz.syurui)
        form_no.w_pattern.Text = Trim(temp_bz.pattern)
        form_no.w_syubetu.Text = Trim(temp_bz.syubetu)
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size4.Text = Trim(temp_bz.size4)
        If Trim(temp_bz.size5) = "" Then
            form_no.w_size5.Text = "R" '.Textを追加
        Else
            form_no.w_size5.Text = Trim(temp_bz.size5) '.Textを追加
        End If
        form_no.w_size6.Text = Trim(temp_bz.size6) '.Textを追加
        form_no.w_size7.Text = Trim(temp_bz.size7)
        form_no.w_size8.Text = Trim(temp_bz.size8)
        form_no.w_size.Text = Trim(temp_bz.Size)
        form_no.w_size_code.Text = Trim(temp_bz.size_code)
        Call dataset_plant(form_no.w_plant, temp_bz.plant, "save")
        form_no.w_plant_code.Text = Trim(temp_bz.plant_code)

        form_no.w_kikaku.Text = Trim(temp_bz.kikaku)

        For i = 1 To Len(temp_bz.kikaku)
            kkno = 0
            kk.Value = Mid(temp_bz.kikaku, i, 1)

            If kk.Value = "J" Then
                kkno = 1
            ElseIf kk.Value = "E" Then
                kkno = 2
            ElseIf kk.Value = "F" Then
                kkno = 3
            ElseIf kk.Value = "I" Then
                kkno = 4
            ElseIf kk.Value = "G" Then
                kkno = 5
            ElseIf kk.Value = "A" Then
                kkno = 6
            End If

            If i = 1 Then
                'form_no.w_kikaku1.ListIndex = kkno
                form_no.w_kikaku1.Text = form_no.w_kikaku1.GetItemText(form_no.w_kikaku1.Items(kkno)) '20100624コード変更
            ElseIf i = 2 Then
                'form_no.w_kikaku2.ListIndex = kkno
                form_no.w_kikaku2.Text = form_no.w_kikaku2.GetItemText(form_no.w_kikaku2.Items(kkno))
            ElseIf i = 3 Then
                'form_no.w_kikaku3.ListIndex = kkno
                form_no.w_kikaku3.Text = form_no.w_kikaku3.GetItemText(form_no.w_kikaku3.Items(kkno))
            ElseIf i = 4 Then
                'form_no.w_kikaku4.ListIndex = kkno
                form_no.w_kikaku4.Text = form_no.w_kikaku4.GetItemText(form_no.w_kikaku4.Items(kkno))
            ElseIf i = 5 Then
                'form_no.w_kikaku5.ListIndex = kkno
                form_no.w_kikaku5.Text = form_no.w_kikaku5.GetItemText(form_no.w_kikaku5.Items(kkno))
            ElseIf i = 6 Then
                'form_no.w_kikaku6.ListIndex = kkno
                form_no.w_kikaku6.Text = form_no.w_kikaku6.GetItemText(form_no.w_kikaku6.Items(kkno))
            End If

        Next i

        If (temp_bz.tos_moyou = 0) Or (temp_bz.tos_moyou = 1) Then
            'form_no.w_tos_moyou.ListIndex = temp_bz.tos_moyou
            'form_no.w_tos_moyou.Text = form_no.w_tos_moyou.GetItemText(form_no.w_tos_moyou.Items(temp_bz.tos_moyou)) '20100624コード変更
            form_no.w_tos_moyou.Text = VB6.GetItemString(form_no.w_tos_moyou, temp_bz.tos_moyou)
        End If
        If (temp_bz.side_moyou = 0) Or (temp_bz.side_moyou = 1) Then
            '    form_no.w_side_moyou.ListIndex = Str$(temp_bz.side_moyou)
            'form_no.w_side_moyou.ListIndex = temp_bz.side_moyou
            'form_no.w_side_moyou.Text = form_no.w_side_moyou.GetItemText(form_no.w_side_moyou.Items(temp_bz.side_moyou))
            form_no.w_side_moyou.Text = VB6.GetItemString(form_no.w_side_moyou, temp_bz.side_moyou)
        End If
        If (temp_bz.side_kenti = 0) Or (temp_bz.side_kenti = 1) Then
            'form_no.w_side_kenti.ListIndex = temp_bz.side_kenti
            'form_no.w_side_kenti.Text = form_no.w_side_kenti.GetItemText(form_no.w_side_kenti.Items(temp_bz.side_kenti))
            form_no.w_side_kenti.Text = VB6.GetItemString(form_no.w_side_kenti, temp_bz.side_kenti)
        End If
        If (temp_bz.peak_mark = 0) Or (temp_bz.peak_mark = 1) Then
            'form_no.w_peak_mark.ListIndex = temp_bz.peak_mark
            'form_no.w_peak_mark.Text = form_no.w_peak_mark.GetItemText(form_no.w_peak_mark.Items(temp_bz.peak_mark))
            form_no.w_peak_mark.Text = VB6.GetItemString(form_no.w_peak_mark, temp_bz.peak_mark)
        End If
        If (temp_bz.nasiji = 0) Or (temp_bz.nasiji = 1) Or (temp_bz.nasiji = 2) Or (temp_bz.nasiji = 3) Then
            'form_no.w_nasiji.ListIndex = temp_bz.nasiji
            'form_no.w_nasiji.Text = form_no.w_nasiji.GetItemText(form_no.w_nasiji.Items(temp_bz.nasiji))
            form_no.w_nasiji.Text = VB6.GetItemString(form_no.w_nasiji, temp_bz.nasiji)
        End If

        form_no.w_hm_num.Text = Str(temp_bz.hm_num)
        If temp_bz.hm_num > 0 Then
            '列と行の総数を設定します。
            form_no.MSFlexGrid1.Cols = 5
            form_no.MSFlexGrid1.Rows = Int((temp_bz.hm_num - 1) / (form_no.MSFlexGrid1.Cols - 1)) + 2

            For i = 1 To form_no.MSFlexGrid1.Cols - 1
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
                'form_no.MSFlexGrid1.FixedAlignment(i) = 2
                form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
            Next i
            For i = 1 To form_no.MSFlexGrid1.Rows - 1
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
            Next i

            i = 0
            For i = 1 To temp_bz.hm_num Step 4
                If i > temp_bz.hm_num Then Exit For
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i), Int(i / 4) + 1, 1)
                If (i + 1) > temp_bz.hm_num Then Exit For
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i + 1), Int(i / 4) + 1, 2)
                If (i + 1) > temp_bz.hm_num Then Exit For
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i + 2), Int(i / 4) + 1, 3)
                If (i + 1) > temp_bz.hm_num Then Exit For
                w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_bz.hm_name(i + 3), Int(i / 4) + 1, 4)
            Next i
        End If

    End Sub


    Sub dataset_F_GZLOOK()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        'MsgBox "刻印図面内容確認画面にデータをセットします"
        '20100705コード追加　.Textを追加
        form_no.w_flag_delete.Text = CStr(temp_gz.flag_delete)
        form_no.w_id.Text = temp_gz.id
        form_no.w_no1.Text = temp_gz.no1
        form_no.w_no2.Text = temp_gz.no2
        form_no.w_comment.Text = temp_gz.comment
        form_no.w_dep_name.Text = temp_gz.dep_name
        form_no.w_entry_name.Text = temp_gz.entry_name
        form_no.w_entry_date.Text = temp_gz.entry_date
        'form_no.w_gm_num = temp_gz.gm_num
        form_no.w_gm_num.Text = CStr(temp_gz.gm_num) '20100705コード変更


        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_gz.gm_num - 1) / 5) + 2

        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_gz.gm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_gz.gm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub


    Sub dataset_F_HZDELE()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        ' -> watanabe edit 2013.06.03
        'Dim err_flg(100) As String
        Dim err_flg(200) As String
        ' <- watanabe edit 2013.06.03

        Dim w_str(100) As Object

        'MsgBox "編集文字図面削除画面にデータをセットします"

        '20100705コード変更 '.Textを追加
        form_no.w_id.Text = temp_hz.id
        form_no.w_no1.Text = temp_hz.no1
        form_no.w_no2.Text = temp_hz.no2
        form_no.w_comment.Text = temp_hz.comment
        form_no.w_dep_name.Text = temp_hz.dep_name
        form_no.w_entry_name.Text = temp_hz.entry_name
        form_no.w_entry_date.Text = temp_hz.entry_date
        form_no.w_hm_num.Text = CStr(temp_hz.hm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_hz.hm_num - 1) / 5) + 2

        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        ir = 1
        ic = 1
        For i = 1 To temp_hz.hm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hz.hm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub

    Sub dataset_F_HZLOOK()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        'MsgBox "編集文字図面内容確認画面にデータをセットします"
        '2010コード変更'.Textを追加
        form_no.w_flag_delete.Text = CStr(temp_hz.flag_delete)
        form_no.w_id.Text = temp_hz.id
        form_no.w_no1.Text = temp_hz.no1
        form_no.w_no2.Text = temp_hz.no2
        form_no.w_comment.Text = temp_hz.comment
        form_no.w_dep_name.Text = temp_hz.dep_name
        form_no.w_entry_name.Text = temp_hz.entry_name
        form_no.w_entry_date.Text = temp_hz.entry_date
        form_no.w_hm_num.Text = CStr(temp_hz.hm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        form_no.MSFlexGrid1.Rows = Int((temp_hz.hm_num - 1) / 5) + 2
        

        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        ir = 1
        ic = 1
        For i = 1 To temp_hz.hm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hz.hm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub


    Sub dataset_F_HMLOOK()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short

        'MsgBox "編集文字内容確認画面にデータをセットします"
        '2010コード変更'.Textを追加
        form_no.w_flag_delete.Text = CStr(temp_hm.flag_delete)
        form_no.w_id.Text = Trim(temp_hm.id)
        ' form_no.w_no = Trim(temp_hm.no)
        form_no.w_spell.Text = Trim(temp_hm.spell)

        '----- .NET 移行 -----
        'form_no.w_width.Text = VB6.Format(temp_hm.width, "#####0.0000")
        'form_no.w_high.Text = VB6.Format(temp_hm.high, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_hm.ang, "#####0.0000")

        form_no.w_width.Text = temp_hm.width.ToString("#####0.0000")
        form_no.w_high.Text = temp_hm.high.ToString("#####0.0000")
        form_no.w_ang.Text = temp_hm.ang.ToString("#####0.0000")

        form_no.w_comment.Text = Trim(temp_hm.comment)
        form_no.w_dep_name.Text = Trim(Right(temp_hm.dep_name, 4))
        form_no.w_entry_name.Text = Trim(temp_hm.entry_name)
        form_no.w_entry_date.Text = Trim(temp_hm.entry_date)
        form_no.w_gm_num.Text = Trim(CStr(temp_hm.gm_num))
        form_no.w_haiti_pic.Text = Trim(CStr(temp_hm.haiti_pic))
        form_no.w_haiti_sitei.Text = Trim(CStr(temp_hm.haiti_sitei))

        form_no.w_hz_id.Text = Trim(temp_hm.hz_id) '.Textを追加
        form_no.w_hz_no1.Text = Trim(temp_hm.hz_no1)
        form_no.w_hz_no2.Text = Trim(temp_hm.hz_no2)


        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        ' form_no.MSFlexGrid1.Rows = temp_hm.gm_num / (form_no.MSFlexGrid1.Cols - 1) + 2
        form_no.MSFlexGrid1.Rows = Int((temp_hm.gm_num - 1) / (form_no.MSFlexGrid1.Cols - 1)) + 2

        '20100706コード変更------------------->以下のコードはフォームのLoad時に設定しています
        ' 列幅の設定
        'form_no.MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 1)
        'form_no.MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 4)
        'form_no.MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 4)
        'form_no.MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 4)
        'form_no.MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 4)
        'form_no.MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.width) - 150) / 21 * 4)
        '-------------------------------------

        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_hm.gm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hm.gm_name(i), ir, ic)
            ic = ic + 1
        Next i

    End Sub



    Sub dataset_F_GMDELE()

        '20100705コード変更'.Textを追加

        '----- .NET 移行 -----
        'form_no.w_high.Text = VB6.Format(temp_gm.high, "#####0.0000")
        'form_no.w_width.Text = VB6.Format(temp_gm.width, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_gm.ang, "#####0.0000")
        'form_no.w_moji_high.Text = VB6.Format(temp_gm.moji_high, "#####0.0000")
        'form_no.w_moji_shift.Text = VB6.Format(temp_gm.moji_shift, "#####0.0000")

        form_no.w_high.Text = temp_gm.high.ToString("#####0.0000")
        form_no.w_width.Text = temp_gm.width.ToString("#####0.0000")
        form_no.w_ang.Text = temp_gm.ang.ToString("#####0.0000")
        form_no.w_moji_high.Text = temp_gm.moji_high.ToString("#####0.0000")
        form_no.w_moji_shift.Text = temp_gm.moji_shift.ToString("#####0.0000")

        Select Case temp_gm.org_hor 'Cに固定
            Case "C"
                form_no.w_org_hor.Text = "Center"
            Case "L"
                form_no.w_org_hor.Text = "Left end"
            Case "R"
                form_no.w_org_hor.Text = "Right end"
            Case Else
                'Debug.Print "水平原点位置ｴﾗｰ"
                form_no.w_org_hor.Text = ""
        End Select
        Select Case temp_gm.org_ver 'Bに固定
            Case "C"
                form_no.w_org_ver.Text = "Center"
            Case "T"
                form_no.w_org_ver.Text = "Top"
            Case "B"
                form_no.w_org_ver.Text = "Lower end"
            Case Else
                'Debug.Print "垂直原点位置ｴﾗｰ"
                form_no.w_org_ver.Text = ""
        End Select

        '20100705コード変更'.Textを追加

        '----- .NET 移行 -----
        'form_no.w_hem_width.Text = VB6.Format(temp_gm.hem_width, "#0.00")
        'form_no.w_hatch_ang.Text = VB6.Format(temp_gm.hatch_ang, "#####0.0000")
        'form_no.w_hatch_width.Text = VB6.Format(temp_gm.hatch_width, "#####0.0000")
        'form_no.w_hatch_space.Text = VB6.Format(temp_gm.hatch_space, "#####0.0000")
        'form_no.w_hatch_x.Text = VB6.Format(temp_gm.hatch_x, "#####0.0000")
        'form_no.w_hatch_y.Text = VB6.Format(temp_gm.hatch_y, "#####0.0000")
        'form_no.w_base_r.Text = VB6.Format(temp_gm.base_r, "#####0.0000")

        form_no.w_hem_width.Text = temp_gm.hem_width.ToString("#0.00")
        form_no.w_hatch_ang.Text = temp_gm.hatch_ang.ToString("#####0.0000")
        form_no.w_hatch_width.Text = temp_gm.hatch_width.ToString("#####0.0000")
        form_no.w_hatch_space.Text = temp_gm.hatch_space.ToString("#####0.0000")
        form_no.w_hatch_x.Text = temp_gm.hatch_x.ToString("#####0.0000")
        form_no.w_hatch_y.Text = temp_gm.hatch_y.ToString("#####0.0000")
        form_no.w_base_r.Text = temp_gm.base_r.ToString("#####0.0000")

        form_no.w_old_font_name.Text = temp_gm.old_font_name
        form_no.w_old_font_class.Text = temp_gm.old_font_class
        form_no.w_old_name.Text = temp_gm.old_name
        form_no.w_haiti_pic.Text = temp_gm.haiti_pic
        form_no.w_comment.Text = temp_gm.comment
        form_no.w_dep_name.Text = Right(temp_gm.dep_name, 6)
        form_no.w_entry_name.Text = temp_gm.entry_name
        form_no.w_entry_date.Text = temp_gm.entry_date

    End Sub


    Sub dataset_F_HMDELE()

        Dim ir As Short
        Dim ic As Short
        Dim i As Short
        Dim w_ret As Short
        Dim w_file As String
        Dim TiffFile As String
        'MsgBox "編集文字削除画面にデータをセットします"

        On Error Resume Next ' エラーのトラップを留保します。
        Err.Clear()

        '20100705コード変更'.Textを追加

        '----- .NET 移行 -----
        'form_no.w_width.Text = VB6.Format(temp_hm.width, "#####0.0000")
        'form_no.w_high.Text = VB6.Format(temp_hm.high, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_hm.ang, "#####0.0000")

        form_no.w_width.Text = temp_hm.width.ToString("#####0.0000")
        form_no.w_high.Text = temp_hm.high.ToString("#####0.0000")
        form_no.w_ang.Text = temp_hm.ang.ToString("#####0.0000")

        form_no.w_spell.Text = temp_hm.spell
        form_no.w_haiti_pic.Text = temp_hm.haiti_pic
        form_no.w_comment.Text = temp_hm.comment
        form_no.w_dep_name.Text = Right(temp_hm.dep_name, 6)
        form_no.w_entry_name.Text = temp_hm.entry_name
        form_no.w_entry_date.Text = temp_hm.entry_date
        form_no.w_gm_num.Text = CStr(temp_hm.gm_num)

        'Brand Ver.5 TIFF->BMP 変更 start
        ' TiffFile = TIFFDir & Trim(form_no.w_mojicd) & ".tif"
        ' 'Tiffﾌｧｲﾙ表示
        ' w_file = Dir(TiffFile)
        ' If w_file <> "" Then
        '     form_no.ImgThumbnail1.Image = TiffFile
        '     form_no.ImgThumbnail1.ThumbWidth = 500
        '     form_no.ImgThumbnail1.ThumbHeight = 200
        ' Else
        '     MsgBox "TIFFﾌｧｲﾙが見つかりません", vbCritical
        ' End If
        TiffFile = TIFFDir & Trim(form_no.w_mojicd) & ".bmp"
        'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 457 '20100701コード変更
            form_no.ImgThumbnail1.Height = 193 '20100701コード変更
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
        End If
        'Brand Ver.5 TIFF->BMP 変更 end


        '原始文字登録表のセット
        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 6
        ' form_no.MSFlexGrid1.Rows = temp_hm.gm_num / (form_no.MSFlexGrid1.Cols - 1) + 2
        form_no.MSFlexGrid1.Rows = Int((temp_hm.gm_num - 1) / (form_no.MSFlexGrid1.Cols - 1)) + 2

        'MsgBox "MSFlexGrid1.Cols,Rows=" & form_no.MSFlexGrid1.Cols & "," & form_no.MSFlexGrid1.Rows

        ' 列幅の設定
        '20100705コード変更
        form_no.MSFlexGrid1.set_ColWidth(0, (form_no.MSFlexGrid1.width - 150) / 21 * 1)
        form_no.MSFlexGrid1.set_ColWidth(1, (form_no.MSFlexGrid1.width - 150) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(2, (form_no.MSFlexGrid1.width - 150) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(3, (form_no.MSFlexGrid1.width - 150) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(4, (form_no.MSFlexGrid1.width - 150) / 21 * 4)
        form_no.MSFlexGrid1.set_ColWidth(5, (form_no.MSFlexGrid1.width - 150) / 21 * 4)
        For i = 1 To form_no.MSFlexGrid1.Cols - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), 0, i)
            'form_no.MSFlexGrid1.FixedAlignment(i) = 2
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i
        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i

        i = 0
        ir = 1
        ic = 1
        For i = 1 To temp_hm.gm_num
            If ic >= form_no.MSFlexGrid1.Cols Then
                ic = 1
                ir = ir + 1
            End If
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hm.gm_name(i), ir, ic)
            ic = ic + 1
        Next i

        end_sql()
    End Sub

    Sub dataset_F_GMLOOK()

        '20100705コード変更'.Textを追加
        form_no.w_flag_delete.Text = CStr(temp_gm.flag_delete)
        form_no.w_id.Text = temp_gm.id

        '----- .NET 移行 -----
        'form_no.w_high.Text = VB6.Format(temp_gm.high, "#####0.0000")
        'form_no.w_width.Text = VB6.Format(temp_gm.width, "#####0.0000")
        'form_no.w_ang.Text = VB6.Format(temp_gm.ang, "#####0.0000")
        'form_no.w_moji_high.Text = VB6.Format(temp_gm.moji_high, "#####0.0000")
        'form_no.w_moji_shift.Text = VB6.Format(temp_gm.moji_shift, "#####0.0000")

        form_no.w_high.Text = temp_gm.high.ToString("#####0.0000")
        form_no.w_width.Text = temp_gm.width.ToString("#####0.0000")
        form_no.w_ang.Text = temp_gm.ang.ToString("#####0.0000")
        form_no.w_moji_high.Text = temp_gm.moji_high.ToString("#####0.0000")
        form_no.w_moji_shift.Text = temp_gm.moji_shift.ToString("#####0.0000")

        Select Case temp_gm.org_hor 'Cに固定
            Case "C"
                form_no.w_org_hor.Text = "Center"
            Case "L"
                form_no.w_org_hor.Text = "Left end"
            Case "R"
                form_no.w_org_hor.Text = "Right end"
            Case Else
                'Debug.Print "水平原点位置ｴﾗｰ"
                form_no.w_org_hor.Text = ""
        End Select
        Select Case temp_gm.org_ver 'Bに固定
            Case "C"
                form_no.w_org_ver.Text = "Center"
            Case "T"
                form_no.w_org_ver.Text = "Top"
            Case "B"
                form_no.w_org_ver.Text = "Lower end"
            Case Else
                'Debug.Print "垂直原点位置ｴﾗｰ"
                form_no.w_org_ver.Text = ""
        End Select

        '以下　20100705　コード変更　＞　.Textを追加

        '----- .NET 移行 -----
        'form_no.w_org_x.Text = VB6.Format(temp_gm.org_x, "#####0.00")
        'form_no.w_org_y.Text = VB6.Format(temp_gm.org_y, "#####0.00")

        'form_no.w_left_bottom_x.Text = VB6.Format(temp_gm.left_bottom_x, "#####0.0000")
        'form_no.w_left_bottom_y.Text = VB6.Format(temp_gm.left_bottom_y, "#####0.0000")
        'form_no.w_right_bottom_x.Text = VB6.Format(temp_gm.right_bottom_x, "#####0.0000")
        'form_no.w_right_bottom_y.Text = VB6.Format(temp_gm.right_bottom_y, "#####0.0000")
        'form_no.w_right_top_x.Text = VB6.Format(temp_gm.right_top_x, "#####0.0000")
        'form_no.w_right_top_y.Text = VB6.Format(temp_gm.right_top_y, "#####0.0000")
        'form_no.w_left_top_x.Text = VB6.Format(temp_gm.left_top_x, "#####0.0000")
        'form_no.w_left_top_y.Text = VB6.Format(temp_gm.left_top_y, "#####0.0000")

        'form_no.w_hem_width.Text = VB6.Format(temp_gm.hem_width, "#0.00")
        'form_no.w_hatch_ang.Text = VB6.Format(temp_gm.hatch_ang, "#####0.0000")
        'form_no.w_hatch_width.Text = VB6.Format(temp_gm.hatch_width, "#####0.0000")
        'form_no.w_hatch_space.Text = VB6.Format(temp_gm.hatch_space, "#####0.0000")
        'form_no.w_hatch_x.Text = VB6.Format(temp_gm.hatch_x, "#####0.0000")
        'form_no.w_hatch_y.Text = VB6.Format(temp_gm.hatch_y, "#####0.0000")
        'form_no.w_base_r.Text = VB6.Format(temp_gm.base_r, "#####0.0000")

        form_no.w_org_x.Text = temp_gm.org_x.ToString("#####0.00")
        form_no.w_org_y.Text = temp_gm.org_y.ToString("#####0.00")

        form_no.w_left_bottom_x.Text = temp_gm.left_bottom_x.ToString("#####0.0000")
        form_no.w_left_bottom_y.Text = temp_gm.left_bottom_y.ToString("#####0.0000")
        form_no.w_right_bottom_x.Text = temp_gm.right_bottom_x.ToString("#####0.0000")
        form_no.w_right_bottom_y.Text = temp_gm.right_bottom_y.ToString("#####0.0000")
        form_no.w_right_top_x.Text = temp_gm.right_top_x.ToString("#####0.0000")
        form_no.w_right_top_y.Text = temp_gm.right_top_y.ToString("#####0.0000")
        form_no.w_left_top_x.Text = temp_gm.left_top_x.ToString("#####0.0000")
        form_no.w_left_top_y.Text = temp_gm.left_top_y.ToString("#####0.0000")

        form_no.w_hem_width.Text = temp_gm.hem_width.ToString("#0.00")
        form_no.w_hatch_ang.Text = temp_gm.hatch_ang.ToString("#####0.0000")
        form_no.w_hatch_width.Text = temp_gm.hatch_width.ToString("#####0.0000")
        form_no.w_hatch_space.Text = temp_gm.hatch_space.ToString("#####0.0000")
        form_no.w_hatch_x.Text = temp_gm.hatch_x.ToString("#####0.0000")
        form_no.w_hatch_y.Text = temp_gm.hatch_y.ToString("#####0.0000")
        form_no.w_base_r.Text = temp_gm.base_r.ToString("#####0.0000")

        form_no.w_old_font_name.Text = temp_gm.old_font_name
        form_no.w_old_font_class.Text = temp_gm.old_font_class
        form_no.w_old_name.Text = temp_gm.old_name
        form_no.w_haiti_pic.Text = temp_gm.haiti_pic
        form_no.w_comment.Text = temp_gm.comment
        form_no.w_dep_name.Text = Right(temp_gm.dep_name, 4)
        form_no.w_entry_name.Text = temp_gm.entry_name
        form_no.w_entry_date.Text = temp_gm.entry_date

    End Sub



    Sub dataset_F_HZSAVE()

        'Dim ir As Short '20100616移植削除
        'Dim ic As Short '20100616移植削除
        Dim i As Short
        Dim w_ret As Short
        Dim flg As Short

        ' -> watanabe edit 2013.06.03
        'Dim err_flg(100) As String
        Dim err_flg(200) As String
        ' <- watanabe edit 2013.06.03

        Dim w_str(100) As Object
        Dim DBTableNameHm As Object

        'MsgBox "編集文字図面登録画面にデータをセットします"

        ' Brand Ver.3 変更
        ' DBTableNameHm = DBName & "..hm_kanri"
        '20100705コード変更'.Textを追加
        DBTableNameHm = DBName & "..hm_kanri1"
        form_no.w_no1.Text = Trim(temp_hz.no1)
        form_no.w_no2.Text = Trim(temp_hz.no2)
        form_no.w_comment.Text = Trim(temp_hz.comment)
        form_no.w_dep_name.Text = Trim(temp_hz.dep_name)
        form_no.w_entry_name.Text = Trim(temp_hz.entry_name)
        form_no.w_entry_date.Text = Trim(temp_hz.entry_date)
        form_no.w_hm_num.Text = CStr(temp_hz.hm_num)

        '列と行の総数を設定します。
        form_no.MSFlexGrid1.Cols = 5
        form_no.MSFlexGrid1.Rows = Int((temp_hz.hm_num - 1) / 2) + 2


        ' -> watanabe add VerUP(2011)
        For i = 0 To form_no.MSFlexGrid1.Rows - 1
            form_no.MSFlexGrid1.set_RowHeight(i, 300)
        Next i

        form_no.MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 1)
        form_no.MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 2)
        form_no.MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 6)
        form_no.MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 2)
        form_no.MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(form_no.MSFlexGrid1.Width) - 100) / 18 * 6)

        For i = 0 To 4
            form_no.MSFlexGrid1.set_FixedAlignment(i, 2)
        Next i

        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "NO", 0, 0)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "error", 0, 1)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "Editing characters code", 0, 2)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "error", 0, 3)
        w_ret = Set_Grid_Data(form_no.MSFlexGrid1, "Editing characters code", 0, 4)
        ' <- watanabe add VerUP(2011)


        For i = 1 To form_no.MSFlexGrid1.Rows - 1
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, Str(i), i, 0)
        Next i
        For i = 1 To temp_hz.hm_num Step 2
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hz.hm_name(i), Int((i - 1) / 2) + 1, 2)
            If (i + 1) > temp_hz.hm_num Then Exit For
            w_ret = Set_Grid_Data(form_no.MSFlexGrid1, temp_hz.hm_name(i + 1), Int((i - 1) / 2) + 1, 4)
        Next i

        '編集文字データチェック(既に他の編集文字図面に使用されていればエラー)
        flg = 0
        If open_mode <> "Revision number" Then
            For i = 1 To Val(Trim(form_no.w_hm_num.Text))
                err_flg(i) = ""
                w_ret = exist_hm_hz(DBTableNameHm, temp_hz.hm_name(i), temp_hz.no1, temp_hz.no2)


                If w_ret = -1 Then
                    MsgBox("SQL error", MsgBoxStyle.Critical, "SQL error")
                    GoTo error_section
                ElseIf w_ret = 1 Then
                    '            MsgBox "編集文字コード[" & temp_hz.hm_name(i) & "]は既に他の編集文字図面で使用されていますので登録出来ません", vbCritical, "編集文字図面新規登録ｴﾗｰ"
                    err_flg(i) = "100"
                    flg = 1
                ElseIf w_ret = 2 Then
                    '            MsgBox "編集文字コード[" & temp_hz.hm_name(i) & "]が存在しません", vbCritical, "編集文字図面新規登録ｴﾗｰ"
                    err_flg(i) = "200"
                    flg = 1
                ElseIf w_ret = 3 Then
                    '            MsgBox "編集文字コード[" & temp_hz.hm_name(i) & "]は自分自身が登録済みです", , "debug"
                End If
            Next i
            For i = 1 To 63
                w_str(i + 9) = "'" & temp_hz.hm_name(i) & " '"
            Next i
            If flg = 1 Then
                For i = 1 To temp_hz.hm_num Step 2
                    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i), ((i - 1) / 2) + 1, 1)
                    If (i + 1) > temp_hz.hm_num Then Exit For
                    w_ret = Set_Grid_Data(form_no.MSFlexGrid1, err_flg(i + 1), ((i - 1) / 2) + 1, 3)
                Next i
                GoTo error_section
            End If
        End If

        Exit Sub

error_section:
        '画面ロック
        If open_mode <> "modify" Then '----- 12/12 1997 yamamoto add -----
            form_no.Command1.Enabled = False
            form_no.Command2.Enabled = False
            form_no.Command4.Enabled = False
            form_no.w_no1.Enabled = False
            form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629コード変更
            form_no.w_no2.Enabled = False
            form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_comment.Enabled = False
            form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_dep_name.Enabled = False
            form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_name.Enabled = False
            form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
            form_no.w_entry_date.Enabled = False
            form_no.w_entry_date.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
        End If

    End Sub

    '2015/1/28 moriya add start
    Sub dataset_F_TMPMARK()
        Dim i As Short

        form_no.w_syurui.Text = ""
        Dim tmp_str As String = "" '20100624追加コード
        For i = 0 To 2
            tmp_str = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(i)) '20100624コード変更
            If temp_bz.syurui = tmp_str Then '20100624コード変更
                'form_no.w_syurui.ListIndex = i
                form_no.w_syurui.Text = tmp_str '20100624コード変更
                Exit For
            End If
        Next i

        '20100705コード変更'.Textを追加
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size5.Text = Trim(temp_bz.size5)
        form_no.w_size6.Text = Trim(temp_bz.size6)

    End Sub
    '2015/1/28 moriya add end


    Sub dataset_F_TMPKAJU()
        Dim i As Short

        form_no.w_syurui.Text = ""
        Dim tmp_str As String = "" '20100624追加コード
        For i = 0 To 2
            tmp_str = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(i)) '20100624コード変更
            If temp_bz.syurui = tmp_str Then '20100624コード変更
                'form_no.w_syurui.ListIndex = i
                form_no.w_syurui.Text = tmp_str '20100624コード変更
                Exit For
            End If
        Next i

        '20100705コード変更'.Textを追加
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size5.Text = Trim(temp_bz.size5)
        form_no.w_size6.Text = Trim(temp_bz.size6)

    End Sub

    Sub dataset_F_TMPMAXLD()
        Dim i As Short

        form_no.w_syurui.Text = ""
        Dim tmp_str As String = "" '20100624コード変更
        For i = 0 To 2
            tmp_str = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(i)) '20100624コード変更
            If temp_bz.syurui = tmp_str Then
                'form_no.w_syurui.ListIndex = i
                form_no.w_syurui.Text = tmp_str '20100624コード変更
                Exit For
            End If
        Next i

        '20100705コード変更'.Textを追加
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size5.Text = Trim(temp_bz.size5)
        form_no.w_size6.Text = Trim(temp_bz.size6)

    End Sub

    Sub dataset_F_TMPSERI()
        Dim i As Short

        form_no.w_syurui.Text = ""
        Dim tmp_str As String = "" '20100624追加コード
        For i = 0 To 2
            tmp_str = form_no.w_syurui.GetItemText(form_no.w_syurui.Items(i)) '20100624コード変更
            If temp_bz.syurui = tmp_str Then
                'form_no.w_syurui.ListIndex = i
                form_no.w_syurui.Text = tmp_str '20100624追加コード
                Exit For
            End If
        Next i

        '20100705コード変更'.Textを追加
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size5.Text = Trim(temp_bz.size5)
        form_no.w_size6.Text = Trim(temp_bz.size6)

        form_no.w_plant.Text = ""
        form_no.w_plant_code.Text = ""
        Select Case temp_bz.plant_code
            Case "CX"
                'form_no.w_plant.ListIndex = 0
                form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(0)) '20100624コード変更
            Case "N3"
                'form_no.w_plant.ListIndex = 1
                form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(1))
            Case "UY"
                'form_no.w_plant.ListIndex = 2
                form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(2))
            Case "CH"
                'form_no.w_plant.ListIndex = 3
                form_no.w_font.Text = form_no.w_font.GetItemText(form_no.w_font.Items(3))
        End Select

        form_no.w_tmp_seri_width = TmpSerialWidth
        form_no.w_tmp_seri_move = TmpSerialMove

    End Sub
    Sub dataset_F_TMPSIZE()

        '20100705コード変更'.Textを追加
        form_no.w_size1.Text = Trim(temp_bz.size1)
        form_no.w_size2.Text = Trim(temp_bz.size2)
        form_no.w_size3.Text = Trim(temp_bz.size3)
        form_no.w_size5.Text = Trim(temp_bz.size5)
        form_no.w_size6.Text = Trim(temp_bz.size6)

    End Sub


    Sub dataset_kikaku(ByRef wk_ctrl As System.Windows.Forms.Control, ByRef wk_kikaku_code As String)

        Select Case wk_kikaku_code
            Case "J"
                wk_ctrl.Text = "JIS"
            Case "E"
                wk_ctrl.Text = "ECE"
            Case "F"
                wk_ctrl.Text = "FMVSS"
            Case "I"
                wk_ctrl.Text = "INMETRO"
            Case "G"
                wk_ctrl.Text = "GULF"
            Case "A"
                wk_ctrl.Text = "ADR"
        End Select

    End Sub

    Sub dataset_nasiji(ByRef wk_ctrl As System.Windows.Forms.Control, ByRef wk_flag As Short)
        'wk_flag = 0:コントロールwk_ctrlに「無し」と表示
        '          1:コントロールwk_ctrlに「放電(N20)」と表示
        '          2:コントロールwk_ctrlに「放電(N30)」と表示
        '          3:コントロールwk_ctrlに「サンドブラスト」と表示
        Select Case wk_flag
            Case 0
                wk_ctrl.Text = "Null"
            Case 1
                wk_ctrl.Text = "Electric discharge (N20)"
            Case 2
                wk_ctrl.Text = "Electric discharge (N30)"
            Case 3
                wk_ctrl.Text = "Sand blast"
        End Select
    End Sub

    'Sub dataset_plant(ByRef wk_ctrl As System.Windows.Forms.Control, ByRef wk_plant As String, ByRef now_posi As String)'20100616移植追加
    Sub dataset_plant(ByRef wk_ctrl As Object, ByRef wk_plant As String, ByRef now_posi As String)

        If UCase(Trim(now_posi)) = "SAVE" Then

            Select Case wk_plant
                Case "TT"
                    '       wk_ctrl.Text = "仙台"
                    'wk_ctrl.ListIndex = 0
                    wk_ctrl.Text = wk_ctrl.GetItemText(wk_ctrl.Items(0)) '20100624コード変更

                    ' -> watanabe edit VerUP(2011)
                    'Case "NW"
                Case "KW"
                    ' <- watanabe edit VerUP(2011)

                    '       wk_ctrl.Text = "桑名"
                    'wk_ctrl.ListIndex = 1
                    wk_ctrl.Text = wk_ctrl.GetItemText(wk_ctrl.Items(1))
                Case "CS"
                    '       wk_ctrl.Text = "正新"
                    'wk_ctrl.ListIndex = 2
                    wk_ctrl.Text = wk_ctrl.GetItemText(wk_ctrl.Items(2))
                Case "CH"
                    '       wk_ctrl.Text = "上海"
                    'wk_ctrl.ListIndex = 3
                    wk_ctrl.Text = wk_ctrl.GetItemText(wk_ctrl.Items(3))
            End Select

        Else

            Select Case wk_plant
                Case "TT"
                    wk_ctrl.Text = "Sendai"
                    '        wk_ctrl.ListIndex = 0

                    ' -> watanabe edit VerUP(2011)
                    'Case "NW"
                Case "KW"
                    ' <- watanabe edit VerUP(2011)

                    wk_ctrl.Text = "Kuwana"
                    '        wk_ctrl.ListIndex = 1
                Case "CS"
                    wk_ctrl.Text = "Cheng shin"
                    '        wk_ctrl.ListIndex = 2
                Case "CH"
                    wk_ctrl.Text = "Shanghai"
                    '        wk_ctrl.ListIndex = 3
            End Select

        End If

    End Sub

    Sub dataset_F_TMP_PTNCODE()

        Dim i As Short


        If Trim(temp_bz.pattern) = "" Then Exit Sub


        '  form_no.w_type.Text = Left(Trim(temp_bz.pattern), 1)
        form_no.w_ptncode.Text = Trim(temp_bz.pattern)

        '  If Trim(form_no.w_type.Text) = "" Then Exit Sub
        If Trim(form_no.w_ptncode.Text) = "" Then Exit Sub

        '  If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", Left(Trim(temp_bz.pattern), 1), 0) = 0 Then
        '     form_no.w_type.Text = ""
        '     MsgBox "タイプが不正です｡ 表示できません", 64
        '
        '  End If

        For i = 1 To Len(Trim(form_no.w_ptncode.Text))
            If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ+-", Mid(Trim(form_no.w_ptncode.Text), i, 1), 0) = 0 Then
                form_no.w_ptncode.Text = ""
                MsgBox("Pattern code is invalid. Can not be displayed.", 64)
                Exit For
            End If
        Next i

    End Sub

    Sub dataset_F_TMP_PTNCODE2()

        Dim i As Short


        If Trim(temp_bz.pattern) = "" Then Exit Sub


        '  form_no.w_type.Text = Left(Trim(temp_bz.pattern), 1)
        form_no.w_ptncode.Text = Trim(temp_bz.pattern)

        '  If Trim(form_no.w_type.Text) = "" Then Exit Sub
        If Trim(form_no.w_ptncode.Text) = "" Then Exit Sub

        '  If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", Left(Trim(temp_bz.pattern), 1), 0) = 0 Then
        '     form_no.w_type.Text = ""
        '     MsgBox "タイプが不正です｡ 表示できません", 64
        '
        '  End If

        For i = 1 To Len(Trim(form_no.w_ptncode.Text))
            If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ+-", Mid(Trim(form_no.w_ptncode.Text), i, 1), 0) = 0 Then
                form_no.w_ptncode.Text = ""
                MsgBox("Pattern code is invalid. Can not be displayed.", 64)
                Exit For
            End If
        Next i

    End Sub
End Module