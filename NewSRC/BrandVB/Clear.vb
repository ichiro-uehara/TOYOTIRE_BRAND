Option Strict Off
Option Explicit On
Module MJ_Clear

#If PJ_BrandVB1 Then
	Sub Clear_F_GMSAVE()
        'Dim F_GMSAVE As Object

        'form_no = F_GMSAVE

        'open_mode = "新規"
        If open_mode = "NEW" Then
            form_no.w_font_name.Text = ""
            form_no.w_name2.Text = ""
            form_no.w_comment.Text = ""
            form_no.w_dep_name.Text = ""
            form_no.w_entry_name.Text = ""
            form_no.w_hem_width.Text = ""
            form_no.w_old_font_name.Text = ""
            form_no.w_old_font_class.Text = ""
            form_no.w_old_name.Text = ""
            form_no.w_base_r.Text = ""
            form_no.w_font_class1.Text = ""
            form_no.w_name1.Text = ""
        Else
        End If

    End Sub

#End If


#If PJ_BrandVB2 Then

    Sub Clear_F_HZDELE()
        Dim lp As Object
        'Dim F_HZDELE As Object

        form_no = F_HZDELE

        form_no.w_id = "HE"
        form_no.w_no1 = ""
        form_no.w_no2 = ""
        form_no.w_comment = ""
        form_no.w_dep_name = ""
        form_no.w_entry_name = ""
        form_no.w_entry_date = ""
        form_no.w_hm_num = ""

        '初期設定 (GRID)
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp

    End Sub


	Sub Clear_F_ZSEARCH_BRAND()
		Dim lp As Object
        'Dim F_ZSEARCH_BRAND As Object
		
		form_no = F_ZSEARCH_BRAND
		
		form_no.w_pattern.Text = ""
		form_no.w_size1.Text = ""
		form_no.w_size2.Text = ""
		form_no.w_size3.Text = ""
		form_no.w_size4.Text = ""
		form_no.w_size5.Text = ""
		form_no.w_size6.Text = ""
		form_no.w_kanri_no.Text = ""
		form_no.w_entry_name.Text = ""
		form_no.w_entry_date_0.Text = ""
		form_no.w_entry_date_1.Text = ""
		form_no.w_total.Text = ""
		
		form_no.MSFlexGrid1.Rows = 2
		form_no.MSFlexGrid1.Cols = 7 '----- 12/11 1997 yamamoto change 5→7 -----
		For lp = 0 To form_no.MSFlexGrid1.Cols - 1
			form_no.MSFlexGrid1.Row = 1
			form_no.MSFlexGrid1.Col = lp
			form_no.MSFlexGrid1.Text = ""
		Next lp
		
	End Sub
	
	Sub Clear_F_ZSEARCH_YOUSO()
		Dim lp As Object
		'Dim F_ZSEARCH_YOUSO As Object
		
		form_no = F_ZSEARCH_YOUSO
		
		form_no.w_mojicd.Text = ""
		form_no.w_total.Text = ""
		'Combo
        'form_no.w_taisho.Text = form_no.w_taisho.List(0)'20100701コード削除

		form_no.MSFlexGrid1.Rows = 2
		form_no.MSFlexGrid1.Cols = 7 '----- 12/11 1997 yamamoto change 5→7 -----
		For lp = 0 To form_no.MSFlexGrid1.Cols - 1
			form_no.MSFlexGrid1.Row = 1
			form_no.MSFlexGrid1.Col = lp
			form_no.MSFlexGrid1.Text = ""
		Next lp
		
	End Sub
	
	Sub Clear_F_ZSEARCH_BANGO()
		Dim lp As Object
		'Dim F_ZSEARCH_BANGO As Object
		
		form_no = F_ZSEARCH_BANGO
		
		form_no.w_id.Text = ""
		form_no.w_no1.Text = ""
		form_no.w_no2.Text = ""
		form_no.w_total.Text = ""
		
		form_no.MSFlexGrid1.Rows = 2
		form_no.MSFlexGrid1.Cols = 7
		For lp = 0 To form_no.MSFlexGrid1.Cols - 1
			form_no.MSFlexGrid1.Row = 1
			form_no.MSFlexGrid1.Col = lp
			form_no.MSFlexGrid1.Text = ""
		Next lp
		
	End Sub

#End If


#If PJ_BrandVB3 Then

	Sub Clear_F_TMP_UTQG()
        'Dim F_TMP_UTQG As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_UTQG

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		
		form_no.w_treadwear.Text = ""
		form_no.w_traction.Text = ""
		form_no.w_temperature.Text = ""
		
		form_no.w_hm_name.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_PLATE()
		Dim i As Object
		'Dim F_TMP_PLATE As Object
		
		Dim w_str As String


        ' -> watanabe add VerUP(2011)
        w_str = ""
        ' -> watanabe add VerUP(2011)

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_PLATE

        'form_no.w_type.ListIndex = 0
        form_no.w_type.SelectedIndex = 0 '20100628コード変更
		
		'(Brand Ver.3 変更)
		If form_no.w_type.Text = "Screw nothing" Then
			w_str = "PLATE1"
		ElseIf form_no.w_type.Text = "front screw" Then 
			w_str = "PLATE2"
		ElseIf form_no.w_type.Text = "Back screw" Then 
			w_str = "PLATE3"
		ElseIf form_no.w_type.Text = "front and back screw" Then 
			w_str = "PLATE4"
		End If

        For i = 1 To MaxSelNum

            ' -> watanabe edit VerUP(2011)
            'If Tmp_hm_word(i) = w_str Then
            If Tmp_hm_word(i) = form_no.w_type.Text Then
                ' <- watanabe edit VerUP(2011)

                form_no.w_hm_name.Text = Tmp_hm_code(i)
                Exit For
            End If
        Next i
		
		' -> watanabe add 2007.03
		form_no.w_plate_w.Text = Tmp_plate_w
		form_no.w_plate_h.Text = Tmp_plate_h
		form_no.w_plate_r.Text = Tmp_plate_r
		form_no.w_plate_n.Text = Tmp_plate_n
		' <- watanabe add 2007.03


        ' -> watanabe add VerUP(2011)
        Dim w_file As String
        Dim TiffFile As String

        If Trim(form_no.w_hm_name.Text) = "" Then Exit Sub

        TiffFile = TIFFDir & form_no.w_hm_name.Text & ".bmp"

        'BMPﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 457
            form_no.ImgThumbnail1.Height = 193
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
        End If
        ' <- watanabe add VerUP(2011)

    End Sub
	
	Sub Clear_F_TMP_SIZE(ByRef clr_level As Short)
        'Dim F_TMP_SIZE As Object
        '  パラメータ
        '  clr_level :0 = 全項目クリア
        '  clr_level  1 = 編集文字項目のみクリア

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_SIZE

        If clr_level = 0 Then
			form_no.w_size1.Text = ""
			form_no.w_size2.Text = ""
			form_no.w_size3.Text = ""
			form_no.w_size4.Text = ""
			form_no.w_size5.Text = ""
			form_no.w_size6.Text = ""
			form_no.w_hm_num.Text = ""
			form_no.w_hight.Text = ""
			form_no.w_width.Text = ""
			form_no.w_ang.Text = ""
		End If
		'Combo
		form_no.w_hm_name.Items.Clear()
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_MAXLOAD(ByRef clr_level As Short)
        'Dim F_TMP_MAXLOAD As Object
        '  パラメータ
        '  clr_level :0 = 全項目クリア
        '  clr_level  1 = 規格値、作図値のみクリア

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_MAXLOAD

        If clr_level = 0 Then
			form_no.w_size1.Text = ""
			form_no.w_size2.Text = ""
			form_no.w_size3.Text = ""
			form_no.w_size4.Text = ""
			form_no.w_size5.Text = ""
			form_no.w_size6.Text = ""
            'form_no.w_kikaku.ListIndex = 0
            form_no.w_kikaku.SelectedIndex = 0 '20100628コード変更
            'form_no.w_syurui.ListIndex = 0
            form_no.w_syurui.SelectedIndex = 0 '20100628コード変更
            form_no.w_font.SelectedIndex = 0 '20100628コード変更
			form_no.w_type.Text = ""
			form_no.w_hm_name.Text = ""
			'Brand Ver.5 TIFF->BMP 変更 start
			'      form_no.ImgThumbnail1.Image = ""
			form_no.ImgThumbnail1.Image = Nothing
			'Brand Ver.5 TIFF->BMP 変更 end
		End If
		
		form_no.w_kajyu.Text = ""
		form_no.w_kikaku_max_load_kg.Text = ""
		form_no.w_kikaku_max_load_lbs.Text = ""
		form_no.w_kikaku_max_press_kpa.Text = ""
		form_no.w_kikaku_max_press_psi.Text = ""
		form_no.w_max_load_kg.Text = ""
		form_no.w_max_load_lbs.Text = ""
		form_no.w_max_press_kpa.Text = ""
		form_no.w_max_press_psi.Text = ""
		
	End Sub
	
	
	
	Sub Clear_F_TMP_SERIARU(ByRef clr_level As Short)
        'Dim F_TMP_SERIARU As Object
        '  パラメータ
        '  clr_level :0 = 全項目クリア
        '  clr_level  1 = 編集文字項目のみクリア

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_SERIARU

        If clr_level = 0 Then
			form_no.w_size1.Text = ""
			form_no.w_size2.Text = ""
			form_no.w_size3.Text = ""
			form_no.w_size4.Text = ""
			form_no.w_size5.Text = ""
			form_no.w_size6.Text = ""
            form_no.w_syurui.SelectedIndex = 0 '20100628コード変更
            form_no.w_font.SelectedIndex = 0 '20100628コード変更
			form_no.w_plant.Text = ""
			form_no.w_hm_name.Text = ""
		End If
		form_no.w_size_code.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_ENO()
        'Dim F_TMP_ENO As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_ENO

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		form_no.w_shonin.Text = ""
		form_no.w_hm_name.Text = ""
		
        '2011/12/08 uriu added
        form_no.chk_shonin.CheckState = 0
        form_no.w_s.Text = ""
        form_no.chk_s.CheckState = 0
        form_no.w_r.Text = ""
        form_no.chk_r.CheckState = 0

        'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
		'  If form_no.w_type.Text = "E4" Then
		'      form_no.w_hm_name.Text = E4
		'  ElseIf form_no.w_type.Text = "E5" Then
		'      form_no.w_hm_name.Text = E5
		'  End If
		
	End Sub
	
	Sub Clear_F_TMP_PLY()
        'Dim F_TMP_PLY As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_PLY

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		form_no.w_tread.Text = ""
		form_no.w_tread1.Text = ""
		form_no.w_tread2.Text = ""
		form_no.w_tread3.Text = ""
		form_no.w_sidewall.Text = ""
		form_no.w_hm_name.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_PLY2()
        'Dim F_TMP_PLY2 As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_PLY2

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		form_no.w_tread1.Text = ""
		form_no.w_tread2.Text = ""
		form_no.w_tread3.Text = ""
		form_no.w_sidewall1.Text = ""
		form_no.w_sidewall2.Text = ""
		form_no.w_hm_name.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_ETC()
		Dim i As Object
        'Dim F_TMP_ETC As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_ETC

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		For i = 1 To 10
			form_no.w_etc(i).Text = ""
		Next i
		form_no.w_hm_name.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	
	Sub Clear_F_TMP_KAJUU(ByRef clr_level As Short)
        'Dim F_TMP_KAJUU As Object
        '  パラメータ
        '  clr_level :0 = 全項目クリア
        '  clr_level  1 = 編集文字項目のみクリア

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_KAJUU

        If clr_level = 0 Then
			form_no.w_size1.Text = ""
			form_no.w_size2.Text = ""
			form_no.w_size3.Text = ""
			form_no.w_size4.Text = ""
			form_no.w_size5.Text = ""
			form_no.w_size6.Text = ""
            'form_no.w_syurui.ListIndex = 0
            form_no.w_syurui.SelectedIndex = 0 '20100628コード変更
            'form_no.w_kikaku.ListIndex = 0
            form_no.w_kikaku.SelectedIndex = 0 '20100701コード変更
			form_no.w_sokudo.Text = ""
		End If
		form_no.w_hm_name.Items.Clear()
		form_no.w_hm_num.Text = ""
		form_no.w_load_index.Text = ""
		form_no.w_hight.Text = ""
		form_no.w_ang.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	Sub Clear_F_TMP_MORUDO()
		'Dim F_TMP_MORUDO As Object
		
		form_no = F_TMP_MORUDO
		
        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		form_no.w_kubun.Text = ""
		form_no.w_no.Text = ""
		form_no.w_hm_name.Text = ""
		
		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
		form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end
		
	End Sub
	
	
	' -> watanabe add 2007.03
	
	Sub Clear_F_TMP_UTQG3()
        'Dim F_TMP_UTQG3 As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_UTQG3

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
		form_no.w_type.Text = ""
		form_no.w_treadwear.Text = ""
		form_no.w_traction.Text = ""
		form_no.w_temperature.Text = ""
		form_no.w_hm_name.Text = ""
		form_no.ImgThumbnail1.Image = Nothing
		
	End Sub
	
	Sub Clear_F_TMP_MAXLOAD3(ByRef clr_level As Short)
        'Dim F_TMP_MAXLOAD3 As Object
        '  パラメータ
        '  clr_level :0 = 全項目クリア
        '  clr_level  1 = 規格値、作図値のみクリア

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_MAXLOAD3

        If clr_level = 0 Then
			form_no.w_size1.Text = ""
			form_no.w_size2.Text = ""
			form_no.w_size3.Text = ""
			form_no.w_size4.Text = ""
			form_no.w_size5.Text = ""
			form_no.w_size6.Text = ""
            'form_no.w_kikaku.ListIndex = 0
            form_no.w_kikaku.SelectedIndex = 0 '20100628コード変更
            'form_no.w_syurui.ListIndex = 0
            form_no.w_syurui.SelectedIndex = 0 '20100628コード変更
            'form_no.w_font.ListIndex = 0
            form_no.w_font.SelectedIndex = 0 '20100628コード変更
			form_no.w_type.Text = ""
			form_no.w_hm_name.Text = ""
			form_no.ImgThumbnail1.Image = Nothing
		End If
		
		form_no.w_kajyu.Text = ""
		form_no.w_kikaku_max_load_kg.Text = ""
		form_no.w_kikaku_max_load_lbs.Text = ""
		form_no.w_kikaku_max_press_kpa.Text = ""
		form_no.w_kikaku_max_press_psi.Text = ""
		form_no.w_max_load_kg.Text = ""
		form_no.w_max_load_lbs.Text = ""
		form_no.w_max_press_kpa.Text = ""
		form_no.w_max_press_psi.Text = ""
		
    End Sub
    Sub Clear_F_TMP_PLY1_3()
        'Dim F_TMP_PLY1_3 As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_PLY1_3

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
        form_no.w_type.Text = ""
        form_no.w_tread.Text = ""
        form_no.w_tread1.Text = ""
        form_no.w_tread2.Text = ""
        form_no.w_tread3.Text = ""
        form_no.w_sidewall.Text = ""
        form_no.w_hm_name.Text = ""
        form_no.ImgThumbnail1.Image = Nothing

    End Sub

    Sub Clear_F_TMP_PLY2_3()
        'Dim F_TMP_PLY2_3 As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_PLY2_3

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
        form_no.w_type.Text = ""
        form_no.w_tread1.Text = ""
        form_no.w_tread2.Text = ""
        form_no.w_tread3.Text = ""
        form_no.w_sidewall1.Text = ""
        form_no.w_sidewall2.Text = ""
        form_no.w_hm_name.Text = ""
        form_no.ImgThumbnail1.Image = Nothing

    End Sub

    Sub Clear_F_TMP_ETC3()
        Dim i As Object
        'Dim F_TMP_ETC3 As Object

        '----- .NET 移行(一旦コメント化) -----
        'form_no = F_TMP_ETC3

        'form_no.w_font.ListIndex = 0
        form_no.w_font.SelectedIndex = 0 '20100628コード変更
        form_no.w_type.Text = ""
        For i = 1 To 10
            form_no.w_etc(i).Text = ""
        Next i
        form_no.w_hm_name.Text = ""
        form_no.ImgThumbnail1.Image = Nothing

    End Sub

    ' <- watanabe add 2007.03

#End If

#If PJ_BrandVB2 Then

	Sub Clear_F_GZDELE()
        Dim lp As Object

        ' -> watanabe del VerUP(2011)
        'Dim F_GZDELE As Object
        ' <- watanabe del VerUP(2011)

		form_no = F_GZDELE

		form_no.w_id = "KO"
		form_no.w_no1 = ""
		form_no.w_no2 = ""
		form_no.w_comment = ""
		form_no.w_dep_name = ""
		form_no.w_entry_name = ""
		form_no.w_entry_date = ""
		form_no.w_gm_num = ""

		'初期設定 (GRID)
        'UPGRADE_ISSUE: Control MSFlexGrid1 は、汎用名前空間 Form 内にあるため、解決できませんでした。
		form_no.MSFlexGrid1.Rows = 2
		form_no.MSFlexGrid1.Cols = 6
		For lp = 0 To form_no.MSFlexGrid1.Cols - 1
			form_no.MSFlexGrid1.Row = 1
			form_no.MSFlexGrid1.Col = lp
			form_no.MSFlexGrid1.Text = ""
		Next lp

	End Sub

#End If

#If PJ_BrandVB1 Then

	Sub Clear_F_GMDELE()

		form_no.w_mojicd.Text = ""
		form_no.w_comment.Text = ""
		form_no.w_dep_name.Text = ""
		form_no.w_entry_name.Text = ""
		form_no.w_entry_date.Text = ""
		form_no.w_high.Text = ""
		form_no.w_width.Text = ""
		form_no.w_ang.Text = ""
		form_no.w_moji_high.Text = ""
		form_no.w_moji_shift.Text = ""
		form_no.w_hem_width.Text = ""
		form_no.w_hatch_ang.Text = ""
		form_no.w_hatch_width.Text = ""
		form_no.w_hatch_space.Text = ""
		form_no.w_hatch_x.Text = ""
		form_no.w_hatch_y.Text = ""
		form_no.w_old_font_name.Text = ""
		form_no.w_old_font_class.Text = ""
		form_no.w_old_name.Text = ""
		form_no.w_org_hor.Text = ""
		form_no.w_org_ver.Text = ""
		form_no.w_base_r.Text = ""

		'Brand Ver.5 TIFF->BMP 変更 start
		'   form_no.ImgThumbnail1.Image = ""
        form_no.ImgThumbnail1.Image = Nothing
		'Brand Ver.5 TIFF->BMP 変更 end


    End Sub

    Sub Clear_F_GMSEARCH()
        Dim lp As Object
        'Dim F_GMSEARCH As Object
        form_no = F_GMSEARCH

        form_no.w_font_name.Text = ""
        form_no.w_name1.Text = ""
        form_no.w_name2.Text = ""
        form_no.w_font_class1.Text = ""
        form_no.w_font_class2.Text = ""
        form_no.w_high.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date_0.Text = ""
        form_no.w_entry_date_1.Text = ""

        'Brand Cad System Ver.3 UP
        form_no.w_old_font.Text = ""

        '----- .NET移行 (MSFlexGrid コメント化) -----
        'form_no.MSFlexGrid1.Rows = 2
        'form_no.MSFlexGrid1.Cols = 22
        'UPGRADE_ISSUE: Control MSFlexGrid1 は、汎用名前空間 Form 内にあるため、解決できませんでした。
        'For lp = 0 To form_no.MSFlexGrid1.Cols - 1
        '    form_no.MSFlexGrid1.Row = 1
        '    form_no.MSFlexGrid1.Col = lp
        '    form_no.MSFlexGrid1.Text = ""
        'Next lp

        form_no.w_total.Text = ""

        '----- .NET移行 (MSFlexGrid コメント化) -----
        'form_no.MSFlexGrid1.Enabled = False
        F_GMSEARCH.DataGridViewList.Rows.Clear()
        F_GMSEARCH.ImgThumbnail1.Image = Nothing
        F_GMSEARCH.w_total.Text = ""

    End Sub


    Sub Clear_F_HMLOOK()
        Dim lp As Object
        'Dim F_HMLOOK As Object
        'form_no = F_HMLOOK

        form_no.w_mojicd.Text = ""
        form_no.w_flag_delete.Text = ""
        form_no.w_id.Text = ""
        form_no.w_spell.Text = ""
        form_no.w_width.Text = ""
        form_no.w_high.Text = ""
        form_no.w_ang.Text = ""
        form_no.w_haiti_pic.Text = ""
        form_no.w_haiti_sitei.Text = ""
        form_no.w_hz_id.Text = ""
        form_no.w_hz_no1.Text = ""
        form_no.w_hz_no2.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date.Text = ""

        form_no.w_gm_num.Text = ""
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp
        'form_no.MSFlexGrid1.Enabled = False

        form_no.ImgThumbnail1.Image = Nothing

    End Sub


    Sub Clear_F_HMSEARCH()
        Dim lp As Object
        'Dim F_HMSEARCH As Object

        'form_no = F_HMSEARCH

        form_no.w_font_name.Text = ""
        form_no.w_no.Text = ""
        form_no.w_spell.Text = ""
        form_no.w_high.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date_0.Text = ""
        form_no.w_entry_date_1.Text = ""

        '----- .NET移行 (MSFlexGrid コメント化) -----
        'form_no.MSFlexGrid1.Rows = 2
        'form_no.MSFlexGrid1.Cols = 13

        'For lp = 0 To form_no.MSFlexGrid1.Cols - 1
        '    form_no.MSFlexGrid1.Row = 1
        '    form_no.MSFlexGrid1.Col = lp
        '    form_no.MSFlexGrid1.Text = ""
        'Next lp
        'form_no.w_total.Text = ""

        '   form_no.MSFlexGrid1.Enabled = False

        F_HMSEARCH.DataGridViewList.Rows.Clear()
        F_HMSEARCH.ImgThumbnail1.Image = Nothing
        F_HMSEARCH.w_total.Text = ""

    End Sub

    Sub Clear_F_HMSEARCH2()
        Dim lp As Object
        'Dim F_HMSEARCH2 As Object

        'form_no = F_HMSEARCH2

        form_no.w_gm_code.Text = ""
        form_no.w_font_name.Text = ""

        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 13


        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp
        form_no.w_total.Text = ""

        '   form_no.MSFlexGrid1.Enabled = False

    End Sub

    Sub Clear_F_GMLOOK()

        form_no.w_mojicd.Text = ""
        form_no.w_flag_delete.Text = ""
        form_no.w_id.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date.Text = ""
        form_no.w_high.Text = ""
        form_no.w_width.Text = ""
        form_no.w_ang.Text = ""
        form_no.w_moji_high.Text = ""
        form_no.w_moji_shift.Text = ""
        form_no.w_hem_width.Text = ""
        form_no.w_hatch_ang.Text = ""
        form_no.w_hatch_width.Text = ""
        form_no.w_hatch_space.Text = ""
        form_no.w_hatch_x.Text = ""
        form_no.w_hatch_y.Text = ""
        form_no.w_old_font_name.Text = ""
        form_no.w_old_font_class.Text = ""
        form_no.w_old_name.Text = ""
        form_no.w_org_hor.Text = ""
        form_no.w_org_ver.Text = ""
        form_no.w_org_x.Text = ""
        form_no.w_org_y.Text = ""
        form_no.w_left_bottom_x.Text = ""
        form_no.w_left_bottom_y.Text = ""
        form_no.w_right_bottom_x.Text = ""
        form_no.w_right_bottom_y.Text = ""
        form_no.w_left_top_x.Text = ""
        form_no.w_left_top_y.Text = ""
        form_no.w_right_top_x.Text = ""
        form_no.w_right_top_y.Text = ""

        form_no.w_base_r.Text = ""
        'form_no.w_haiti_pic = ""
        form_no.w_haiti_pic.Text = "" '20100621移植変更

        'Brand Ver.5 TIFF->BMP 変更 start
        '   form_no.ImgThumbnail1.Image = ""
        'form_no.ImgThumbnail1.Picture = Nothing
        'Brand Ver.5 TIFF->BMP 変更 end
        form_no.ImgThumbnail1.Image = Nothing '20100621移植変更

    End Sub

    Sub Clear_F_HMSAVE()
        'Dim F_HMSAVE As Object

        'form_no = F_HMSAVE

        If open_mode = "NEW" Then
            form_no.w_font_name.Text = ""
            'form_no.w_no.Text = ""
            'form_no.w_haiti_pic.Text = ""
            form_no.w_spell.Text = ""
            form_no.w_comment.Text = ""
            form_no.w_dep_name.Text = ""
            form_no.w_entry_name.Text = ""
            'form_no.w_entry_date.Text = ""
            'form_no.w_width.Text = ""
            form_no.w_high.Text = ""
            form_no.w_ang.Text = ""
            'form_no.w_gm_num.Text = ""
        Else
            'form_no.w_font_name.Text = ""
            'form_no.w_no.Text = ""
            form_no.w_spell.Text = ""
            form_no.w_comment.Text = ""
            form_no.w_dep_name.Text = ""
            form_no.w_entry_name.Text = ""
            'form_no.w_entry_date.Text = ""
            'form_no.w_width.Text = ""
            form_no.w_high.Text = ""
            form_no.w_ang.Text = ""
            'form_no.w_gm_num.Text = ""
        End If

    End Sub

    Sub Clear_F_HMDELE()
        Dim lp As Object
        'Dim F_HMDELE As Object

        'form_no = F_HMDELE

        form_no.w_mojicd.Text = ""
        form_no.w_haiti_pic.Text = ""
        form_no.w_spell.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date.Text = ""
        form_no.w_width.Text = ""
        form_no.w_high.Text = ""
        form_no.w_ang.Text = ""
        form_no.w_gm_num.Text = ""

        '初期設定 (GRID)
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp

    End Sub
#End If

#If PJ_BrandVB2 Then

    Sub Clear_F_GZSAVE()
        If open_mode = "NEW" Then
            form_no.w_no1.Text = ""
        ElseIf open_mode = "modify" Then
            'form_no.w_no1.Text = ""
            'form_no.w_no2.Text = ""
            'form_no.w_gm_num.Text = ""
        Else
            'form_no.w_no1.Text = ""
            'form_no.w_no2.Text = ""
            'form_no.w_gm_num.Text = ""
        End If
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""

    End Sub
    Sub Clear_F_GZLOOK()
        Dim lp As Object
        'Dim F_GZLOOK As Object

        form_no = F_GZLOOK

        form_no.w_flag_delete.Text = ""
        form_no.w_id.Text = "KO"
        form_no.w_no1.Text = ""
        form_no.w_no2.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date.Text = ""
        form_no.w_gm_num.Text = ""

        '初期設定 (GRID)
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp

    End Sub

    Sub Clear_F_HZLOOK()
        Dim lp As Object
        'Dim F_HZLOOK As Object

        form_no = F_HZLOOK

        form_no.w_flag_delete.Text = ""
        form_no.w_id.Text = "HE"
        form_no.w_no1.Text = ""
        form_no.w_no2.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        form_no.w_entry_date.Text = ""
        form_no.w_flag_delete.Text = ""
        form_no.w_hm_num.Text = ""

        '初期設定 (GRID)
        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        For lp = 0 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = lp
            form_no.MSFlexGrid1.Text = ""
        Next lp

    End Sub

    Sub Clear_F_HZSAVE()
        If open_mode = "NEW" Then
            form_no.w_no1.Text = ""
        ElseIf open_mode = "modify" Then
            'form_no.w_no1.Text = ""
            'form_no.w_hm_num.Text = ""
        Else
            'form_no.w_no1.Text = ""
            'form_no.w_no2.Text = ""
            'form_no.w_hm_num.Text = ""
        End If
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""

    End Sub
    Sub Clear_F_BZSAVE()

        If open_mode = "NEW" Then
            form_no.w_no1.Text = ""
        ElseIf open_mode = "modify" Then
            'form_no.w_no1.Text = ""
            'form_no.w_no2.Text = ""
            'form_no.w_hm_num.Text = ""
        Else
            'form_no.w_no1.Text = ""
            'form_no.w_no2.Text = ""
            'form_no.w_hm_num.Text = ""
        End If
        form_no.w_kanri_no.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""
        '   form_no.w_entry_date.Text = ""

        form_no.w_syurui.Text = ""
        form_no.w_syubetu.Text = ""
        form_no.w_pattern.Text = ""

        form_no.w_size1.Text = ""
        form_no.w_size2.Text = ""
        form_no.w_size3.Text = ""
        form_no.w_size4.Text = ""
        form_no.w_size5.Text = "R"
        form_no.w_size6.Text = ""
        form_no.w_size7.Text = ""
        form_no.w_size8.Text = ""
        form_no.w_size.Text = ""
        form_no.w_size_code.Text = ""

        form_no.w_plant.Text = ""
        form_no.w_plant_code.Text = ""


        ' -> watanabe del VerUP(2011)   2007年のバージョンアップで削除されている項目
        'form_no.w_kikaku1.Text = ""
        'form_no.w_kikaku2.Text = ""
        'form_no.w_kikaku3.Text = ""
        'form_no.w_kikaku4.Text = ""
        'form_no.w_kikaku5.Text = ""
        'form_no.w_kikaku6.Text = ""
        'form_no.w_kikaku.Text = ""
        'form_no.w_tos_moyou.Text = ""
        'form_no.w_side_moyou.Text = ""
        'form_no.w_side_kenti.Text = ""
        'form_no.w_peak_mark.Text = ""
        'form_no.w_nasiji.Text = ""
        ' <- watanabe del VerUP(2011)

    End Sub

    Sub Clear_F_BZDELE()
        form_no.w_no1.Text = ""
        form_no.w_no2.Text = ""
        form_no.w_kanri_no.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""

        form_no.w_syurui.Text = ""
        form_no.w_syubetu.Text = ""
        form_no.w_pattern.Text = ""

        form_no.w_size1.Text = ""
        form_no.w_size2.Text = ""
        form_no.w_size3.Text = ""
        form_no.w_size4.Text = ""
        form_no.w_size5.Text = ""
        form_no.w_size6.Text = ""
        form_no.w_size7.Text = ""
        form_no.w_size8.Text = ""
        form_no.w_size.Text = ""
        form_no.w_size_code.Text = ""

        form_no.w_plant.Text = ""
        form_no.w_plant_code.Text = ""

        form_no.w_kikaku1.Text = ""
        form_no.w_kikaku2.Text = ""
        form_no.w_kikaku3.Text = ""
        form_no.w_kikaku4.Text = ""
        form_no.w_kikaku5.Text = ""
        form_no.w_kikaku6.Text = ""
        form_no.w_kikaku.Text = ""

        form_no.w_tos_moyou.Text = ""
        form_no.w_side_moyou.Text = ""
        form_no.w_side_kenti.Text = ""
        form_no.w_peak_mark.Text = ""
        form_no.w_nasiji.Text = ""

        '97.04.24 update n.matsumi
        form_no.w_hm_num.Text = ""

    End Sub

    Sub Clear_F_BZLOOK()

        form_no.w_flag_delete.Text = ""
        form_no.w_no1.Text = ""
        form_no.w_no2.Text = ""
        form_no.w_kanri_no.Text = ""
        form_no.w_comment.Text = ""
        form_no.w_dep_name.Text = ""
        form_no.w_entry_name.Text = ""

        form_no.w_syurui.Text = ""
        form_no.w_syubetu.Text = ""
        form_no.w_pattern.Text = ""

        form_no.w_size1.Text = ""
        form_no.w_size2.Text = ""
        form_no.w_size3.Text = ""
        form_no.w_size4.Text = ""
        form_no.w_size5.Text = ""
        form_no.w_size6.Text = ""
        form_no.w_size7.Text = ""
        form_no.w_size8.Text = ""
        form_no.w_size.Text = ""
        form_no.w_size_code.Text = ""

        form_no.w_plant.Text = ""
        form_no.w_plant_code.Text = ""

        form_no.w_kikaku1.Text = ""
        form_no.w_kikaku2.Text = ""
        form_no.w_kikaku3.Text = ""
        form_no.w_kikaku4.Text = ""
        form_no.w_kikaku5.Text = ""
        form_no.w_kikaku6.Text = ""
        form_no.w_kikaku.Text = ""

        form_no.w_tos_moyou.Text = ""
        form_no.w_side_moyou.Text = ""
        form_no.w_side_kenti.Text = ""
        form_no.w_peak_mark.Text = ""
        form_no.w_nasiji.Text = ""


        ' -> watanabe add VerUP(2011)
        form_no.w_entry_date.Text = ""
        form_no.w_hm_num.Text = ""

        form_no.MSFlexGrid1.Rows = 2
        form_no.MSFlexGrid1.Cols = 6
        Dim ii As Integer
        For ii = 1 To form_no.MSFlexGrid1.Cols - 1
            form_no.MSFlexGrid1.Row = 1
            form_no.MSFlexGrid1.Col = ii
            form_no.MSFlexGrid1.Text = ""
        Next ii
        ' <- watanabe add VerUP(2011)

    End Sub
#End If

End Module