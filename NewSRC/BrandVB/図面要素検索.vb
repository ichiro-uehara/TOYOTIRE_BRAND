Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_ZSEARCH_YOUSO
	Inherits System.Windows.Forms.Form
	
	'�T�v�F�{�^���N���b�N����
	'�����F�L�����Z���t���O�𗧂Ă�
	'----- 1/28 1997 by yamamoto -----
	Private Sub cmd_Cancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Cancel.Click
		
		GL_cancel_flg = 1
		
	End Sub
	
	'�T�v�F�{�^���N���b�N����
	'�����F�ǉ����ځF���b�N�Z�b�g�A�L�����Z�����I�������ƌ����𒆎~����i �O���b�h�̓N���A �j
	Private Sub cmd_Search_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Search.Click
		Dim lp As Object
		Dim j As Object
		Dim num As String
		Dim result As Object
		
		Dim search_word(100) As String
		Dim L_DAT(20) As String
		Dim i As Short
		Dim ic As Object
		Dim ir As Short
		'  Dim w_dir  As String
		Dim w_file As String
		Dim TiffFile As String
		Dim w_moji As New VB6.FixedLengthString(10)
		Dim cnt_max As Short
		Dim srh_cnt As Integer
		Dim w_ret As Short
		Dim index_row As String

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)


		On Error Resume Next
		srh_cnt = 0
		GL_cancel_flg = 0
		
		'MsgBox "�������܂�"
		MSFlexGrid1.Rows = 2
		For i = 0 To MSFlexGrid1.Cols - 1
			w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
		Next i
		MSFlexGrid1.Enabled = False
		
		'�܂������̕\��
		'   If FreePicNum < 1 Then
		'      MsgBox "�󂫃s�N�`��������܂���", vbCritical
		'      Exit Sub
		'   End If
		
		w_moji.Value = Trim(w_mojicd.Text)
        If w_taisho.Text = "Stamp drawing" Then
            'UPGRADE_WARNING: �I�u�W�F�N�g TIFFDirGM �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            'UPGRADE_WARNING: �I�u�W�F�N�g TIFFDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            TIFFDir = TIFFDirGM
        Else
            'UPGRADE_WARNING: �I�u�W�F�N�g TIFFDirHM �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            'UPGRADE_WARNING: �I�u�W�F�N�g TIFFDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            TIFFDir = TIFFDirHM
        End If
		
		'Brand Ver.5 TIFF->BMP �ύX start
		'   TiffFile = TIFFDir & Trim$(w_moji) & ".tif"
		'   'Tiff̧�ٕ\��
		'   w_file = Dir(TiffFile)
		'   If w_file <> "" Then
		'       ImgThumbnail1.Image = TiffFile
		'       ImgThumbnail1.ThumbWidth = 500
		'       ImgThumbnail1.ThumbHeight = 200
		'   Else
		'       MsgBox "TIFF̧�ق�������܂���", vbCritical
		'       Exit Sub
		'   End If
		'UPGRADE_WARNING: �I�u�W�F�N�g TIFFDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		TiffFile = TIFFDir & Trim(w_moji.Value) & ".bmp"
		'BMP̧�ٕ\��
		'UPGRADE_WARNING: Dir �ɐV�������삪�w�肳��Ă��܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"' ���N���b�N���Ă��������B
		w_file = Dir(TiffFile)
		If w_file <> "" Then
			ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            ImgThumbnail1.Width = 457 '500 '20100701�R�[�h�ύX
            ImgThumbnail1.Height = 193 '200 '20100701�R�[�h�ύX
		Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
			Exit Sub
		End If
		'Brand Ver.5 TIFF->BMP �ύX start
		
		Select Case w_taisho.Text
            Case "Stamp drawing"
                ' Brand Ver.3 �ύX
                '         DBTableName = DBName & "..gz_kanri"
                DBTableName = DBName & "..gz_kanri1"
                DBTableName2 = DBName & "..gz_kanri2"
                cnt_max = 63
            Case "Editing characters drawing"
                ' Brand Ver.3 �ύX
                '         DBTableName = DBName & "..hz_kanri"
                DBTableName = DBName & "..hz_kanri1"
                DBTableName2 = DBName & "..hz_kanri2"
                cnt_max = 63
            Case "Brand drawing"
                ' Brand Ver.3 �ύX
                '         DBTableName = DBName & "..bz_kanri"
                DBTableName = DBName & "..bz_kanri1"
                DBTableName2 = DBName & "..bz_kanri2"
                cnt_max = 100
        End Select
		
		init_sql()
		
		' Brand Ver.3 �ύX
		'   search_word(0) = " WHERE flag_delete = 0 "
		'   search_word(0) = search_word(0) & "AND ( "
		'   For i = 1 To cnt_max - 1
		'      Select Case w_taisho.Text
		'         Case "����}��"
		'            search_word(i) = "( gm_name" & Format$(i, "000") & " = '" & w_moji & "' ) OR "
		'         Case "�ҏW�����}��", "�u�����h�}��"
		'            search_word(i) = "( hm_name" & Format$(i, "000") & " = '" & w_moji & "' ) OR "
		'      End Select
		'   Next i
		
		'   Select Case w_taisho.Text
		'      Case "����}��"
		'         search_word(cnt_max) = "( gm_name" & Format$(cnt_max, "000") & " = '" & w_moji & "' ) )"
		'      Case "�ҏW�����}��", "�u�����h�}��"
		'         search_word(cnt_max) = "( hm_name" & Format$(cnt_max, "000") & " = '" & w_moji & "' ) )"
		'   End Select
		
		
		'  Brand Ver.3 �ύX
		'  �ް��ް��Y��������\��
		'   result = SqlCmd(SqlConn, "SELECT COUNT(*) FROM " & DBTableName)
		'   For i = 0 To cnt_max
		'      result = SqlCmd(SqlConn, search_word(i))
		'   Next i
		'   result = SqlExec(SqlConn)
		'   result = SqlResults(SqlConn)
		
		'   If result = SUCCEED Then
		'      Do Until SqlNextRow(SqlConn) = NOMOREROWS
		'         num$ = SqlData$(SqlConn, 1)
		'         w_total = num$
		'      Loop
		'   Else
		'      MsgBox "�����Ɏ��s���܂���"
		'      GoTo error_section
		'   End If
		
		
        ' -> watanabe edit VerUP(2011)
        '      ' Brand Ver.3 �ύX
        ''    result = SqlCmd(SqlConn, "SELECT DISTINCT id, no1, no2")
        ''UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        'result = SqlCmd(SqlConn, "SELECT COUNT(*) ")
        ''UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        'result = SqlCmd(SqlConn, " FROM " & DBTableName2)
        'If w_taisho.Text = "����}��" Then
        '	'UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        '	result = SqlCmd(SqlConn, " WHERE gm_name = '" & w_moji.Value & "' ")
        'Else
        '	'UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        '	result = SqlCmd(SqlConn, " WHERE hm_name = '" & w_moji.Value & "' ")
        'End If
        ''UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        'result = SqlExec(SqlConn)
        ''UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        'result = SqlResults(SqlConn)
        ''    srh_cnt = 0
        ''    If result = SUCCEED Then
        ''        Do Until SqlNextRow(SqlConn) = NOMOREROWS
        ''            srh_cnt = srh_cnt + 1
        ''        Loop
        ''    End If
        ''    w_total = srh_cnt
        ''UPGRADE_WARNING: �I�u�W�F�N�g result �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        'If result = SUCCEED Then
        '	Do Until SqlNextRow(SqlConn) = NOMOREROWS
        '		num = SqlData(SqlConn, 1)
        '		w_total.Text = num
        '	Loop 
        'End If


        '�����R�}���h�쐬
        sqlcmd = "SELECT COUNT(*) "
        sqlcmd = sqlcmd & " FROM " & DBTableName2
        If w_taisho.Text = "Stamp drawing" Then
            sqlcmd = sqlcmd & " WHERE gm_name = '" & w_moji.Value & "' "
        Else
            sqlcmd = sqlcmd & " WHERE hm_name = '" & w_moji.Value & "' "
        End If

        '����
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        w_total.Text = "-1"
        Do Until Rs.EOF

            If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                num = CStr(Val(Rs.rdoColumns(0).Value))
            Else
                num = "-1"
            End If
            w_total.Text = num

            Rs.MoveNext()
        Loop

        Rs.Close()

        If w_total.Text = "-1" Then
            MsgBox("Failed to find.")
            GoTo error_section
        End If
        ' <- watanabe edit VerUP(2011)


		If CDbl(w_total.Text) = 0 Then
			Select Case w_taisho.Text
                Case "Stamp drawing"
                    MsgBox("There is no carved seal drawing corresponding.")
                Case "Editing characters drawing"
                    MsgBox("There is no editing characters drawing the appropriate.")
                Case "Brand drawing"
                    MsgBox("There is no brand drawing corresponding.")
            End Select
			GoTo error_section
		Else
			'UPGRADE_WARNING: �I�u�W�F�N�g AskNum �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            If CDbl(w_total.Text) > AskNum Then
                w_ret = MsgBox("There is " & w_total.Text & " data. Would you like to view?", MsgBoxStyle.YesNo, "Confirmation")
                If w_ret = MsgBoxResult.No Then
                    MsgBox("Canceled the search.", , "Cancel")
                    w_total.Text = ""
                    GoTo error_section
                End If
            End If
		End If
		
		'UPGRADE_WARNING: �I�u�W�F�N�g co_rockset_F_ZSEARCH() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		w_ret = co_rockset_F_ZSEARCH(2, 1)
		MSFlexGrid1.Redraw = False
		
		
		'��د�ނɌ������e�\��
		If CDbl(w_total.Text) > 0 Then
			MSFlexGrid1.Rows = Int((CDbl(w_total.Text) - 1) / 2) + 2
		Else
			MSFlexGrid1.Rows = 2
			For i = 0 To MSFlexGrid1.Cols - 1
				w_ret = Set_Grid_Data(MSFlexGrid1, "", 1, i)
			Next i
		End If
		
		MSFlexGrid1.set_RowHeight(-1, 300)
		index_row = "; NO "


		'�����R�}���h�쐬
		sqlcmd = "SELECT DISTINCT id, no1, no2"
        sqlcmd = sqlcmd & " FROM " & DBTableName2
        If w_taisho.Text = "Stamp drawing" Then
            sqlcmd = sqlcmd & " WHERE gm_name = '" & w_moji.Value & "' "
        Else
            sqlcmd = sqlcmd & " WHERE hm_name = '" & w_moji.Value & "' "
        End If
        sqlcmd = sqlcmd & " ORDER BY id, no1, no2"

        '����
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        Rs.MoveFirst()

        i = 0
        ic = 0
        ir = 0
        Do Until Rs.EOF

            System.Windows.Forms.Application.DoEvents()
            If GL_cancel_flg = 1 Then GoTo cancel_end_section

            i = i + 1
            ic = (i - 1) - Int((i - 1) / 2) * 2 + 1
            ir = Int((i - 1) / 2) + 1
            w_ret = Set_Grid_Data(MSFlexGrid1, "", ir, (ic - 1) * 3 + 1)
            w_ret = Set_Grid_Data(MSFlexGrid1, "��", ir, (ic - 1) * 3 + 2)
            For j = 1 To 3
                If IsDBNull(Rs.rdoColumns(j - 1).Value) = False Then
                    L_DAT(j) = Rs.rdoColumns(j - 1).Value
                Else
                    L_DAT(j) = ""
                End If
            Next j

            Select Case w_taisho.Text
                Case "Stamp drawing", "Editing characters drawing"
                    w_ret = Set_Grid_Data(MSFlexGrid1, L_DAT(1) & "-" & L_DAT(2) & "-" & L_DAT(3), ir, 3 + (ic - 1) * 3)
                Case "Brand drawing"
                    w_ret = Set_Grid_Data(MSFlexGrid1, L_DAT(1) & "-" & L_DAT(2) & "-" & L_DAT(3), ir, 3 + (ic - 1) * 3)
            End Select

            srh_cnt = srh_cnt + 1
            w_total.Text = CStr(srh_cnt)
			If (srh_cnt + 1) Mod 2 = 0 Then
				'----- .NET �ڍs -----
				'index_row = index_row & "|" & VB6.Format((srh_cnt + 1) / 2)
				index_row = index_row & "|" & ((srh_cnt + 1) / 2).ToString()
			End If

			Rs.MoveNext()
        Loop

        Rs.Close()
        ' <- watanabe edit VerUP(2011)


        MSFlexGrid1.FormatString = index_row
        MSFlexGrid1.set_FixedAlignment(0, 4)
        MSFlexGrid1.Redraw = True
        MSFlexGrid1.Enabled = True

        'UPGRADE_WARNING: �I�u�W�F�N�g co_rockset_F_ZSEARCH() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        w_ret = co_rockset_F_ZSEARCH(2, 0)
        'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
        form_main.Text2.Text = ""
        CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")
        MSFlexGrid1.Focus()

error_section:

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        Exit Sub

cancel_end_section:

        ' -> watanabe add VerUP(2011)
        On Error Resume Next
        Err.Clear()
        Rs.Close()
        ' <- watanabe add VerUP(2011)

        end_sql()
        w_total.Text = ""
        MSFlexGrid1.Rows = 2
        MSFlexGrid1.Cols = 7 '----- 12/11 1997 yamamoto change 5��7 -----
        For lp = 0 To MSFlexGrid1.Cols - 1
            MSFlexGrid1.Row = 1
            'UPGRADE_WARNING: �I�u�W�F�N�g lp �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
            MSFlexGrid1.Col = lp
            MSFlexGrid1.Text = ""
        Next lp
        MSFlexGrid1.Redraw = True
        'UPGRADE_WARNING: �I�u�W�F�N�g co_rockset_F_ZSEARCH() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        w_ret = co_rockset_F_ZSEARCH(2, 0)
        MsgBox("Search has been canceled.", 64, "Cancel")

    End Sub
	
	Private Sub cmd_Clear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_Clear.Click

        w_taisho.SelectedIndex = 0 '20100701�ǉ��R�[�h
		Call Clear_F_ZSEARCH_YOUSO()
		
	End Sub
	
	Private Sub cmd_End_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_End.Click
		
		form_no.Close()
		'  form_main.Show
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
                .HelpContext = 901
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub cmd_ZumenRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmd_ZumenRead.Click
		Dim w_ret As Object
		
		Dim syori_cnt As Short
		Dim w_str As String
		Dim w_mess As String
		Dim ZumenName As String
		Dim error_no As String
		Dim i, j As Short
		Dim time_start As Date
		Dim time_now As Date
		Dim w_err As String
		Dim w_name As String
        'Dim w_dir As String '20100616�ڐA�폜
        'Dim w_id As String '20100616�ڐA�폜


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w_str = ""
        w_name = ""
        ' <- watanabe add VerUP(2011)


		syori_cnt = 0
		
		For i = 1 To MSFlexGrid1.Rows - 1
			For j = 1 To 2
				If syori_cnt >= FreePicNum Then
                    MsgBox("There are no free pictures." & Chr(13) & "Number of empty pictures=" & FreePicNum, MsgBoxStyle.Critical, "CAD reading error")
					Exit Sub
				End If
				'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				w_ret = Get_Grid_Data(MSFlexGrid1, w_err, i, 3 * (j - 1) + 1)
				'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 3 * (j - 1) + 2)
				'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
				w_ret = Get_Grid_Data(MSFlexGrid1, w_name, i, 3 * (j - 1) + 3)
				If w_str = "��" And w_err = "" Then
					ZumenName = w_name
					If VB.Left(Trim(ZumenName), 2) = "KO" Then
						'UPGRADE_WARNING: �I�u�W�F�N�g KokuinDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						w_mess = KokuinDir & ZumenName
					ElseIf VB.Left(Trim(ZumenName), 2) = "HE" Then 
						'UPGRADE_WARNING: �I�u�W�F�N�g HensyuZumenDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						w_mess = HensyuZumenDir & ZumenName
					ElseIf VB.Left(Trim(ZumenName), 4) = "AT-B" Then 
						'UPGRADE_WARNING: �I�u�W�F�N�g BrandDir �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						w_mess = BrandDir & ZumenName
					Else
                        MsgBox("Input error.")
						'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						form_main.Text2.Text = ""
						Exit Sub
					End If
					'             w_mess = ZumenName
					'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
					w_ret = PokeACAD("MDLREAD", w_mess)
					'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
					w_ret = RequestACAD("MDLREAD")
					
					time_start = Now
					Do 
						time_now = Now
						'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						If Trim(form_main.Text2.Text) = "" Then
							If System.Date.FromOADate(time_now.ToOADate - time_start.ToOADate) > System.Date.FromOADate(timeOutSecond) Then
                                MsgBox("Time-out error.", 64, "ERROR")
								'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
                                w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
								'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								w_ret = RequestACAD("ERROR")
								Exit Sub
							End If
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						ElseIf VB.Left(Trim(form_main.Text2.Text), 7) = "OK-DATA" Then 
							'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							w_ret = Set_Grid_Data(MSFlexGrid1, "0", i, 3 * (j - 1) + 1)
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
							Exit Do
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
						ElseIf VB.Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then 
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
							'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							w_ret = Set_Grid_Data(MSFlexGrid1, error_no, i, 3 * (j - 1) + 1)
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
							Exit Do
						Else
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                            MsgBox("Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
							'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
							w_ret = Set_Grid_Data(MSFlexGrid1, "?", i, 3 * (j - 1) + 1)
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
							Exit Sub
						End If
					Loop 
					'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
					w_ret = Get_Grid_Data(MSFlexGrid1, w_str, i, 3 * (j - 1) + 1)
					'�}�ʓǂݍ��݂n�j
					If Val(w_str) = 0 Then
						cmd_Search.Enabled = False
						cmd_Clear.Enabled = False
						cmd_Help.Enabled = False
						cmd_ZumenRead.Enabled = False
						w_mojicd.Enabled = False
						w_mojicd.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						w_taisho.Enabled = False
						w_taisho.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
						MSFlexGrid1.Enabled = False
					End If
					syori_cnt = syori_cnt + 1
				End If
			Next j
		Next i
		
		If syori_cnt = 0 Then
            MsgBox("Data to be read is not selected.")
		Else
			''''     MsgBox "CAD�Ǎ��݊���"
		End If
		
	End Sub
	
	'�T�v�F�t�H�[�����[�h
	'�����F�ǉ����ځF���b�N�Z�b�g
	'----- 1/28 1997 by yamamoto -----
	Private Sub F_ZSEARCH_YOUSO_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_ret As Short
		Dim index_col As String
		
        form_no = Me
        temp_gz.Initilize() '20100702�ǉ��R�[�h
        temp_hz.Initilize() '20100702�ǉ��R�[�h
        temp_bz.Initilize() '20100702�ǉ��R�[�h
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B
		
		Call Clear_F_ZSEARCH_YOUSO()
		
		MSFlexGrid1.Rows = 2
		MSFlexGrid1.Cols = 7
		
		' �s�����̐ݒ�
		MSFlexGrid1.set_RowHeight(-1, 300)
		
        index_col = "^NO|^error|^Read|^Drawing name|^error|^Read|^Drawing name"
		MSFlexGrid1.FormatString = index_col
		
		' �񕝂̐ݒ�
		MSFlexGrid1.set_ColWidth(0, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(1, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(2, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(3, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 4)
		MSFlexGrid1.set_ColWidth(4, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(5, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 1)
		MSFlexGrid1.set_ColWidth(6, (VB6.PixelsToTwipsX(MSFlexGrid1.Width) - 300) / 13 * 4)
		
		MSFlexGrid1.Enabled = False
		
		'�����ޯ��
		w_taisho.Items.Clear()
        w_taisho.Items.Add("Stamp drawing")
        w_taisho.Items.Add("Editing characters drawing")
        w_taisho.Items.Add("Brand drawing")
		w_taisho.SelectedIndex = 0
		
		'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
		form_main.Text2.Text = ""
		CommunicateMode = comFreePic
		w_ret = RequestACAD("PICEMPTY")
		
		'UPGRADE_WARNING: �I�u�W�F�N�g co_rockset_F_ZSEARCH() �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		w_ret = co_rockset_F_ZSEARCH(2, 0)
		
	End Sub

	'----- .NET�ڍs (ToDo:DataGridView�̃C�x���g�ɕύX) -----
#If False Then
	Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		Dim w_ret As Object
		
		Dim w_col As Short
		Dim w_row As Short
		Dim w_err As String
		Dim w1 As String
		Dim w2 As String
		Dim i As Short


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w1 = ""
        w2 = ""
        ' <- watanabe add VerUP(2011)


		If eventArgs.KeyAscii <> 32 Then Exit Sub
		
		w_col = Val(CStr(MSFlexGrid1.Col))
		w_row = Val(CStr(MSFlexGrid1.Row))
		
		MSFlexGrid1.Redraw = False
		
		If w_col = 2 Or w_col = 5 Then
			'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			w_ret = Get_Grid_Data(MSFlexGrid1, w_err, w_row, w_col - 1)
			If w_err = "" Then
				If MSFlexGrid1.Text = "��" Then
					MSFlexGrid1.Text = "��"
				ElseIf MSFlexGrid1.Text = "��" Then 
					For i = 1 To MSFlexGrid1.Rows - 1
						'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
						w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
						If w1 = "" And w2 = "��" Then
							w_ret = Set_Grid_Data(MSFlexGrid1, "��", i, 2)
						End If
						w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 4)
						w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 5)
						If w1 = "" And w2 = "��" Then
							w_ret = Set_Grid_Data(MSFlexGrid1, "��", i, 5)
						End If
					Next i
					MSFlexGrid1.Text = "��"
				End If
			End If
		End If
		
		MSFlexGrid1.Redraw = True
		
	End Sub
#End If

	'----- .NET�ڍs (ToDo:DataGridView�̃C�x���g�ɕύX) -----
#If False Then
	Private Sub MSFlexGrid1_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseDownEvent) Handles MSFlexGrid1.MouseDownEvent
		Dim w_ret As Object
		
		Dim w_col As Short
		Dim w_row As Short
		Dim w_err As String
		Dim w1 As String
		Dim w2 As String
		Dim i As Short


        ' -> watanabe add VerUP(2011)
        w_err = ""
        w1 = ""
        w2 = ""
        ' <- watanabe add VerUP(2011)


		w_col = Val(CStr(MSFlexGrid1.Col))
		w_row = Val(CStr(MSFlexGrid1.Row))
		
		MSFlexGrid1.Redraw = False
		
		If w_col = 2 Or w_col = 5 Then
			'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
			w_ret = Get_Grid_Data(MSFlexGrid1, w_err, w_row, w_col - 1)
			If w_err = "" Then
				If MSFlexGrid1.Text = "��" Then
					MSFlexGrid1.Text = "��"
				ElseIf MSFlexGrid1.Text = "��" Then 
					For i = 1 To MSFlexGrid1.Rows - 1
						'UPGRADE_WARNING: �I�u�W�F�N�g w_ret �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
						w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 1)
						w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 2)
						If w1 = "" And w2 = "��" Then
                            w_ret = Set_Grid_Data(MSFlexGrid1, "��", i, 2)
						End If
						w_ret = Get_Grid_Data(MSFlexGrid1, w1, i, 4)
						w_ret = Get_Grid_Data(MSFlexGrid1, w2, i, 5)
						If w1 = "" And w2 = "��" Then
							w_ret = Set_Grid_Data(MSFlexGrid1, "��", i, 5)
						End If
					Next i
					MSFlexGrid1.Text = "��"
				End If
			End If
		End If
		
		MSFlexGrid1.Redraw = True
		
	End Sub
#End If

	Private Sub w_mojicd_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_mojicd.Leave
		
		'UPGRADE_ISSUE: Control w_mojicd �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
		form_no.w_mojicd.Text = UCase(Trim(form_no.w_mojicd.Text))
		
	End Sub
	
	Private Sub w_taisho_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_taisho.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		If KeyAscii = 13 Then GoTo EventExitSub
		Call Combo_Sousa(w_taisho, KeyAscii)
		KeyAscii = 0
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
End Class