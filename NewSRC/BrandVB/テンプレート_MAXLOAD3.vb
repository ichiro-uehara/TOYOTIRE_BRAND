Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class F_TMP_MAXLOAD3
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim result As Integer
		Dim a As String
		Dim b As String
		Dim c As String
		Dim d As String
		Dim e As String
		Dim w_ret As Short

        ' -> watanabe add VerUP(2011)
        Dim sqlcmd As String
        Dim Rs As RDO.rdoResultset
        ' <- watanabe add VerUP(2011)

		On Error Resume Next
        Err.Clear()
		
		'/* ���̓`�F�b�N */
		If check_F_TMP_MAXLOAD(1) <> 0 Then
			Exit Sub
		Else
			'// SQL ���� ����
			init_sql()
			

            ' -> watanabe edit VerUP(2011)
            '         If w_kikaku.Text = "JATMA" Then
            '             result = sqlcmd(SqlConn, "SELECT standard_load_index, ")
            '             result = sqlcmd(SqlConn, " standard_max_load_kg , standard_max_load_lbs, ")
            '             result = sqlcmd(SqlConn, " standard_max_press_kpa, standard_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..jatma")
            '
            '         ElseIf w_kikaku.Text = "TRA(�y��)" Then
            '             result = sqlcmd(SqlConn, "SELECT light_load_index, ")
            '             result = sqlcmd(SqlConn, " light_max_load_kg, light_max_load_lbs,")
            '             result = sqlcmd(SqlConn, " light_max_press_kpa, light_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..tra")
            '
            '         ElseIf w_kikaku.Text = "TRA(�W��)" Then
            '             result = sqlcmd(SqlConn, "SELECT standard_load_index, ")
            '             result = sqlcmd(SqlConn, " standard_max_load_kg, standard_max_load_lbs,")
            '             result = sqlcmd(SqlConn, " standard_max_press_kpa, standard_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..tra")
            '
            '         ElseIf w_kikaku.Text = "TRA(����)" Then
            '             result = sqlcmd(SqlConn, "SELECT extra_load_index, ")
            '             result = sqlcmd(SqlConn, " extra_max_load_kg, extra_max_load_lbs,")
            '             result = sqlcmd(SqlConn, " extra_max_press_kpa, extra_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..tra")
            '
            '         ElseIf w_kikaku.Text = "ETRTO(�W��)" Then
            '             result = sqlcmd(SqlConn, "SELECT standard_load_index, ")
            '             result = sqlcmd(SqlConn, " standard_max_load_kg, standard_max_load_lbs,")
            '             result = sqlcmd(SqlConn, " standard_max_press_kpa, standard_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..etrto")
            '
            '         ElseIf w_kikaku.Text = "ETRTO(����)" Then
            '             result = sqlcmd(SqlConn, "SELECT extra_load_index, ")
            '             result = sqlcmd(SqlConn, " extra_max_load_kg, extra_max_load_lbs,")
            '             result = sqlcmd(SqlConn, " extra_max_press_kpa, extra_max_press_psi")
            '             result = sqlcmd(SqlConn, " FROM " & STANDARD_DBName & "..etrto")
            '         End If
            '
            '         result = SqlCmd(SqlConn, " WHERE ( syurui = '" & Trim(form_no.w_syurui.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size1 = '" & Trim(form_no.w_size1.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size2 = '" & Trim(form_no.w_size2.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size3 = '" & Trim(form_no.w_size3.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size4 = '" & Trim(form_no.w_size4.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size5 = '" & Trim(form_no.w_size5.Text) & "' AND")
            'result = SqlCmd(SqlConn, " size6 = '" & Trim(form_no.w_size6.Text) & "')")
            'result = SqlExec(SqlConn)
            'result = SqlResults(SqlConn)
            'If result = SUCCEED Then
            '	If SqlNextRow(SqlConn) = REGROW Then
            '		a = SqlData(SqlConn, 1)
            '		b = SqlData(SqlConn, 2)
            '		c = SqlData(SqlConn, 3)
            '		d = SqlData(SqlConn, 4)
            '		e = SqlData(SqlConn, 5)
            '
            '		If a = "" Then
            '			MsgBox("�Y������K�i�l������܂���", MsgBoxStyle.Critical, "DATA NOT FOUND")
            '			GoTo end_section
            '		End If
            '
            '		form_no.w_kajyu.Text = a
            '		form_no.w_kikaku_max_load_kg.Text = b
            '		form_no.w_kikaku_max_load_lbs.Text = c
            '		form_no.w_kikaku_max_press_kpa.Text = d
            '		form_no.w_kikaku_max_press_psi.Text = e
            '
            '		If form_no.w_max_load_kg.Enabled = True Then
            '			form_no.w_max_load_kg.Text = b
            '		End If
            '		If form_no.w_max_load_lbs.Enabled = True Then
            '			form_no.w_max_load_lbs.Text = c
            '		End If
            '		If form_no.w_max_press_kpa.Enabled = True Then
            '			form_no.w_max_press_kpa.Text = "300"
            '		End If
            '		If form_no.w_max_press_psi.Enabled = True Then
            '			form_no.w_max_press_psi.Text = "44"
            '		End If
            '
            '		CommunicateMode = comFreePic
            '		w_ret = RequestACAD("PICEMPTY")
            '	Else
            '		MsgBox("�Y������^�C���T�C�Y��������܂���", MsgBoxStyle.Critical, "DATA NOT FOUND")
            '		GoTo end_section
            '	End If
            'Else
            '	MsgBox("�ް��ް�SELECT�װ", MsgBoxStyle.Critical)
            '	GoTo end_section
            'End If
            '
            'end_section:


            '�����R�}���h�쐬
            sqlcmd = ""
            If w_kikaku.Text = "JATMA" Then
                sqlcmd = sqlcmd & "SELECT standard_load_index, "
                sqlcmd = sqlcmd & " standard_max_load_kg , standard_max_load_lbs, "
                sqlcmd = sqlcmd & " standard_max_press_kpa, standard_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..jatma"

            ElseIf w_kikaku.Text = "TRA (lightweight)" Then
                sqlcmd = sqlcmd & "SELECT light_load_index, "
                sqlcmd = sqlcmd & " light_max_load_kg, light_max_load_lbs,"
                sqlcmd = sqlcmd & " light_max_press_kpa, light_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..tra"

            ElseIf w_kikaku.Text = "TRA (standard)" Then
                sqlcmd = sqlcmd & "SELECT standard_load_index, "
                sqlcmd = sqlcmd & " standard_max_load_kg, standard_max_load_lbs,"
                sqlcmd = sqlcmd & " standard_max_press_kpa, standard_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..tra"

            ElseIf w_kikaku.Text = "TRA (special)" Then
                sqlcmd = sqlcmd & "SELECT extra_load_index, "
                sqlcmd = sqlcmd & " extra_max_load_kg, extra_max_load_lbs,"
                sqlcmd = sqlcmd & " extra_max_press_kpa, extra_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..tra"

            ElseIf w_kikaku.Text = "ETRTO(Standard)" Then
                sqlcmd = sqlcmd & "SELECT standard_load_index, "
                sqlcmd = sqlcmd & " standard_max_load_kg, standard_max_load_lbs,"
                sqlcmd = sqlcmd & " standard_max_press_kpa, standard_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..etrto"

            ElseIf w_kikaku.Text = "ETRTO(Special)" Then
                sqlcmd = sqlcmd & "SELECT extra_load_index, "
                sqlcmd = sqlcmd & " extra_max_load_kg, extra_max_load_lbs,"
                sqlcmd = sqlcmd & " extra_max_press_kpa, extra_max_press_psi"
                sqlcmd = sqlcmd & " FROM " & STANDARD_DBName & "..etrto"
            End If

            sqlcmd = sqlcmd & " WHERE ( syurui = '" & Trim(form_no.w_syurui.Text) & "' AND"
            sqlcmd = sqlcmd & " size1 = '" & Trim(form_no.w_size1.Text) & "' AND"
            sqlcmd = sqlcmd & " size2 = '" & Trim(form_no.w_size2.Text) & "' AND"
            sqlcmd = sqlcmd & " size3 = '" & Trim(form_no.w_size3.Text) & "' AND"
            sqlcmd = sqlcmd & " size4 = '" & Trim(form_no.w_size4.Text) & "' AND"
            sqlcmd = sqlcmd & " size5 = '" & Trim(form_no.w_size5.Text) & "' AND"
            sqlcmd = sqlcmd & " size6 = '" & Trim(form_no.w_size6.Text) & "')"

            '����
            On Error GoTo error_section
            Err.Clear()
            Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            On Error Resume Next
            Err.Clear()

            Rs.MoveFirst()

            If GL_T_RDO.Con.RowsAffected() > 0 Then

                If IsDBNull(Rs.rdoColumns(0).Value) = False Then
                    a = Rs.rdoColumns(0).Value
                Else
                    a = ""
                End If

                If IsDBNull(Rs.rdoColumns(1).Value) = False Then
                    b = Rs.rdoColumns(1).Value
                Else
                    b = ""
                End If

                If IsDBNull(Rs.rdoColumns(2).Value) = False Then
                    c = Rs.rdoColumns(2).Value
                Else
                    c = ""
                End If

                If IsDBNull(Rs.rdoColumns(3).Value) = False Then
                    d = Rs.rdoColumns(3).Value
                Else
                    d = ""
                End If

                If IsDBNull(Rs.rdoColumns(4).Value) = False Then
                    e = Rs.rdoColumns(4).Value
                Else
                    e = ""
                End If

                If a = "" Then
                    MsgBox("There is no standard value corresponding.", MsgBoxStyle.Critical, "DATA NOT FOUND")
                    GoTo end_section
                End If

                form_no.w_kajyu.Text = a
                form_no.w_kikaku_max_load_kg.Text = b
                form_no.w_kikaku_max_load_lbs.Text = c
                form_no.w_kikaku_max_press_kpa.Text = d
                form_no.w_kikaku_max_press_psi.Text = e

                If form_no.w_max_load_kg.Enabled = True Then
                    form_no.w_max_load_kg.Text = b
                End If
                If form_no.w_max_load_lbs.Enabled = True Then
                    form_no.w_max_load_lbs.Text = c
                End If
                If form_no.w_max_press_kpa.Enabled = True Then
                    form_no.w_max_press_kpa.Text = "300"
                End If
                If form_no.w_max_press_psi.Enabled = True Then
                    form_no.w_max_press_psi.Text = "44"
                End If

                CommunicateMode = comFreePic
                w_ret = RequestACAD("PICEMPTY")

            Else
                MsgBox("There is no tire size corresponding.", MsgBoxStyle.Critical, "DATA NOT FOUND")
                GoTo end_section
            End If

end_section:

            Rs.Close()
            ' <- watanabe edit VerUP(2011)

            end_sql()
        End If


        ' -> watanabe add VerUP(2011)
        Exit Sub

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        Rs.Close()
        end_sql()
        ' <- watanabe add VerUP(2011)

    End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		Call Clear_F_TMP_MAXLOAD3(0)
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        InitFlag = False '20100628�ǉ��R�[�h
        form_no.Close()
		End
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
        On Error Resume Next
        Err.Clear()
        Dim oCommonDialog As Object
        oCommonDialog = CreateObject("MSComDlg.CommonDialog")

        If Err.Number = 0 Then
            With oCommonDialog
                .HelpCommand = cdlHelpContext
                .HelpFile = "c:\VBhelp\BRAND.HLP"
                .HelpContext = 806
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		Dim w_mess As String
		Dim w_cmd As String
		Dim w_str As String
		Dim w_ret As Short
		Dim pic_no As Short
		Dim gm_no As Short
		Dim gm_alph As String
		Dim type_no As Short
		Dim cmd_no As Short
		
		Dim key_value As String
		Dim tmp_str As String
		
		Dim grp_num As Short
		Dim top_dumy_num As Short
		Dim top_hmcode As String
		
		Dim grp_datum_no() As Short
		Dim grp_dist_x() As Double
		Dim grp_dist_y() As Double
		Dim grp_dumy_num() As Short
		Dim grp_hmcode() As String
		
		Dim ZumenName As String
		Dim change_num As Short
		Dim sub_num As Short
		
		Dim hexdata As String
		Dim str_dbl As New VB6.FixedLengthString(16)
		Dim str_int As New VB6.FixedLengthString(8)
		
		Dim error_no As String
		Dim i As Short
		Dim j As Short


        ' -> watanabe add VerUP(2011)
        w_mess = ""
        ' <- watanabe add VerUP(2011)

		
		'/* ���̓`�F�b�N */
		If check_F_TMP_MAXLOAD(2) <> 0 Then
			Exit Sub
		End If
		
		'** ��ʏ��𑗐M **
		Call temp_bz_get(2)
		Call bz_spec_set(w_mess)
		w_ret = PokeACAD("SPECADD", w_mess)
		w_ret = RequestACAD("SPECADD")
		System.Windows.Forms.Application.DoEvents()
		
		init_sql()
		form_no.Enabled = False
		F_MSG.Show()
		
		type_no = 0
		For i = 1 To MaxSelNum
			If (w_type.Text = Tmp_hm_word(i)) Then
				If (Tmp_prcs_code(i) = "MAXLOAD1") Then
					type_no = 1
				ElseIf (Tmp_prcs_code(i) = "MAXLOAD2") Then 
					type_no = 2
				ElseIf (Tmp_prcs_code(i) = "MAXLOAD3") Then 
					type_no = 3
				End If
			End If
		Next i
		
		If type_no = 0 Then
            MsgBox("The designation of the type is wrong.", 64, "Configuration file error")
			GoTo error_section
		End If
		
		If chk_max_load_kg.CheckState = 0 Then
			If type_no = 1 Or type_no = 2 Then
				For i = 1 To Len(form_no.w_max_load_kg.Text)
					w_str = Mid(form_no.w_max_load_kg.Text, i, 1)
					If IsNumeric(w_str) Then
						If Val(w_str) >= 0 And Val(w_str) < 10 Then
							If GensiNUM(Val(w_str)) = "" Then
                                MsgBox("A substituted primitive letter for input KG is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
								GoTo error_section
							End If
						End If
					ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
						If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                            MsgBox("A substituted primitive letter for input kg is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				Next i
			End If
		End If
		
		If chk_max_load_lbs.CheckState = 0 Then
			If type_no = 1 Or type_no = 3 Then
				For i = 1 To Len(form_no.w_max_load_lbs.Text)
					w_str = Mid(form_no.w_max_load_lbs.Text, i, 1)
					If IsNumeric(w_str) Then
						If Val(w_str) >= 0 And Val(w_str) < 10 Then
							If GensiNUM(Val(w_str)) = "" Then
                                MsgBox("A substituted primitive letter for input LBS is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
								GoTo error_section
							End If
						End If
					ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
						If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                            MsgBox("A substituted primitive letter for input LBS is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				Next i
			End If
		End If
		
		If chk_max_press_kpa.CheckState = 0 Then
			If type_no = 1 Or type_no = 2 Then
				For i = 1 To Len(form_no.w_max_press_kpa.Text)
					w_str = Mid(form_no.w_max_press_kpa.Text, i, 1)
					If IsNumeric(w_str) Then
						If Val(w_str) >= 0 And Val(w_str) < 10 Then
							If GensiNUM(Val(w_str)) = "" Then
                                MsgBox("A substituted primitive letter for input KPA is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
								GoTo error_section
							End If
						End If
					ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
						If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                            MsgBox("A substituted primitive letter for input kg is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				Next i
			End If
		End If
		
		If chk_max_press_psi.CheckState = 0 Then
			If type_no = 1 Or type_no = 3 Then
				For i = 1 To Len(form_no.w_max_press_psi.Text)
					w_str = Mid(form_no.w_max_press_psi.Text, i, 1)
					If IsNumeric(w_str) Then
						If Val(w_str) >= 0 And Val(w_str) < 10 Then
							If GensiNUM(Val(w_str)) = "" Then
                                MsgBox("A substituted primitive letter for input PSI is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
								GoTo error_section
							End If
						End If
					ElseIf Asc("A") <= Asc(w_str) And Asc(w_str) <= Asc("Z") Then 
						If GensiALPH(Asc(w_str) - Asc("A")) = "" Then
                            MsgBox("A substituted primitive letter for input PSI is not set to the configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
							GoTo error_section
						End If
					End If
				Next i
			End If
		End If
		
		If FreePicNum < 2 Then
            MsgBox("The number of pictures is not enough." & Chr(13) & "Number of empty pictures =" & FreePicNum)
			GoTo error_section
		End If
		
		
		' �O���[�v�f�[�^�̕����擾
		key_value = Tmp_hm_group(w_type.SelectedIndex + 1)
		
        ' �O���[�v��
		tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
		If IsNumeric(tmp_str) = False Then
            MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
			GoTo error_section
		End If
		grp_num = CShort(tmp_str)
		key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		
        ' �擪�_�~�[����
		tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
		If IsNumeric(tmp_str) = False Then
            MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
			GoTo error_section
		End If
		top_dumy_num = CShort(tmp_str)
		key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		
        ' �擪�ҏW�����R�[�h�擾
		If grp_num = 1 Then
			top_hmcode = Trim(key_value)
		Else
			If InStr(key_value, "|") = 0 Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			top_hmcode = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
		End If
		
		
        ' �ǉ��ҏW�����f�[�^�擾
		ReDim grp_datum_no(grp_num)
		ReDim grp_dist_x(grp_num)
		ReDim grp_dist_y(grp_num)
		ReDim grp_dumy_num(grp_num)
		ReDim grp_hmcode(grp_num)
		For i = 0 To grp_num - 2
            ' ��s
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_datum_no(i) = CShort(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' ����X
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dist_x(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' ����Y
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dist_y(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' �_�~�[����
			tmp_str = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
			If IsNumeric(tmp_str) = False Then
                MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
				GoTo error_section
			End If
			grp_dumy_num(i) = CDbl(tmp_str)
			key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			
            ' �擪�ҏW�����R�[�h�擾
			If i = (grp_num - 2) Then
				grp_hmcode(i) = Trim(key_value)
			Else
				If InStr(key_value, "|") = 0 Then
                    MsgBox("Configuration file (" & Tmp_Utqg3_ini & ") error" & Chr(13) & "Setting of the selected type is incorrect.")
					GoTo error_section
				End If
				grp_hmcode(i) = Trim(VB.Left(key_value, InStr(key_value, "|") - 1))
				key_value = Trim(Mid(key_value, InStr(key_value, "|") + 1))
			End If
			
			'MsgBox CStr(i + 2) & Chr(13) & grp_datum_no(i) & Chr(13) & grp_dist_x(i) & Chr(13) & grp_dist_y(i) & Chr(13) & grp_dumy_num(i) & Chr(13) & grp_hmcode(i)
			
		Next i
		
		' �擪�ҏW�����쐬
		change_num = 0
		
		'// �u�����[�h�̑��M
		w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
		w_ret = RequestACAD("CHNGMODE")
		
		'// �ҏW�������M
		pic_no = what_pic_from_hmcode(top_hmcode)
		If pic_no < 1 Then GoTo error_section
		ZumenName = "HM-" & VB.Left(Trim(top_hmcode), 6)

        '----- .NET �ڍs -----
        'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
        w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("HMCODE", w_mess)
		
		
		'// ���n�������M
		cmd_no = 1
		
		'[[ KG ]]
		If top_dumy_num > change_num Then
			If type_no = 1 Or type_no = 2 Then
				If chk_max_load_kg.CheckState = 0 Then
					For i = 1 To Len(form_no.w_max_load_kg.Text)
						gm_no = Val(Mid(form_no.w_max_load_kg.Text, i, 1))
						pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
						If pic_no < 1 Then GoTo error_section
						ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                        '----- .NET �ڍs -----
                        'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                        w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                        w_cmd = "GMCODE" & cmd_no
						w_ret = PokeACAD(w_cmd, w_mess)
					Next i
				Else
					w_mess = ""
					w_cmd = "HOLDGM" & cmd_no
					w_ret = PokeACAD(w_cmd, w_mess)
				End If
				change_num = change_num + 1
				cmd_no = cmd_no + 1
			End If
		End If
		
		'[[ LBS ]]
		If top_dumy_num > change_num Then
			If type_no = 1 Or type_no = 3 Then
				If chk_max_load_lbs.CheckState = 0 Then
					For i = 1 To Len(form_no.w_max_load_lbs.Text)
						gm_no = Val(Mid(form_no.w_max_load_lbs.Text, i, 1))
						pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
						If pic_no < 1 Then GoTo error_section
						ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                        '----- .NET �ڍs -----
                        'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                        w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                        w_cmd = "GMCODE" & cmd_no
						w_ret = PokeACAD(w_cmd, w_mess)
					Next i
				Else
					w_mess = ""
					w_cmd = "HOLDGM" & cmd_no
					w_ret = PokeACAD(w_cmd, w_mess)
				End If
				change_num = change_num + 1
				cmd_no = cmd_no + 1
			End If
		End If
		
		'[[ KPA ]]
		If top_dumy_num > change_num Then
			If type_no = 1 Or type_no = 2 Then
				If chk_max_press_kpa.CheckState = 0 Then
					For i = 1 To Len(form_no.w_max_press_kpa.Text)
						gm_no = Val(Mid(form_no.w_max_press_kpa.Text, i, 1))
						pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
						If pic_no < 1 Then GoTo error_section
						ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                        '----- .NET �ڍs -----
                        'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                        w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                        w_cmd = "GMCODE" & cmd_no
						w_ret = PokeACAD(w_cmd, w_mess)
					Next i
				Else
					w_mess = ""
					w_cmd = "HOLDGM" & cmd_no
					w_ret = PokeACAD(w_cmd, w_mess)
				End If
				change_num = change_num + 1
				cmd_no = cmd_no + 1
			End If
		End If
		
		'[[ PSI ]]
		If top_dumy_num > change_num Then
			If type_no = 1 Or type_no = 3 Then
				If chk_max_press_psi.CheckState = 0 Then
					For i = 1 To Len(form_no.w_max_press_psi.Text)
						gm_no = Val(Mid(form_no.w_max_press_psi.Text, i, 1))
						pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
						If pic_no < 1 Then GoTo error_section
						ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                        '----- .NET �ڍs -----
                        'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                        w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                        w_cmd = "GMCODE" & cmd_no
						w_ret = PokeACAD(w_cmd, w_mess)
					Next i
				Else
					w_mess = ""
					w_cmd = "HOLDGM" & cmd_no
					w_ret = PokeACAD(w_cmd, w_mess)
				End If
				change_num = change_num + 1
				cmd_no = cmd_no + 1
			End If
		End If


        ' -> watanabe add VerUP(2011)
        CommunicateMode = comTmpWait
        ' <- watanabe add VerUP(2011)

        '// �I���̑��M
		w_ret = RequestACAD("TMPCHANG")
		
        ' CAD�����I���`�F�b�N
		If check_cad_run = False Then
			GoTo error_section
		End If
		
        '// ��}���s�o�h�b�ێ��̑��M
		w_ret = RequestACAD("TMPTOPPIC")
		
        ' CAD�����I���`�F�b�N
		If check_cad_run = False Then
			GoTo error_section
		End If

        ' -> watanabe add VerUP(2011)
        CommunicateMode = comNone
        ' <- watanabe add VerUP(2011)

		
        ' �O���[�v�������[�v
		For j = 0 To grp_num - 2

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

            ' �O��f�[�^�N���A
			w_ret = RequestACAD("TMPDATCLR")
			
            ' CAD�����I���`�F�b�N
			If check_cad_run = False Then
				GoTo error_section
			End If

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)

			
            ' �O���[�v�ҏW�����쐬
			sub_num = 0
			
			'// �u�����[�h�̑��M
			w_ret = PokeACAD("CHNGMODE", VB.Left(Trim(ReplaceMode), 1))
			w_ret = RequestACAD("CHNGMODE")
			
			'�i�}�ʖ��j���M
			pic_no = what_pic_from_hmcode(grp_hmcode(j))
			If pic_no < 1 Then GoTo error_section
			ZumenName = "HM-" & VB.Left(Trim(grp_hmcode(j)), 6)

            '----- .NET �ڍs -----
            'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & HensyuDir & ZumenName
            w_mess = Val(CStr(pic_no)).ToString("000") & HensyuDir & ZumenName

            w_ret = PokeACAD("HMCODE", w_mess)
			
			
			'[[ KG ]]
			If grp_dumy_num(j) > sub_num Then
				If (type_no = 1 And change_num = 0) Or (type_no = 2 And change_num = 0) Then
					If chk_max_load_kg.CheckState = 0 Then
						For i = 1 To Len(form_no.w_max_load_kg.Text)
							gm_no = Val(Mid(form_no.w_max_load_kg.Text, i, 1))
							pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
							If pic_no < 1 Then GoTo error_section
							ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                            '----- .NET �ڍs -----
                            'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                            w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                            w_cmd = "GMCODE" & (sub_num + 1)
							w_ret = PokeACAD(w_cmd, w_mess)
						Next i
					Else
						w_mess = ""
						w_cmd = "HOLDGM" & (sub_num + 1)
						w_ret = PokeACAD(w_cmd, w_mess)
					End If
					change_num = change_num + 1
					sub_num = sub_num + 1
				End If
			End If
			
			'[[ LBS ]]
			If grp_dumy_num(j) > sub_num Then
				If (type_no = 1 And change_num = 1) Or (type_no = 3 And change_num = 0) Then
					If chk_max_load_lbs.CheckState = 0 Then
						For i = 1 To Len(form_no.w_max_load_lbs.Text)
							gm_no = Val(Mid(form_no.w_max_load_lbs.Text, i, 1))
							pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
							If pic_no < 1 Then GoTo error_section
							ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                            '----- .NET �ڍs -----
                            'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                            w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                            w_cmd = "GMCODE" & (sub_num + 1)
							w_ret = PokeACAD(w_cmd, w_mess)
						Next i
					Else
						w_mess = ""
						w_cmd = "HOLDGM" & (sub_num + 1)
						w_ret = PokeACAD(w_cmd, w_mess)
					End If
					change_num = change_num + 1
					sub_num = sub_num + 1
				End If
			End If
			
			'[[ KPA ]]
			If grp_dumy_num(j) > sub_num Then
				If (type_no = 1 And change_num = 2) Or (type_no = 2 And change_num = 1) Then
					If chk_max_press_kpa.CheckState = 0 Then
						For i = 1 To Len(form_no.w_max_press_kpa.Text)
							gm_no = Val(Mid(form_no.w_max_press_kpa.Text, i, 1))
							pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
							If pic_no < 1 Then GoTo error_section
							ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                            '----- .NET �ڍs -----
                            'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                            w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                            w_cmd = "GMCODE" & (sub_num + 1)
							w_ret = PokeACAD(w_cmd, w_mess)
						Next i
					Else
						w_mess = ""
						w_cmd = "HOLDGM" & (sub_num + 1)
						w_ret = PokeACAD(w_cmd, w_mess)
					End If
					change_num = change_num + 1
					sub_num = sub_num + 1
				End If
			End If
			
			'[[ PSI ]]
			If grp_dumy_num(j) > sub_num Then
				If (type_no = 1 And change_num = 3) Or (type_no = 3 And change_num = 1) Then
					If chk_max_press_psi.CheckState = 0 Then
						For i = 1 To Len(form_no.w_max_press_psi.Text)
							gm_no = Val(Mid(form_no.w_max_press_psi.Text, i, 1))
							pic_no = what_pic_from_gmcode(GensiNUM(gm_no))
							If pic_no < 1 Then GoTo error_section
							ZumenName = "GM-" & Mid(GensiNUM(gm_no), 1, 6)

                            '----- .NET �ڍs -----
                            'w_mess = VB6.Format(Val(CStr(pic_no)), "000") & GensiDir & ZumenName
                            w_mess = Val(CStr(pic_no)).ToString("000") & GensiDir & ZumenName

                            w_cmd = "GMCODE" & (sub_num + 1)
							w_ret = PokeACAD(w_cmd, w_mess)
						Next i
					Else
						w_mess = ""
						w_cmd = "HOLDGM" & (sub_num + 1)
						w_ret = PokeACAD(w_cmd, w_mess)
					End If
					change_num = change_num + 1
					sub_num = sub_num + 1
				End If
			End If


            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

			'// �I���̑��M
			w_ret = RequestACAD("TMPCHANG")
			
            ' CAD�����I���`�F�b�N
			If check_cad_run = False Then
				GoTo error_section
			End If

			'// ��}���s�o�h�b�ێ��̑��M
			w_ret = RequestACAD("TMPADDPIC")
			
            ' CAD�����I���`�F�b�N
			If check_cad_run = False Then
				GoTo error_section
			End If
			
            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)


            ' �O���[�v��
			hexdata = ""
			w_ret = InttoHex(grp_datum_no(j), str_int.Value)
			hexdata = hexdata & str_int.Value
			
			w_ret = DbltoHex(grp_dist_x(j), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value
			
			w_ret = DbltoHex(grp_dist_y(j), str_dbl.Value)
			hexdata = hexdata & str_dbl.Value


            ' -> watanabe add VerUP(2011)
            CommunicateMode = comTmpWait
            ' <- watanabe add VerUP(2011)

			w_ret = PokeACAD("TMPGRPDAT", hexdata)
			w_ret = RequestACAD("TMPGRPADD")
			
            ' CAD�����I���`�F�b�N
			If check_cad_run = False Then
				GoTo error_section
			End If

            ' -> watanabe add VerUP(2011)
            CommunicateMode = comNone
            ' <- watanabe add VerUP(2011)

        Next j
		
		' VB�I��
		CommunicateMode = comNone
		end_sql()
		End
		
error_section: 
		On Error Resume Next
		
		end_sql()
		
		F_MSG.Close()
		form_no.Enabled = True
	End Sub
	
	Private Sub F_TMP_MAXLOAD3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        ' -> watanabe del VerUP(2011)
        'Dim aa As String
        ' <- watanabe del VerUP(2011)

        Dim w_w_str As String
		Dim ret As Short
		Dim i As Short
		
        form_no = Me

		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B
		
		'�K�p�K�i
		w_kikaku.Items.Clear()
		w_kikaku.Items.Add("JATMA")
        w_kikaku.Items.Add("TRA (lightweight)")
        w_kikaku.Items.Add("TRA (standard)")
        w_kikaku.Items.Add("TRA (special)")
        w_kikaku.Items.Add("ETRTO(Standard)")
        w_kikaku.Items.Add("ETRTO(Special)")
		
		'�^�C�����
		w_syurui.Items.Clear()
		w_syurui.Items.Add("PC")
		w_syurui.Items.Add("LT")
		w_syurui.Items.Add("TB")
		
		'�t�H���g
        form_no.w_font.Items.Clear()
		For i = 1 To Tmp_font_cnt
			If Trim(Tmp_font_word(i)) = "" Then
				Exit For
			Else
                form_no.w_font.Items.Add(Tmp_font_word(i))
			End If
		Next i
		
		'�^�C�v
		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & Trim(Tmp_Maxload3_ini)
		ret = set_read6(w_w_str, "max_load3", 1)
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
		Call Clear_F_TMP_MAXLOAD3(0)
		
        CommunicateMode = comSpecData
        RequestACAD("SPECDATA")

		If Trim(w_syurui.Text) = "" Then
			w_syurui.Text = "PC"
		End If

        InitFlag = True '20100628�ǉ��R�[�h
	End Sub
	
    'UPGRADE_WARNING: �C�x���g w_font.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
	Private Sub w_font_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font.SelectedIndexChanged
		Dim i As Short
		Dim read_flg As Short
		Dim w_w_str As String
		Dim ret As Short

        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

		read_flg = 0
		For i = 1 To Tmp_font_cnt + 1
			If Tmp_font_word(i) = w_font.Text Then
				w_w_str = Environ("ACAD_SET")
                w_w_str = Trim(w_w_str) & Trim(Tmp_Maxload3_ini)
				ret = set_read6(w_w_str, "Maxload3", i)
				If ret = False Then
                    MsgBox(Tmp_Maxload3_ini & "File reading error.", 64, "BrandVB error")
					Exit Sub
				Else
					read_flg = 1
					Exit For
				End If
			End If
		Next i
		
		If read_flg = 0 Then
            MsgBox("Font type of data that are selected, not set configuration file (" & Tmp_Maxload3_ini & ")", 64, "Configuration file error")
			Exit Sub
		End If
		
		'�^�C�v
        form_no.w_type.Items.Clear()
		For i = 1 To MaxSelNum
			If Tmp_hm_word(i) = "" Then
				Exit For
			Else
                form_no.w_type.Items.Add(Tmp_hm_word(i))
			End If
		Next i
		
        form_no.w_type.Text = ""
        form_no.w_hm_name.Text = ""
        form_no.ImgThumbnail1.Image = Nothing
		
		Call Clear_F_TMP_MAXLOAD3(1)
		
	End Sub
	
    'UPGRADE_WARNING: �C�x���g w_hm_name.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
	Private Sub w_hm_name_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_hm_name.TextChanged
		Dim w_text As String
		Dim TiffFile As String
		Dim w_file As String

        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        On Error Resume Next
		
		Err.Clear()
		
		w_text = w_hm_name.Text
		TiffFile = TIFFDir & w_hm_name.Text & ".bmp"
        If Trim(w_text) = "" Then Exit Sub

		'bmp̧�ٕ\��
		w_file = Dir(TiffFile)
		If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 457 '500 '2010�R�[�h�ύX
            form_no.ImgThumbnail1.Height = 193 '200 '2010�R�[�h�ύX
        Else
            MsgBox("BMP file can not be found.", MsgBoxStyle.Critical)
            form_no.ImgThumbnail1.Image = Nothing
		End If
		
	End Sub
	
	'UPGRADE_WARNING: �C�x���g w_kikaku.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B 
	Private Sub w_kikaku_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku.SelectedIndexChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If
        Call Clear_F_TMP_MAXLOAD3(1)
	End Sub
	
	'UPGRADE_WARNING: �C�x���g w_kikaku.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
	Private Sub w_kikaku_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_kikaku.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If
        Call Clear_F_TMP_MAXLOAD3(1)
	End Sub
	
    'UPGRADE_WARNING: �C�x���g w_size1.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B 
    Private Sub w_size1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_size1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size1.Leave
        'UPGRADE_ISSUE: Control w_size1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B 
        form_no.w_size1.Text = UCase(Trim(form_no.w_size1.Text))
    End Sub

    'UPGRADE_WARNING: �C�x���g w_size2.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_size2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_size2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size2.Leave
        'UPGRADE_ISSUE: Control w_size2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B
        form_no.w_size2.Text = UCase(Trim(form_no.w_size2.Text))
    End Sub

    'UPGRADE_WARNING: �C�x���g w_size3.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B 
    Private Sub w_size3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_size3_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size3.Leave
        'UPGRADE_ISSUE: Control w_size3 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B
        form_no.w_size3.Text = UCase(Trim(form_no.w_size3.Text))
    End Sub

    Private Sub w_size4_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size4.Leave
        'UPGRADE_ISSUE: Control w_size4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B 
        form_no.w_size4.Text = UCase(Trim(form_no.w_size4.Text))
    End Sub

    'UPGRADE_WARNING: �C�x���g w_size5.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_size5_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_size5_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_size5.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii > 32 Then
            If (KeyAscii = CDbl("100")) Or (KeyAscii = CDbl("114")) Then
                KeyAscii = KeyAscii - 32
            ElseIf (KeyAscii <> CDbl("45")) And (KeyAscii <> CDbl("68")) And (KeyAscii <> CDbl("82")) And (KeyAscii <> CDbl("42")) Then
                KeyAscii = 0
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub w_size5_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size5.Leave
        'UPGRADE_ISSUE: Control w_size5 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
        form_no.w_size5.Text = UCase(Trim(form_no.w_size5.Text))
    End Sub

    'UPGRADE_WARNING: �C�x���g w_size6.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_size6_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_size6_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_size6.Leave
        form_no.w_size6.Text = UCase(Trim(form_no.w_size6.Text))
    End Sub

    'UPGRADE_WARNING: �C�x���g w_syurui.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_syurui_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_syurui.TextChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If Trim(w_syurui.Text) = "" Then
            If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
                Call Clear_F_TMP_MAXLOAD3(1)
            End If
        End If
    End Sub

    'UPGRADE_WARNING: �C�x���g w_syurui.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_syurui_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_syurui.SelectedIndexChanged
        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If
    End Sub

    Private Sub w_syurui_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_syurui.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then GoTo EventExitSub
        Call Combo_Sousa(w_syurui, KeyAscii)
        KeyAscii = 0
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    'UPGRADE_WARNING: �C�x���g w_type.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_type_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.TextChanged

        ' -> watanabe del VerUP(2011)
        'Dim flg As Short
        ' <- watanabe del VerUP(2011)

        Dim i As Short

        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        If Trim(w_type.Text) = "" Then
            form_no.w_hm_name.Text = ""
            form_no.ImgThumbnail1.Image = Nothing
            If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
                Call Clear_F_TMP_MAXLOAD3(1)
            End If
            Exit Sub
        End If

        For i = 1 To MaxSelNum
            If (w_type.Text = Tmp_hm_word(i)) Then
                If (Tmp_prcs_code(i) = "MAXLOAD1") Then
                    form_no.w_max_load_kg.Enabled = True
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629�R�[�h�ύX
                    form_no.w_max_load_lbs.Enabled = True
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_kpa.Enabled = True
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_psi.Enabled = True
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)

                    form_no.chk_max_load_kg.Enabled = True
                    form_no.chk_max_load_lbs.Enabled = True
                    form_no.chk_max_press_kpa.Enabled = True
                    form_no.chk_max_press_psi.Enabled = True

                ElseIf (Tmp_prcs_code(i) = "MAXLOAD2") Then
                    form_no.w_max_load_kg.Enabled = True
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_load_lbs.Enabled = False
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_press_kpa.Enabled = True
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_psi.Enabled = False
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

                    form_no.chk_max_load_kg.Enabled = True
                    form_no.chk_max_load_lbs.Enabled = False
                    form_no.chk_max_press_kpa.Enabled = True
                    form_no.chk_max_press_psi.Enabled = False
                    form_no.chk_max_load_lbs.CheckState = 0
                    form_no.chk_max_press_psi.CheckState = 0

                ElseIf (Tmp_prcs_code(i) = "MAXLOAD3") Then
                    form_no.w_max_load_kg.Enabled = False
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_load_lbs.Enabled = True
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_kpa.Enabled = False
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_press_psi.Enabled = True
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)

                    form_no.chk_max_load_kg.Enabled = False
                    form_no.chk_max_load_lbs.Enabled = True
                    form_no.chk_max_press_kpa.Enabled = False
                    form_no.chk_max_press_psi.Enabled = True
                    form_no.chk_max_load_kg.CheckState = 0
                    form_no.chk_max_press_kpa.CheckState = 0
                End If
            End If
        Next i

        For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = w_type.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                Exit For
            End If
        Next i

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If

    End Sub

    'UPGRADE_WARNING: �C�x���g w_type.SelectedIndexChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
    Private Sub w_type_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_type.SelectedIndexChanged

        ' -> watanabe del VerUP(2011)
        'Dim w_str As String
        ' <- watanabe del VerUP(2011)

        Dim i As Short

        If InitFlag = False Then '20100628�ǉ��R�[�h
            Exit Sub
        End If

        For i = 1 To MaxSelNum
            If (w_type.Text = Tmp_hm_word(i)) Then
                If (Tmp_prcs_code(i) = "MAXLOAD1") Then
                    form_no.w_max_load_kg.Enabled = True
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005) '20100629�R�[�h�ύX
                    form_no.w_max_load_lbs.Enabled = True
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_kpa.Enabled = True
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_psi.Enabled = True
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)

                    form_no.chk_max_load_kg.Enabled = True
                    form_no.chk_max_load_lbs.Enabled = True
                    form_no.chk_max_press_kpa.Enabled = True
                    form_no.chk_max_press_psi.Enabled = True

                ElseIf (Tmp_prcs_code(i) = "MAXLOAD2") Then
                    form_no.w_max_load_kg.Enabled = True
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_load_lbs.Enabled = False
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_press_kpa.Enabled = True
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_psi.Enabled = False
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)

                    form_no.chk_max_load_kg.Enabled = True
                    form_no.chk_max_load_lbs.Enabled = False
                    form_no.chk_max_press_kpa.Enabled = True
                    form_no.chk_max_press_psi.Enabled = False
                    form_no.chk_max_load_lbs.CheckState = 0
                    form_no.chk_max_press_psi.CheckState = 0

                ElseIf (Tmp_prcs_code(i) = "MAXLOAD3") Then
                    form_no.w_max_load_kg.Enabled = False
                    form_no.w_max_load_kg.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_load_lbs.Enabled = True
                    form_no.w_max_load_lbs.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
                    form_no.w_max_press_kpa.Enabled = False
                    form_no.w_max_press_kpa.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
                    form_no.w_max_press_psi.Enabled = True
                    form_no.w_max_press_psi.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)

                    form_no.chk_max_load_kg.Enabled = False
                    form_no.chk_max_load_lbs.Enabled = True
                    form_no.chk_max_press_kpa.Enabled = False
                    form_no.chk_max_press_psi.Enabled = True
                    form_no.chk_max_load_kg.CheckState = 0
                    form_no.chk_max_press_kpa.CheckState = 0
                End If
            End If
        Next i

        For i = 1 To MaxSelNum
            If Tmp_hm_word(i) = w_type.Text Then
                w_hm_name.Text = Tmp_hm_code(i)
                Exit For
            End If
        Next i

        If (Trim(w_kikaku_max_load_kg.Text) <> "") Or (Trim(w_kikaku_max_load_lbs.Text) <> "") Or (Trim(w_kikaku_max_press_kpa.Text) <> "") Or (Trim(w_kikaku_max_press_psi.Text) <> "") Then
            Call Clear_F_TMP_MAXLOAD3(1)
        End If

    End Sub

    Private Sub w_type_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles w_type.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then GoTo EventExitSub
        Call Combo_Sousa(w_type, KeyAscii)
        KeyAscii = 0
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
End Class