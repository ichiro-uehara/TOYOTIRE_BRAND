Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_MAIN1
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MAIN1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Short
		Dim w_w_str As String

		form_main = Me

		'----- .NET�ڍs (StartPosition�v���p�e�B��CenterScreen�őΉ�) -----
		'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B

		'97.04.23 n.matsumi update start ...............................
#If DEBUG Then
		'2014/8/18 moriya update start
		'w_w_str = "C:\ACAD19_02\BrandV5\uenv\BR_Set.ini"
		'w_w_str = "\\ihp0d7\Acad\VER19\uenv\BR_Set.ini"
		'2014/8/18 moriya update end

		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & "BR_Set.ini"
#Else
		w_w_str = Environ("ACAD_SET")
		w_w_str = Trim(w_w_str) & "BR_Set.ini"
#End If
		ret = set_read(w_w_str)

        '    ret = config_read("..\Files\BrandVB.cfg")
        'n.m    ret = set_read("..\Files\BrandVB.set")
		'97.04.23 n.matsumi update ended ...............................
		
		If ret = False Then
            MsgBox("Error reading initialization file (BR_Set.ini)", MsgBoxStyle.Information, "error")
			GoTo error_section
		End If
		'    ret = env_get()
		'    If ret = False Then
		'         GoTo error_section
		'    End If
		'   text2.LinkTimeout = timeOutSecond * 10

		ret = init_cad()
		Select Case ret
			Case -1
				MsgBox("Fail to connect with the AdvanceCad.", MsgBoxStyle.Information)
				GoTo error_section

			Case errNoAppResponded
				MsgBox("AdvanceCad has not been started.", MsgBoxStyle.Information)
				MsgBox("It is a communication error. It is finished.")
				GoTo error_section
		End Select

		ret = init_sql()
		If ret = False Then
            MsgBox("Cannot be connected to the SQL server.", MsgBoxStyle.Information)
			GoTo error_section
		End If

		'----- .NET�ڍs (DDE�ʐM�R�����g��) -----
		CommunicateMode = comWinName
		ret = RequestACAD("WINNAME")

		'�u���n�������� / �ҏW���������v�Ńf�o�b�O�iToDo:�f�o�b�O��폜�j
		'CommunicateMode = 2
		'Text2.Text = "GMSEARCH"
		'Text2.Text = "HMSEARCH1"
		'�f�o�b�O�iToDo:�f�o�b�O��폜�j

		Exit Sub

error_section: 
		
        MsgBox("To exit", MsgBoxStyle.Critical, "Error end")
		End
		
	End Sub

    Private Sub Form_Terminate_Renamed()

        'SQL�ڑ���۰�ނ��܂�

        ' -> watanabe edit VerUP(2011)
        'SqlExit()
        end_sql()
        ' <- watanabe edit VerUP(2011)

    End Sub

	Private Sub LINK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles LINK.Click
		
		'   ret = init_cad
		'   Select Case ret
		'       Case False
		'           MsgBox "AdvanceCad�Ƃ̐ڑ��Ɏ��s���܂���", 64
		'       Case errNoAppResponded
		'           MsgBox "AdvanceCad�͋N������Ă��܂���"
		'   End Select
		
	End Sub
	
	Private Sub POKE_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles POKE.Click
		
		'On Error Resume Next
		' form_main.text2.LinkPoke
		' If Err Then MsgBox Error
		
	End Sub
	
	Private Sub REQUEST_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles REQUEST.Click
		
		'/// TEST VERSION
		'On Error Resume Next
		'  Text2.LinkItem = "WINNAME"
		' text2.LinkItem = form_main.text2.Text
		' text2.LinkRequest
		' NotifyFlag = False
		
	End Sub
	
    'UPGRADE_WARNING: �C�x���g Text2.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
	Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
		
		Dim w_ret As Short
		Dim Command_Line As String
		Dim hex_data As String
		Static hIndex As Short
		Dim w_w_str As String
		Dim ret As Short

		If form_main.Text2.Text = "" Then Exit Sub

		Command_Line = Trim(form_main.Text2.Text)
        '  output_command_line (Command_Line) '----- 12/11 1997 yamamoto add (debug)-----

        '   �K�i�`�F�b�N�G���[
		If VB.Left(Command_Line, 6) = "ERRORZ" Then
            MsgBox("There was an error in standard check.", MsgBoxStyle.Critical, "ERROR FROM ACAD")
			F_MSG3.Close()
			form_no.Enabled = True
			Exit Sub
			End
		End If
		
		If VB.Left(Command_Line, 5) = "ERROR" Then
			MsgBox(Command_Line, MsgBoxStyle.Critical, "ERROR FROM ACAD")
			If Mid(ScreenName, 1, 3) = "TMP" Then
				On Error Resume Next
				F_MSG.Close()
				form_no.Enabled = True
				'97.04.23 n.matsumi update start ..............................
			Else
				End
				'97.04.23 n.matsumi update ended ..............................
			End If
			Exit Sub
		End If
		
		If Trim(form_main.SRflag.Text) = "SEND" Then Exit Sub
		
		Select Case CommunicateMode
			'���M�҂��Ȃ�
			Case comNone
				If VB.Left(Command_Line, 6) = "VBKILL" Then
					'           MsgBox "VBKILL��M���܂���" & Chr(13) & "BrandVB���I�����܂�"
					End
				Else
                    '            MsgBox Command_Line & "��M���܂���"
				End If
				'��ʖ��҂���
			Case comWinName
				'================
				'�e�[�u�����̎擾
				'================
				
				If VB.Left(Command_Line, 2) = "GM" Then
					DBTableName = DBName & "..gm_kanri" '���n�����Ǘ�
					TIFFDir = TIFFDirGM
				ElseIf VB.Left(Command_Line, 2) = "HM" Then
					DBTableName = DBName & "..hm_kanri1" '�ҏW�����Ǘ�(��{��)
					DBTableName2 = DBName & "..hm_kanri2" '�ҏW�����Ǘ�(������)
					TIFFDir = TIFFDirHM
				ElseIf VB.Left(Command_Line, 2) = "GZ" Then
					DBTableName = DBName & "..gz_kanri1" '����}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..gz_kanri2" '����}�ʊǗ�(������)
				ElseIf VB.Left(Command_Line, 2) = "HZ" Then
					DBTableName = DBName & "..hz_kanri1" '�ҏW�����}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..hz_kanri2" '�ҏW�����}�ʊǗ�(������)
				ElseIf VB.Left(Command_Line, 2) = "BZ" Then
					DBTableName = DBName & "..bz_kanri1" '�u�����h�}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..bz_kanri2" '�u�����h�}�ʊǗ�(������)
				End If
				If Len(Command_Line) < 8 Then
					ScreenName = Command_Line & Space(8 - Len(Command_Line))
				End If
				
				'================
				'��ʌĂяo��
				'================
				'/// ���n�����o�^���
				If VB.Left(Command_Line, 6) = "GMSAVE" Then
					'               ScreenName = "GMSAVE  "
					CommunicateMode = comNone
					sel_win1.Show()
					sel_win1.mode.Text = "Primitive character"
					form_main.Text2.Text = ""
					'// ���n�����폜���
				ElseIf VB.Left(Command_Line, 6) = "GMDELE" Then
					'ScreenName = "GMDELE  "
					'CommunicateMode = comNone
					'F_GMDELE.Show()
					'form_main.Text2.Text = ""
					'// ���n�����������
				ElseIf VB.Left(Command_Line, 8) = "GMSEARCH" Then 
					ScreenName = "GMSEARCH"
					CommunicateMode = comFreePic
					'----- .NET�ڍs (�b��I��99���Ƃ��� �� ToDo:�d�l�m�F��ɑΉ�)-----
					FreePicNum = 0

					F_GMSEARCH.Show()
					form_main.Text2.Text = ""
					w_ret = RequestACAD("PICEMPTY")
					'// ���n�������e�m�F���
				ElseIf VB.Left(Command_Line, 6) = "GMLOOK" Then
					'ScreenName = "GMLOOK  "
					'CommunicateMode = comNone
					'F_GMLOOK.Show()
					'form_main.Text2.Text = ""
					'// ���n�����Ǎ���
				ElseIf VB.Left(Command_Line, 6) = "GMREAD" Then
					'CommunicateMode = comFreePic
					'FreePicNum = 0
					'               open_mode = "Primitive character"
					'F_CADREAD.Show()
					'form_main.Text2.Text = ""
					'w_ret = RequestACAD("PICEMPTY")
					'// �ҏW�����o�^���
				ElseIf VB.Left(Command_Line, 6) = "HMSAVE" Then
					ScreenName = "HMSAVE  "
					CommunicateMode = comNone
					sel_win1.Show()
					sel_win1.mode.Text = "Editing characters"
					form_main.Text2.Text = ""
					'// �ҏW�����폜���
				ElseIf VB.Left(Command_Line, 6) = "HMDELE" Then
					'ScreenName = "HMDELE  "
					'CommunicateMode = comNone
					'F_HMDELE.Show()
					'form_main.Text2.Text = ""
					'// �ҏW�����������
				ElseIf VB.Left(Command_Line, 9) = "HMSEARCH1" Then
					ScreenName = "HMSEARCH1"
					CommunicateMode = comFreePic
					'----- .NET�ڍs (�b��I��99���Ƃ��� �� ToDo:�d�l�m�F��ɑΉ�)-----
					'FreePicNum = 0

					F_HMSEARCH.Show()
					form_main.Text2.Text = ""
					w_ret = RequestACAD("PICEMPTY")
					'Brand Ver.4 �ǉ�
					'// �ҏW�����������
				ElseIf VB.Left(Command_Line, 9) = "HMSEARCH2" Then
					'ScreenName = "HMSEARCH2"
					'CommunicateMode = comFreePic
					'FreePicNum = 0
					'F_HMSEARCH2.Show()
					'form_main.Text2.Text = ""
					'w_ret = RequestACAD("PICEMPTY")
					'�ǉ��I��
					'// �ҏW�������e�m�F���
				ElseIf VB.Left(Command_Line, 6) = "HMLOOK" Then
					'ScreenName = "HMLOOK  "
					'CommunicateMode = comNone
					'F_HMLOOK.Show()
					'form_main.Text2.Text = ""
					'// �ҏW�����Ǎ���
				ElseIf VB.Left(Command_Line, 6) = "HMREAD" Then
					'CommunicateMode = comFreePic
					'               FreePicNum = 0
					'               open_mode = "Editing characters"
					'F_CADREAD.Show()
					'form_main.Text2.Text = ""
					'w_ret = RequestACAD("PICEMPTY")
				Else
                    MsgBox("That isn't ready yet. [" & VB.Left(Command_Line, 8) & "]")
					End
				End If
				'�����f�[�^�����҂���
			Case comSpecData
				
				Select Case ScreenName
					'���n������
					
                    Case "GMSAVE  "

						If VB.Left(Command_Line, 7) = "SPEC101" Then
							CommunicateMode = comNone
							hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
							temp_gm_set(hex_data)
							dataset_F_GMSAVE()
						Else
							MsgBox("It is not a primitive character property data [" & Command_Line & "]")
						End If
						form_main.Text2.Text = ""

					Case "HMSAVE  "

						If VB.Left(Command_Line, 7) = "SPEC201" Then
							form_main.Text2.Text = ""
							If Mid(Command_Line, 8, 1) = "-" Then
								hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
								w_ret = temp_hm_set(0, hex_data)
							ElseIf IsNumeric(Mid(Command_Line, 8, 1)) Then
								hIndex = CShort(Mid(Command_Line, 8, 1))
								hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
								w_ret = temp_hm_set(hIndex, hex_data)
							End If
						Else
							MsgBox("It is not the editing characters property data [" & Command_Line & "]")
						End If

					Case Else
						MsgBox("Do not understand. . . (" & ScreenName & ")," & Len(ScreenName))
				End Select
				
			Case comFreePic
				If (VB.Left(Command_Line, 8) = "PICEMPTY") Then

                    ' -> watanabe edit 2013.05.29
                    'FreePicNum = Val(Mid(Command_Line, 9, 2))
                    FreePicNum = Val(Mid(Command_Line, 9, 3))
                    ' <- watanabe edit 2013.05.29

					' MsgBox "�󂫃s�N�`����" & FreePicNum
					
                    ' -> watanabe edit 2013.05.29
                    'If FreePicNum > 50 Then FreePicNum = 50
                    If FreePicNum > 130 Then FreePicNum = 130
                    ' <- watanabe edit 2013.05.29

                    CommunicateMode = comNone
					form_main.Text2.Text = ""
				Else
                    MsgBox("Not a free picture information [" & Command_Line & "]")
					End
				End If
				
			Case Else
                MsgBox("communicateMode error")
		End Select
	End Sub
	
	Private Sub Text2_LinkClose()
		Dim Connected As Object
		
		Connected = False
		
	End Sub
	
	Private Sub Text2_LinkError(ByRef LinkErr As Short)
		Dim Msg As Object
        Msg = "DDE communication error"
		MsgBox(Msg)
	End Sub
	
	Private Sub Text2_LinkNotify()
        'Dim NotifyFlag As Object
		If Not NotifyFlag Then
            MsgBox("Can get the new data from the DDE source.")
			NotifyFlag = True
		End If
	End Sub
	
	Private Sub Vbsql1_Error(ByVal SqlConn As Integer, ByVal Severity As Integer, ByVal ErrorNum As Integer, ByVal ErrorStr As String, ByVal OSErrorNum As Integer, ByVal OSErrorStr As String, ByRef RetCode As Integer)
		MsgBox("DB-Library Error: " & Str(ErrorNum) & " " & ErrorStr)
	End Sub
	
	Private Sub Vbsql1_Message(ByVal SqlConn As Integer, ByVal Message As Integer, ByVal State As Integer, ByVal Severity As Integer, ByVal MsgStr As String, ByVal ServerNameStr As String, ByVal ProcNameStr As String, ByVal Line As Integer)
		' If Severity > 1 Then
		'   MsgBox ("SQL Server Error: " + Str$(Message&) + " " + MsgStr$)
		' End If
	End Sub
End Class