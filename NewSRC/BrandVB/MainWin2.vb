Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_MAIN2
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MAIN2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Short
		Dim w_w_str As String
		
		form_main = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B
		

#If DEBUG Then
        '20100623�ڐA�ύX
        '2014/8/18 moriya update start
        'w_w_str = "C:\ACAD19_02\BrandV5\uenv\BR_Set.ini"
        w_w_str = "\\ihp0d7\Acad\VER19\uenv\BR_Set.ini"
        '2014/8/18 moriya update end
        ret = set_read(w_w_str)

#Else
        '97.04.23 n.matsumi update start ...............................
        w_w_str = Environ("ACAD_SET")
		'    MsgBox "�����ݒ�̧��1:" & w_w_str, 64
        w_w_str = Trim(w_w_str) & "BR_Set.ini"
		ret = set_read(w_w_str)
        '    MsgBox "�����ݒ�̧��2:" & w_w_str, 64

		'ret = config_read("..\Files\BrandVB.cfg")
        'n.m    ret = set_read("..\Files\BrandVB.set")
         '97.04.23 n.matsumi update ended ...............................

#End If
		
		If ret = False Then
            MsgBox("Error reading initialization file (BR_Set.ini)", MsgBoxStyle.Information, "error")
			GoTo error_section
		End If
		'*****12/8 1997 yamamoto start****
		'    ret = env_get()
		'    If ret = False Then
		'         GoTo error_section
		'    End If
		'*****12/8 1997 yamamoto end******
		'   text2.LinkTimeout = timeOutSecond * 10
		
		ret = init_cad
		Select Case ret
			Case -1
                MsgBox("Fail to connect with the AdvanceCad.", MsgBoxStyle.Information)
				GoTo error_section
			Case errNoAppResponded
                MsgBox("AdvanceCad has not been started.", MsgBoxStyle.Information)
                MsgBox("It is a communication error. It is finished.")
				GoTo error_section
		End Select
		
		ret = init_sql
		If ret = False Then
            MsgBox("Cannot be connected to the SQL server.", MsgBoxStyle.Information)
			GoTo error_section
		End If

        CommunicateMode = comWinName
		ret = RequestACAD("WINNAME")
		
		Exit Sub
		
error_section: 
		
        MsgBox("To exit", MsgBoxStyle.Critical, "Error end")
		End
		
	End Sub
	
	'UPGRADE_NOTE: Form_Terminate �� Form_Terminate_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	'UPGRADE_WARNING: F_MAIN2 �C�x���g Form.Terminate �ɂ͐V�������삪�܂܂�܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"' ���N���b�N���Ă��������B
	Private Sub Form_Terminate_Renamed()

        'SQL�ڑ���۰�ނ��܂�

        ' -> watanabe edit VerUP(2011)
        'SqlExit()
        Call end_sql()
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
	
	
	'UPGRADE_WARNING: �C�x���g Text2.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"' ���N���b�N���Ă��������B
	Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
		Dim i As Object
		Dim w_ret As Short
		Dim Command_Line As String
		Dim hex_data As String
		Static hIndex As Short
		Dim w_w_str As String
		Dim ret As Short
		
		'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
		If form_main.Text2.Text = "" Then Exit Sub
		
		'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
		Command_Line = Trim(form_main.Text2.Text)
		'output_command_line (Command_Line) '----- 12/11 1997 yamamoto add (debug)
		
		'�K�i�`�F�b�N�G���[
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
			Else
				End
			End If
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: Control SRflag �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
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
				ElseIf VB.Left(Command_Line, 2) = "HM" Then 
					' Brand Ver.3 �ύX
					'          DBTableName = DBName & "..hm_kanri"  '�ҏW�����Ǘ�
					DBTableName = DBName & "..hm_kanri1" '�ҏW�����Ǘ�(��{��)
					DBTableName2 = DBName & "..hm_kanri2" '�ҏW�����Ǘ�(������)
				ElseIf VB.Left(Command_Line, 2) = "GZ" Then 
					' Brand Ver.3 �ύX
					'          DBTableName = DBName & "..gz_kanri"  '����}�ʊǗ�
					DBTableName = DBName & "..gz_kanri1" '����}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..gz_kanri2" '����}�ʊǗ�(������)
				ElseIf VB.Left(Command_Line, 2) = "HZ" Then 
					' Brand Ver.3 �ύX
					'          DBTableName = DBName & "..hz_kanri"  '�ҏW�����}�ʊǗ�
					DBTableName = DBName & "..hz_kanri1" '�ҏW�����}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..hz_kanri2" '�ҏW�����}�ʊǗ�(������)
				ElseIf VB.Left(Command_Line, 2) = "BZ" Then 
					' Brand Ver.3 �ύX
					'          DBTableName = DBName & "..bz_kanri"  '�u�����h�}�ʊǗ�
					DBTableName = DBName & "..bz_kanri1" '�u�����h�}�ʊǗ�(��{��)
					DBTableName2 = DBName & "..bz_kanri2" '�u�����h�}�ʊǗ�(������)
				End If
				
				If Len(Command_Line) < 8 Then
					ScreenName = Command_Line & Space(8 - Len(Command_Line))
				End If
				
				'================
				'��ʌĂяo��
				'================
				'// ����}�ʓo�^���-------------------------------------------------------------
				If VB.Left(Command_Line, 6) = "GZSAVE" Then
					ScreenName = "GZSAVE  "
					CommunicateMode = comNone
					sel_win2.Show()
                    sel_win2.mode.Text = "Stamp drawing"
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// ����}�ʍ폜���
				ElseIf VB.Left(Command_Line, 6) = "GZDELE" Then 
					ScreenName = "GZDELE  "
					CommunicateMode = comNone
					F_GZDELE.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// ����}�ʓ��e�m�F���
				ElseIf VB.Left(Command_Line, 6) = "GZLOOK" Then 
					ScreenName = "GZLOOK  "
					CommunicateMode = comNone
					F_GZLOOK.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �ҏW�����}�ʓo�^���-------------------------------------------------------
				ElseIf VB.Left(Command_Line, 6) = "HZSAVE" Then 
					ScreenName = "HZSAVE  "
					CommunicateMode = comNone
					sel_win2.Show()
                    sel_win2.mode.Text = "Editing characters drawing"
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �ҏW�����}�ʍ폜���
				ElseIf VB.Left(Command_Line, 6) = "HZDELE" Then 
					ScreenName = "HZDELE  "
					CommunicateMode = comNone
					F_HZDELE.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �ҏW�����}�ʓ��e�m�F���
				ElseIf VB.Left(Command_Line, 6) = "HZLOOK" Then 
					ScreenName = "HZLOOK  "
					CommunicateMode = comNone
					F_HZLOOK.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �u�����h�}�ʓo�^���------------------------------------------------------
				ElseIf VB.Left(Command_Line, 6) = "BZSAVE" Then 
					ScreenName = "BZSAVE  "
					CommunicateMode = comNone
					sel_win2.Show()
                    sel_win2.mode.Text = "Brand drawing"
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �u�����h�}�ʍ폜���
				ElseIf VB.Left(Command_Line, 6) = "BZDELE" Then 
					ScreenName = "BZDELE  "
					CommunicateMode = comNone
					F_BZDELE.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �u�����h�}�ʓ��e�m�F���
				ElseIf VB.Left(Command_Line, 6) = "BZLOOK" Then 
					ScreenName = "BZLOOK  "
					CommunicateMode = comNone
					F_BZLOOK.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
					
					'// �}�ʔԍ��������--------------------------------------------------------
				ElseIf VB.Left(Command_Line, 8) = "NOSEARCH" Then 
					ScreenName = "NOSEARCH"
					CommunicateMode = comNone
					F_ZSEARCH_BANGO.Show()
					'form_main.text2.Text = ""  Form_Load��PICEMPTY��Reqest���Ă��邽��
					
					'// �}�ʗv�f�������
				ElseIf VB.Left(Command_Line, 8) = "ELSEARCH" Then 
					ScreenName = "ELSEARCH"
					CommunicateMode = comNone
					F_ZSEARCH_YOUSO.Show()
					'form_main.text2.Text = ""  Form_Load��PICEMPTY��Reqest���Ă��邽��
					
					'// �u�����h�������
				ElseIf VB.Left(Command_Line, 8) = "BZSEARCH" Then 
					ScreenName = "BZSEARCH"
					CommunicateMode = comNone
					F_ZSEARCH_BRAND.Show()
					'form_main.text2.Text = ""  Form_Load��PICEMPTY��Reqest���Ă��邽��
					
					'// �}�ʌďo�����
				ElseIf VB.Left(Command_Line, 7) = "ZMNCALL" Then 
					ScreenName = "ZMNCALL"
					CommunicateMode = comNone
					F_ZMNCALL.Show()
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
				Else
                    MsgBox("That isn't ready yet. [" & VB.Left(Command_Line, 8) & "]")
					End
				End If
				
				'�����f�[�^�����҂���
			Case comSpecData
				
				Select Case ScreenName
					
					Case "GZSAVE  " '����}�� �o�^----------------------------------------------
						If VB.Left(Command_Line, 6) = "GMCODE" Then

                            'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            w_ret = temp_gz_set(hex_data)
                            dataset_F_GZSAVE()

							'�������i�`�F�b�N(Brand CAD System Ver.3 UP )
							For i = 1 To temp_gz.gm_num
								'UPGRADE_WARNING: �I�u�W�F�N�g i �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								If IsNumeric(Mid(Trim(temp_gz.gm_name(i)), 3, 4)) = False Then
                                    MsgBox("Can not register for individual parts are included.", 64)
									'��ʃ��b�N
									'UPGRADE_ISSUE: Control Command1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command1.Enabled = False
									'UPGRADE_ISSUE: Control Command2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command2.Enabled = False
									'UPGRADE_ISSUE: Control Command4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command4.Enabled = False
									'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_no1.Enabled = False
									'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_no2.Enabled = False
									'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_comment.Enabled = False
									'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_dep_name.Enabled = False
									'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_entry_name.Enabled = False
									'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									Exit For
								End If
							Next i

                            CommunicateMode = comNone

                        ElseIf VB.Left(Command_Line, 7) = "ZMNNAME" Then
                            'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            w_ret = zumen_no_set(hex_data)
                            If w_ret <> True Then
                                MsgBox("There is no Stamp drawing.", MsgBoxStyle.Critical, "Zumen not Found")
                                End
                            End If
                            CommunicateMode = comSpecData
                            RequestACAD("GMCODE")
						Else
                            MsgBox("Not the Stamp drawing data.[" & Command_Line & "]")
						End If
						
					Case "HZSAVE  " '�ҏW�����}�� �o�^----------------------------------------------
						If VB.Left(Command_Line, 6) = "HMCODE" Then
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
							hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
							w_ret = temp_hz_set(hex_data)
							dataset_F_HZSAVE()
							'�������i�`�F�b�N(Brand CAD System Ver.3 UP )
							For i = 1 To temp_hz.hm_num
								'UPGRADE_WARNING: �I�u�W�F�N�g i �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
								If IsNumeric(Mid(Trim(temp_hz.hm_name(i)), 3, 4)) = False Then
                                    MsgBox("Can not register for individual parts are included.", 64)
									'��ʃ��b�N
									'UPGRADE_ISSUE: Control Command1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command1.Enabled = False
									'UPGRADE_ISSUE: Control Command2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command2.Enabled = False
									'UPGRADE_ISSUE: Control Command4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.Command4.Enabled = False
									'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_no1.Enabled = False
									'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_no2.Enabled = False
									'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_comment.Enabled = False
									'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_dep_name.Enabled = False
									'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
									form_no.w_entry_name.Enabled = False
									'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
									Exit For
								End If
							Next i
							CommunicateMode = comNone
						ElseIf VB.Left(Command_Line, 7) = "ZMNNAME" Then 
							'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
							form_main.Text2.Text = ""
							hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
							w_ret = zumen_no_set_hz(hex_data)
							If w_ret <> True Then
                                MsgBox("There is no Editing characters drawing.", MsgBoxStyle.Critical, "Zumen not Found")
								End
                            End If
                            CommunicateMode = comSpecData
							RequestACAD("HMCODE")
						Else
                            MsgBox("It is not in the Editing characters drawing data [" & Command_Line & "]")
						End If
						
					Case "BZSAVE  " '�u�����h�o�^-------------------------------------------------------

                        If VB.Left(Command_Line, 7) = "SPEC501" Then

                            'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_BZSAVE()

                            If open_mode = "NEW" Then
                                temp_bz.no1 = ""
                                'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                form_no.w_no1.Text = ""
                                temp_bz.no2 = "00"
                                'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                form_no.w_no2.Text = "00"
                                'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                form_no.w_no2.Enabled = False
                                'UPGRADE_ISSUE: Control w_kanri_no �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                form_no.w_kanri_no.Text = ""

                                CommunicateMode = comSpecData
                                RequestACAD("HMCODE")
                            Else
                                CommunicateMode = comSpecData
                                RequestACAD("ZMNNAME")
                            End If


                        ElseIf VB.Left(Command_Line, 6) = "HMCODE" Then
                            'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            w_ret = temp_bz_set(1, hex_data)
                            dataset_F_BZSAVE()

                            '�������i�`�F�b�N(Brand CAD System Ver.3 UP )
                            For i = 1 To temp_bz.hm_num
                                'UPGRADE_WARNING: �I�u�W�F�N�g i �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
                                If IsNumeric(Mid(Trim(temp_bz.hm_name(i)), 3, 4)) = False Then
                                    MsgBox("Can not register for individual parts are included.", 64)
                                    '��ʃ��b�N
                                    'UPGRADE_ISSUE: Control Command1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.Command1.Enabled = False
                                    'UPGRADE_ISSUE: Control Command2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.Command2.Enabled = False
                                    'UPGRADE_ISSUE: Control Command4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.Command4.Enabled = False
                                    'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no1.Enabled = False
                                    'UPGRADE_ISSUE: Control w_no1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no2.Enabled = False
                                    'UPGRADE_ISSUE: Control w_no2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_no2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kanri_no �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kanri_no.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kanri_no �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kanri_no.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_comment.Enabled = False
                                    'UPGRADE_ISSUE: Control w_comment �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_dep_name.Enabled = False
                                    'UPGRADE_ISSUE: Control w_dep_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_entry_name.Enabled = False
                                    'UPGRADE_ISSUE: Control w_entry_name �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_syurui �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_syurui.Enabled = False
                                    'UPGRADE_ISSUE: Control w_syurui �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_syurui.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_pattern �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_pattern.Enabled = False
                                    'UPGRADE_ISSUE: Control w_pattern �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_pattern.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_syubetu �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_syubetu.Enabled = False
                                    'UPGRADE_ISSUE: Control w_syubetu �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_syubetu.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size1.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size2.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size3 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size3.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size3 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size4.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size5 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size5.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size5 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size6 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size6.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size6 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size7 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size7.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size7 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size7.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_size8 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size8.Enabled = False
                                    'UPGRADE_ISSUE: Control w_size8 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_size8.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_plant �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_plant.Enabled = False
                                    'UPGRADE_ISSUE: Control w_plant �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_plant.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku1.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku1 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku1.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku2.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku2.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku3 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku3.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku3 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku3.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku4.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku4 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku4.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku5 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku5.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku5 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku5.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_kikaku6 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku6.Enabled = False
                                    'UPGRADE_ISSUE: Control w_kikaku6 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_kikaku6.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_tos_moyou �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_tos_moyou.Enabled = False
                                    'UPGRADE_ISSUE: Control w_tos_moyou �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_tos_moyou.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_peak_mark �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_peak_mark.Enabled = False
                                    'UPGRADE_ISSUE: Control w_peak_mark �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_peak_mark.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_side_moyou �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_side_moyou.Enabled = False
                                    'UPGRADE_ISSUE: Control w_side_moyou �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_side_moyou.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_side_kenti �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_side_kenti.Enabled = False
                                    'UPGRADE_ISSUE: Control w_side_kenti �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_side_kenti.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    'UPGRADE_ISSUE: Control w_nasiji �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_nasiji.Enabled = False
                                    'UPGRADE_ISSUE: Control w_nasiji �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                                    form_no.w_nasiji.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
                                    Exit For
                                End If
                            Next i
                            CommunicateMode = comNone

                        ElseIf VB.Left(Command_Line, 7) = "ZMNNAME" Then
                            'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            w_ret = zumen_no_set_bz(hex_data)
                            If w_ret <> True Then
                                MsgBox("There is no brand drawing.", MsgBoxStyle.Critical, "Zumen not Found")
                                End
                            End If
                            CommunicateMode = comSpecData
                            RequestACAD("HMCODE")

                        ElseIf VB.Left(Command_Line, 10) = "SPECADD OK" Then
                            'ACAD����I��
                        ElseIf VB.Left(Command_Line, 13) = "ZUMEN SAVE OK" Then
                            'ACAD����I��
                        Else
                            MsgBox("It is not a brand drawing data [" & Command_Line & "]")
                        End If
						
					Case Else
                        MsgBox("Do not understand. . .(" & ScreenName & ")," & Len(ScreenName))
				End Select
				
			Case comFreePic
				If (VB.Left(Command_Line, 8) = "PICEMPTY") Then

                    ' -> watanabe edit 2013.06.03
                    'FreePicNum = Val(Mid(Command_Line, 9, 2))
                    'If FreePicNum > 50 Then FreePicNum = 50
                    FreePicNum = Val(Mid(Command_Line, 9, 3))
                    If FreePicNum > 130 Then FreePicNum = 130
                    ' <- watanabe edit 2013.06.03

                    CommunicateMode = comNone
					'UPGRADE_ISSUE: Control Text2 �́A�ėp���O��� Form ���ɂ��邽�߁A�����ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"' ���N���b�N���Ă��������B
					form_main.Text2.Text = ""
				Else
                    MsgBox("Not a free picture information [" & Command_Line & "]")
					End
				End If
				
			Case Else
                MsgBox("communicateMode error")
		End Select
	End Sub
	
	'UPGRADE_ISSUE: TextBox �C�x���g Text2.LinkClose �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub Text2_LinkClose()
		Dim Connected As Object
		
		'UPGRADE_WARNING: �I�u�W�F�N�g Connected �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
		Connected = False
		
	End Sub
	
	
	'UPGRADE_ISSUE: TextBox �C�x���g Text2.LinkError �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub Text2_LinkError(ByRef LinkErr As Short)
		Dim Msg As Object
		'UPGRADE_WARNING: �I�u�W�F�N�g Msg �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
        Msg = "DDE communication error"
		MsgBox(Msg)
	End Sub
	
	
	'UPGRADE_ISSUE: TextBox �C�x���g Text2.LinkNotify �̓A�b�v�O���[�h����܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"' ���N���b�N���Ă��������B
	Private Sub Text2_LinkNotify()
        'Dim NotifyFlag As Object
		If Not NotifyFlag Then
            MsgBox("Can get the new data from the DDE source.")
			'UPGRADE_WARNING: �I�u�W�F�N�g NotifyFlag �̊���v���p�e�B�������ł��܂���ł����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' ���N���b�N���Ă��������B
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