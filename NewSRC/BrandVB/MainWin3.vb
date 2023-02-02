Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_MAIN3
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MAIN3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Short
		Dim w_w_str As String

        ' -> watanabe del VerUP(2011)
        'Dim w_ret As Short
        ' <- watanabe del VerUP(2011)

		form_main = Me
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。

#If DEBUG Then
        '20100623移植変更
        '2014/8/18 moriya update start
        'w_w_str = "C:\ACAD19_02\BrandV5\uenv\BR_Set.ini"
        w_w_str = "\\ihp0d7\Acad\VER19\uenv\BR_Set.ini"
        '2014/8/18 moriya update end
        ret = set_read(w_w_str)

#Else
		'97.04.23 n.matsumi update start ...............................
		w_w_str = Environ("ACAD_SET")
		'    MsgBox "初期設定ﾌｧｲﾙ1:" & w_w_str, 64
		w_w_str = Trim(w_w_str) & "BR_Set.ini"
		ret = set_read(w_w_str)
		'    MsgBox "初期設定ﾌｧｲﾙ2:" & w_w_str, 64
		
		'ret = config_read("..\Files\BrandVB.cfg")
        'n.m    ret = set_read("..\Files\BrandVB.set")
		'97.04.23 n.matsumi update ended ...............................

#End If

		If ret = False Then
            MsgBox("Error reading initialization file (BR_Set.ini)", MsgBoxStyle.Information, "error")
			GoTo error_section
		End If
		
		'*****12/8 1997 yamamoto start*****
		'    ret = env_get()
		'    If ret = False Then
		'         GoTo error_section
		'    End If
		'*****12/8 1997 yamamoto end*****
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
	
    Private Sub Form_Terminate_Renamed()
        'SQL接続をｸﾛｰｽﾞします

        ' -> watanabe edit VerUP(2011)
        'SqlExit()
        Call end_sql()
        ' <- watanabe edit VerUP(2011)

    End Sub
	
	Private Sub LINK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles LINK.Click
		
		'   ret = init_cad
		'   Select Case ret
		'       Case False
		'           MsgBox "AdvanceCadとの接続に失敗しました", 64
		'       Case errNoAppResponded
		'           MsgBox "AdvanceCadは起動されていません"
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
	
    'UPGRADE_WARNING: イベント Text2.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
		Dim w_ret As Short
		Dim Command_Line As String
		Dim hex_data As String
		Static hIndex As Short
		Dim w_w_str As String
        Dim ret As Short


        If form_main.Text2.Text = "" Then Exit Sub
		
        Command_Line = Trim(form_main.Text2.Text)
		'   output_command_line (Command_Line) '----- 12/11 1997 yamamoto add(debug) -----
		
		'   規格チェックエラー
		If VB.Left(Command_Line, 6) = "ERRORZ" Then
            MsgBox("There was an error in standard check.", MsgBoxStyle.Critical, "ERROR FROM ACAD")
			F_MSG3.Close()
			form_no.Enabled = True
			Exit Sub
			End
		End If
		
        TIFFDir = TIFFDirHM
		
		
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
		
		
        If Trim(form_main.SRflag.Text) = "SEND" Then Exit Sub
		
		Select Case CommunicateMode
			
			'送信待ちなし
			Case comNone
				If VB.Left(Command_Line, 6) = "VBKILL" Then
					'                MsgBox "VBKILL受信しました" & Chr(13) & "BrandVBを終了します"
					End
				Else
                    '            MsgBox Command_Line & "受信しました"
				End If
				
				'画面名待ち時
            Case comWinName
                '================
                'テーブル名の取得
                '================
                If VB.Left(Command_Line, 2) = "GM" Then
                    DBTableName = DBName & "..gm_kanri" '原始文字管理
                ElseIf VB.Left(Command_Line, 2) = "HM" Then
                    ' Brand Ver.3 変更
                    '          DBTableName = DBName & "..hm_kanri"  '編集文字管理
                    DBTableName = DBName & "..hm_kanri1" '編集文字管理(基本部)
                    DBTableName2 = DBName & "..hm_kanri2" '編集文字管理(文字部)
                ElseIf VB.Left(Command_Line, 2) = "GZ" Then
                    ' Brand Ver.3 変更
                    '          DBTableName = DBName & "..gz_kanri"  '刻印図面管理
                    DBTableName = DBName & "..gz_kanri1" '刻印図面管理(基本部)
                    DBTableName2 = DBName & "..gz_kanri2" '刻印図面管理(文字部)
                ElseIf VB.Left(Command_Line, 2) = "HZ" Then
                    ' Brand Ver.3 変更
                    '          DBTableName = DBName & "..hz_kanri"  '編集文字図面管理
                    DBTableName = DBName & "..hz_kanri1" '編集文字図面管理(基本部)
                    DBTableName2 = DBName & "..hz_kanri2" '編集文字図面管理(文字部)
                ElseIf VB.Left(Command_Line, 2) = "BZ" Then
                    ' Brand Ver.3 変更
                    '          DBTableName = DBName & "..bz_kanri"  'ブランド図面管理
                    DBTableName = DBName & "..bz_kanri1" 'ブランド図面管理(基本部)
                    DBTableName2 = DBName & "..bz_kanri2" 'ブランド図面管理(文字部)
                End If
                If Len(Command_Line) < 8 Then
                    ScreenName = Command_Line & Space(8 - Len(Command_Line))
                End If

                '================
                '画面呼び出し
                '================
                '// テンプレート(サイズ)画面
                If VB.Left(Command_Line, 8) = "TMPSIZE1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'FreePicNum = 0
                    'CommunicateMode = comNone
                    'ScreenName = "TMPSIZE1"
                    'F_TMP_SIZE.Show()

                    '// テンプレート(荷重指数)画面
                ElseIf VB.Left(Command_Line, 8) = "TMPKAJU1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Load1_ini)
                    'ret = set_read2(w_w_str, "load1")
                    'FreePicNum = 0
                    'ScreenName = "TMPKAJU1"
                    'F_TMP_KAJUU.Show()

                    '// テンプレート(セリアル)画面
                ElseIf VB.Left(Command_Line, 7) = "TMPSERI" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Serial1_ini)
                    'ret = set_read2(w_w_str, "serial1")
                    'FreePicNum = 0
                    'ScreenName = "TMPSERI "
                    'F_TMP_SERIARU.Show()

                    '// テンプレート(モールド番号)画面
                ElseIf VB.Left(Command_Line, 7) = "TMPMOLD" Then
                    w_w_str = Environ("ACAD_SET")
                    w_w_str = Trim(w_w_str) & Trim(Tmp_Mold_no1_ini)
                    ret = set_read2(w_w_str, "mold_no1")
                    ScreenName = "TMPMOLD "
                    CommunicateMode = comNone
                    F_TMP_MORUDO.Show()
                    'form_main.text2.Text = ""  Form_LoadでPICEMPTYをReqestしているため

                    '// テンプレート(ENO)画面
                ElseIf VB.Left(Command_Line, 6) = "TMPENO" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_E_no1_ini)
                    'ret = set_read2(w_w_str, "e_no1")
                    'ScreenName = "TMPENO  "
                    'CommunicateMode = comNone
                    'F_TMP_ENO.Show()

                    '// テンプレート(PTNCODE)画面
                ElseIf VB.Left(Command_Line, 8) = "TMPPTRN1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Pattern1_ini)
                    'ret = set_read2(w_w_str, "pattern_code1")
                    'ScreenName = "TMPPTN1"
                    'CommunicateMode = comNone
                    'F_TMP_PTNCODE.Show()

                    '// テンプレート(UTQG)画面
                    ' -> watanabe edit 2007.03
                    '            ElseIf Left$(Command_Line, 7) = "TMPUTQG" Then
                ElseIf VB.Left(Command_Line, 8) = "TMPUTQG1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg1_ini)
                    'ret = set_read2(w_w_str, "utqg1")
                    'ScreenName = "TMPUTQG1"
                    'CommunicateMode = comNone
                    'F_TMP_UTQG.Show()

                    '// テンプレート(MAXLOAD)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPMAXLD1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Maxload1_ini)
                    'ret = set_read2(w_w_str, "max_load1")
                    'ScreenName = "TMPMAXLD1"
                    'CommunicateMode = comNone
                    'F_TMP_MAXLOAD.Show()

                    '// テンプレート(PLY1)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPPLY1_1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Ply1_ini)
                    'ret = set_read2(w_w_str, "ply1")
                    'ScreenName = "TMPPRY1_1"
                    'CommunicateMode = comNone
                    'F_TMP_PLY.Show()

                    '// テンプレート(PLY2)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPPLY2_1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Ply2_ini)
                    'ret = set_read2(w_w_str, "ply2")
                    'ScreenName = "TMPPRY2_1"
                    'CommunicateMode = comNone
                    'F_TMP_PLY2.Show()

                    '// テンプレート(ETC)画面
                ElseIf VB.Left(Command_Line, 7) = "TMPETC1" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_ETC_ini)
                    'ret = set_read2(w_w_str, "etc")
                    'ScreenName = "TMPETC1"
                    'CommunicateMode = comNone
                    'F_TMP_ETC.Show()

                    '// テンプレート(ﾌﾟﾚｰﾄ)画面
                ElseIf VB.Left(Command_Line, 8) = "TMPPLATE" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'FreePicNum = 0
                    'ScreenName = "TMPPLATE"
                    'F_TMP_PLATE.Show()

                    '// テンプレート(PTNCODE) タイプ2 画面
                ElseIf VB.Left(Command_Line, 8) = "TMPPTRN2" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Pattern2_ini)
                    'ret = set_read2(w_w_str, "pattern_code2")
                    'ScreenName = "TMPPTN2"
                    'CommunicateMode = comNone
                    'F_TMP_PTNCODE2.Show()

                    '// テンプレート(SIZE) タイプ2 画面
                ElseIf VB.Left(Command_Line, 8) = "TMPSIZE2" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Size2_ini)
                    'ret = set_read2(w_w_str, "size2")
                    'ScreenName = "TMPSIZE2"
                    'CommunicateMode = comNone
                    'F_TMP_SIZE2.Show()

                    '// テンプレート(LOAD) タイプ2 (S) 画面
                ElseIf VB.Left(Command_Line, 8) = "TMPKAJ2S" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Load2S_ini)
                    'ret = set_read2(w_w_str, "load2S")
                    'ScreenName = "TMPKAJ2S"
                    'CommunicateMode = comNone
                    'F_TMP_KAJUU2S.Show()

                    '// テンプレート(LOAD) タイプ2 (D) 画面
                ElseIf VB.Left(Command_Line, 8) = "TMPKAJ2D" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Load2D_ini)
                    'ret = set_read2(w_w_str, "load2D")
                    'ScreenName = "TMPKAJ2D"
                    'CommunicateMode = comNone
                    'F_TMP_KAJUU2D.Show()

                    '// テンプレート(LT) タイプ2 画面
                ElseIf VB.Left(Command_Line, 6) = "TMPLT2" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Lt2_ini)
                    'ret = set_read2(w_w_str, "lt2")
                    'ScreenName = "TMPLT2"
                    'CommunicateMode = comNone
                    'F_TMP_LT2.Show()

                    '// テンプレート(PR) タイプ2 画面
                ElseIf VB.Left(Command_Line, 6) = "TMPPR2" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Pr2_ini)
                    'ret = set_read2(w_w_str, "pr2")
                    'ScreenName = "TMPPR2"
                    'CommunicateMode = comNone
                    'F_TMP_PR2.Show()

                    '// テンプレート(PSI) タイプ2 画面
                ElseIf VB.Left(Command_Line, 7) = "TMPPSI2" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Psi2_ini)
                    'ret = set_read2(w_w_str, "psi2")
                    'ScreenName = "TMPPSI2"
                    'CommunicateMode = comNone
                    'F_TMP_PSI2.Show()

                    '// テンプレート(UTQG3)画面
                ElseIf VB.Left(Command_Line, 8) = "TMPUTQG3" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Utqg3_ini)
                    'ret = set_read2(w_w_str, "utqg3")
                    'ScreenName = "TMPUTQG3"
                    'CommunicateMode = comNone
                    'F_TMP_UTQG3.Show()

                    '// テンプレート(MAXLOAD3)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPMAXLD3" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Maxload3_ini)
                    'ret = set_read2(w_w_str, "max_load3")
                    'ScreenName = "TMPMAXLD3"
                    'CommunicateMode = comNone
                    'F_TMP_MAXLOAD3.Show()

                    '// テンプレート(PLY1_3)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPPLY1_3" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Ply1_3_ini)
                    'ret = set_read2(w_w_str, "ply1_3")
                    'ScreenName = "TMPPRY1_3"
                    'CommunicateMode = comNone
                    'F_TMP_PLY1_3.Show()

                    '// テンプレート(PLY2_3)画面
                ElseIf VB.Left(Command_Line, 9) = "TMPPLY2_3" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_Ply2_3_ini)
                    'ret = set_read2(w_w_str, "ply2_3")
                    'ScreenName = "TMPPRY2_3"
                    'CommunicateMode = comNone
                    'F_TMP_PLY2_3.Show()

                    '// テンプレート(ETC)画面
                ElseIf VB.Left(Command_Line, 7) = "TMPETC3" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'w_w_str = Environ("ACAD_SET")
                    'w_w_str = Trim(w_w_str) & Trim(Tmp_ETC3_ini)
                    'ret = set_read2(w_w_str, "etc3")
                    'ScreenName = "TMPETC3"
                    'CommunicateMode = comNone
                    'F_TMP_ETC3.Show()

                    '// テンプレート(Mark)画面
                ElseIf VB.Left(Command_Line, 7) = "TMPMARK" Then
                    '----- .NET 移行(一旦コメント化) -----
                    'ScreenName = "TMPMARK"
                    'CommunicateMode = comNone
                    'F_TMP_MARK.Show()

                Else
                    MsgBox("That isn't ready yet. [" & VB.Left(Command_Line, 8) & "]")
                    End
                End If

                '特性データ到着待ち時
            Case comSpecData

                Select Case ScreenName
                    '----- 1997 yamamoto -----
                    Case "TMPPTN1"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMP_PTNCODE()
                        End If
                        '----- yamamoto end ------
                    Case "TMPSIZE1"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPSIZE()
                        End If

                    Case "TMPKAJU1"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPKAJU()
                        End If

                    Case "TMPSERI "
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPSERI()
                        End If

                        ' -> watanabe add 2007.03
                        '                Case "TMPMAXLD"
                    Case "TMPMAXLD1"
                        ' <- watanabe add 2007.03
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPMAXLD()
                        End If

                        ' -> watanabe add 2007.03
                    Case "TMPMAXLD3"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPMAXLD()
                        End If
                        ' <- watanabe add 2007.03

                        '(Brand System Ver.3 追加)
                    Case "TMPPTN2"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMP_PTNCODE2()
                        End If

                        '(Brand System Ver.3 追加)
                    Case "TMPSIZE2"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPSIZE()
                        End If

                        '(Brand System Ver.3 追加)
                    Case "TMPKAJ2S"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPKAJU()
                        End If
                        '2015/1/28 moriya add start
                    Case "TMPMARK"
                        If VB.Left(Command_Line, 7) = "SPEC501" Then
                            form_main.Text2.Text = ""
                            hex_data = Mid(Command_Line, 9, Len(Command_Line) - 8)
                            open_mode = ""
                            w_ret = temp_bz_set(0, hex_data)
                            dataset_F_TMPMARK()
                        End If

                        '2015/1/28 moriya add end
                    Case Else
                        MsgBox("Do not understand. . . (" & ScreenName & ")," & Len(ScreenName))
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
					form_main.Text2.Text = ""
				Else
                    MsgBox("Not a free picture information [" & Command_Line & "]")
					End
				End If

                '2014/12/15 moriya add start
            Case comMark
                Exit Sub
                '2014/12/15 moriya add end

            Case comPTNCODE
                Exit Sub

                ' -> watanabe add VerUP(2011)
            Case comTmpWait
                Exit Sub
                ' <- watanabe add VerUP(2011)

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