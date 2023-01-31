Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_MAIN1
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MAIN1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim ret As Short
		Dim w_w_str As String

		form_main = Me

		'----- .NET移行 (StartPositionプロパティをCenterScreenで対応) -----
		'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。

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

		'----- .NET移行 (DDE通信コメント化) -----
		CommunicateMode = comWinName
		ret = RequestACAD("WINNAME")

		'「原始文字検索 / 編集文字検索」でデバッグ（ToDo:デバッグ後削除）
		'CommunicateMode = 2
		'Text2.Text = "GMSEARCH"
		'Text2.Text = "HMSEARCH1"
		'デバッグ（ToDo:デバッグ後削除）

		Exit Sub

error_section: 
		
        MsgBox("To exit", MsgBoxStyle.Critical, "Error end")
		End
		
	End Sub

    Private Sub Form_Terminate_Renamed()

        'SQL接続をｸﾛｰｽﾞします

        ' -> watanabe edit VerUP(2011)
        'SqlExit()
        end_sql()
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
        '  output_command_line (Command_Line) '----- 12/11 1997 yamamoto add (debug)-----

        '   規格チェックエラー
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
			'送信待ちなし
			Case comNone
				If VB.Left(Command_Line, 6) = "VBKILL" Then
					'           MsgBox "VBKILL受信しました" & Chr(13) & "BrandVBを終了します"
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
					TIFFDir = TIFFDirGM
				ElseIf VB.Left(Command_Line, 2) = "HM" Then
					DBTableName = DBName & "..hm_kanri1" '編集文字管理(基本部)
					DBTableName2 = DBName & "..hm_kanri2" '編集文字管理(文字部)
					TIFFDir = TIFFDirHM
				ElseIf VB.Left(Command_Line, 2) = "GZ" Then
					DBTableName = DBName & "..gz_kanri1" '刻印図面管理(基本部)
					DBTableName2 = DBName & "..gz_kanri2" '刻印図面管理(文字部)
				ElseIf VB.Left(Command_Line, 2) = "HZ" Then
					DBTableName = DBName & "..hz_kanri1" '編集文字図面管理(基本部)
					DBTableName2 = DBName & "..hz_kanri2" '編集文字図面管理(文字部)
				ElseIf VB.Left(Command_Line, 2) = "BZ" Then
					DBTableName = DBName & "..bz_kanri1" 'ブランド図面管理(基本部)
					DBTableName2 = DBName & "..bz_kanri2" 'ブランド図面管理(文字部)
				End If
				If Len(Command_Line) < 8 Then
					ScreenName = Command_Line & Space(8 - Len(Command_Line))
				End If
				
				'================
				'画面呼び出し
				'================
				'/// 原始文字登録画面
				If VB.Left(Command_Line, 6) = "GMSAVE" Then
					'               ScreenName = "GMSAVE  "
					CommunicateMode = comNone
					sel_win1.Show()
					sel_win1.mode.Text = "Primitive character"
					form_main.Text2.Text = ""
					'// 原始文字削除画面
				ElseIf VB.Left(Command_Line, 6) = "GMDELE" Then
					'ScreenName = "GMDELE  "
					'CommunicateMode = comNone
					'F_GMDELE.Show()
					'form_main.Text2.Text = ""
					'// 原始文字検索画面
				ElseIf VB.Left(Command_Line, 8) = "GMSEARCH" Then 
					ScreenName = "GMSEARCH"
					CommunicateMode = comFreePic
					'----- .NET移行 (暫定的に99件とする ⇒ ToDo:仕様確認後に対応)-----
					FreePicNum = 0

					F_GMSEARCH.Show()
					form_main.Text2.Text = ""
					w_ret = RequestACAD("PICEMPTY")
					'// 原始文字内容確認画面
				ElseIf VB.Left(Command_Line, 6) = "GMLOOK" Then
					'ScreenName = "GMLOOK  "
					'CommunicateMode = comNone
					'F_GMLOOK.Show()
					'form_main.Text2.Text = ""
					'// 原始文字読込み
				ElseIf VB.Left(Command_Line, 6) = "GMREAD" Then
					'CommunicateMode = comFreePic
					'FreePicNum = 0
					'               open_mode = "Primitive character"
					'F_CADREAD.Show()
					'form_main.Text2.Text = ""
					'w_ret = RequestACAD("PICEMPTY")
					'// 編集文字登録画面
				ElseIf VB.Left(Command_Line, 6) = "HMSAVE" Then
					ScreenName = "HMSAVE  "
					CommunicateMode = comNone
					sel_win1.Show()
					sel_win1.mode.Text = "Editing characters"
					form_main.Text2.Text = ""
					'// 編集文字削除画面
				ElseIf VB.Left(Command_Line, 6) = "HMDELE" Then
					'ScreenName = "HMDELE  "
					'CommunicateMode = comNone
					'F_HMDELE.Show()
					'form_main.Text2.Text = ""
					'// 編集文字検索画面
				ElseIf VB.Left(Command_Line, 9) = "HMSEARCH1" Then
					ScreenName = "HMSEARCH1"
					CommunicateMode = comFreePic
					'----- .NET移行 (暫定的に99件とする ⇒ ToDo:仕様確認後に対応)-----
					'FreePicNum = 0

					F_HMSEARCH.Show()
					form_main.Text2.Text = ""
					w_ret = RequestACAD("PICEMPTY")
					'Brand Ver.4 追加
					'// 編集文字検索画面
				ElseIf VB.Left(Command_Line, 9) = "HMSEARCH2" Then
					'ScreenName = "HMSEARCH2"
					'CommunicateMode = comFreePic
					'FreePicNum = 0
					'F_HMSEARCH2.Show()
					'form_main.Text2.Text = ""
					'w_ret = RequestACAD("PICEMPTY")
					'追加終了
					'// 編集文字内容確認画面
				ElseIf VB.Left(Command_Line, 6) = "HMLOOK" Then
					'ScreenName = "HMLOOK  "
					'CommunicateMode = comNone
					'F_HMLOOK.Show()
					'form_main.Text2.Text = ""
					'// 編集文字読込み
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
				'特性データ到着待ち時
			Case comSpecData
				
				Select Case ScreenName
					'原始文字名
					
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

					' MsgBox "空きピクチャ＝" & FreePicNum
					
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