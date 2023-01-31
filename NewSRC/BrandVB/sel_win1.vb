Option Strict Off
Option Explicit On
Friend Class sel_win1
	Inherits System.Windows.Forms.Form
	
	Private Sub Command_Click()
		
		Select Case mode.Text
            Case "Primitive character"
				form_no = F_GMSAVE
			Case "Editing characters"
				form_no = F_HMSAVE
		End Select
		form_no.Show()
		Me.Close()
		
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		open_mode = Command1.Text

		Call Command_Click()

	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		open_mode = Command2.Text
		
		Call Command_Click()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		End
	End Sub
	
	Private Sub sel_win1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'----- .NET�ڍs (StartPosition�v���p�e�B��CenterScreen�őΉ�) -----
		'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B

	End Sub
	
    'UPGRADE_WARNING: �C�x���g mode.TextChanged �́A�t�H�[�������������ꂽ�Ƃ��ɔ������܂��B
	Private Sub mode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mode.TextChanged
		
		Select Case mode.Text
            Case "Primitive character", "Editing characters"
                Command1.Visible = True
				Command1.Text = "NEW"

				'----- .NET�ڍs  -----
				'Command1.Left = VB6.TwipsToPixelsX(480)
				Command1.Left = ConvTwipToPixel(Me, 480)

				'----- .NET�ڍs  -----
				'Command1.Width = VB6.TwipsToPixelsX(1935)
				Command1.Width = ConvTwipToPixel(Me, 1935)

				Command2.Visible = True
				Command2.Text = "Change"

				'----- .NET�ڍs  -----
				'Command2.Left = VB6.TwipsToPixelsX(3000)
				Command2.Left = ConvTwipToPixel(Me, 3000)

				'----- .NET�ڍs  -----
				'Command2.Width = VB6.TwipsToPixelsX(1935)
				Command2.Width = ConvTwipToPixel(Me, 1935)

		End Select
		
	End Sub
End Class