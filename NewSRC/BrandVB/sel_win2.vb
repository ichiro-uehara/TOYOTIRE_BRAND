Option Strict Off
Option Explicit On
Friend Class sel_win2
	Inherits System.Windows.Forms.Form
	
	Private Sub Command_Click()
		
		Select Case mode.Text
            Case "Stamp drawing"
                form_no = F_GZSAVE
            Case "Editing characters drawing"
                form_no = F_HZSAVE
            Case "Brand drawing"
                form_no = F_BZSAVE
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
	
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		
		open_mode = Command4.Text
		
		Call Command_Click()
		
	End Sub
	
	Private Sub sel_win2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' フォームを画面の水平方向にセンタリングします。
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' フォームを画面の縦方向にセンタリングします。
		
	End Sub
	
    'UPGRADE_WARNING: イベント mode.TextChanged は、フォームが初期化されたときに発生します。
	Private Sub mode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mode.TextChanged
		
		Select Case mode.Text
            Case "Primitive character", "Editing characters"
                Command1.Visible = True
                Command1.Text = "NEW"
                Command1.Left = VB6.TwipsToPixelsX(480)
                Command1.Width = VB6.TwipsToPixelsX(1935)

                Command2.Visible = True
                Command2.Text = "Change"
                Command2.Left = VB6.TwipsToPixelsX(3000)
                Command2.Width = VB6.TwipsToPixelsX(1935)

                Command4.Visible = False
            Case "Stamp drawing", "Editing characters drawing", "Brand drawing"
                Command1.Visible = True
                Command1.Text = "NEW"
                Command1.Left = VB6.TwipsToPixelsX(480)
                Command1.Width = VB6.TwipsToPixelsX(1092)

                Command2.Visible = True
                Command2.Text = "Revision number"
                Command2.Left = VB6.TwipsToPixelsX(2160)
                Command2.Width = VB6.TwipsToPixelsX(1092)

                Command4.Visible = True
                Command4.Text = "modify"
                Command4.Left = VB6.TwipsToPixelsX(3840)
                Command4.Width = VB6.TwipsToPixelsX(1092)
        End Select
		
	End Sub
End Class