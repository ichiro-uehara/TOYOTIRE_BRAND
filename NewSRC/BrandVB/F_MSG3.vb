Option Strict Off
Option Explicit On
Friend Class F_MSG3
	Inherits System.Windows.Forms.Form
	
	Private Sub F_MSG3_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B
		Cursor = System.Windows.Forms.Cursors.WaitCursor
	End Sub
	
	
	
	Private Sub Label2_Click()
		
	End Sub
End Class