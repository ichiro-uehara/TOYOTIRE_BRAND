Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class F_HMSAVE
	Inherits System.Windows.Forms.Form
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim result As Object
		
		Dim w_ret As Short
		Dim w_mess As String
		Dim ZumenName As String
		Dim TiffFile As String
		Dim tmpTiffFile As String
		Dim w_file As String
		
		On Error Resume Next ' �G���[�̃g���b�v�𗯕ۂ��܂��B
		Err.Clear()
		
		init_sql()
		
		If check_F_HMSAVE <> 0 Then
			end_sql()
			Exit Sub
		Else
            If open_mode = "NEW" Then
                result = hm_insert()
            Else
                result = hm_update()
            End If
			If result = FAIL Then
                MsgBox("Failed to register the editing characters.", 64, "registration error")
			Else
                MsgBox("Registered the editing characters.")
                ZumenName = "HM-" & Trim(form_no.w_font_name.Text)
				
				'Brand Ver.5 TIFF->BMP �ύX start
				'        TiffFile = TIFFDir & Trim(form_no.w_font_name) & Left$(Trim(form_no.w_no), 2) & ".tif"
				'        tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".tif"
				'        FileCopy tmpTiffFile, TiffFile
				'        If Err.Number <> 0 Then
				'           MsgBox "error_no: " & Str(Err.Number) & Err.Description, , "file error"
				'        End If
				'
				'        'Tiff̧�ٕ\��
				'        w_file = Dir(TiffFile)
				'        If w_file <> "" Then
				'            form_no.ImgThumbnail1.Image = tmpTiffFile
				'            form_no.ImgThumbnail1.ThumbWidth = 500
				'            form_no.ImgThumbnail1.ThumbHeight = 200
				'        Else
				'            MsgBox "TIFF̧�ق�������܂���", vbCritical, "File not found"
				'        End If
                TiffFile = TIFFDir & Trim(form_no.w_font_name.Text) & VB.Left(Trim(form_no.w_no.Text), 2) & ".bmp"
				tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
				FileCopy(tmpTiffFile, TiffFile)
				If Err.Number <> 0 Then
					MsgBox("error_no: " & Str(Err.Number) & Err.Description,  , "file error")
				End If
				
				'Tiff̧�ٕ\��
                w_file = Dir(TiffFile)
				If w_file <> "" Then
                    form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
					form_no.ImgThumbnail1.ScaleWidth = 500
					form_no.ImgThumbnail1.ScaleHeight = 200
				Else
                    MsgBox("BMP file can not be found.", MsgBoxStyle.Critical, "File not found")
				End If
				'Brand Ver.5 TIFF->BMP �ύX end
				
				'�����f�[�^���M
				w_ret = temp_hm_get()
				'�i�}�ʖ��A�z�uPIC�j���M

				'----- .NET �ڍs -----
				'w_mess = VB6.Format(VB.Left(form_no.w_haiti_pic.Text, 3), "000") & HensyuDir & ZumenName
				w_mess = String.Format("000", Strings.Left(form_no.w_haiti_pic.Text, 3)) & HensyuDir & ZumenName

				w_ret = PokeACAD("ACADSAVE", w_mess)
				w_ret = RequestACAD("ACADSAVE")
				
				'��ʃ��b�N
                form_no.Command1.Enabled = False
				form_no.Command2.Enabled = False
				form_no.Command4.Enabled = False
				form_no.w_font_name.Enabled = False
                form_no.w_font_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX
				form_no.w_spell.Enabled = False
                form_no.w_spell.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				form_no.w_comment.Enabled = False
                form_no.w_comment.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				form_no.w_dep_name.Enabled = False
                form_no.w_dep_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				form_no.w_entry_name.Enabled = False
                form_no.w_entry_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				form_no.w_high.Enabled = False
                form_no.w_high.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
				form_no.w_ang.Enabled = False
                form_no.w_ang.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0)
			End If
			
		End If
		
		end_sql()
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		
		Call Clear_F_HMSAVE()
		
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		
		form_no.Close()
		End
		
		'form1.Show
		
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
                .HelpContext = 400
                .ShowHelp()
            End With
        End If
	End Sub
	
	Private Sub F_HMSAVE_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim tmpTiffFile As String
		Dim w_ret As Object
		
        Dim aa As String

        ' -> watanabe add VerUP(2011)
        aa = ""
        ' <- watanabe add VerUP(2011)

        form_no = Me
        temp_hm.Initilize() '20100702�ǉ��R�[�h

        If open_mode = "NEW" Then
            w_ret = PokeACAD("SAVEMODE", "FRESH")
            RequestACAD("SAVEMODE")

            'Brand Ver.5 TIFF->BMP �ύX start
            '     'Tiff̧��
            '     tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".tif"
            '     form_no.ImgThumbnail1.Image = tmpTiffFile
            '     form_no.ImgThumbnail1.ThumbWidth = 500
            '     form_no.ImgThumbnail1.ThumbHeight = 200
            'BMP̧��
            tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(tmpTiffFile)
            form_no.ImgThumbnail1.Width = System.Drawing.Image.FromFile(tmpTiffFile).Width '500 '20100701�R�[�h�ύX
            form_no.ImgThumbnail1.Height = System.Drawing.Image.FromFile(tmpTiffFile).Height '200 '20100701�R�[�h�ύX
            'Brand Ver.5 TIFF->BMP �ύX end
            Call true_date(aa)
            temp_hm.entry_date = aa
        Else
            w_ret = PokeACAD("SAVEMODE", "MODIFY")
            RequestACAD("SAVEMODE")
            form_no.w_font_name.Enabled = False
            form_no.w_font_name.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0C0C0) '20100629�R�[�h�ύX

            'Brand Ver.5 TIFF->BMP �ύX start
            '     tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".tif"
            '     form_no.ImgThumbnail1.Image = tmpTiffFile
            '     form_no.ImgThumbnail1.ThumbWidth = 500
            '     form_no.ImgThumbnail1.ThumbHeight = 200
            tmpTiffFile = TMPTIFFDir & TmpTIFFName & ".bmp"
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(tmpTiffFile)
            form_no.ImgThumbnail1.Width = System.Drawing.Image.FromFile(tmpTiffFile).Width '500 '20100701�R�[�h�ύX
            form_no.ImgThumbnail1.Height = System.Drawing.Image.FromFile(tmpTiffFile).Height '200 '20100701�R�[�h�ύX
            'Brand Ver.5 TIFF->BMP �ύX end

            Call true_date(aa)
            temp_hm.entry_date = aa

        End If

		'----- .NET�ڍs (StartPosition�v���p�e�B��CenterScreen�őΉ�) -----
		'Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Width)) / 2) ' �t�H�[������ʂ̐��������ɃZ���^�����O���܂��B
		'Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Height)) / 2) ' �t�H�[������ʂ̏c�����ɃZ���^�����O���܂��B

		form_no.Text1.Text = open_mode
		
		Call Clear_F_HMSAVE()

		CommunicateMode = comSpecData

		RequestACAD("SPECDATA")

	End Sub

	'----- .NET�ڍs (ToDo:DataGridView�̃C�x���g�ɕύX) -----
#If False Then
	Private Sub MSFlexGrid1_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_KeyPressEvent) Handles MSFlexGrid1.KeyPressEvent
		
        MsgBox("You can not change the key input.", 64)
		eventArgs.KeyAscii = 0
		
	End Sub
#End If

	Private Sub w_comment_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_comment.Leave
		
		form_no.w_comment.Text = apos_check(form_no.w_comment.Text)
		
	End Sub
	
	Private Sub w_dep_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_dep_name.Leave
		
		form_no.w_dep_name.Text = UCase(Trim(form_no.w_dep_name.Text))
		
	End Sub
	
	Private Sub w_font_name_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_font_name.Leave
		
		form_no.w_font_name.Text = UCase(Trim(form_no.w_font_name.Text))
		
	End Sub
	
	Private Sub w_spell_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles w_spell.Leave
		
		form_no.w_spell.Text = apos_check(form_no.w_spell.Text)
		
	End Sub

	'----- .NET�ڍs  -----
	'DataGridViewList CellPainting�C�x���g
	'�s�ԍ���`�悷��
	Private Sub DataGridViewList_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridViewList.CellPainting

		Try

			If e.ColumnIndex < 0 And e.RowIndex >= 0 Then
				'�Z����`�悷��
				e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

				'�s�ԍ���`�悷��͈͂����肷��
				Dim indexRect As Rectangle = e.CellBounds

				indexRect.Inflate(-2, -2)
				'�s�ԍ���`�悷��
				TextRenderer.DrawText(e.Graphics,
					(e.RowIndex + 1).ToString(),
					e.CellStyle.Font,
					indexRect,
					e.CellStyle.ForeColor,
					TextFormatFlags.Right Or TextFormatFlags.VerticalCenter)
				'�`�悪�����������Ƃ�m�点��
				e.Handled = True
			End If

		Catch ex As Exception

			MessageBox.Show(ex.Message, "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error)

		End Try

	End Sub
End Class