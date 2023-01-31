Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module MJ_Date
	Sub true_date(ByRef outdata As String)
		Dim mm, dd As String
		
        If Len(CStr(Month(Today))) = 1 Then
            mm = "0" & CStr(Month(Today))
        Else
            mm = CStr(Month(Today))
        End If

        If Len(CStr(VB.Day(Today))) = 1 Then
            dd = "0" & CStr(VB.Day(Today))
        Else
            dd = CStr(VB.Day(Today))
        End If
        outdata = CStr(Year(Today)) & mm & dd

	End Sub
End Module