Option Strict Off
Option Explicit On
Module MJ_DLL
	'Declare Function LCopy Lib "branddll.dll" (ByRef lpString As Byte, ByVal iMaxLength As Long) As Long
	'Declare Function SCopy Lib "branddll.dll" (ByVal lpString As String, ByVal iMaxLength As Long) As Single
	
	'short変数(2byte) -> ヘキサ文字(4byte)
	Declare Function ShttoHex Lib "BrandDLL.DLL" (ByVal insht As Short, ByVal lpString As String) As Integer
	'int変数(4byte) -> ヘキサ文字(8byte)
	Declare Function InttoHex Lib "BrandDLL.DLL" (ByVal inint As Integer, ByVal lpString As String) As Integer
	'float変数(4byte) -> ヘキサ文字(8byte)
	Declare Function FlttoHex Lib "BrandDLL.DLL" (ByVal inflt As Single, ByVal lpString As String) As Integer
	'double変数(8byte) -> ヘキサ文字(16byte)
	Declare Function DbltoHex Lib "BrandDLL.DLL" (ByVal indbl As Double, ByVal lpString As String) As Integer
	'ヘキサ文字(4byte) -> short変数(2byte)
	Declare Function HextoSht Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outsht As Short) As Integer
	'ヘキサ文字(8byte) ->  int変数(4byte)
	Declare Function HextoInt Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outint As Integer) As Integer
	'ヘキサ文字(8byte) -> float変数(4byte)
	Declare Function HextoFlt Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outflt As Single) As Integer
	'ヘキサ文字(16byte)  -> double変数(8byte)
	Declare Function HextoDbl Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outdbl As Double) As Integer
	
	
	'////Declare Function ChrtoHex Lib "BrandDLL.DLL" (byval bitdat as string ) as
	'////Declare Function HextoChr Lib "BrandDLL.DLL" (ByVal bitdat As Integer) As String
End Module