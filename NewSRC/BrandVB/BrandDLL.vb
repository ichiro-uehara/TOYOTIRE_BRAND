Option Strict Off
Option Explicit On
Module MJ_DLL
	'Declare Function LCopy Lib "branddll.dll" (ByRef lpString As Byte, ByVal iMaxLength As Long) As Long
	'Declare Function SCopy Lib "branddll.dll" (ByVal lpString As String, ByVal iMaxLength As Long) As Single
	
	'short�ϐ�(2byte) -> �w�L�T����(4byte)
	Declare Function ShttoHex Lib "BrandDLL.DLL" (ByVal insht As Short, ByVal lpString As String) As Integer
	'int�ϐ�(4byte) -> �w�L�T����(8byte)
	Declare Function InttoHex Lib "BrandDLL.DLL" (ByVal inint As Integer, ByVal lpString As String) As Integer
	'float�ϐ�(4byte) -> �w�L�T����(8byte)
	Declare Function FlttoHex Lib "BrandDLL.DLL" (ByVal inflt As Single, ByVal lpString As String) As Integer
	'double�ϐ�(8byte) -> �w�L�T����(16byte)
	Declare Function DbltoHex Lib "BrandDLL.DLL" (ByVal indbl As Double, ByVal lpString As String) As Integer
	'�w�L�T����(4byte) -> short�ϐ�(2byte)
	Declare Function HextoSht Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outsht As Short) As Integer
	'�w�L�T����(8byte) ->  int�ϐ�(4byte)
	Declare Function HextoInt Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outint As Integer) As Integer
	'�w�L�T����(8byte) -> float�ϐ�(4byte)
	Declare Function HextoFlt Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outflt As Single) As Integer
	'�w�L�T����(16byte)  -> double�ϐ�(8byte)
	Declare Function HextoDbl Lib "BrandDLL.DLL" (ByVal lpString As String, ByRef outdbl As Double) As Integer
	
	
	'////Declare Function ChrtoHex Lib "BrandDLL.DLL" (byval bitdat as string ) as
	'////Declare Function HextoChr Lib "BrandDLL.DLL" (ByVal bitdat As Integer) As String
End Module