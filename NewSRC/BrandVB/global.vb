Option Strict Off
Option Explicit On
Module MJ_Global

    ' -> watanabe del VerUP(2011)
    'ウインドウズのプログラムを呼び出すための宣言
    'Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    ' <- watanabe del VerUP(2011)

    'Public form_no As System.Windows.Forms.Form
    Public form_no As Object
    'Public form_main As System.Windows.Forms.Form
    Public form_main As Object
	
	Public dummy_text As String
	Public open_mode As String
	Public ScreenName As String
	
	'Global oSQLServer As New SQLOLE.SQLServer
	'Global oDatabase As SQLOLE.DataBase
	'Global oTable As Object
	
	Public read_type As New VB6.FixedLengthString(6)

    '// SQL変数

    '----- .NET移行 (SqlConn の初期値()を設定)-----
    Public SqlConn As Integer
    Public SqlCur As Integer
	
	
	'AdvanceCADとの通信用変数
	Public TransAppName As String
	Public TransTopic As String
	Public TransItem As String
	
	'通信モード変数
	Public CommunicateMode As Byte

    '空きピクチャ数
    '----- .NET移行 (暫定的に99件とする ⇒ ToDo:仕様確認後に対応)-----
    Public FreePicNum As Byte = 99


    Public Const comNone As Short = 1 '処理待ちなし
	Public Const comWinName As Short = 2 '画面名待ち
	Public Const comSpecData As Short = 3 '特性データ待ち
	Public Const comFreePic As Short = 4 '空きピクチャ待ち
	Public Const comCodeHyo As Short = 5 'コード表データ待ち
	Public Const batchTypeRead As Short = 6
	Public Const batchTmpRead As Short = 7
	Public Const batchHmRead As Short = 8
	'----- 12/16 1997 yamamoto start-----
	Public Const comPTNCODE As Short = 9 'パターンコード画面PICEMPTY待ち
	'----- 12/16 1997 yamamoto end -------
	'Brand Ver.3 追加
	Public Const comCodeHyo2 As Short = 10

    ' -> watanabe add VerUP(2011)
    Public Const comTmpWait As Short = 11   'テンプレート変換待機
    ' <- watanabe add VerUP(2011)
    Public Const comMark As Short = 12 '次のピクチャ読み込み '2014/12/15 moriya add
	
	'エラー定数
	Public Const errNoAppResponded As Short = 282
    Public Const errDDERefused As Short = 285

    '20100618移植追加 ヘルプ用
    Public Const cdlHelpContext As Integer = &H1&
    Public InitFlag As Boolean = False '20100628追加コード

    
    'ファイルディレクトリ （設定ファイルより）'20100708 定義Obj->Str変更
    Public GensiDir As String
    Public HensyuDir As String
    Public KokuinDir As String
    Public HensyuZumenDir As String
    Public BrandDir As String
    Public TIFFDir As String
    Public TIFFDirGM As String
    Public TIFFDirHM As String
    Public HelpFileName As String

    'テンプレート設定ファイル名（設定ファイルより）'20100708 定義Obj->Str変更
    Public Tmp_Size1_ini As String
    Public Tmp_Load1_ini As String
    Public Tmp_Pattern1_ini As String
    Public Tmp_Serial1_ini As String
    Public Tmp_Mold_no1_ini As String
    Public Tmp_E_no1_ini As String
    Public Tmp_Utqg1_ini As String
    Public Tmp_Maxload1_ini As String
    Public Tmp_Ply1_ini As String
    Public Tmp_Ply2_ini As String
    Public Tmp_ETC_ini As String
    Public Tmp_Size2_ini As String
    Public Tmp_Load2S_ini As String
    Public Tmp_Load2D_ini As String
    Public Tmp_Pattern2_ini As String
    Public Tmp_Lt2_ini As String
    Public Tmp_Pr2_ini As String
    Public Tmp_Psi2_ini As String
    Public Tmp_Plate_ini As String
    Public Tmp_MARK_ini As String       '2014/12/15 moriya add (1行)

    ' -> watanabe add 2007.03  '20100708 定義Obj->Str変更
    Public Tmp_Utqg3_ini As String
    Public Tmp_Maxload3_ini As String
    Public Tmp_Ply1_3_ini As String
    Public Tmp_Ply2_3_ini As String
    Public Tmp_ETC3_ini As String
	' <- watanabe add 2007.03
	
	
	'テンプレート タイプ２ 作成用共通データ
	Public Tmp2_Dummy_HM As String
	
	'データベース定数 （設定ファイルより）
	Public DBServer As String
	Public DBLoginID As String
	Public DBpasswd As String
	Public DBexample As String
	Public DBName As String
	Public DBTableName As String
	Public STANDARD_DBName As String
	'Brand Ver.3 追加
	Public DBTableName2 As String
	
	'システム定数 （設定ファイルより）
	'タイムアウト秒数
	Public timeOutSecond As Short
	'ACAD通信DDE
	Public ACADTransAppName As String
	Public ACADTransTopic As String
	Public ACADTransItem As String
	'TIFFﾌｧｲﾙ
	Public TMPTIFFDir As String
    Public TmpTIFFName As String
	
	'テンプレート設定 （設定ファイルより）
	Public Const MaxSelNum As Short = 256
	Public Tmp_font_cnt As Short
	Public Tmp_font_word(MaxSelNum) As String
    Public Tmp_font_size(MaxSelNum) As Double
	Public Tmp_font_block(MaxSelNum) As Short
	Public ReplaceMode As Object
	Public Tmp_hm_word(MaxSelNum) As String
	Public Tmp_hm_code(MaxSelNum) As String
	' -> watanabe add 2007.03
	Public Tmp_hm_group(MaxSelNum) As String
	' <- watanabe add 2007.03
	Public Tmp_rule_word(MaxSelNum) As String
	Public Tmp_rule_type(MaxSelNum) As String
	Public Tmp_rule_x(MaxSelNum) As Double
	Public Tmp_rule_y(MaxSelNum) As Double
	Public GensiKIGO(128) As String
	Public GensiALPH(26) As String
	Public GensiNUM(10) As String
    Public AskNum As Object
	'Brand Ver.4 追加
	Public Tmp_prcs_code(MaxSelNum) As String
	Public GensiALPHS(26) As String '小文字の入力
	
	' -> watanabe add 2007.06
	Public Tmp_brd_no As Short
	' <- watanabe add 2007.06
	
	
	'ｾﾘｱﾙの寸法
	Public TmpSerialWidth As Double 'セリアルのコード文字全体幅
	Public TmpSerialMove As Double 'セリアルのコード文字プレート間
	
	' -> watanabe Add 2007.03
	' プレート変型データ
	Public Tmp_plate_w As Double '幅
	Public Tmp_plate_h As Double '高さ
	Public Tmp_plate_r As Double 'コーナーＲ
	Public Tmp_plate_n As Double 'ネジ位置
	' <- watanabe Add 2007.03
	
	
	Structure GM_KANRI ' 原始文字
		Dim flag_delete As Byte ' データ型の要素を定義します。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        '<VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id() As Char
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=6)> Public font_name As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public font_class1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public font_class2 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public name1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public name2 As String
		Dim high As Double
		Dim width As Double
		Dim ang As Double
		Dim moji_high As Double
		Dim moji_shift As Double
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public org_hor As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public org_ver As String
		Dim org_x As Double
		Dim org_y As Double
		Dim left_bottom_x As Double
		Dim left_bottom_y As Double
		Dim right_bottom_x As Double
		Dim right_bottom_y As Double
		Dim right_top_x As Double
		Dim right_top_y As Double
		Dim left_top_x As Double
		Dim left_top_y As Double
		Dim hem_width As Double
		Dim hatch_ang As Double
		Dim hatch_width As Double
		Dim hatch_space As Double
		Dim hatch_x As Double
		Dim hatch_y As Double
		Dim base_r As Double
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=6)> Public old_font_name As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public old_font_class As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public old_name As String
		Dim haiti_pic As Short
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public gz_id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public gz_no1 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public gz_no2 As String
		Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
		'    entry_date As Long
		Dim entry_date As String
	End Structure
	
	Structure HM_KANRI ' 編集文字
		Dim flag_delete As Byte ' データ型の要素を定義します。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=6)> Public font_name As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no As String
		Dim spell As String
		Dim haiti_sitei As Short
		Dim gm_num As Short
		Dim width As Double
		Dim high As Double
		Dim ang As Double
		Dim haiti_pic As Short
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public hz_id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public hz_no1 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public hz_no2 As String
		Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim gm_name() As String '20100616移植追加
        Public Sub Initilize() '20100616移植追加
            ReDim gm_name(500)
        End Sub
    End Structure

    Structure GZ_KANRI ' 刻印図面
        Dim flag_delete As Byte ' データ型の要素を定義します。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public no1 As String
        '<VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no2() As Char
        Dim no2 As String '20100618型変更
        Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim gm_num As Short
        'Dim gm_name(65) As String*10
        Dim gm_name() As String '20100616移植追加
        Public Sub Initilize() '20100616移植追加
            ' -> watanabe edit 2013.05.29
            'ReDim gm_name(65)
            ReDim gm_name(130)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure HZ_KANRI ' 編集文字図面
        Dim flag_delete As Byte ' データ型の要素を定義します。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public no1 As String
        '<VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no2() As Char
        Dim no2 As String '20100618型変更
        Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim hm_num As Short
        'Dim hm_name(65) As String*8
        Dim hm_name() As String '20100616移植追加
        Public Sub Initilize() '20100616移植追加
            ' -> watanabe edit 2013.05.29
            'ReDim hm_name(65)
            ReDim hm_name(130)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure BZ_KANRI ' ブランド図面
        Dim flag_delete As Byte
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public id As String
        ' -> watanabe edit 2007.03
        '    no1 As String * 4
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public no1 As String
        ' <- watanabe edit 2007.03
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no2 As String
        <VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=8)> Public kanri_no As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public syurui As String
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=3)> Public syubetu As String
        <VBFixedString(6), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=6)> Public pattern As String
        <VBFixedString(21), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=21)> Public Size As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size2 As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size3 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size4 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size5 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public size6 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public size7 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public size8 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public size_code As String
        <VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=10)> Public kikaku As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public plant As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public plant_code As String
        Dim tos_moyou As Short
        Dim side_moyou As Short
        Dim side_kenti As Short
        Dim peak_mark As Short
        Dim nasiji As Short
        Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim hm_num As Short
        'UPGRADE_ISSUE: 宣言の型がサポートされていません: 固定長文字列の配列 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"' をクリックしてください。
        'Dim hm_name(100) As String*8
        Dim hm_name() As String '20100616移植追加
        Public Sub Initilize() '20100616移植追加
            ' -> watanabe edit 2013.05.29
            'ReDim hm_name(100)
            ReDim hm_name(260)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure T_SIZE_CODE
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public syurui As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size2 As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size3 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size4 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size5 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public size6 As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public size_code As String
    End Structure

    Structure T_JETMA
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public syurui As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size2 As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size3 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size4 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size5 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public size6 As String
        Dim standard_load_index As Byte
    End Structure

    Structure T_TRA
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public syurui As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size2 As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size3 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size4 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size5 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public size6 As String
        Dim standard_load_index As Byte
        Dim standard_max_load_kg As Short
        Dim standard_max_load_lbs As Short
        Dim standard_max_press_kpa As Short
        Dim standard_max_press_psi As Byte
        Dim extra_load_index As Byte
        Dim extra_max_load_kg As Short
        Dim extra_max_load_lbs As Short
        Dim extra_max_press_kpa As Short
        Dim extra_max_press_psi As Byte
    End Structure

    Structure T_ETRTO
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。 
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public syurui As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size1 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size2 As String
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public size3 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size4 As String
        <VBFixedString(1), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=1)> Public size5 As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public size6 As String
        Dim standard_load_index As Byte
        Dim extra_load_index As Byte
    End Structure

    Structure TYPE_TEST
        'UPGRADE_WARNING: 固定長文字列のサイズはバッファに合わせる必要があります。
        '20100706コード変更　...変換でChar型になっていたものをSrring型に修正
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public start_data As String
        Dim i_data As Short
        '    l_data As Long
        Dim f_data As Single
        Dim d_data As Double
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=3)> Public end_data As String
    End Structure

    Public temp_gm As GM_KANRI
    'UPGRADE_WARNING: 構造体の配列は、使用する前に初期化する必要があります。>mainのLoadで初期化
    Public temp_hm As HM_KANRI
	Public temp_gz As GZ_KANRI
	Public temp_hz As HZ_KANRI
    Public temp_bz As BZ_KANRI
	Public temp_size As T_SIZE_CODE
	Public temp_jetma As T_JETMA
	Public temp_tra As T_TRA
	Public temp_etrto As T_ETRTO
	Public grid_num As Short
	
	'----- 12/8 1997 yamamoto start -----
	'BrandVB2でﾛｽﾄﾌｫｰｶｽ、ｾｯﾄﾌｫｰｶｽ問題点のためﾌﾗｸﾞを追加します
	Public w_no1_flg As Short 'フラグ(OK = 0,NG = 1)
	Public w_no2_flg As Short 'フラグ(OK = 0,NG = 1)
	
	'ﾌﾗｸﾞの追加（画面ｸﾞﾘｯﾄﾞﾃﾞｰﾀ）
	Public grd_clear_flg As Short '(ｸﾞﾘｯﾄﾞｸﾘｱ = 0、not_clear = 1 )
	
	'----- 1/27 1998 yamamoto --------
	'ｷｬﾝｾﾙﾌﾗｸﾞ
	Public GL_cancel_flg As Short 'キャンセルフラグ ( 0 = 初期値 or OFF  1 = ON ）


    ' -> watanabe edit VerUP(2011)

    '----- .NET移行 (ADO接続) -----
    'Public GL_T_RDO As T_RDO_Struct 'ＲＤＯ接続用
    'Public GL_RDOEnv As RDO.rdoEnvironment

    Public GL_T_ADO As T_ADO_Struct 'ADO接続用

    Public Const DEF_FUNC_NORMAL As Integer = 1 '関数戻り値
    Public Const DEF_FUNC_ERROR As Integer = 0 '関数戻り値

    Public Const SUCCEED As Integer = 1
    Public Const FAIL As Integer = 0
    ' <- watanabe edit VerUP(2011)

End Module