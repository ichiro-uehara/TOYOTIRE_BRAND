Option Strict Off
Option Explicit On
Module MJ_Global

    ' -> watanabe del VerUP(2011)
    '�E�C���h�E�Y�̃v���O�������Ăяo�����߂̐錾
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

    '// SQL�ϐ�

    '----- .NET�ڍs (SqlConn �̏����l()��ݒ�)-----
    Public SqlConn As Integer
    Public SqlCur As Integer
	
	
	'AdvanceCAD�Ƃ̒ʐM�p�ϐ�
	Public TransAppName As String
	Public TransTopic As String
	Public TransItem As String
	
	'�ʐM���[�h�ϐ�
	Public CommunicateMode As Byte

    '�󂫃s�N�`����
    '----- .NET�ڍs (�b��I��99���Ƃ��� �� ToDo:�d�l�m�F��ɑΉ�)-----
    Public FreePicNum As Byte = 99


    Public Const comNone As Short = 1 '�����҂��Ȃ�
	Public Const comWinName As Short = 2 '��ʖ��҂�
	Public Const comSpecData As Short = 3 '�����f�[�^�҂�
	Public Const comFreePic As Short = 4 '�󂫃s�N�`���҂�
	Public Const comCodeHyo As Short = 5 '�R�[�h�\�f�[�^�҂�
	Public Const batchTypeRead As Short = 6
	Public Const batchTmpRead As Short = 7
	Public Const batchHmRead As Short = 8
	'----- 12/16 1997 yamamoto start-----
	Public Const comPTNCODE As Short = 9 '�p�^�[���R�[�h���PICEMPTY�҂�
	'----- 12/16 1997 yamamoto end -------
	'Brand Ver.3 �ǉ�
	Public Const comCodeHyo2 As Short = 10

    ' -> watanabe add VerUP(2011)
    Public Const comTmpWait As Short = 11   '�e���v���[�g�ϊ��ҋ@
    ' <- watanabe add VerUP(2011)
    Public Const comMark As Short = 12 '���̃s�N�`���ǂݍ��� '2014/12/15 moriya add
	
	'�G���[�萔
	Public Const errNoAppResponded As Short = 282
    Public Const errDDERefused As Short = 285

    '20100618�ڐA�ǉ� �w���v�p
    Public Const cdlHelpContext As Integer = &H1&
    Public InitFlag As Boolean = False '20100628�ǉ��R�[�h

    
    '�t�@�C���f�B���N�g�� �i�ݒ�t�@�C�����j'20100708 ��`Obj->Str�ύX
    Public GensiDir As String
    Public HensyuDir As String
    Public KokuinDir As String
    Public HensyuZumenDir As String
    Public BrandDir As String
    Public TIFFDir As String
    Public TIFFDirGM As String
    Public TIFFDirHM As String
    Public HelpFileName As String

    '�e���v���[�g�ݒ�t�@�C�����i�ݒ�t�@�C�����j'20100708 ��`Obj->Str�ύX
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
    Public Tmp_MARK_ini As String       '2014/12/15 moriya add (1�s)

    ' -> watanabe add 2007.03  '20100708 ��`Obj->Str�ύX
    Public Tmp_Utqg3_ini As String
    Public Tmp_Maxload3_ini As String
    Public Tmp_Ply1_3_ini As String
    Public Tmp_Ply2_3_ini As String
    Public Tmp_ETC3_ini As String
	' <- watanabe add 2007.03
	
	
	'�e���v���[�g �^�C�v�Q �쐬�p���ʃf�[�^
	Public Tmp2_Dummy_HM As String
	
	'�f�[�^�x�[�X�萔 �i�ݒ�t�@�C�����j
	Public DBServer As String
	Public DBLoginID As String
	Public DBpasswd As String
	Public DBexample As String
	Public DBName As String
	Public DBTableName As String
	Public STANDARD_DBName As String
	'Brand Ver.3 �ǉ�
	Public DBTableName2 As String
	
	'�V�X�e���萔 �i�ݒ�t�@�C�����j
	'�^�C���A�E�g�b��
	Public timeOutSecond As Short
	'ACAD�ʐMDDE
	Public ACADTransAppName As String
	Public ACADTransTopic As String
	Public ACADTransItem As String
	'TIFF̧��
	Public TMPTIFFDir As String
    Public TmpTIFFName As String
	
	'�e���v���[�g�ݒ� �i�ݒ�t�@�C�����j
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
	'Brand Ver.4 �ǉ�
	Public Tmp_prcs_code(MaxSelNum) As String
	Public GensiALPHS(26) As String '�������̓���
	
	' -> watanabe add 2007.06
	Public Tmp_brd_no As Short
	' <- watanabe add 2007.06
	
	
	'�رق̐��@
	Public TmpSerialWidth As Double '�Z���A���̃R�[�h�����S�̕�
	Public TmpSerialMove As Double '�Z���A���̃R�[�h�����v���[�g��
	
	' -> watanabe Add 2007.03
	' �v���[�g�ό^�f�[�^
	Public Tmp_plate_w As Double '��
	Public Tmp_plate_h As Double '����
	Public Tmp_plate_r As Double '�R�[�i�[�q
	Public Tmp_plate_n As Double '�l�W�ʒu
	' <- watanabe Add 2007.03
	
	
	Structure GM_KANRI ' ���n����
		Dim flag_delete As Byte ' �f�[�^�^�̗v�f���`���܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
	
	Structure HM_KANRI ' �ҏW����
		Dim flag_delete As Byte ' �f�[�^�^�̗v�f���`���܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        Dim gm_name() As String '20100616�ڐA�ǉ�
        Public Sub Initilize() '20100616�ڐA�ǉ�
            ReDim gm_name(500)
        End Sub
    End Structure

    Structure GZ_KANRI ' ����}��
        Dim flag_delete As Byte ' �f�[�^�^�̗v�f���`���܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public no1 As String
        '<VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no2() As Char
        Dim no2 As String '20100618�^�ύX
        Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim gm_num As Short
        'Dim gm_name(65) As String*10
        Dim gm_name() As String '20100616�ڐA�ǉ�
        Public Sub Initilize() '20100616�ڐA�ǉ�
            ' -> watanabe edit 2013.05.29
            'ReDim gm_name(65)
            ReDim gm_name(130)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure HZ_KANRI ' �ҏW�����}��
        Dim flag_delete As Byte ' �f�[�^�^�̗v�f���`���܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public id As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public no1 As String
        '<VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public no2() As Char
        Dim no2 As String '20100618�^�ύX
        Dim comment As String
        <VBFixedString(2), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=2)> Public dep_name As String
        <VBFixedString(4), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=4)> Public entry_name As String
        Dim entry_date As String
        Dim hm_num As Short
        'Dim hm_name(65) As String*8
        Dim hm_name() As String '20100616�ڐA�ǉ�
        Public Sub Initilize() '20100616�ڐA�ǉ�
            ' -> watanabe edit 2013.05.29
            'ReDim hm_name(65)
            ReDim hm_name(130)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure BZ_KANRI ' �u�����h�}��
        Dim flag_delete As Byte
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        'UPGRADE_ISSUE: �錾�̌^���T�|�[�g����Ă��܂���: �Œ蒷������̔z�� �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="934BD4FF-1FF9-47BD-888F-D411E47E78FA"' ���N���b�N���Ă��������B
        'Dim hm_name(100) As String*8
        Dim hm_name() As String '20100616�ڐA�ǉ�
        Public Sub Initilize() '20100616�ڐA�ǉ�
            ' -> watanabe edit 2013.05.29
            'ReDim hm_name(100)
            ReDim hm_name(260)
            ' <- watanabe edit 2013.05.29
        End Sub
    End Structure

    Structure T_SIZE_CODE
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B 
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
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
        'UPGRADE_WARNING: �Œ蒷������̃T�C�Y�̓o�b�t�@�ɍ��킹��K�v������܂��B
        '20100706�R�[�h�ύX�@...�ϊ���Char�^�ɂȂ��Ă������̂�Srring�^�ɏC��
        <VBFixedString(5), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=5)> Public start_data As String
        Dim i_data As Short
        '    l_data As Long
        Dim f_data As Single
        Dim d_data As Double
        <VBFixedString(3), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=3)> Public end_data As String
    End Structure

    Public temp_gm As GM_KANRI
    'UPGRADE_WARNING: �\���̂̔z��́A�g�p����O�ɏ���������K�v������܂��B>main��Load�ŏ�����
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
	'BrandVB2��۽�̫����A���̫������_�̂����׸ނ�ǉ����܂�
	Public w_no1_flg As Short '�t���O(OK = 0,NG = 1)
	Public w_no2_flg As Short '�t���O(OK = 0,NG = 1)
	
	'�׸ނ̒ǉ��i��ʸ�د���ް��j
	Public grd_clear_flg As Short '(��د�޸ر = 0�Anot_clear = 1 )
	
	'----- 1/27 1998 yamamoto --------
	'��ݾ��׸�
	Public GL_cancel_flg As Short '�L�����Z���t���O ( 0 = �����l or OFF  1 = ON �j


    ' -> watanabe edit VerUP(2011)

    '----- .NET�ڍs (ADO�ڑ�) -----
    'Public GL_T_RDO As T_RDO_Struct '�q�c�n�ڑ��p
    'Public GL_RDOEnv As RDO.rdoEnvironment

    Public GL_T_ADO As T_ADO_Struct 'ADO�ڑ��p

    Public Const DEF_FUNC_NORMAL As Integer = 1 '�֐��߂�l
    Public Const DEF_FUNC_ERROR As Integer = 0 '�֐��߂�l

    Public Const SUCCEED As Integer = 1
    Public Const FAIL As Integer = 0
    ' <- watanabe edit VerUP(2011)

End Module