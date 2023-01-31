Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices
Module MJ_Communicate

    Public NotifyFlag As Integer '20100615�ڐA�ǉ�
    'Private Const vbLinkNone As Integer = 0 '20100616�ڐA�ǉ�
    'Private Const vbLinkManual As Integer = 2 '20100616�ڐA�ǉ�

    '-----------------------------------------------------------------------------
    '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h
    Private Const XCLASS_BOOL As Integer = &H1000
    Private Const XTYPF_NOBLOCK As Integer = &H2
    Private Const XTYP_CONNECT As Integer = &H60 Or XCLASS_BOOL Or XTYPF_NOBLOCK
    Private Const CP_WINANSI As Integer = 1004
    Private Const SW_RESTORE As Integer = 9
    Private Const DDE_FACK As Integer = &H8000
    Private Const XCLASS_FLAGS As Integer = &H4000
    Private Const XCLASS_DATA As Integer = &H2000
    Private Const XTYP_EXECUTE As Integer = &H50 Or XCLASS_FLAGS
    Private Const XTYP_POKE As Integer = &H90 Or XCLASS_FLAGS
    Private Const XTYP_REQUEST As Integer = &HB0 Or XCLASS_DATA
    Private Const DNS_REGISTER As Integer = &H1
    Private Const DNS_UNREGISTER As Integer = &H2

    Private Declare Function DdeInitializeA Lib "user32" (ByRef pidInst As Integer, ByVal pfnCallback As DdeCallbackDelegate, ByVal afCmd As Integer, ByVal ulRes As Integer) As Short
    Private Declare Function DdeCreateStringHandleA Lib "user32" (ByVal idInst As Integer, <MarshalAs(UnmanagedType.LPStr)> ByVal psz As String, ByVal iCodePage As Integer) As Integer
    Private Declare Function DdeConnect Lib "user32" (ByVal idInst As Integer, ByVal hszService As Integer, ByVal hszTopic As Integer, ByVal pCC As Integer) As Integer
    Private Declare Function DdeNameService Lib "user32" (ByVal idInst As Integer, ByVal hsz1 As Integer, ByVal hsz2 As Integer, ByVal afCmd As Integer) As Integer
    Private Declare Function DdeFreeStringHandle Lib "user32" (ByVal idInst As Integer, ByVal hsz As Integer) As Integer
    Private Declare Function DdeQueryStringA Lib "user32" (ByVal idInst As Integer, ByVal hsz As Integer, ByRef psz As Byte, ByVal cchMax As Integer, ByVal iCodePage As Integer) As Integer
    Private Declare Function DdeQueryStringLen Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Integer, ByVal hsz As Integer, ByVal psz As Integer, ByVal cchMax As Integer, ByVal iCodePage As Integer) As Integer
    Private Declare Function DdeUninitialize Lib "user32" (ByVal idInst As Integer) As Integer
    Private Declare Function DdeGetData Lib "user32" (ByVal hData As Integer, ByRef pData As Byte, ByVal nSize As Integer, ByVal nOffset As Integer) As Integer
    Private Declare Function DdeGetDataLen Lib "user32" Alias "DdeGetData" (ByVal hData As Integer, ByVal pDest As Integer, ByVal nSize As Integer, ByVal nOffset As Integer) As Integer
    Private Declare Function DdeCreateDataHandle Lib "user32" (ByVal idInst As Integer, ByVal lpSrc As String, ByVal cb As Integer, ByVal cbOff As Integer, ByVal hszItem As Integer, ByVal wFmt As Integer, ByVal afCmd As Integer) As Integer
    Private Declare Function DdeClientTransaction Lib "user32" (ByVal sData As String, ByVal cbData As Integer, ByVal hConv As Integer, ByVal hszItem As Integer, ByVal wFmt As Integer, ByVal wType As Integer, ByVal dwTimeout As Integer, ByRef pdwResult As Integer) As Integer
    Private Declare Function DdeFreeDataHandle Lib "user32" (ByVal hData As Integer) As Integer

    Private Declare Function DdeGetLastError Lib "user32" (ByVal idInst As Integer) As Integer 'error�m�F�p

    '***DdeDisconnect�͑��̃��W���[�����炻�̂܂܌Ăяo���\��������***
    Public Declare Function DdeDisconnect Lib "user32" (ByVal hConv As Integer) As Integer

    Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Integer) As Integer
    Private Declare Function EmptyClipboard Lib "user32" () As Integer
    Private Declare Function CloseClipboard Lib "user32" () As Integer
    Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Integer) As Integer
    Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Integer, ByVal hData As Integer) As Integer
    Private Declare Function EnumClipboardFormats Lib "user32" (ByVal nIndex As Integer) As Integer

    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Integer) As Integer
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hmem As Integer) As Integer
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hmem As Integer) As Integer
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef lpDest As Byte, ByVal lpSrc As Integer, ByVal nBytes As Integer)
    Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal lpDest As Integer, ByRef lpSrc As Byte, ByVal nBytes As Integer)
    Private Declare Function lstrlenA Lib "kernel32" (<MarshalAs(UnmanagedType.LPStr)> ByVal s As String) As Integer

    Private Declare Function RegisterClipboardFormatA Lib "user32" (<MarshalAs(UnmanagedType.LPStr)> ByVal lpszName As String) As Integer

    Private Delegate Function DdeCallbackDelegate(ByVal uType As Integer, ByVal uFmt As Integer, ByVal hConv As Integer, ByVal hszTopic As Integer, ByVal hszItem As Integer, ByVal hData As Integer, ByVal dwData1 As Integer, ByVal dwData2 As Integer) As Integer

    Dim idInst As Integer
    Dim hszDDEService As Integer
    Dim hszDDETopic As Integer
    Dim sServiceName As String
    Dim sTopicName As String

    Private text_ctl As Object
    Private hConv_g As Integer

    Const vbCFText As Integer = 1
    '-----------------------------------------------------------------------------

    '//DDE�ݸ���蓮�ōs���܂�
    'Function CreateLink(ByRef Ctl As System.Windows.Forms.Control, ByRef appname As String, ByRef topic As String, ByRef item As String) As Short
    Function CreateLink(ByRef Ctl As Object, ByRef appname As String, ByRef topic As String, ByRef item As String) As Short '20100616�ڐA�ǉ�
        On Error Resume Next

        text_ctl = Ctl

        '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h====================================
        If DdeInitializeA(idInst, AddressOf DdeCallback, 0, 0) = 0 Then
            hszDDEService = DdeCreateStringHandleA(idInst, "BrandVB", CP_WINANSI)
            hszDDETopic = DdeCreateStringHandleA(idInst, "Topic1", CP_WINANSI)
            DdeNameService(idInst, hszDDEService, 0, DNS_REGISTER)
            sServiceName = appname '���T�[�r�X���ƃg�s�b�N���̕�����͒ʐM�J�n�ɔ����đޔ����Ă���
            sTopicName = topic
        End If

        hConv_g = DdeConnect(appname, topic)

        '=======================================================================
        'Ctl.LinkMode = vbLinkNone
        'Ctl.LinkTopic = appname & "|" & topic
        'Ctl.LinkItem = item
        'Ctl.LinkMode = vbLinkManual

        ' -> watanabe edit VerUP(2011)
        'CreateLink = Err.Number
        'If Err.Number = 0 Then
        '    'Ctl.LinkRequest()
        'End If
        If hConv_g = 0 Then
            CreateLink = -1
        Else
            CreateLink = Err.Number
        End If
        ' <- watanabe edit VerUP(2011)

    End Function

    '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h
    Function DdeCallback(ByVal uType As Integer, ByVal uFmt As Integer, ByVal hConv As Integer, ByVal hszTopic As Integer, ByVal hszItem As Integer, ByVal hData As Integer, ByVal dwData1 As Integer, ByVal dwData2 As Integer) As Integer
        On Error Resume Next
        Select Case uType
            Case XTYP_CONNECT 'DDE�ʐM�̊m��
                Dim s1 As String = DdeString(hszTopic)  '�g�s�b�N��
                Dim s2 As String = DdeString(hszItem)   '���̂Ƃ��̓T�[�r�X��

                '�T�[�r�X���ƃg�s�b�N������v����Ƃ���True(1)��Ԃ�
                If s2 & "|" & s1 = sServiceName & "|" & sTopicName Then Return 1

            Case XTYP_REQUEST 'LinkRequest���\�b�h�����s����܂���
                'LinkItem�v���p�e�B�ɐݒ肳�ꂽ�A�C�e�������擾
                Dim sItem As String = DdeString(hszItem)

                'VB.NET�ł̓t�H�[����ɒ��ڔz�u���ꂽ�R���g���[������
                'Controls�R���N�V�����Ɋ܂܂�Ȃ��̂Œ���
                Dim sReturn As String = text_ctl.Text

                '�f�[�^�Ƃ��ĕ������Ԃ��܂��B
                Return DdeCreateDataHandle(idInst, sReturn, lstrlenA(sReturn) + 1, 0, hszItem, vbCFText, 0)

            Case XTYP_POKE 'LinkPoke���\�b�h�����s����܂���
                '�f�[�^�𕶎���Ƃ��Ď󂯎��܂�
                Dim n As Integer = DdeGetDataLen(hData, 0, 0, 0)
                Dim bData(n - 1) As Byte
                DdeGetData(hData, bData(0), n, 0)
                Dim sData As String = System.Text.Encoding.Default.GetString(bData, 0, n - 1)
                DdeFreeDataHandle(hData)

                'LinkItem�v���p�e�B�ɐݒ肳�ꂽ�A�C�e�������擾
                Dim sItem As String = DdeString(hszItem)

                'VB.NET�ł̓t�H�[����ɒ��ڔz�u���ꂽ�R���g���[������
                'Controls�R���N�V�����Ɋ܂܂�Ȃ��̂Œ���
                'text_ctl.Text = sData

                '���������̕�
                Return DDE_FACK
        End Select
    End Function

    '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h
    Private Function DdeString(ByVal hsz As Integer) As String
        Dim n As Integer = DdeQueryStringLen(idInst, hsz, 0, 0, CP_WINANSI) + 1
        Dim b(n - 1) As Byte
        DdeQueryStringA(idInst, hsz, b(0), n, CP_WINANSI)
        Return System.Text.Encoding.Default.GetString(b, 0, n - 1)
    End Function

    '-------------< DDE REQUEST >-----------------
    '//mss : �A�C�e������
    Function RequestACAD(ByRef mss As String) As Short
        On Error Resume Next

        If hConv_g = 0 Then Exit Function

        Dim hszItem As Integer = DdeCreateStringHandleA(idInst, mss, CP_WINANSI)
        Dim dwResult As Integer
        Dim hData As Integer = DdeClientTransaction("", 0, hConv_g, hszItem, vbCFText, XTYP_REQUEST, timeOutSecond * 1000, dwResult)
        If hData = 0 Then Return 1 'error!!
        Dim n As Integer = DdeGetDataLen(hData, 0, 0, 0)
        Dim b(n - 1) As Byte
        DdeGetData(hData, b(0), n, 0)
        DdeFreeDataHandle(hData)

        Dim ret_str As String = System.Text.Encoding.Default.GetString(b, 0, n - 1)
        text_ctl.Text = ret_str

        NotifyFlag = False

    End Function

    '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h
    'DDE�N���C�A���g�Ƃ��ĒʐM���J�n���܂��B
    'ServiceName: �O���A�v���P�[�V������
    'TopicName: �g�s�b�N��
    '���������Ƃ���DDE��b(conversation)�̃n���h����Ԃ��܂��B
    'VB6��LinkMode�v���p�e�B��vbLinkManual�ɂ����Ƃ��ɑΉ����܂��B
    '�ʐM�I�����͒���DdeDisconnect API�����s���Ă��������B
    Function DdeConnect(ByVal ServiceName As String, ByVal TopicName As String) As Integer
        Dim hsz1 As Integer = DdeCreateStringHandleA(idInst, ServiceName, CP_WINANSI)
        Dim hsz2 As Integer = DdeCreateStringHandleA(idInst, TopicName, CP_WINANSI)
        Dim hConv As Integer = DdeConnect(idInst, hsz1, hsz2, 0)
        DdeFreeStringHandle(idInst, hsz1)
        DdeFreeStringHandle(idInst, hsz2)
        
        Dim err As Integer = 0
        err = DdeGetLastError(idInst)

        If hConv = 0 Then MsgBox(ServiceName & "|" & TopicName & ControlChars.CrLf & err & "There is no response from the external application for the start of this DDE.", MsgBoxStyle.Exclamation)
        Return hConv
    End Function

    '-----------< DDE POKE > -----------------------------------
    '//iname : �A�C�e������
    '//mss   : POKE�f�[�^�{��
    Function PokeACAD(ByRef iname As String, ByRef mss As String) As Short
        On Error Resume Next

        If hConv_g = 0 Then Exit Function

        form_main.SRflag.Text = "SEND"
        DdePoke(hConv_g, iname, mss)

        If Err.Number Then MsgBox(ErrorToString())

        form_main.SRflag.Text = ""

    End Function

    '!!!!!!!!!!!!!!!!!!!!!!!!!!�ڐA�R�[�h
    'DDE�T�[�o�[�ɑ΂��ĕ�������������݂܂��B
    'hConv: DdeConnect�̖߂�l
    'ItemName: �A�C�e�����i�l���������ރe�L�X�g�{�b�N�X�̖��O�j
    'PokeText: �������ޕ�����
    'VB6��LinkPoke���\�b�h�ɑΉ����܂��B
    Sub DdePoke(ByVal hConv As Integer, ByVal ItemName As String, ByVal PokeText As String)
        Dim hszItem As Integer = DdeCreateStringHandleA(idInst, ItemName, CP_WINANSI)
        Dim dwResult As Integer
        DdeClientTransaction(PokeText, lstrlenA(PokeText) + 1, hConv, hszItem, vbCFText, XTYP_POKE, 100, dwResult)
        DdeFreeStringHandle(idInst, hszItem)
    End Sub

    '//DDE�ݸ��ؒf���܂�
    'Private Sub Disconnect(ByRef Ctl As System.Windows.Forms.Control)
    Private Sub Disconnect(ByRef Ctl As Object) '20100616�ڐA�ǉ�
        On Error Resume Next

        'InitFlag = False '20100628�ǉ��R�[�h

        DdeDisconnect(hConv_g)

        If idInst <> 0 Then
            DdeNameService(idInst, hszDDEService, 0, DNS_UNREGISTER)
            DdeFreeStringHandle(idInst, hszDDEService)
            DdeFreeStringHandle(idInst, hszDDETopic)
            DdeUninitialize(idInst)
            idInst = 0
        End If

        'Dim tempTimeOutVal As Object
        'tempTimeOutVal = Ctl.LinkTimeout
        'Ctl.LinkTimeout = 1
        'Ctl.LinkMode = vbLinkNone
        'Ctl.LinkTimeout = tempTimeOutVal
    End Sub


    ' -> watanabe edit VerUP(2011)
    '
    '    '-----------------< SQL �̏�����(�n�o�����j >---------------------------------
    '    Function init_sql() As Short
    '        Dim Login As Object
    '        Dim result As Object
    '        On Error GoTo init_sql_error_section
    '
    '        init_sql = False
    '
    '        'Get a Login record and set login attributes.
    '        'Initialize VBSQL.
    '        If SqlInit() = "" Then
    '            MsgBox("VBSQL has not been initialized.", MsgBoxStyle.Critical, "SQL�װ")
    '            Exit Function
    '        End If
    '
    '        result = SqlSetLoginTime(5)
    '
    '        If result = FAIL Then
    '            MsgBox("��ѱ�Đݒ�װ", MsgBoxStyle.Critical, "SQL�װ")
    '            Exit Function
    '        End If
    '
    '        Login = SqlLogin()
    '        result = SqlSetLUser(Login, DBLoginID)
    '        result = SqlSetLPwd(Login, DBpasswd)
    '        result = SqlSetLApp(Login, DBexample)
    '
    '        'Get a connection for communicating with SQL Server.
    '        SqlConn = SqlOpen(Login, DBServer)
    '
    '        If SqlConn = 0 Then
    '            GoTo init_sql_error_section
    '        End If
    '
    '        init_sql = True
    '        GoTo init_sql_end_section
    '
    'init_sql_error_section:
    '        Resume Next
    '        MsgBox("SQL INITIALIZE�װ", MsgBoxStyle.Critical, "SQL�װ")
    '        init_sql = False
    '
    'init_sql_end_section:
    '    End Function
    '
    '    '---<SQL �b�k�n�r�d >---------
    '    Function end_sql() As Short
    '        SqlExit()
    '        SqlConn = 0
    '    End Function


    '-----------------< SQL �̏�����(�n�o�����j >---------------------------------
    Function init_sql() As Short
        Dim IRet As Integer
        Dim ErrMsg As String

        If SqlConn = 1 Then
            init_sql = True
            Exit Function
        End If

        On Error GoTo init_sql_error_section

        '�߂�l������
        init_sql = False

        '�ڑ��t���O��OFF�ɐݒ�
        SqlConn = 0

        '�G���[���b�Z�[�W������
        ErrMsg = ""

        '----- .NET�ڍs (ADO�ڑ�) -----
        ''RDO�ڑ��p�\���̏�����
        'IRet = VBRDO_T_RDOInit(GL_T_RDO)

        ''RDO�ڑ����ݒ�
        'GL_T_RDO.DSN = Trim(DBName)         ''VBSQL�ł�DSN���g�p���Ă��炸���݂̑��̐ڑ���DB�����g�p���Ă���̂ŁADB�����g�p
        'GL_T_RDO.UID = Trim(DBLoginID)
        'GL_T_RDO.PWD = Trim(DBpasswd)
        'GL_T_RDO.DBName = Trim(DBName)
        'GL_T_RDO.Server = Trim(DBServer)

        ''DSN�쐬
        'If VBRDO_DSNRegistry(GL_T_RDO) = DEF_FUNC_ERROR Then
        '    ErrMsg = DEF_MSG_E9001 & " [ VBRDO_DSNRegistry ] "
        '    GoTo init_sql_error_section
        'End If

        ''RDO�ڑ�
        'If VBRDO_OpenEnv(GL_RDOEnv) = DEF_FUNC_ERROR Then
        '    ErrMsg = DEF_MSG_E9002 & " [ VBRDO_Connect ] "
        '    GoTo init_sql_error_section
        'End If
        'If VBRDO_Connect(GL_RDOEnv, GL_T_RDO) = DEF_FUNC_ERROR Then
        '    ErrMsg = DEF_MSG_E9002 & " [ VBRDO_Connect ] "
        '    GoTo init_sql_error_section
        'End If

        '---------------------------------------------------------------------------------

        'ADO�ڑ��p�\���̏�����
        IRet = VBADO_T_ADOInit(GL_T_ADO)

        'ADO�ڑ����ݒ�
        GL_T_ADO.DSN = Trim(DBName)         'VBSQL�ł�DSN���g�p���Ă��炸���݂̑��̐ڑ���DB�����g�p���Ă���̂ŁADB�����g�p
        GL_T_ADO.UID = Trim(DBLoginID)
        GL_T_ADO.PWD = Trim(DBpasswd)
        GL_T_ADO.DBName = Trim(DBName)
        GL_T_ADO.Server = Trim(DBServer)

        'ADO�ڑ�
        If VBADO_Connect(GL_T_ADO) = DEF_FUNC_ERROR Then
            ErrMsg = DEF_MSG_E9002 & " [ VBADO_Connect ] "
            GoTo init_sql_error_section
        End If

        '----- .NET�ڍs (ADO�ڑ�) -----

        '�ڑ��t���O��ON�ɐݒ�
        SqlConn = 1

        '�߂�l�Z�b�g
        init_sql = True

        Exit Function

init_sql_error_section:
        If ErrMsg = "" Then
            ErrMsg = Err.Description
        End If

        On Error Resume Next
        MsgBox(ErrMsg, MsgBoxStyle.Critical, "SQL error")
        init_sql = False
    End Function

    '---<SQL �b�k�n�r�d >---------
    Function end_sql() As Short
        Dim IRet As Integer

        If SqlConn = 0 Then
            end_sql = True
            Exit Function
        End If

        '----- .NET�ڍs (ADO�ڑ�) -----
        ''RDO�ڑ��ؒf
        'IRet = VBRDO_Discon(GL_T_RDO)
        'IRet = VBRDO_CloseEnv(GL_RDOEnv)

        '-------------------------------

        'ADO�ڑ��ؒf
        IRet = VBADO_Discon(GL_T_ADO)

        '----- .NET�ڍs (ADO�ڑ�) -----

        '�ڑ��t���O��OFF�ɐݒ�
        SqlConn = 0

    End Function
    ' <- watanabe edit VerUP(2011)



    '--------< Advance CAD �ڑ� >---------------
    Function init_cad() As Short
        Dim ConnectTxt As Short

        On Error GoTo init_cad_error_section

        ConnectTxt = CreateLink(form_main.Text2, ACADTransAppName, ACADTransTopic, ACADTransItem)
        'ConnectTxt = CreateLink(form_main.text2, ACADTransAppName, ACADTransTopic, "INITIALIZE")


        If ConnectTxt = errNoAppResponded Then
            init_cad = errNoAppResponded
        ElseIf ConnectTxt = 0 Then
            init_cad = 0
            '    optDataType(conDestText).Value = True
        Else
            init_cad = ConnectTxt
        End If

        ' init_cad = True
        GoTo init_cad_end_section

init_cad_error_section:
        Resume Next
        init_cad = -1

init_cad_end_section:
    End Function
    ' -> watanabe add 2007.03

    '--------< Advance CAD �����I���m�F >---------------
    Function check_cad_run() As Boolean
        Dim error_no As Object
        Dim w_ret As Object
        Dim time_start As Object
        Dim time_now As Object

        On Error GoTo error_section

        ' �߂�l������
        check_cad_run = False


        ' �ϊ������I���m�F
        time_start = Now
        Do
            time_now = Now

            If Trim(form_main.Text2.Text) = "" Then
                If time_now - time_start > timeOutSecond Then
                    MsgBox("Time-out error", 64, "ERROR")
                    w_ret = PokeACAD("ERROR", "TIMEOUT " & timeOutSecond & " seconds have passed.")
                    w_ret = RequestACAD("ERROR")
                    GoTo error_section
                End If

            ElseIf Left(Trim(form_main.Text2.Text), 6) = "VBKILL" Then
                form_main.Text2.Text = ""
                Exit Do

            ElseIf Left(Trim(form_main.Text2.Text), 5) = "ERROR" Then
                error_no = Mid(Trim(form_main.Text2.Text), 6, 3)
                MsgBox("An error was returned.  ERROR_NO=" & error_no, 64, "Error of the return value of the ACAD")
                form_main.Text2.Text = ""
                GoTo error_section

            Else
                MsgBox("Return code is invalid." & Chr(13) & Trim(form_main.Text2.Text), 64, "Error of the return value of the ACAD")
                form_main.Text2.Text = ""
                GoTo error_section
            End If
        Loop

        ' �߂�l�ݒ�
        check_cad_run = True

        Exit Function

error_section:
    End Function

    ' <- watanabe add 2007.03
End Module