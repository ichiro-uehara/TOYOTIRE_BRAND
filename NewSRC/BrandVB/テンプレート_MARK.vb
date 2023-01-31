Public Class F_TMP_MARK
    Public Tmp_Family(MaxSelNum) As String
    Public Tmp_Zuban(MaxSelNum) As String
    Public Tmp_Nengo(MaxSelNum) As String
    Public Const DEF_Standard As String = "DB"
    Public Const DEF_StandardINI As String = "BR_Set.ini"
    Dim LL_StandardADO As Object '2014/12/16 moriya add
    Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

    Private Sub テンプレート_MARK_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim w_ret As Object
        Dim ret As Object
        Dim w_w_str As String
        Dim i, IRet As Short
        Dim LL_tmp As New T_ADO_Struct '2014/12/16 moriya add
        Dim wkfname As String
        Dim AcadDir As String

        LL_StandardADO = LL_tmp

        ' ''AcadDir = New String(Chr(0), 255)
        ' ''IRet = GetEnvironmentVariable("ACAD_SET", AcadDir, Len(AcadDir))
        ' ''If IRet = 0 Then
        ' ''    Exit Sub
        ' ''End If

        ' ''AcadDir = AcadDir.Substring(0, InStr(1, AcadDir, Chr(0), CompareMethod.Binary) - 1)
        ' ''If AcadDir.Substring(AcadDir.Length - 1, 1) <> "\" Then
        ' ''    AcadDir = AcadDir & "\"
        ' ''End If
        ' ''wkfname = Trim(AcadDir) & DEF_StandardINI

        '' ''DB設定
        '' ''初期処理
        ' ''VBADO_T_ADOInit(LL_StandardADO)
        '' ''ADO接続用初期設定ﾌｧｲﾙ読込
        ' ''set_read(wkfname)

        ' ''With LL_StandardADO
        ' ''    .DSN = Trim(STANDARD_DBName) 'ADO接続用ﾃﾞｰﾀｿｰｽ名
        ' ''    '.UID = Trim(DBLoginID) 'ADO接続用ﾕｰｻﾞID
        ' ''    '.PWD = Trim(DBpasswd) 'ADO接続用ﾊﾟｽﾜｰﾄﾞ
        ' ''    .UID = "sa" 'ADO接続用ﾕｰｻﾞID
        ' ''    .PWD = "IHDB66" 'ADO接続用ﾊﾟｽﾜｰﾄﾞ
        ' ''    .DBName = Trim(STANDARD_DBName) 'ADO接続用ﾃﾞｰﾀﾍﾞｰｽ名
        ' ''    .Server = Trim(DBServer) 'ADO接続用ｻｰﾊﾞ-
        ' ''End With

        form_no = Me

        'Brand Ver.4 追加
        itemset()

        'タイプ
        '(Brand Ver.3 変更)
        w_w_str = Environ("ACAD_SET")
        w_w_str = Trim(w_w_str) & Trim(Tmp_MARK_ini)
        ret = set_read5(w_w_str, "Mark", 1)

        'ファイルから取得したデータを格納
        For i = 1 To MaxSelNum - 1
            Tmp_Family(i - 1) = Tmp_prcs_code(i)
            Tmp_Zuban(i - 1) = Tmp_hm_word(i)
            Tmp_Nengo(i - 1) = Tmp_hm_code(i)
        Next

        form_main.Text2.Text = ""
        CommunicateMode = comFreePic
        w_ret = RequestACAD("PICEMPTY")

        '入力規制
        AddHandler Me.w_size3.KeyPress, AddressOf Me.OnlyNum_jisu_KeyPress
        AddHandler Me.w_kajyu.KeyPress, AddressOf Me.OnlyNum_jisu_KeyPress

        'bitmapの画面表示モードを設定
        ImgThumbnail1.SizeMode = PictureBoxSizeMode.Zoom
        ImgThumbnail2.SizeMode = PictureBoxSizeMode.Zoom

        CommunicateMode = comSpecData
        RequestACAD("SPECDATA")

        'InitFlag = True '20100628追加コード
    End Sub

    'Tire type選択時の入力規制処理
    Private Sub w_syurui_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles w_syurui.SelectedIndexChanged
        If Me.w_syurui.SelectedItem.ToString() = "PC" Then
            w_type.Enabled = False 'Tire type
            w_type.BackColor = Color.Gainsboro

            w_version.Enabled = True 'Load version
            w_sokudo.Enabled = True 'Speed symbol
            w_size4.Enabled = True 'Speed

            w_version.BackColor = Color.White
            w_sokudo.BackColor = Color.White
            w_size4.BackColor = Color.White

        ElseIf (Me.w_syurui.SelectedItem.ToString() = "LT") Or (Me.w_syurui.SelectedItem.ToString() = "TB") Then
            w_type.Enabled = True 'Tire type
            w_type.BackColor = Color.White

            w_version.Enabled = False 'Load version
            w_sokudo.Enabled = False 'Speed symbol
            w_size4.Enabled = False 'Speed

            w_version.BackColor = Color.Gainsboro
            w_sokudo.BackColor = Color.Gainsboro
            w_size4.BackColor = Color.Gainsboro

        Else
            w_type.Enabled = True 'Tire type
            w_version.Enabled = True 'Load version
            w_sokudo.Enabled = True 'Speed symbol
            w_size4.Enabled = True 'Speed

            w_type.BackColor = Color.White
            w_version.BackColor = Color.White
            w_sokudo.BackColor = Color.White
            w_size4.BackColor = Color.White

        End If
    End Sub

    Private Sub Command4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command4.Click
        form_no.Close()
        F_MAIN3.Close()
    End Sub

    'Clear押下時の処理
    Private Sub Command3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command3.Click

        w_size1.Text = ""
        w_size2.Text = ""
        w_size3.Text = ""
        w_size4.Text = ""
        w_size5.Text = ""
        w_size6.Text = ""
        w_kajyu.Text = ""
        w_kaju.Text = ""

        w_syurui.Text = ""
        w_type.Text = ""
        w_version.Text = ""
        w_sokudo.Text = ""

        '色も全て元に戻す
        w_size1.BackColor = Color.White
        w_size2.BackColor = Color.White
        w_size3.BackColor = Color.White
        w_size4.BackColor = Color.White
        w_size5.BackColor = Color.White
        w_size6.BackColor = Color.White
        w_kajyu.BackColor = Color.White
        w_kaju.BackColor = Color.White

        w_sokudo.BackColor = Color.White

        '全て入力可能に戻す
        w_size1.Enabled = True
        w_size2.Enabled = True
        w_size3.Enabled = True
        w_size4.Enabled = True
        w_size5.Enabled = True
        w_size6.Enabled = True
        w_kajyu.Enabled = True
        w_kaju.Enabled = True

        w_sokudo.Enabled = True


        'Result枠内をクリア
        ResultClear()

        itemset()

    End Sub

    Sub ResultClear()
        w_Family.Text = ""
        w_Reg1.Text = ""
        w_Reg2.Text = ""
        w_hm_name_1.Text = ""
        w_hm_name_2.Text = ""

        form_no.ImgThumbnail1.Image = Nothing
        form_no.ImgThumbnail2.Image = Nothing
    End Sub

    Sub itemset()
        '適用規格
        w_syurui.Items.Clear()
        w_syurui.Items.Add("PC")
        w_syurui.Items.Add("LT")
        w_syurui.Items.Add("TB")

        w_type.Items.Clear()
        w_type.Items.Add("T/L")
        w_type.Items.Add("T/T")

        'タイヤ種類
        w_version.Items.Clear()
        w_version.Items.Add("Normal(SL)")
        w_version.Items.Add("Reinforced(XL)")

        w_syurui.BackColor = Color.White
        w_type.BackColor = Color.White
        w_version.BackColor = Color.White

        w_syurui.Enabled = True
        w_type.Enabled = True
        w_version.Enabled = True
    End Sub

    'Substitution Read押下時の処理
    Private Sub Command2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command2.Click
        Dim gm_no As Object
        Dim ZumenName As Object
        Dim pic_no As Object
        Dim w_ret As Object
        Dim w_str As Object
        Dim w_mess As String
        Dim i As Short

        '検索していない場合
        If w_Family.Text = "" Or w_Reg1.Text = "" Or w_Reg2.Text = "" _
            Or w_hm_name_1.Text = "" Or w_hm_name_2.Text = "" Then
            MsgBox("Not Found MARK.", MsgBoxStyle.Critical)
            Exit Sub
        End If

        '呼び出しモデルが1つに指定されていない場合はエラー表示（例：F19）
        If w_Reg1.Text = "-" Or w_Reg2.Text = "-" _
            Or w_hm_name_1.Text = "-" Or w_hm_name_2.Text = "-" Then
            MsgBox("Not Found MARK.", MsgBoxStyle.Critical)
            Exit Sub
        End If


        '// 置換モードの送信
        w_ret = PokeACAD("CHNGMODE", Trim(ReplaceMode).Substring(0, 1))
        w_ret = RequestACAD("CHNGMODE")

        'テンプレートとなる図面を送信
        pic_no = what_pic_from_hmcode(form_no.w_hm_name_1.Text)
        If pic_no < 1 Then GoTo error_section
        ZumenName = "HM-" & Trim(form_no.w_hm_name_1.Text).Substring(0, 6)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
        w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("HMCODE", w_mess)

        '// 次の処理を行うためにVBはそのままにする。
        CommunicateMode = comMark
        w_ret = RequestACAD("TMPCHANG")

        '風船作成
        CommunicateMode = comMark
        w_ret = RequestACAD("TMPBALLOON")

        '2つめ
        'テンプレートとなる図面を送信
        pic_no = what_pic_from_hmcode(form_no.w_hm_name_2.Text)
        If pic_no < 1 Then GoTo error_section
        ZumenName = "HM-" & Trim(form_no.w_hm_name_2.Text).Substring(0, 6)

        '----- .NET 移行 -----
        'w_mess = VB6.Format(Val(pic_no), "000") & HensyuDir & ZumenName
        w_mess = Val(pic_no).ToString("000") & HensyuDir & ZumenName

        w_ret = PokeACAD("HMCODE", w_mess) 'HMCODE(置換文字)扱いで読み込むと、枠線を消せる

        '// 終了の送信
        CommunicateMode = comMark
        'w_ret = RequestACAD("ACADREAD")
        w_ret = RequestACAD("TMPCHANG")

        '風船作成
        w_ret = PokeACAD("HMCODE", "602\\") '条件分岐に使用_風船60_2を作成してVB終了
        CommunicateMode = comNone
        w_ret = RequestACAD("TMPBALLOON")

        Exit Sub

error_section:
        On Error Resume Next
    End Sub

    'Search押下時の処理
    Private Sub Command1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command1.Click
        Dim Family As String
        Dim Reg1, Reg2 As String
        Dim hm_name_1, hm_name_2 As String
        Dim iret As Integer
        Dim TiffFile As String
        Dim w_file As Object

        'Tire typeによって処理を分ける
        Select Case w_syurui.Text
            Case "PC"
                If w_size1.Text.Contains("LT") = False And w_size6.Text.Contains("C") = False And w_size6.Text.Contains("LT") = False Then
                    iret = SearchPC(Family, Reg1, Reg2, hm_name_1, hm_name_2, w_syurui.Text)
                    If iret <> 0 Then
                        Exit Sub
                    End If
                Else
                    'Result枠内をクリア
                    ResultClear()
                    '選ばれていない場合はメッセージ表示
                    MsgBox("Check Tire type.", MsgBoxStyle.Critical)
                    w_syurui.Select()
                    w_syurui.SelectAll()
                End If
            Case "LT"
                If w_size1.Text.Contains("LT") = True Or w_size6.Text.Contains("C") = True Or w_size6.Text.Contains("LT") = True Then
                    iret = SearchLTTB(Family, Reg1, Reg2, hm_name_1, hm_name_2, w_syurui.Text)
                    If iret <> 0 Then
                        Exit Sub
                    End If
                Else
                    'Result枠内をクリア
                    ResultClear()
                    '選ばれていない場合はメッセージ表示
                    MsgBox("Check Tire type.", MsgBoxStyle.Critical)
                    w_syurui.Select()
                    w_syurui.SelectAll()
                End If
            Case "TB"
                iret = SearchLTTB(Family, Reg1, Reg2, hm_name_1, hm_name_2, w_syurui.Text)
                If iret <> 0 Then
                    Exit Sub
                End If
            Case Else
                'Result枠内をクリア
                ResultClear()
                '選ばれていない場合はメッセージ表示
                MsgBox("Select Tire type.", MsgBoxStyle.Critical)
                w_syurui.Select()
                w_syurui.SelectAll()

                Exit Sub
        End Select

        w_Family.Text = "(" & Family & ")"

        w_Reg1.Text = Reg1
        w_Reg2.Text = Reg2

        w_hm_name_1.Text = hm_name_1
        w_hm_name_2.Text = hm_name_2

        '左側のBMPを表示
        TiffFile = TIFFDir & hm_name_1 & ".bmp"

        'Tiffﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            form_no.ImgThumbnail1.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail1.Width = 455
            form_no.ImgThumbnail1.Height = 193
        Else
            form_no.ImgThumbnail1.Image = Nothing
        End If

        '右側のBMPを表示
        TiffFile = TIFFDir & hm_name_2 & ".bmp"

        'Tiffﾌｧｲﾙ表示
        w_file = Dir(TiffFile)
        If w_file <> "" Then
            form_no.ImgThumbnail2.Image = System.Drawing.Image.FromFile(TiffFile)
            form_no.ImgThumbnail2.Width = 272
            form_no.ImgThumbnail2.Height = 193
        Else
            form_no.ImgThumbnail2.Image = Nothing
        End If

    End Sub

    'Tire type = PCのときのマーク検索
    Function SearchPC(ByRef Family As String, ByRef Reg1 As String, ByRef Reg2 As String, ByRef hm_name_1 As String, ByRef hm_name_2 As String, ByVal Category As String) As Integer
        Dim sqlcmd, sqlcmd1, sqlcmd2, sqlcmd3, sqlcmd4 As String
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim size3, sokudo, symbol, secwidth As String
        Dim iret As Integer
        Dim Rs As RDO.rdoResultset

        size3 = w_size3.Text
        sokudo = w_sokudo.Text

        '指定条件の速度記号が表す値を検索
        sqlcmd = "select S_Volue from standard..SPEED WHERE S_Symbol = '" & sokudo & "'"

        ' ''ds.Clear()
        ' ''ADO_DB_Search(sqlcmd, "Customers1", LL_StandardADO, ds)

        ' ''dt = ds.Tables("Customers1")

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        If GL_T_RDO.Con.RowsAffected() = 0 Then
            'Result枠内をクリア
            ResultClear()
            '検索条件の指定がおかしい
            MsgBox("Wrong volue (Speed symbol).", MsgBoxStyle.Critical)
            w_sokudo.Select()
            w_sokudo.SelectAll()

            SearchPC = -1
            Exit Function
        ElseIf IsDBNull(Rs.rdoColumns(0).Value) = True Then
            'Result枠内をクリア
            ResultClear()
            '検索条件に不備
            MsgBox("Set search Key value again.", MsgBoxStyle.Critical)
            w_sokudo.Select()
            w_sokudo.SelectAll()

            SearchPC = -1
            Exit Function
        Else
            '検索結果を格納
            symbol = Rs.rdoColumns(0).Value
        End If

        ' ''If dt Is Nothing Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件に不備
        ' ''    MsgBox("Set search Key value again.", MsgBoxStyle.Critical)
        ' ''    w_sokudo.Select()
        ' ''    w_sokudo.SelectAll()

        ' ''    SearchPC = -1
        ' ''    Exit Function
        ' ''ElseIf dt.Rows.Count = 0 Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件の指定がおかしい
        ' ''    MsgBox("Wrong volue (Speed symbol).", MsgBoxStyle.Critical)
        ' ''    w_sokudo.Select()
        ' ''    w_sokudo.SelectAll()

        ' ''    SearchPC = -1
        ' ''    Exit Function
        ' ''End If

        ' ''symbol = dt.Rows(0).Item(0)

        secwidth = ""
        secwidth = Trim(w_size3.Text.ToString)
        sqlcmd = ""

        'セクション幅の値により、条件を変更する
        If secwidth.Length = 3 Then
            If secwidth.Substring(2, 1) = "5" Then
                'シリーズ82特有のセクション幅指定の場合は以下の条件を使用する
                '条件設定（テーブル接合）
                sqlcmd1 = "Select * from standard..IMARKPC INNER JOIN standard..SPEED ON standard..IMARKPC.sokudo_1 =standard..SPEED.S_Symbol where "
                '条件設定（構造1,構造2,カテゴリ,シリーズ82フラグ）
                sqlcmd2 = "size5 ='" & w_size5.Text & "' AND sversion ='" & w_version.Text & "' AND Category = '" & Category & "' AND Series82FLG = 1 "
                '条件設定（速度_検索条件の値のジャスト、以上、以下全てを検索）
                sqlcmd3 = "(((sokudo_1='" & sokudo & "') or (sokudo_2='" & sokudo & "') or (sokudo_3='" & sokudo & "')) and (sokudoFLG =0) ) OR ((S_Volue>=" & symbol & ") AND (sokudoFLG=-1)) OR ((S_Volue<=" & symbol & ") AND (sokudoFLG=1))"

                sqlcmd = sqlcmd1 & sqlcmd2 & " AND (" & sqlcmd3 & ")"
            End If
        End If

        If sqlcmd = "" Then
            'シリーズ82特有のセクション幅指定ではない場合
            '条件設定（テーブル接合）
            sqlcmd1 = "Select * from standard..IMARKPC INNER JOIN standard..SPEED ON standard..IMARKPC.sokudo_1 =standard..SPEED.S_Symbol where "
            '条件設定（構造1,構造2,カテゴリ）
            sqlcmd2 = "size5 ='" & w_size5.Text & "' AND sversion ='" & w_version.Text & "' AND Category = '" & Category & "'"
            '条件設定（偏平比_検索条件の値のジャスト、以上、以下全てを検索）
            'sqlcmd3 = "((size3_1=" & size3 & " or size3_2 = " & size3 & ") AND size3FLG = 0) OR ((size3_1>=" & size3 & ")AND (size3FLG=-1)) OR ((size3_1<=" & size3 & ")AND (size3FLG=1))"
            sqlcmd3 = "(((size3_1=" & size3 & ") or (size3_2 = " & size3 & ")) AND (size3FLG = 0)) OR ((size3_1>=" & size3 & ") AND (size3FLG=-1)) OR ((size3_1<=" & size3 & ") AND (size3FLG=1))"
            '条件設定（速度_検索条件の値のジャスト、以上、以下全てを検索）
            sqlcmd4 = "(((sokudo_1='" & sokudo & "') or (sokudo_2='" & sokudo & "') or (sokudo_3='" & sokudo & "')) and (sokudoFLG =0) ) OR ((S_Volue>=" & symbol & ") AND (sokudoFLG=-1)) OR ((S_Volue<=" & symbol & ") AND (sokudoFLG=1))"

            sqlcmd = sqlcmd1 & sqlcmd2 & " AND (" & sqlcmd3 & ") AND (" & sqlcmd4 & ")"
        End If

        'ds.Clear()
        'ADO_DB_Search(sqlcmd, "Customers2", LL_StandardADO, ds)
        'dt = ds.Tables("Customers2")
        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        If GL_T_RDO.Con.RowsAffected() > 1 Then
            'Result枠内をクリア
            ResultClear()
            '検索条件の指定がおかしい(複数候補あり)
            MsgBox("Not Found Only One Mark.", MsgBoxStyle.Critical)
            SearchPC = -1
            Exit Function
        ElseIf GL_T_RDO.Con.RowsAffected() = 0 Then
            'Result枠内をクリア
            ResultClear()
            '検索条件の指定がおかしい(候補なし)
            MsgBox("Not Found Mark.", MsgBoxStyle.Critical)
            SearchPC = -1
            Exit Function
        End If

        ' ''If dt Is Nothing Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件に不備
        ' ''    MsgBox("Set search Key value again.", MsgBoxStyle.Critical)
        ' ''    SearchPC = -1
        ' ''    Exit Function
        ' ''ElseIf dt.Rows.Count > 1 Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件の指定がおかしい(複数候補あり)
        ' ''    MsgBox("Not Found Only One Mark.", MsgBoxStyle.Critical)
        ' ''    SearchPC = -1
        ' ''    Exit Function
        ' ''ElseIf dt.Rows.Count = 0 Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件の指定がおかしい(候補なし)
        ' ''    MsgBox("Not Found Mark.", MsgBoxStyle.Critical)
        ' ''    SearchPC = -1
        ' ''    Exit Function
        ' ''End If

        '初期化
        Family = ""
        Reg1 = ""
        Reg2 = ""
        hm_name_1 = ""
        hm_name_2 = ""

        '' ''値を格納
        ' ''If IsDBNull(dt.Rows(0).Item("Fname")) = False Then
        ' ''    Family = Trim(dt.Rows(0).Item("Fname").ToString)
        ' ''End If

        ' ''If IsDBNull(dt.Rows(0).Item("Register_No")) = False Then
        ' ''    Reg1 = Trim(dt.Rows(0).Item("Register_No").ToString)
        ' ''End If

        ' ''If IsDBNull(dt.Rows(0).Item("Nengo")) = False Then
        ' ''    Reg2 = Trim(dt.Rows(0).Item("Nengo").ToString)
        ' ''End If

        '値を格納
        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            Family = Trim(Rs.rdoColumns(0).Value.ToString)
        End If

        If IsDBNull(Rs.rdoColumns(25).Value) = False Then
            Reg1 = Trim(Rs.rdoColumns(25).Value.ToString)
        End If

        If IsDBNull(Rs.rdoColumns(26).Value) = False Then
            Reg2 = Trim(Rs.rdoColumns(26).Value.ToString)
        End If


        iret = Array.IndexOf(Tmp_Family, Family)
        If iret <> -1 Then
            hm_name_1 = Tmp_Zuban(iret)
            hm_name_2 = Tmp_Nengo(iret)
        End If

        '各値が空欄のとき
        If Reg1 = "" Then
            Reg1 = "-"
        End If

        If Reg2 = "" Then
            Reg2 = "-"
        End If

        If hm_name_1 = "" Then
            hm_name_1 = "-"
        End If

        If hm_name_2 = "" Then
            hm_name_2 = "-"
        End If

        SearchPC = 0

        Exit Function

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        Rs.Close()
        end_sql()

    End Function

    'Tire type = LT　または TBのときのマーク検索
    Function SearchLTTB(ByRef Family As String, ByRef Reg1 As String, ByRef Reg2 As String, ByRef hm_name_1 As String, ByRef hm_name_2 As String, ByVal Category As String) As Integer
        Dim sqlcmd As String
        Dim ds, ds2 As New DataSet
        Dim dt, dt2 As New DataTable
        Dim iret As Integer
        Dim Rs, Rs2 As RDO.rdoResultset
        Dim GL_T_RDO2 As T_RDO_Struct 'ＲＤＯ接続用

        GL_T_RDO2 = GL_T_RDO

        '指定条件の速度記号が表す値を検索
        If Len(w_kajyu.Text) > 0 Then
            sqlcmd = "select * from standard..IMARKTBLT WHERE size5 = '" & w_size5.Text & "' AND Category = '" & Category & "' AND kajyu_Min <= " & w_kajyu.Text & " AND kajyu_Max >= " & w_kajyu.Text & " AND (ttype = '" & w_type.Text & "' OR ttype='')"
        Else
            MsgBox("Check Load index!", vbOKOnly)
            Exit Function
        End If

        ' ''ds.Clear()
        ' ''ADO_DB_Search(sqlcmd, "Customers1", LL_StandardADO, ds)

        ' ''dt = ds.Tables("Customers1")

        ' ''If dt Is Nothing Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件に不備
        ' ''    MsgBox("Set search Key value again.", MsgBoxStyle.Critical)
        ' ''    SearchLTTB = -1
        ' ''    Exit Function
        ' ''ElseIf dt.Rows.Count > 1 Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件の指定がおかしい(複数候補あり)
        ' ''    MsgBox("Not Found Only One Mark.", MsgBoxStyle.Critical)
        ' ''    SearchLTTB = -1
        ' ''    Exit Function
        ' ''ElseIf dt.Rows.Count = 0 Then
        ' ''    'Result枠内をクリア
        ' ''    ResultClear()
        ' ''    '検索条件の指定がおかしい(候補なし)
        ' ''    MsgBox("Not Found Mark.", MsgBoxStyle.Critical)
        ' ''    SearchLTTB = -1
        ' ''    Exit Function
        ' ''End If

        '検索
        On Error GoTo error_section
        Err.Clear()
        Rs = GL_T_RDO.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
        On Error Resume Next
        Err.Clear()

        Rs.MoveFirst()

        If GL_T_RDO.Con.RowsAffected() = 0 Then
            'Result枠内をクリア
            ResultClear()
            '検索条件の指定がおかしい(候補なし)
            MsgBox("Not Found Mark.", MsgBoxStyle.Critical)
            SearchLTTB = -1
            Exit Function
        ElseIf GL_T_RDO.Con.RowsAffected() > 1 Then
            'Result枠内をクリア
            ResultClear()
            '検索条件の指定がおかしい(複数候補あり)
            MsgBox("Not Found Only One Mark.", MsgBoxStyle.Critical)
            SearchLTTB = -1
            Exit Function
        End If

        'LTの場合、ロードインデックス値がTB範囲内にないことを確認
        If Category = "LT" Then
            '指定条件の速度記号が表す値を検索
            sqlcmd = "select * from standard..IMARKTBLT WHERE Category = 'TB' AND kajyu_Min <= " & w_kajyu.Text & " AND kajyu_Max >= " & w_kajyu.Text

            ' ''ds2.Clear()

            ' ''ADO_DB_Search(sqlcmd, "Customers2", LL_StandardADO, ds2)

            ' ''dt2 = ds2.Tables("Customers2")

            '' ''TB内のロードインデックス値であれば（検索結果があった場合）エラーとする
            ' ''If dt2.Rows.Count <> 0 Then
            ' ''    'Result枠内をクリア
            ' ''    ResultClear()
            ' ''    '検索条件の指定がおかしい(TB内のロードインデックス値)
            ' ''    MsgBox("Load index is TB(When LT Data Selected).", MsgBoxStyle.Critical)
            ' ''    SearchLTTB = -1
            ' ''    Exit Function
            ' ''End If

            '検索
            On Error GoTo error_section
            Err.Clear()
            Rs2 = GL_T_RDO2.Con.OpenResultset(sqlcmd, RDO.ResultsetTypeConstants.rdOpenKeyset, RDO.LockTypeConstants.rdConcurRowVer)
            On Error Resume Next
            Err.Clear()

            Rs2.MoveFirst()

            'TB内のロードインデックス値であれば（検索結果があった場合）エラーとする
            If GL_T_RDO2.Con.RowsAffected() <> 0 Then
                'Result枠内をクリア
                ResultClear()
                '検索条件の指定がおかしい(TB内のロードインデックス値)
                MsgBox("Load index is TB(When LT Data Selected).", MsgBoxStyle.Critical)
                SearchLTTB = -1
                Exit Function
            End If
        End If


        '初期化
        Family = ""
        Reg1 = ""
        Reg2 = ""
        hm_name_1 = ""
        hm_name_2 = ""

        '値を格納
        ' ''If IsDBNull(dt.Rows(0).Item("Fname")) = False Then
        ' ''    Family = Trim(dt.Rows(0).Item("Fname").ToString)
        ' ''End If

        ' ''If IsDBNull(dt.Rows(0).Item("Register_No")) = False Then
        ' ''    Reg1 = Trim(dt.Rows(0).Item("Register_No").ToString)
        ' ''End If

        ' ''If IsDBNull(dt.Rows(0).Item("Nengo")) = False Then
        ' ''    Reg2 = Trim(dt.Rows(0).Item("Nengo").ToString)
        ' ''End If
        If IsDBNull(Rs.rdoColumns(0).Value) = False Then
            Family = Trim(Rs.rdoColumns(0).Value.ToString)
        End If

        If IsDBNull(Rs.rdoColumns(5).Value) = False Then
            Reg1 = Trim(Rs.rdoColumns(5).Value.ToString)
        End If

        If IsDBNull(Rs.rdoColumns(6).Value) = False Then
            Reg2 = Trim(Rs.rdoColumns(6).Value.ToString)
        End If



        iret = Array.IndexOf(Tmp_Family, Family)
        If iret <> -1 Then
            hm_name_1 = Tmp_Zuban(iret)
            hm_name_2 = Tmp_Nengo(iret)
        End If

        '各値が空欄のとき
        If Trim(Reg1) = "" Then
            Reg1 = "-"
        End If

        If Trim(Reg2) = "" Then
            Reg2 = "-"
        End If

        If Trim(hm_name_1) = "" Then
            hm_name_1 = "-"
        End If

        If Trim(hm_name_2) = "" Then
            hm_name_2 = "-"
        End If

        SearchLTTB = 0

        Exit Function

error_section:
        On Error Resume Next
        MsgBox("database select error.", MsgBoxStyle.Critical)
        Err.Clear()
        'Rs.Close()
        end_sql()

    End Function

    '対象値のNothingチェック
    Function NothingCheck(ByVal obj As Object) As Integer
        If obj Is Nothing Then
            NothingCheck = 0
        Else
            NothingCheck = obj
        End If
    End Function

    'テキストボックスに値が入力されたときの処理（実数値のみ入力可）
    Sub OnlyNum_jisu_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim tmp_val() As String

        If (e.KeyChar < "0"c Or e.KeyChar > "9"c) And e.KeyChar <> vbBack And _
            e.KeyChar <> "."c And e.KeyChar <> "-"c Then
            e.Handled = True
        End If

        '文字列の先頭が0で、"0"または"0."という表記ではない場合は、変更前の値に戻して処理を抜ける
        If (sender.text.Length > 0) Then
            tmp_val = Split(sender.text, ".")
            '小数点が2つ以上存在しないようにする
            If tmp_val.Length = 2 And (e.KeyChar = ".") Then
                e.Handled = True
            End If
        End If

        'マイナスが先頭以外に存在しないようにする
        If sender.text.Length > 0 Then
            If ((sender.text.Substring(0, 1) = "-") And (e.KeyChar = "-")) Then
                e.Handled = True
            End If
        End If
    End Sub
End Class