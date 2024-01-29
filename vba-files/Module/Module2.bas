Attribute VB_Name = "Module2"
Option Explicit

Sub GoodBye()
    MsgBox "GoodBye!"
End Sub

'■RPAでCSVファイルが出力されたかチェック
Function prc11_check_exest_csv()
prc_sub_no = prc_no & "-01" 'RPAで出力されたCSVがデスクトップ端末上にあるか
    csv_name = "sysData_" & Right(Replace(cur_day, "/", ""), 6) & ".csv"

    If Dir(DESKTOP_PATH & "download\" & csv_name) = "" Then
        Call write_log(csv_name & "ファイルがデスクトップ端末内に保存されていません")
        prc11_check_exest_csv = False
    Else
        'MsAccessのリンクテーブル用に名称から日付を削除したファイルを作る
        If Dir(DESKTOP_PATH & "sysData.csv") <> "" Then Kill DESKTOP_PATH & "sysData.csv"
        DoEvents
        FileCopy DESKTOP_PATH & "download\" & csv_name, DESKTOP_PATH & "sysData.csv"

        'Call write_log("") '要らない
        prc11_check_exest_csv = True
    End If
End Function



'■前日生成分のファイルをコピー、作業当日の名称で保存する
'※ver1.4～ 営業部サーバーのファイルではなく自端末のファイルをコピーして使うよう変更
Function prc12_copy_preday_report() As Boolean
    On Error GoTo err_handler

'前日生成のファイル名を取得
'前日生成したファイルがあるフォルダをDir関数に設定する
prc_sub_no = prc_no & "-01"
    Dim xls As Variant

    '>>>ここから 稼働１日目の処理が甘かったorz
    Select Case working_day_cnt
    Case 1  '検証済（４月更新時を除く）
        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then '前月実績
            xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & cur_yyyymm

        Else    '当月実績は、前月のフォルダを参照する

            Select Case Right(cur_yyyymm, 2)
            Case "01"
                 xls = Dir(DESKTOP_PATH & "create\" & CLng(Left(cur_yyyymm, 4)) - 1 & "12\*当月見込エリア別(" & fy_yyyy & ").xlsx")
                 cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & cur_yyyymm - 89

            Case "04"
                cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy - 1 & "年度") & CLng(cur_yyyymm) - 1

            Case Else
                cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & CLng(cur_yyyymm) - 1

            End Select
        End If

    Case 2, 3
        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then '前月実績
            xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "月).xlsx")
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & cur_yyyymm
        Else    '当月実績は、前月のフォルダを参照する
            'xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 1 & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
            xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
            'cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & CLng(cur_yyyymm) - 1
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & CLng(cur_yyyymm)
        End If

    Case Else
        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
        cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & cur_yyyymm

    End Select

'    If (working_day_cnt = 2 Or working_day_cnt = 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "月).xlsx")
'
'    ElseIf working_day_cnt = 1 Then
'        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'            Select Case Month(cur_day)
'            Case 4       '４月の稼働１日目
'                xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy - 1 & ").xlsx")
'            Case Else
'                xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
'            End Select
'        ElseIf (format(cur_day, "yyyymm") = cur_yyyymm) Then
'            Select Case Month(cur_day)
'            Case 1
'                xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 89 & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
'            Case Else
'                xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 1 & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
'            End Select
'        End If
'    Else
'        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*当月見込エリア別(" & fy_yyyy & ").xlsx")
'    End If
    '<<<ここまで

    'フォルダの中から最終日付のファイル名を取得（これはDドラで検索）
    Dim last_date As Integer: last_date = 0
    Do Until xls = ""
'        Debug.Print xls

        If CInt(Left(xls, 4)) = format(cur_day, "mmdd") Then 'デバッグ用、当日ファイルが既にある場合は削除
            Kill DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls

        ElseIf last_date < CInt(Left(xls, 4)) Then
            last_date = CInt(Left(xls, 4))
            xls_name = xls

        End If
        xls = Dir()
    Loop

'前日処理したファイルを共有サイトから取得する
prc_sub_no = prc_no & "-02"
'    Dim file_source As String: file_source = DESKTOP_PATH & "create\" & xls_name

    '>>>ここから 稼働１日目の処理が甘かったorz
    cur_sync_site_path = cur_sync_site_path & "\" & xls_name

'    If working_day_cnt = 1 And (format(cur_day, "yyyymm") = cur_yyyymm) Then
'        If Month(cur_day) = 1 Then
'            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & CLng(cur_yyyymm) & "\" & xls_name
'        Else
'           cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & CLng(cur_yyyymm) - 1 & "\" & xls_name
'        End If
'    Else
'        cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy年度", fy_yyyy & "年度") & cur_yyyymm & "\" & xls_name
'    End If
    '<<<ここまで

    Dim file_destination As String
    If (working_day_cnt <= 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
        xls_name = format(cur_day, "mmdd") & "当月見込エリア別(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "月).xlsx"
    Else
        xls_name = format(cur_day, "mmdd") & "当月見込エリア別(" & fy_yyyy & ").xlsx"
    End If
    file_destination = DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls_name

    '>>>ここから
    If Dir(DESKTOP_PATH & "create\" & cur_yyyymm, vbDirectory) = "" Then
        MkDir DESKTOP_PATH & "create\" & cur_yyyymm
        DoEvents
    End If
    FileCopy cur_sync_site_path, file_destination
    'FileCopy file_source, file_destination
    '<<<ここまで

    Call write_log("前日生成ファイルのコピー完了")
    prc12_copy_preday_report = True
Exit Function

err_handler:
    Call write_log("error")
'    Resume
    prc12_copy_preday_report = False
End Function

'【バックアップ】
''''■前日生成分のファイルをコピー、作業当日の名称で保存する
''''※ver1.4～ 営業部サーバーのファイルではなく自端末のファイルをコピーして使うよう変更
'''Function prc12_copy_preday_report() As Boolean
'''    On Error GoTo err_handler
'''
'''prc_sub_no = prc_no & "-01"
'''    Dim xls As Variant
'''    If (working_day_cnt = 2 Or working_day_cnt = 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'''        xls = Dir(DESKTOP_PATH & "create\*当月見込エリア別(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "月).xlsx")
'''    ElseIf working_day_cnt = 1 And Month(cur_day) = 4 Then      '４月の稼働１日目
'''        xls = Dir(DESKTOP_PATH & "create\*当月見込エリア別(" & fy_yyyy - 1 & ").xlsx")
'''    Else
'''        xls = Dir(DESKTOP_PATH & "create\*当月見込エリア別(" & fy_yyyy & ").xlsx")
'''    End If
'''
'''    Dim last_date As Integer: last_date = 0
'''    Do Until xls = ""
'''        If CInt(Left(xls, 4)) = format(cur_day, "mmdd") Then 'デバッグ用、当日ファイルが既にある場合は削除
'''            Kill DESKTOP_PATH & "create\" & xls
'''
'''        ElseIf last_date < CInt(Left(xls, 4)) Then
'''            last_date = CInt(Left(xls, 4))
'''            xls_name = xls
'''
'''        End If
'''        xls = Dir()
'''    Loop
'''
'''prc_sub_no = prc_no & "-02"
'''    Dim file_source As String: file_source = DESKTOP_PATH & "create\" & xls_name
'''
'''    Dim file_destination As String
'''    If (working_day_cnt <= 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'''        xls_name = format(cur_day, "mmdd") & "当月見込エリア別(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "月).xlsx"
'''    Else
'''        xls_name = format(cur_day, "mmdd") & "当月見込エリア別(" & fy_yyyy & ").xlsx"
'''    End If
'''    file_destination = DESKTOP_PATH & "create\" & xls_name
'''
'''    FileCopy file_source, file_destination
'''
'''    Call write_log("前日生成ファイルのコピー完了")
'''    prc12_copy_preday_report = True
'''Exit Function
'''
'''err_handler:
'''    Call write_log("error")
'''    prc12_copy_preday_report = False
'''End Function

'■
Function prc21_exec_accessDB() As Boolean
'Function step111_Exec_Acc() As Boolean
'--- 変更履歴   ---
'   ver.1.1     2018/10/31  katouk48    最終動作確認完了
'   ver.1.4     2021/12/14  katok21：メンテ
'------------------

    On Error GoTo err_handler

prc_sub_no = prc_no & "-01"
    Dim obj_access As Object: Set obj_access = CreateObject("Access.Application")
    Dim qd As Object
    Dim str_sql As String

    Const acTable = 0
    Const acViewNormal = 0
    Const acEdit = 1

prc_sub_no = prc_no & "-02"
    With obj_access
        .OpenCurrentDatabase (ThisWorkbook.path & "\" & DB_NAME)    'AccessDBを開く
        .DoCmd.SetWarnings False    'アラートを表示させない
        .Visible = True

        prc_sub_no = prc_no & "-03"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl10_納期回答sys"
        .DoCmd.DeleteObject acTable, "tbl11_納期回答sys_copy"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        .DoCmd.OpenQuery "qry10_納期回答sysデータテーブル作成", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry11_コピーテーブル作成", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry12_販売金額更新", acViewNormal, acEdit

        prc_sub_no = prc_no & "-04"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl20_総本_売上高"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        'cur_yyyymm = "202112"   '※デバッグ用
        str_sql = "SELECT a.品種コード, h.日本語名１ AS 品種名, m1.地域ID, m1.エリア, " & _
               "a.本部コード, s1.日本語名１ AS 本部名, a.支社コード, s2.日本語名１ AS 支社名, a.支店コード, s3.日本語名１ AS 支店名," & _
               "a.営業所コード, s4.日本語名１ AS 営業所名, a.ルートコード, s5.日本語名１ AS ルート名, " & _
               "a.事業所コード, s6.事業所名, " & _
               "a.売上年月, Sum(a.売上高＿前年) AS 売上高＿前年, Sum(a.売上高＿実績) AS 売上高＿実績 " & _
        "INTO tbl20_総本_売上高 " & _
        "FROM ((((((( SOUHON_全予実事業所別FACT＿当期 a " & _
        "INNER JOIN SOUHON_品種名称 h ON a.品種コード = h.コード ) " & _
        "INNER JOIN SOUHON_組織名称 s1 ON a.本部コード = s1.コード ) " & _
        "INNER JOIN SOUHON_組織名称 s2 ON a.支社コード = s2.コード ) " & _
        "INNER JOIN SOUHON_組織名称 s3 ON a.支店コード = s3.コード ) " & _
        "INNER JOIN SOUHON_組織名称 s4 ON a.営業所コード = s4.コード ) " & _
        "INNER JOIN SOUHON_事業所マスタ＿当月 s6 ON a.事業所コード = s6.事業所コード ) " & _
        "LEFT JOIN SOUHON_組織名称 s5 ON a.[ルートコード] = s5.コード ) " & _
        "LEFT JOIN mst_支社_新旧対比 m1 ON a.支社コード = m1.統轄支店コード " & _
        "GROUP BY a.品種コード, h.日本語名１, m1.地域ID, m1.エリア, " & _
                 "a.本部コード, s1.日本語名１, a.支社コード, s2.日本語名１, a.支店コード, s3.日本語名１, " & _
                 "a.営業所コード, s4.日本語名１, a.事業所コード, s6.事業所名, a.ルートコード, s5.日本語名１, " & _
                 "a.売上年月 " & _
        "HAVING (a.品種コード='T41334')" ' AND a.売上年月='" & cur_yyyymm & "');"
        Set qd = .CurrentDb.QueryDefs("qry21_総本売上高_取得")
        qd.Sql = str_sql
        .DoCmd.OpenQuery "qry21_総本売上高_取得", acViewNormal, acEdit


        prc_sub_no = prc_no & "-05"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl30_総本_受注残"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        str_sql = "SELECT j.本部コード, j.統轄支店コード, j.統轄支店名, j.営業所コード, j.営業所名, a.事業所コード, a.事業所名, " & _
                       "a.品種コード, a.品種名, a.年月, Sum(a.売上計画) AS 売上計画, Sum(a.売上高) AS 売上高, Sum(a.受注残) AS 受注残 " & _
                "INTO tbl30_総本_受注残 " & _
                "FROM SOUHON_受注分析受注残＿事業所 a " & _
                "INNER JOIN SOUHON_事業所マスタ＿当月 j ON a.事業所コード = j.事業所コード " & _
                "WHERE a.Vレベルコード = 'V00100' " & _
                "GROUP BY j.本部コード, j.統轄支店コード, j.統轄支店名, j.営業所コード, j.営業所名, a.事業所コード, a.事業所名, a.品種コード, a.品種名, a.年月 " & _
                "HAVING a.品種コード = 'T41334' AND a.年月 = '" & cur_yyyymm & "';"
        Set qd = .CurrentDb.QueryDefs("qry31_総本受注残_取得")
        qd.Sql = str_sql
        .DoCmd.OpenQuery "qry31_総本受注残_取得", acViewNormal, acEdit

        .DoCmd.SetWarnings True

        prc_sub_no = prc_no & "-06"
        .Quit
    End With


prc_sub_no = prc_no & "-99"
    Set obj_access = Nothing

    Call write_log("AccessDB更新完了")
    prc21_exec_accessDB = True
Exit Function

err_handler:
    Call write_log("error")
    prc21_exec_accessDB = False
End Function


'◆稼働１日目・３日目用
Function prc22_update_past_record() As Boolean
'--- 変更履歴   ---
'   ver.1.4     2021/12/20  katok21：新規設定
'------------------

    On Error GoTo err_handler

prc_sub_no = prc_no & "-01"
    Dim obj_access As Object: Set obj_access = CreateObject("Access.Application")

    Const acTable = 0
    Const acViewNormal = 0
    Const acEdit = 1

prc_sub_no = prc_no & "-02"
    With obj_access
        .OpenCurrentDatabase (ThisWorkbook.path & "\" & DB_NAME)    'AccessDBを開く
        .DoCmd.SetWarnings False    'アラートを表示させない
        .Visible = True

        prc_sub_no = prc_no & "-03"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl40_(月初用)過去実績"
        .DoCmd.DeleteObject acTable, "tbl41_tbl40+支社エリア付与"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        prc_sub_no = prc_no & "-04"
        .DoCmd.OpenQuery "qry41_組替過去実績_当期", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry42_組替過去実績_前１", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry43_組替過去実績_前２", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry44_組替過去実績_前３", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry45_組替過去実績_前４", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry46_組替過去実績_前５", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry47_組替過去実績_前６", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry48_tbl40+支社エリア付与", acViewNormal, acEdit

        prc_sub_no = prc_no & "-05"
        .DoCmd.SetWarnings True
        .Quit
    End With


prc_sub_no = prc_no & "-99"
    Set obj_access = Nothing

    Call write_log("稼働" & format(working_day_cnt, "00") & "日目：過去実績 更新完了")
    prc22_update_past_record = True
Exit Function

err_handler:
    Call write_log("error")
    prc22_update_past_record = False
End Function



'■納期回答Sys「基データ」シート貼り付け、「販売金額」の値変更
Function prc31_input_rawdata() As Boolean
'Function step121_Paste_RawData() As Boolean
'Function step122_Update_Price() As Boolean
'--- 変更履歴   ---
'   ver.1.0     2018/10/25  katouk48    安斎さん最終打ち合わせ後、追加要望
'       「基データ」シートへの転記は不要かと思ったけども、確認用に必要だって
'   ver.1.1     2018/10/31  katouk48    最終動作確認完了
'   ver.1.4     2021/12/14  katok21：メンテ、２つのプロシージャ統合
'------------------

    On Error GoTo err_handler

'    xls_name = "1213当月見込エリア別(2021).xlsx"    '※デバッグ用

prc_sub_no = prc_no & "-01"
    Dim bk As Workbook
    Dim is_open As Boolean: is_open = False
    For Each bk In Workbooks
        If bk.Name = xls_name Then
            is_open = True
            Exit For
        End If
    Next bk
    If is_open = False Then Workbooks.Open DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls_name
    Workbooks(xls_name).Activate
    DoEvents    'なんとなく

prc_sub_no = prc_no & "-02"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")

    Dim str_sql As String
    str_sql = "SELECT * FROM [tbl10_納期回答sys] WHERE キャンセル日 IS NULL "

    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-03"
    Dim r_num As Long

    Workbooks(xls_name).Worksheets("基データ").Activate
    r_num = Range("A2").End(xlDown).Row
    Rows("2:" & r_num).ClearContents

    Range("A2").CopyFromRecordset ref_access.rs
    Set ref_access = Nothing

prc_sub_no = prc_no & "-04"
    Call set_datapage_configuration

prc_sub_no = prc_no & "-05"
    Const c_num_E As Integer = 5     '管理
    Const c_num_H As Integer = 8     '発注内容区分
    Const c_num_AD As Integer = 30   '販売金額

    r_num = 2
    Do Until Cells(r_num, 1).Value = ""
        If Left(Cells(r_num, c_num_E).Value, 1) = "F" And _
            (Cells(r_num, c_num_H).Value = "2" Or Cells(r_num, c_num_H).Value = "3") Then

            Cells(r_num, c_num_AD).Formula = "=CK" & r_num & "*1.43"
        End If
        r_num = r_num + 1
    Loop

prc_sub_no = prc_no & "-99"
    Call write_log("[基データ]シート データ貼付けOK")
    prc31_input_rawdata = True
Exit Function

not_connect_db:
    Set ref_access = Nothing

    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc31_input_rawdata = False
End Function



'■６シート（当月全FAX注文～前日注文）へデータ貼り付け
Function prc32_input_toku_sys_data() As Boolean
'Function step131_Input_TokuSysData() As Boolean
'--- 変更履歴   ---
'   ver.1.1.0   2018/10/30  katouk48：確認済
'   ver.1.1.1   2018/11/01  katouk48：
'   ver.1.4     2021/12/14  katok21：メンテ
'------------------

    On Error GoTo err_handler

'    cur_day = Date
'    cur_yyyymm = "202112"    '※デバッグ用
'    xls_name = "1213当月見込エリア別(2021).xlsx"    '※デバッグ用

    Workbooks(xls_name).Activate

'抽出条件に使用する日付を変数にセット
prc_sub_no = prc_no & "-01"
    Dim r_num As Long, c_num As Integer
    Dim fld As Object
    Dim str_sql As String

    '１ヶ月前の最終日をセット
    Dim pre_month_lastday As Date
    pre_month_lastday = DateSerial(Left(cur_yyyymm, 4), Right(cur_yyyymm, 2), 1) - 1
    '当月の最終日をセット
    Dim cur_month_lastday As Date
    cur_month_lastday = DateSerial(Left(cur_yyyymm, 4), Right(cur_yyyymm, 2), 1)
    cur_month_lastday = DateAdd("m", 1, cur_month_lastday) - 1
    '翌月、翌々月の最終日をセット
    Dim nxt1_month_lastday As Date, nxt2_month_lastday As Date
    nxt1_month_lastday = DateAdd("m", 1, cur_month_lastday) - 1
    nxt2_month_lastday = DateAdd("m", 1, nxt1_month_lastday) - 1

'総本サーバー接続
prc_sub_no = prc_no & "-02"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")


'①「当月全FAX注文」シート入力
prc_sub_no = prc_no & "-10"
    Worksheets("当月全FAX注文").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-11"
    str_sql = "SELECT [*発注日], 発注元, 発注内容区分, 旧支社コード, [*販売金額], 確定出荷日, EOC伝票処理日, 統轄支店コード, 統轄支店名 " & _
            "FROM tbl11_納期回答sys_copy " & _
            "WHERE (確定出荷日 >= #" & pre_month_lastday & "# AND 確定出荷日 < #" & cur_month_lastday & "#) "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-12"
    c_num = 1
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close

prc_sub_no = prc_no & "-13"
    Call set_datapage_configuration


'②「当月注残」シート入力
prc_sub_no = prc_no & "-20"
    Worksheets("当月注残").Select
    Cells.ClearContents

    If format(cur_day, "yyyymm") = cur_yyyymm And working_day_cnt < 3 Then
        prc_sub_no = prc_no & "-21"
        Worksheets("当月全FAX注文").Cells.Copy Destination:=Worksheets("当月注残").Range("A1")

    Else
        prc_sub_no = prc_no & "-22"
        Call ref_access.open_rs(str_sql)    '貼り付けるデータは"prc_no.32-11"のまま変更なし
        If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

        prc_sub_no = prc_no & "-23"
        c_num = 1
        For Each fld In ref_access.rs.fields
            Cells(1, c_num).Value = fld.Name
            c_num = c_num + 1
        Next fld
        Range("A2").CopyFromRecordset ref_access.rs
        ref_access.rs.Close

        prc_sub_no = prc_no & "-24"
        r_num = 1
        Do Until Cells(r_num, 1).Value = ""
            If Cells(r_num, 6).Value <= (cur_day - 1) Then  '「確定出荷日」が作業日前日より前なら
                If Cells(r_num, 7).Value <> cur_day Then    '「EOC伝票処理日」が作業日当日で無ければ
                    Rows(r_num).Delete
                Else
                    r_num = r_num + 1
                End If
            Else
                r_num = r_num + 1
            End If
        Loop
    End If

prc_sub_no = prc_no & "-25"
    Call set_datapage_configuration


'③「翌月納期確定」シート入力
prc_sub_no = prc_no & "-30"
    Worksheets("翌月納期確定").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-31"
    str_sql = "SELECT [*発注日], 発注元, 発注内容区分, 旧支社コード, [*販売金額], 確定出荷日, EOC伝票処理日, 統轄支店コード, 統轄支店名 " & _
            "FROM tbl11_納期回答sys_copy " & _
            "WHERE (確定出荷日 >= #" & cur_month_lastday & "# AND 確定出荷日 < #" & nxt1_month_lastday & "#) "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-32"
    c_num = 1
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close

prc_sub_no = prc_no & "-33"
    Call set_datapage_configuration


'④「翌々月」シート入力
prc_sub_no = prc_no & "-40"
    Worksheets("翌々月").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-41"
    str_sql = "SELECT [*発注日], 発注元, 発注内容区分, 旧支社コード, [*販売金額], 確定出荷日, EOC伝票処理日, 統轄支店コード, 統轄支店名 " & _
            "FROM tbl11_納期回答sys_copy " & _
            "WHERE (確定出荷日 >= #" & nxt1_month_lastday & "# AND 確定出荷日 < #" & nxt2_month_lastday & "#) "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-42"
    c_num = 1
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close

prc_sub_no = prc_no & "-43"
    Call set_datapage_configuration


'⑤「納期未確定」シート入力
prc_sub_no = prc_no & "-50"
    Worksheets("納期未確定").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-51"
    str_sql = "SELECT [*発注日], 発注元, 発注内容区分, 旧支社コード, [*販売金額], 確定出荷日, EOC伝票処理日, 統轄支店コード, 統轄支店名 " & _
            "FROM tbl11_納期回答sys_copy " & _
            "WHERE 確定出荷日 IS NULL "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-52"
    c_num = 1
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close

prc_sub_no = prc_no & "-53"
    Call set_datapage_configuration


'⑥「前日注文」シート入力
prc_sub_no = prc_no & "-60"
    Worksheets("前日注文").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-61"
    str_sql = "SELECT [*発注日], 発注元, 発注内容区分, 旧支社コード, [*販売金額], 確定出荷日, EOC伝票処理日, 統轄支店コード, 統轄支店名 " & _
            "FROM tbl11_納期回答sys_copy " & _
            "WHERE [*発注日] = #" & (cur_day - 1) & "# "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-62"
    c_num = 1
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close

prc_sub_no = prc_no & "-63"
    Call set_datapage_configuration


'後片付け
prc_sub_no = prc_no & "-99"
    Set ref_access = Nothing

    Call write_log("納期回答sys データ貼付け完了")
    prc32_input_toku_sys_data = True
Exit Function

not_connect_db:
    Set ref_access = Nothing

    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Set ref_access = Nothing

    Call write_log("error")
    prc32_input_toku_sys_data = False
End Function

