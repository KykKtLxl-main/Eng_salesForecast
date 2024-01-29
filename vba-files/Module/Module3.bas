Attribute VB_Name = "Module3"
Option Explicit

'■概要
Function prc41_input_TEMSSresult() As Boolean
'Function step141_Input_LHT() As Boolean
'Function step142_Input_HanbaiJgyo() As Boolean
'Function step143_Input_Tokuju() As Boolean
'Function step144_Input_LBT() As Boolean
'Function step145_Input_LWT() As Boolean
'Function step151_Input_ExtDairitenRoot() As Boolean
'Function step152_Input_ExtTokkenRoot() As Boolean
'Function step161_Input_LHTkanto() As Boolean
'--- 変更履歴   ---
'   ver.1.1     2018/10/31  katouk48    最終動作確認完了
'   ver.1.4     2021/12/14  katok21：メンテ⇒
'                    事業所別実績シートを追加することにしたので、マクロ実行ではなく数式での参照に変更
'------------------

    On Error GoTo err_handler

'    Dim str_sql As String
    Workbooks(xls_name).Activate
    Worksheets("TEMSS実績").Select
    Range("A1").Value = cur_yyyymm

'prc_sub_no = prc_no & "-01"
'    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
'    Call ref_access.connect("access")


''□Ｎ列「ＬＨＴ」欄
'prc_sub_no = prc_no & "-10"
'    Range("N2:N11").ClearContents
'
'prc_sub_no = prc_no & "-11"
'    str_sql = "SELECT SUM(売上高＿実績) AS 売上高 " & _
'        "FROM [Tbl20_総本_売上高] " & _
'        "WHERE 本部コード = 'P00300' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-12"
'    Range("N2").CopyFromRecordset ref_access.rs
'    ref_access.rs.Close

'prc_sub_no = prc_no & "-13"
'    str_sql = "SELECT エリア, sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00300' " & _
'            "GROUP BY 地域ID, エリア " & _
'            "ORDER BY 地域ID ASC"
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-14"
'    Dim ssya As Range: Set ssya = Range("M3:" & Range("M3").End(xlDown).Address)
'    Dim r As Range
'    For Each r In ssya
'        If ref_access.rec_cnt <= 1 Then
'            r.Offset(0, 1).Value = 0
'        Else
'            Do Until ref_access.rs.EOF
'                If Replace(r.Value, "支社", "") = ref_access.rs.fields("エリア") Then
'                    r.Offset(0, 1).Value = ref_access.rs.fields("売上高")
'                    Exit Do
'                Else
'                    ref_access.rs.MoveNext
'                End If
'            Loop
'
'            If r.Offset(0, 1).Value = "" Then r.Offset(0, 1).Value = 0
'            ref_access.rs.movefirst
'        End If
'    Next r
'    ref_access.rs.Close
    

''□Ｐ列「販売事業部」欄
'prc_sub_no = prc_no & "-20"
'    Range("P2:P11").ClearContents
'
'prc_sub_no = prc_no & "-21"
'    str_sql = "SELECT SUM(売上高＿実績) AS 売上高 " & _
'              "FROM [Tbl20_総本_売上高] WHERE 支社コード = 'N01420' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-22"
'    Range("P2").CopyFromRecordset ref_access.rs
'    ref_access.rs.Close
    
    
''□Ｒ列「特需」欄
'prc_sub_no = prc_no & "-30"
'    Range("R2,Q3:R11").ClearContents
'
'prc_sub_no = prc_no & "-31"
'    str_sql = "SELECT SUM(売上高＿実績) AS 売上高 " & _
'              "FROM [Tbl20_総本_売上高] WHERE 支社コード = 'N00245' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-32"
'    If ref_access.rs.fields("売上高") = 0 Or IsNull(ref_access.rs.fields("売上高")) Then
'        Range("R2").Value = 0
'    Else
'        Range("R2").CopyFromRecordset ref_access.rs
'    End If
'    ref_access.rs.Close
    
    
''□Ｔ列「LBT」欄
'prc_sub_no = prc_no & "-40"
'    Range("T2:T11").ClearContents
'
'prc_sub_no = prc_no & "-41"
'    str_sql = "SELECT SUM(売上高＿実績) AS 売上高 " & _
'              "FROM [Tbl20_総本_売上高] WHERE 本部コード = 'P00400' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-42"
'    If ref_access.rs.fields("売上高") = 0 Or IsNull(ref_access.rs.fields("売上高")) Then
'        Range("T2").Value = 0
'    Else
'        Range("T2").CopyFromRecordset ref_access.rs
'    End If
'    ref_access.rs.Close
    
    
''□Ｖ列「LWT」欄
'prc_sub_no = prc_no & "-50"
'    Range("V2:V11").ClearContents
'
'prc_sub_no = prc_no & "-51"
'    str_sql = "SELECT SUM(売上高＿実績) AS 売上高 " & _
'              "FROM [Tbl20_総本_売上高] WHERE 本部コード = 'P00200' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-52"
'    If ref_access.rs.fields("売上高") = 0 Or IsNull(ref_access.rs.fields("売上高")) Then
'        Range("V2").Value = 0
'    Else
'        Range("V2").CopyFromRecordset ref_access.rs
'    End If
'    ref_access.rs.Close
'
'prc_sub_no = prc_no & "-53"
'    str_sql = "SELECT エリア, sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00200' " & _
'            "GROUP BY 地域ID, エリア " & _
'            "ORDER BY 地域ID ASC"
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-54"
'    If ref_access.rs.fields("売上高") = 0 Or IsNull(ref_access.rs.fields("売上高")) Then
'        Range("V2").Value = 0
'    Else
'        Range("V2").CopyFromRecordset ref_access.rs
'        ref_access.rs.Close
'
'        prc_sub_no = prc_no & "-55"
'        Set ssya = Range("V3:" & Range("V3").End(xlDown).Address)
'        For Each r In ssya
'            If ref_access.rec_cnt <= 1 Then
'                r.Offset(0, 1).Value = 0
'            Else
'                Do Until ref_access.rs.EOF
'                    If Replace(r.Value, "支社", "") = ref_access.rs.fields("エリア") Then
'                        r.Offset(0, 1).Value = ref_access.rs.fields("売上高")
'                        Exit Do
'                    Else
'                        ref_access.rs.MoveNext
'                    End If
'                Loop
'
'                If r.Offset(0, 1).Value = "" Then r.Offset(0, 1).Value = 0
'                ref_access.rs.movefirst
'            End If
'        Next r
'        ref_access.rs.Close
'    End If


''□Ｃ列「EXT代理店」欄
'prc_sub_no = prc_no & "-60"
'    Range("C2:C11").ClearContents
'
'prc_sub_no = prc_no & "-61"
'    str_sql = "SELECT Sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00300' AND ルートコード = 'J10003' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-62"
'    Range("C2").CopyFromRecordset ref_access.rs
'    ref_access.rs.Close
'
'prc_sub_no = prc_no & "-63"
'    str_sql = "SELECT エリア, sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00300' " & _
'            "GROUP BY 地域ID, エリア " & _
'            "ORDER BY 地域ID ASC"
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-64"
'    Set ssya = Range("C3:" & Range("C3").End(xlDown).Address)
'    For Each r In ssya
'        If ref_access.rec_cnt <= 1 Then
'            r.Offset(0, 1).Value = 0
'        Else
'            Do Until ref_access.rs.EOF
'                If Replace(r.Value, "支社", "") = ref_access.rs.fields("エリア") Then
'                    r.Offset(0, 1).Value = ref_access.rs.fields("売上高")
'                    Exit Do
'                Else
'                    ref_access.rs.MoveNext
'                End If
'            Loop
'
'            If r.Offset(0, 1).Value = "" Then r.Offset(0, 1).Value = 0
'            ref_access.rs.movefirst
'        End If
'    Next r
'    ref_access.rs.Close
    
    
''□Ｆ列「特建」欄
'prc_sub_no = prc_no & "-70"
'    Range("F2:F9").ClearContents
'
'prc_sub_no = prc_no & "-71"
'    str_sql = "SELECT Sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00300' AND ルートコード = 'J10013' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-72"
'    Range("F2").CopyFromRecordset ref_access.rs
'    ref_access.rs.Close
'
'prc_sub_no = prc_no & "-73"
'    str_sql = "SELECT エリア, Sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE 本部コード = 'P00300' AND ルートコード = 'J10013' " & _
'            "GROUP BY 地域ID, エリア " & _
'            "ORDER BY 地域ID ASC;"
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-74"
'    Set ssya = Range("E3:" & Range("E3").End(xlDown).Address)
'    For Each r In ssya
'        If ref_access.rec_cnt <= 1 Then
'            r.Offset(0, 1).Value = 0
'        Else
'            Do Until ref_access.rs.EOF
'                If Replace(r.Value, "支社", "") = ref_access.rs.fields("エリア") Then
'                    r.Offset(0, 1).Value = ref_access.rs.fields("売上高")
'                    Exit Do
'                Else
'                    ref_access.rs.MoveNext
'                End If
'            Loop
'
'            If r.Offset(0, 1).Value = "" Then r.Offset(0, 1).Value = 0
'            ref_access.rs.movefirst
'        End If
'    Next r
'    ref_access.rs.Close
    
    
'□A13セル～「関東エクステリア支店」欄 -----------
prc_sub_no = prc_no & "-80"
    Range("A15:C42").ClearContents
'
prc_sub_no = prc_no & "-81"
    Dim ref_excel As ClassAdo: Set ref_excel = New ClassAdo
    Call ref_excel.connect("excel", DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls_name)

    Dim str_sql As String
'    str_sql = "SELECT Sum(売上高＿実績) AS 売上高 " & _
'            "FROM [Tbl20_総本_売上高] " & _
'            "WHERE エリア = '関東' AND ルートコード = 'J10003' "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
'
'prc_sub_no = prc_no & "-82"
'    Range("C14").CopyFromRecordset ref_access.rs
'    ref_access.rs.Close

prc_sub_no = prc_no & "-83"
    str_sql = "SELECT 支社名, 営業所名, Sum(売上高＿実績) AS 売上高 " & _
            "FROM [事業所別実績$] " & _
            "WHERE エリア = '関東' AND ルートコード = 'J10003' AND 売上年月 = '" & cur_yyyymm & "'" & _
            "GROUP BY 地域ID, エリア, 支社名, 営業所コード, 営業所名 " & _
            "ORDER BY 地域ID ASC, 営業所コード ASC;"
    Call ref_excel.open_rs(str_sql)
    If ref_excel.is_cn = False Or ref_excel.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-84"
    Range("A15").CopyFromRecordset ref_excel.rs
    ref_excel.rs.Close
'--------------------------------------------------
    
prc_sub_no = prc_no & "-99"
'    Set ref_excel = Nothing

    Call write_log("[TEMSS実績]シート データ貼付けOK")
    prc41_input_TEMSSresult = True
    Exit Function

not_connect_db:
'    Set ref_excel = Nothing
    
    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc41_input_TEMSSresult = False
    
End Function



'■概要
Function prc42_input_juchuzan() As Boolean
'Function step171_Input_Juchuzan() As Boolean
'--- 変更履歴   ---
'   ver.1.1     2018/10/31  katouk48    最終動作確認完了
'   ver.1.4     2021/12/14  katok21：メンテ
'------------------

    On Error GoTo err_handler
    
prc_sub_no = prc_no & "-01"
    Workbooks(xls_name).Activate
    Worksheets("サッシR注残").Activate
    Range("A2:F65536").ClearContents
    Range("F1").Value = cur_day
    
prc_sub_no = prc_no & "-02"
    Dim str_sql As String
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")

prc_sub_no = prc_no & "-11"
    str_sql = "SELECT 品種名, 統轄支店名, 事業所名, 事業所コード, Sum(受注残) AS 受注残 " & _
            "FROM [Tbl30_総本_受注残] " & _
            "WHERE 本部コード = 'P00300' " & _
            "GROUP BY 品種名, 統轄支店コード, 統轄支店名, 事業所名, 事業所コード " & _
            "ORDER BY 統轄支店コード ASC, 事業所コード ASC "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
    
prc_sub_no = prc_no & "-12"
    Dim pre_rec As String

    Dim r_num As Long: r_num = 2
    Do Until ref_access.rs.EOF
        If pre_rec = "" Or pre_rec = ref_access.rs.fields("統轄支店名") Then
            Cells(r_num, 2).Value = ref_access.rs.fields("品種名")
            Cells(r_num, 3).Value = ref_access.rs.fields("統轄支店名")
            Cells(r_num, 4).Value = ref_access.rs.fields("事業所名")
            Cells(r_num, 5).Value = ref_access.rs.fields("事業所コード")
            Cells(r_num, 6).Value = ref_access.rs.fields("受注残") / 1000
            pre_rec = ref_access.rs.fields("統轄支店名")
            ref_access.rs.MoveNext
        Else
            Cells(r_num, 2).Value = ref_access.rs.fields("品種名")
            Cells(r_num, 3).Value = pre_rec
            Cells(r_num, 4).Value = "合計"
            Cells(r_num, 6).Value = Application.WorksheetFunction.SumIf(Columns("C"), pre_rec, Columns("F"))
            pre_rec = ""
        End If

        r_num = r_num + 1
    Loop

prc_sub_no = prc_no & "-13"
    ref_access.rs.movefirst
    Cells(r_num, 2).Value = ref_access.rs.fields("品種名")
    Cells(r_num, 3).Value = pre_rec
    Cells(r_num, 4).Value = "合計"
    Cells(r_num, 6).Value = Application.WorksheetFunction.SumIf(Columns("C"), pre_rec, Columns("F"))
    ref_access.rs.Close


prc_sub_no = prc_no & "-99"
    Set ref_access = Nothing
    
    Call write_log("[サッシR注残]シート データ貼付けOK")
    prc42_input_juchuzan = True
    Exit Function

not_connect_db:
    Set ref_access = Nothing
    
    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc42_input_juchuzan = False
    
End Function



Function prc51_input_result_by_jgyo() As Boolean
'Function step181_forKanto_Kitakanto() As Boolean
'   ver.1.4     2021/12/14  katok21：[事業所別実績]シート新設、中部・関西追加要望のため全事業所の実績シートを作る（面倒になったorz）

prc_sub_no = prc_no & "-10"
    Workbooks(xls_name).Activate
    
    On Error Resume Next
    Worksheets("事業所別実績").Activate
    If ActiveSheet.AutoFilterMode = True Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False
    
    'デバッグ用、「事業所別実績」シートが無ければ生成
    If Err.Number <> "0" Then
        Worksheets.Add after:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = "事業所別実績"
    Else
        Cells.ClearContents
    End If
    On Error GoTo err_handler
    
prc_sub_no = prc_no & "-11"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")

    Dim str_sql As String: str_sql = "SELECT * FROM tbl20_総本_売上高 "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db
    
prc_sub_no = prc_no & "-12"
    Dim c_num As Integer: c_num = 1
    Dim fld As Object
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    ref_access.rs.Close
    
prc_sub_no = prc_no & "-13"
    Call set_datapage_configuration
    
prc_sub_no = prc_no & "-99"
    Set ref_access = Nothing
    
    Call write_log("[事業所別実績]シート データ貼付けOK")
    prc51_input_result_by_jgyo = True
    Exit Function

not_connect_db:
    Set ref_access = Nothing
    
    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc51_input_result_by_jgyo = False
End Function


'◆稼働１日目・３日目用
Function prc52_input_pastresult_by_jgyo() As Boolean
'--- 変更履歴   ---
'   ver.1.4     2021/12/20  katok21：新規設定
'------------------

prc_sub_no = prc_no & "-10"
    Workbooks(xls_name).Activate

    On Error Resume Next
    Worksheets("事業所別実績_過去２年・月別").Activate
    If ActiveSheet.AutoFilterMode = True Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False

    'デバッグ用、「事業所別実績」シートが無ければ生成
    If Err.Number <> "0" Then
        Worksheets.Add after:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = "事業所別実績_過去２年・月別"
    Else
        Cells.ClearContents
    End If
    On Error GoTo err_handler

prc_sub_no = prc_no & "-11"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")

    Dim str_sql As String: str_sql = "SELECT * FROM [tbl41_tbl40+支社エリア付与] "
    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-12"
    Dim c_num As Integer: c_num = 1
    Dim fld As Object
    For Each fld In ref_access.rs.fields
        Cells(1, c_num).Value = fld.Name
        c_num = c_num + 1
    Next fld
    Range("A2").CopyFromRecordset ref_access.rs
    
    Dim rec As Integer: rec = ref_access.rec_cnt
    ref_access.rs.Close


'◇累計値の数式を入力
prc_sub_no = prc_no & "-13"
    Cells(1, c_num).Value = "累計"
    Cells(2, c_num).Formula = formula_sum_past_result
    Cells(2, c_num).Copy Destination:=Range(Cells(2, c_num), Cells(rec + 1, c_num))
    DoEvents    '一応

prc_sub_no = prc_no & "-14"
    Call set_datapage_configuration

prc_sub_no = prc_no & "-99"
    Set ref_access = Nothing

    Call write_log("[過去２年・月別]シート データ貼付けOK")
    prc52_input_pastresult_by_jgyo = True
    Exit Function

not_connect_db:
    Set ref_access = Nothing

    Call write_log("作業用ＤＢ接続エラー")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc52_input_pastresult_by_jgyo = False
End Function


Function formula_sum_past_result() As String
    Select Case Right(cur_yyyymm, 2)
    Case "04", "10"
        formula_sum_past_result = "=sum(J2:J2)"
    Case "05", "11"
        formula_sum_past_result = "=sum(J2:K2)"
    Case "06", "12"
        formula_sum_past_result = "=sum(J2:L2)"
    Case "07", "01"
        formula_sum_past_result = "=sum(J2:M2)"
    Case "08", "02"
        formula_sum_past_result = "=sum(J2:N2)"
    Case "09", "03"
        formula_sum_past_result = "=sum(J2:O2)"
    End Select
End Function


'◆稼働１日目・３日目用
Function prc53_input_month_area_sheet() As Boolean
'--- 変更履歴   ---
'   ver.1.4     2021/12/20  katok21：新規設定
'------------------

    On Error GoTo err_handler

prc_sub_no = prc_no & "-01"
    Worksheets("【公共品見込資料】").Activate
    Range("year_month").Value = Right(fy_yyyy, 2) & base_month
    Range("year_month").Offset(0, 1).Value = IIf(base_month = "04", "上期累計", "下期累計")
    
    Dim pre_month As Integer
    If Right(cur_yyyymm, 2) = "01" Then pre_month = 12 Else pre_month = CInt(Right(cur_yyyymm, 2)) - 1
    Range("year_month").Offset(1, 2).Value = base_month & "-" & pre_month & "月実績"
    Range("year_month").Offset(1, 3).Value = base_month & "-" & Right(cur_yyyymm, 2) & "月計画"
    Range("year_month").Offset(1, 4).Value = base_month & "-" & Right(cur_yyyymm, 2) & "月前年"
    Range("year_month").Offset(1, 5).Value = base_month & "-" & Right(cur_yyyymm, 2) & "月前々年"

prc_sub_no = prc_no & "-02"
    Dim sh As Worksheet, sh_array As Variant
    Set sh_array = Worksheets(Array("北関東【公共品見込資料】", "関東【公共品見込資料】", "中部【公共品見込資料】", "関西【公共品見込資料】", "中四国【公共品見込資料】"))

    For Each sh In sh_array
        sh.Activate
        'デバッグ用
        sh.Range("year").Value = ""
        sh.Range("month").Value = ""
        
        sh.Range("year").Value = Right(fy_yyyy, 2)
        sh.Range("month").Value = base_month
        sh.Range("month").Offset(0, 2).Value = IIf(base_month = "04", "上期累計", "下期累計")
    Next sh

    Call write_log("各エリアシート 年月更新OK")
    prc53_input_month_area_sheet = True
    Exit Function

err_handler:
    Call write_log("error")
    prc53_input_month_area_sheet = False

End Function

Function base_month() As String
    Select Case Right(cur_yyyymm, 2)
    Case "04", "05", "06", "07", "08", "09"
        base_month = "04"
    Case "10", "11", "12", "01", "02", "03"
        base_month = "10"
    End Select
End Function


'■概要
Function prc61_close_report_file() As Boolean
    On Error GoTo err_handler

prc_sub_no = prc_no & "-00"
    Workbooks(xls_name).Activate
    Worksheets("【公共品見込資料】").Select
    
prc_sub_no = prc_no & "-01"
    Range("A1").Value = format(cur_day, "mm/dd時点")
    Range("B2").Value = Right(cur_yyyymm, 2) & "月"
    
prc_sub_no = prc_no & "-02"
    Application.GoTo Range("A1"), True
    Workbooks(xls_name).Save

prc_sub_no = prc_no & "-99"
    Call write_log("[" & xls_name & "] 更新完了")
    prc61_close_report_file = True
Exit Function

err_handler:
    Call write_log("error")
    prc61_close_report_file = False
End Function

Function prc62_save_to_sharesite() As Boolean

prc_sub_no = prc_no & "-01"
    cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy", fy_yyyy)
    If Dir(cur_sync_site_path, vbDirectory) = "" Then
        MkDir cur_sync_site_path
        DoEvents
    End If
    
prc_sub_no = prc_no & "-02"
    cur_sync_site_path = cur_sync_site_path & cur_yyyymm & "\"
    If Dir(cur_sync_site_path, vbDirectory) = "" Then
        MkDir cur_sync_site_path
        DoEvents
    End If
    
prc_sub_no = prc_no & "-03"
    Workbooks(xls_name).SaveAs cur_sync_site_path & xls_name
    Workbooks(xls_name).Close False
    DoEvents

prc_sub_no = prc_no & "-99"
    Call write_log("ファイル共有サイト UPDATE完了")
    prc62_save_to_sharesite = True
Exit Function

err_handler:
    Call write_log("error")
    prc62_save_to_sharesite = False
End Function
