Attribute VB_Name = "Module1"
Option Explicit

'■定数■
'作業用ＤＢ名称
Public Const DB_NAME As String = "B0079_new.accdb"
'デスクトップ端末（作業スペース）のパス
Public Const DESKTOP_PATH As String = "\\Lxjjaq6628\d\◆for_RPA\B0079_EXT公共見込予測\"

'ファイル共有サイトとの同期フォルダ
Public Const SYNC_SITE_PATH As String = _
    "C:\Users\11084631\LIXIL\エンジニアリング推進室 - 03_当月見込エリア別\yyyy年度\"
'    "C:\Users\11084631\LIXIL\エクステリア営業部 共有サイト - 03_当月見込エリア別\yyyy年度\" 'yyyymm\"  '220209：保存先変更
Public cur_sync_site_path As String     'yyyyを勘定年に置換してセットする

Public Const SHARESITE_PATH As String = _
    "https://lixilgroup.sharepoint.com/sites/JPFS1430/01_/00.数字関連/02_見込/03_当月見込エリア別/yyyy年度/yyyymm/"
'    "https://lixilgroup.sharepoint.com/sites/JPFS0481/02_all_lixil/02_08_エンジニアリング営業G/0_本部/02_数字関連/02_見込/03_当月見込エリア別/yyyy年度/yyyymm/" '220209：保存先変更
Public cur_sharesite_path As String     'yyyyを勘定年に置換してセットする

'■変数■
'処理経過箇所を識別するための番号をセット
Public prc_no As String, prc_sub_no As String
'メールで送付するログを書き込む
Public log As String
''稼働１～３日目用、処理回数を文字列でセット
'Public cmd_cnt As String


Public fy_yyyy As String            '勘定年
'Public pre_fy_year As String        '前年勘定年、４月更新時にのみ使用
'Public pre2_month As String         '前々
'Public pre1_month As String         '前月
Public cur_yyyymm As String          '当月
'Public nxt1_month As String         '翌月
'Public nxt2_month As String         '翌々月
'Public pre_day As Date            '稼働日前日
Public cur_day As Date              '作業日
'Public nxt_day As String            '稼働日翌日
Public working_day_cnt As Integer   '月初から作業日までの稼動日数
'public first_day as Boolean         '月始めならTrue
'Public last_day As Boolean          '月末ならTrue


Public csv_name As String   'RPAでダウンロードされたCSVファイル名
Public xls_name As String   '生成するExcelファイル名


Sub cmd_B0079()
'◆
prc_no = "00"
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    log = ""    'デバッグ時使用
    Call write_log("start")
    
'◆総本サーバー接続、稼働日か否か判定
prc_no = "01"
'    cur_day = #12/3/2021# 'Date
    cur_day = Date

    Dim ref_souhon As ClassAdo: Set ref_souhon = New ClassAdo
    Call ref_souhon.connect("souhon")
    Call ref_souhon.workingday_check(format(cur_day, "yyyymmdd"))

    If ref_souhon.is_cn = False Or ref_souhon.is_rs = False Then
        Set ref_souhon = Nothing
        
        Call write_log("総本サーバー接続エラー")
        GoTo err_handler
        
    ElseIf ref_souhon.is_workingday = False Then
        Set ref_souhon = Nothing
        
        Call write_log("非稼働日、処理不要")
        GoTo not_workingday
    Else
        working_day_cnt = ref_souhon.days_cnt
        cur_yyyymm = ref_souhon.ac_month
        Set ref_souhon = Nothing
    
        Call write_log("稼働" & format(working_day_cnt, "00") & "日目、処理実施")
    End If
    
'◆他日付情報の設定
prc_no = "02"
    '稼働３日目は勘定月をマイナス１する
    If working_day_cnt = 3 Then
        If Right(cur_yyyymm, 2) = "01" Then
            cur_yyyymm = CStr(CLng(cur_yyyymm) - 89)
        Else
            cur_yyyymm = CStr(CLng(cur_yyyymm) - 1)
        End If
    End If

    '１～３月は勘定月の前年を勘定年とする
    If CInt(Right(cur_yyyymm, 2)) < 4 Then
        fy_yyyy = CLng(Left(cur_yyyymm, 4)) - 1
    Else
        fy_yyyy = Left(cur_yyyymm, 4)
    End If
        
'◆メイン処理実行
prc_no = "03"
    If main_prc(1) = False Then GoTo err_handler
    
    cur_sharesite_path = Replace(Replace(SHARESITE_PATH, "yyyymm", cur_yyyymm), "yyyy年度", fy_yyyy & "年度")
    log = "●" & cur_yyyymm & "実績&emsp;" & _
            "<a href=""" & cur_sharesite_path & xls_name & """>" & xls_name & "</a>&emsp;更新完了<br>" & log


'◆以降は稼働３日目以降の処理、日付を作業日当月に変更する
prc_no = "04"
    If working_day_cnt > 3 Then GoTo finish
        
    '稼働３日目までの処理、当月分のファイルも作成する
    If Right(cur_yyyymm, 2) = "12" Then
        cur_yyyymm = CStr(CLng(cur_yyyymm) + 89)
    Else
        cur_yyyymm = CStr(CLng(cur_yyyymm) + 1)
    End If
    
    '１～３月は勘定月の前年を勘定年とする
    If CInt(Right(cur_yyyymm, 2)) < 4 Then
        fy_yyyy = CLng(Left(cur_yyyymm, 4)) - 1
    Else
        fy_yyyy = Left(cur_yyyymm, 4)
    End If
    
'◆メイン処理実行（２回目）
prc_no = "05"
    If main_prc(2) = False Then GoTo err_handler
    
    cur_sharesite_path = Replace(Replace(SHARESITE_PATH, "yyyymm", cur_yyyymm), "yyyy年度", fy_yyyy & "年度")
    log = "●" & cur_yyyymm & "実績&emsp;" & _
            "<a href=""" & cur_sharesite_path & xls_name & """>" & xls_name & "</a>&emsp;更新完了<br>" & log
    

'◆終了
prc_no = "99"
finish:
    Call write_log("finish")

    mail_sub = "NEW【正常終了】B0079_EXT公共見込予測"
    Call send_mail(mail_to, mail_cc, mail_sub, log)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

not_workingday:
    Call write_log("finish")

    mail_sub = "NEW【非稼働日】B0079_EXT公共見込予測"
    Call send_mail(mail_to, mail_cc, mail_sub, log)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

err_handler:
    If Err.Number <> 0 Then Call write_log("error")
    Call write_log("stop")
    
    mail_sub = "NEW【ERROR!!】B0079_EXT公共見込予測"
    Call send_mail(mail_to, mail_cc, mail_sub, log, True)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Function main_prc(cnt As Integer) As Boolean
'引数 cnt ：処理回数、稼働１・３日目のみ有効
    prc_no = "11": If prc11_check_exest_csv = False Then GoTo err_handler
    prc_no = "12": If prc12_copy_preday_report = False Then GoTo err_handler
    
    prc_no = "21": If prc21_exec_accessDB = False Then GoTo err_handler
    Application.Wait Now() + TimeValue("00:00:30")
    If cnt = 2 Then prc_no = "22": If prc22_update_past_record = False Then GoTo err_handler

    '◇
    prc_no = "31": If prc31_input_rawdata = False Then GoTo err_handler
    prc_no = "32": If prc32_input_toku_sys_data = False Then GoTo err_handler
    
    '◇総本データの入力
    prc_no = "51": If prc51_input_result_by_jgyo = False Then GoTo err_handler
    prc_no = "41": If prc41_input_TEMSSresult = False Then GoTo err_handler
    prc_no = "42": If prc42_input_juchuzan = False Then GoTo err_handler
    If cnt = 2 Then
        prc_no = "52": If prc52_input_pastresult_by_jgyo = False Then GoTo err_handler
        prc_no = "53": If prc53_input_month_area_sheet = False Then GoTo err_handler
    End If
    
    prc_no = "61": If prc61_close_report_file = False Then GoTo err_handler
    prc_no = "62": If prc62_save_to_sharesite = False Then GoTo err_handler
    
    main_prc = True
    Exit Function
    
err_handler:
    main_prc = False
End Function

Sub write_log(proc As String)
    Select Case proc
    Case "start" ',
        log = log & "------------------------------------------------------------<br>"
        log = log & "■" & StrConv(proc, vbUpperCase) & " ⇒ " & Now & "<br>"

    Case "finish", "stop"
        log = log & "■" & StrConv(proc, vbUpperCase) & " ⇒ " & Now & "<br>"
        log = log & "------------------------------------------------------------<br>"
        
    Case "error"    '想定外エラー時
        log = log & "&emsp;・" & Time & "：ID[" & prc_no & "]&ensp;" & proc & "<br>"
        log = log & "&emsp;&emsp;&emsp;＞[" & Err.Number & "]&ensp;" & Err.Description & "<br>"

    Case Else
        log = log & "&emsp;・" & Time & "：ID[" & prc_no & "]&ensp;" & proc & "<br>"

    End Select
End Sub

Sub set_datapage_configuration()
    Application.ScreenUpdating = True
    DoEvents
    
    Range("A2").Select
    '一旦オートフィルタ・ウインドウ枠の固定を解除する
    If ActiveSheet.AutoFilterMode = True Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False
    
    If ActiveSheet.AutoFilterMode = False Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = False Then ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = False
End Sub
