Attribute VB_Name = "Module1"
Option Explicit

'���萔��
'��Ɨp�c�a����
Public Const DB_NAME As String = "B0079_new.accdb"
'�f�X�N�g�b�v�[���i��ƃX�y�[�X�j�̃p�X
Public Const DESKTOP_PATH As String = "\\Lxjjaq6628\d\��for_RPA\B0079_EXT���������\��\"

'�t�@�C�����L�T�C�g�Ƃ̓����t�H���_
Public Const SYNC_SITE_PATH As String = _
    "C:\Users\11084631\LIXIL\�G���W�j�A�����O���i�� - 03_���������G���A��\yyyy�N�x\"
'    "C:\Users\11084631\LIXIL\�G�N�X�e���A�c�ƕ� ���L�T�C�g - 03_���������G���A��\yyyy�N�x\" 'yyyymm\"  '220209�F�ۑ���ύX
Public cur_sync_site_path As String     'yyyy������N�ɒu�����ăZ�b�g����

Public Const SHARESITE_PATH As String = _
    "https://lixilgroup.sharepoint.com/sites/JPFS1430/01_/00.�����֘A/02_����/03_���������G���A��/yyyy�N�x/yyyymm/"
'    "https://lixilgroup.sharepoint.com/sites/JPFS0481/02_all_lixil/02_08_�G���W�j�A�����O�c��G/0_�{��/02_�����֘A/02_����/03_���������G���A��/yyyy�N�x/yyyymm/" '220209�F�ۑ���ύX
Public cur_sharesite_path As String     'yyyy������N�ɒu�����ăZ�b�g����

'���ϐ���
'�����o�߉ӏ������ʂ��邽�߂̔ԍ����Z�b�g
Public prc_no As String, prc_sub_no As String
'���[���ő��t���郍�O����������
Public log As String
''�ғ��P�`�R���ڗp�A�����񐔂𕶎���ŃZ�b�g
'Public cmd_cnt As String


Public fy_yyyy As String            '����N
'Public pre_fy_year As String        '�O�N����N�A�S���X�V���ɂ̂ݎg�p
'Public pre2_month As String         '�O�X
'Public pre1_month As String         '�O��
Public cur_yyyymm As String          '����
'Public nxt1_month As String         '����
'Public nxt2_month As String         '���X��
'Public pre_day As Date            '�ғ����O��
Public cur_day As Date              '��Ɠ�
'Public nxt_day As String            '�ғ�������
Public working_day_cnt As Integer   '���������Ɠ��܂ł̉ғ�����
'public first_day as Boolean         '���n�߂Ȃ�True
'Public last_day As Boolean          '�����Ȃ�True


Public csv_name As String   'RPA�Ń_�E�����[�h���ꂽCSV�t�@�C����
Public xls_name As String   '��������Excel�t�@�C����


Sub cmd_B0079()
'��
prc_no = "00"
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    log = ""    '�f�o�b�O���g�p
    Call write_log("start")
    
'�����{�T�[�o�[�ڑ��A�ғ������ۂ�����
prc_no = "01"
'    cur_day = #12/3/2021# 'Date
    cur_day = Date

    Dim ref_souhon As ClassAdo: Set ref_souhon = New ClassAdo
    Call ref_souhon.connect("souhon")
    Call ref_souhon.workingday_check(format(cur_day, "yyyymmdd"))

    If ref_souhon.is_cn = False Or ref_souhon.is_rs = False Then
        Set ref_souhon = Nothing
        
        Call write_log("���{�T�[�o�[�ڑ��G���[")
        GoTo err_handler
        
    ElseIf ref_souhon.is_workingday = False Then
        Set ref_souhon = Nothing
        
        Call write_log("��ғ����A�����s�v")
        GoTo not_workingday
    Else
        working_day_cnt = ref_souhon.days_cnt
        cur_yyyymm = ref_souhon.ac_month
        Set ref_souhon = Nothing
    
        Call write_log("�ғ�" & format(working_day_cnt, "00") & "���ځA�������{")
    End If
    
'�������t���̐ݒ�
prc_no = "02"
    '�ғ��R���ڂ͊��茎���}�C�i�X�P����
    If working_day_cnt = 3 Then
        If Right(cur_yyyymm, 2) = "01" Then
            cur_yyyymm = CStr(CLng(cur_yyyymm) - 89)
        Else
            cur_yyyymm = CStr(CLng(cur_yyyymm) - 1)
        End If
    End If

    '�P�`�R���͊��茎�̑O�N������N�Ƃ���
    If CInt(Right(cur_yyyymm, 2)) < 4 Then
        fy_yyyy = CLng(Left(cur_yyyymm, 4)) - 1
    Else
        fy_yyyy = Left(cur_yyyymm, 4)
    End If
        
'�����C���������s
prc_no = "03"
    If main_prc(1) = False Then GoTo err_handler
    
    cur_sharesite_path = Replace(Replace(SHARESITE_PATH, "yyyymm", cur_yyyymm), "yyyy�N�x", fy_yyyy & "�N�x")
    log = "��" & cur_yyyymm & "����&emsp;" & _
            "<a href=""" & cur_sharesite_path & xls_name & """>" & xls_name & "</a>&emsp;�X�V����<br>" & log


'���ȍ~�͉ғ��R���ڈȍ~�̏����A���t����Ɠ������ɕύX����
prc_no = "04"
    If working_day_cnt > 3 Then GoTo finish
        
    '�ғ��R���ڂ܂ł̏����A�������̃t�@�C�����쐬����
    If Right(cur_yyyymm, 2) = "12" Then
        cur_yyyymm = CStr(CLng(cur_yyyymm) + 89)
    Else
        cur_yyyymm = CStr(CLng(cur_yyyymm) + 1)
    End If
    
    '�P�`�R���͊��茎�̑O�N������N�Ƃ���
    If CInt(Right(cur_yyyymm, 2)) < 4 Then
        fy_yyyy = CLng(Left(cur_yyyymm, 4)) - 1
    Else
        fy_yyyy = Left(cur_yyyymm, 4)
    End If
    
'�����C���������s�i�Q��ځj
prc_no = "05"
    If main_prc(2) = False Then GoTo err_handler
    
    cur_sharesite_path = Replace(Replace(SHARESITE_PATH, "yyyymm", cur_yyyymm), "yyyy�N�x", fy_yyyy & "�N�x")
    log = "��" & cur_yyyymm & "����&emsp;" & _
            "<a href=""" & cur_sharesite_path & xls_name & """>" & xls_name & "</a>&emsp;�X�V����<br>" & log
    

'���I��
prc_no = "99"
finish:
    Call write_log("finish")

    mail_sub = "NEW�y����I���zB0079_EXT���������\��"
    Call send_mail(mail_to, mail_cc, mail_sub, log)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

not_workingday:
    Call write_log("finish")

    mail_sub = "NEW�y��ғ����zB0079_EXT���������\��"
    Call send_mail(mail_to, mail_cc, mail_sub, log)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

err_handler:
    If Err.Number <> 0 Then Call write_log("error")
    Call write_log("stop")
    
    mail_sub = "NEW�yERROR!!�zB0079_EXT���������\��"
    Call send_mail(mail_to, mail_cc, mail_sub, log, True)

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Function main_prc(cnt As Integer) As Boolean
'���� cnt �F�����񐔁A�ғ��P�E�R���ڂ̂ݗL��
    prc_no = "11": If prc11_check_exest_csv = False Then GoTo err_handler
    prc_no = "12": If prc12_copy_preday_report = False Then GoTo err_handler
    
    prc_no = "21": If prc21_exec_accessDB = False Then GoTo err_handler
    Application.Wait Now() + TimeValue("00:00:30")
    If cnt = 2 Then prc_no = "22": If prc22_update_past_record = False Then GoTo err_handler

    '��
    prc_no = "31": If prc31_input_rawdata = False Then GoTo err_handler
    prc_no = "32": If prc32_input_toku_sys_data = False Then GoTo err_handler
    
    '�����{�f�[�^�̓���
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
        log = log & "��" & StrConv(proc, vbUpperCase) & " �� " & Now & "<br>"

    Case "finish", "stop"
        log = log & "��" & StrConv(proc, vbUpperCase) & " �� " & Now & "<br>"
        log = log & "------------------------------------------------------------<br>"
        
    Case "error"    '�z��O�G���[��
        log = log & "&emsp;�E" & Time & "�FID[" & prc_no & "]&ensp;" & proc & "<br>"
        log = log & "&emsp;&emsp;&emsp;��[" & Err.Number & "]&ensp;" & Err.Description & "<br>"

    Case Else
        log = log & "&emsp;�E" & Time & "�FID[" & prc_no & "]&ensp;" & proc & "<br>"

    End Select
End Sub

Sub set_datapage_configuration()
    Application.ScreenUpdating = True
    DoEvents
    
    Range("A2").Select
    '��U�I�[�g�t�B���^�E�E�C���h�E�g�̌Œ����������
    If ActiveSheet.AutoFilterMode = True Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False
    
    If ActiveSheet.AutoFilterMode = False Then Range("A2").AutoFilter
    If ActiveWindow.FreezePanes = False Then ActiveWindow.FreezePanes = True
    Application.ScreenUpdating = False
End Sub
