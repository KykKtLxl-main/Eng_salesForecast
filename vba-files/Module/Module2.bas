Attribute VB_Name = "Module2"
Option Explicit

Sub GoodBye()
    MsgBox "GoodBye!"
End Sub

'��RPA��CSV�t�@�C�����o�͂��ꂽ���`�F�b�N
Function prc11_check_exest_csv()
prc_sub_no = prc_no & "-01" 'RPA�ŏo�͂��ꂽCSV���f�X�N�g�b�v�[����ɂ��邩
    csv_name = "sysData_" & Right(Replace(cur_day, "/", ""), 6) & ".csv"

    If Dir(DESKTOP_PATH & "download\" & csv_name) = "" Then
        Call write_log(csv_name & "�t�@�C�����f�X�N�g�b�v�[�����ɕۑ�����Ă��܂���")
        prc11_check_exest_csv = False
    Else
        'MsAccess�̃����N�e�[�u���p�ɖ��̂�����t���폜�����t�@�C�������
        If Dir(DESKTOP_PATH & "sysData.csv") <> "" Then Kill DESKTOP_PATH & "sysData.csv"
        DoEvents
        FileCopy DESKTOP_PATH & "download\" & csv_name, DESKTOP_PATH & "sysData.csv"

        'Call write_log("") '�v��Ȃ�
        prc11_check_exest_csv = True
    End If
End Function



'���O���������̃t�@�C�����R�s�[�A��Ɠ����̖��̂ŕۑ�����
'��ver1.4�` �c�ƕ��T�[�o�[�̃t�@�C���ł͂Ȃ����[���̃t�@�C�����R�s�[���Ďg���悤�ύX
Function prc12_copy_preday_report() As Boolean
    On Error GoTo err_handler

'�O�������̃t�@�C�������擾
'�O�����������t�@�C��������t�H���_��Dir�֐��ɐݒ肷��
prc_sub_no = prc_no & "-01"
    Dim xls As Variant

    '>>>�������� �ғ��P���ڂ̏������Â�����orz
    Select Case working_day_cnt
    Case 1  '���؍ρi�S���X�V���������j
        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then '�O������
            xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ").xlsx")
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & cur_yyyymm

        Else    '�������т́A�O���̃t�H���_���Q�Ƃ���

            Select Case Right(cur_yyyymm, 2)
            Case "01"
                 xls = Dir(DESKTOP_PATH & "create\" & CLng(Left(cur_yyyymm, 4)) - 1 & "12\*���������G���A��(" & fy_yyyy & ").xlsx")
                 cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & cur_yyyymm - 89

            Case "04"
                cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy - 1 & "�N�x") & CLng(cur_yyyymm) - 1

            Case Else
                cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & CLng(cur_yyyymm) - 1

            End Select
        End If

    Case 2, 3
        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then '�O������
            xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "��).xlsx")
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & cur_yyyymm
        Else    '�������т́A�O���̃t�H���_���Q�Ƃ���
            'xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 1 & "\*���������G���A��(" & fy_yyyy & ").xlsx")
            xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) & "\*���������G���A��(" & fy_yyyy & ").xlsx")
            'cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & CLng(cur_yyyymm) - 1
            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & CLng(cur_yyyymm)
        End If

    Case Else
        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ").xlsx")
        cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & cur_yyyymm

    End Select

'    If (working_day_cnt = 2 Or working_day_cnt = 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "��).xlsx")
'
'    ElseIf working_day_cnt = 1 Then
'        If (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'            Select Case Month(cur_day)
'            Case 4       '�S���̉ғ��P����
'                xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy - 1 & ").xlsx")
'            Case Else
'                xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ").xlsx")
'            End Select
'        ElseIf (format(cur_day, "yyyymm") = cur_yyyymm) Then
'            Select Case Month(cur_day)
'            Case 1
'                xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 89 & "\*���������G���A��(" & fy_yyyy & ").xlsx")
'            Case Else
'                xls = Dir(DESKTOP_PATH & "create\" & CLng(cur_yyyymm) - 1 & "\*���������G���A��(" & fy_yyyy & ").xlsx")
'            End Select
'        End If
'    Else
'        xls = Dir(DESKTOP_PATH & "create\" & cur_yyyymm & "\*���������G���A��(" & fy_yyyy & ").xlsx")
'    End If
    '<<<�����܂�

    '�t�H���_�̒�����ŏI���t�̃t�@�C�������擾�i�����D�h���Ō����j
    Dim last_date As Integer: last_date = 0
    Do Until xls = ""
'        Debug.Print xls

        If CInt(Left(xls, 4)) = format(cur_day, "mmdd") Then '�f�o�b�O�p�A�����t�@�C�������ɂ���ꍇ�͍폜
            Kill DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls

        ElseIf last_date < CInt(Left(xls, 4)) Then
            last_date = CInt(Left(xls, 4))
            xls_name = xls

        End If
        xls = Dir()
    Loop

'�O�����������t�@�C�������L�T�C�g����擾����
prc_sub_no = prc_no & "-02"
'    Dim file_source As String: file_source = DESKTOP_PATH & "create\" & xls_name

    '>>>�������� �ғ��P���ڂ̏������Â�����orz
    cur_sync_site_path = cur_sync_site_path & "\" & xls_name

'    If working_day_cnt = 1 And (format(cur_day, "yyyymm") = cur_yyyymm) Then
'        If Month(cur_day) = 1 Then
'            cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & CLng(cur_yyyymm) & "\" & xls_name
'        Else
'           cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & CLng(cur_yyyymm) - 1 & "\" & xls_name
'        End If
'    Else
'        cur_sync_site_path = Replace(SYNC_SITE_PATH, "yyyy�N�x", fy_yyyy & "�N�x") & cur_yyyymm & "\" & xls_name
'    End If
    '<<<�����܂�

    Dim file_destination As String
    If (working_day_cnt <= 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
        xls_name = format(cur_day, "mmdd") & "���������G���A��(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "��).xlsx"
    Else
        xls_name = format(cur_day, "mmdd") & "���������G���A��(" & fy_yyyy & ").xlsx"
    End If
    file_destination = DESKTOP_PATH & "create\" & cur_yyyymm & "\" & xls_name

    '>>>��������
    If Dir(DESKTOP_PATH & "create\" & cur_yyyymm, vbDirectory) = "" Then
        MkDir DESKTOP_PATH & "create\" & cur_yyyymm
        DoEvents
    End If
    FileCopy cur_sync_site_path, file_destination
    'FileCopy file_source, file_destination
    '<<<�����܂�

    Call write_log("�O�������t�@�C���̃R�s�[����")
    prc12_copy_preday_report = True
Exit Function

err_handler:
    Call write_log("error")
'    Resume
    prc12_copy_preday_report = False
End Function

'�y�o�b�N�A�b�v�z
''''���O���������̃t�@�C�����R�s�[�A��Ɠ����̖��̂ŕۑ�����
''''��ver1.4�` �c�ƕ��T�[�o�[�̃t�@�C���ł͂Ȃ����[���̃t�@�C�����R�s�[���Ďg���悤�ύX
'''Function prc12_copy_preday_report() As Boolean
'''    On Error GoTo err_handler
'''
'''prc_sub_no = prc_no & "-01"
'''    Dim xls As Variant
'''    If (working_day_cnt = 2 Or working_day_cnt = 3) And (format(cur_day, "yyyymm") <> cur_yyyymm) Then
'''        xls = Dir(DESKTOP_PATH & "create\*���������G���A��(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "��).xlsx")
'''    ElseIf working_day_cnt = 1 And Month(cur_day) = 4 Then      '�S���̉ғ��P����
'''        xls = Dir(DESKTOP_PATH & "create\*���������G���A��(" & fy_yyyy - 1 & ").xlsx")
'''    Else
'''        xls = Dir(DESKTOP_PATH & "create\*���������G���A��(" & fy_yyyy & ").xlsx")
'''    End If
'''
'''    Dim last_date As Integer: last_date = 0
'''    Do Until xls = ""
'''        If CInt(Left(xls, 4)) = format(cur_day, "mmdd") Then '�f�o�b�O�p�A�����t�@�C�������ɂ���ꍇ�͍폜
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
'''        xls_name = format(cur_day, "mmdd") & "���������G���A��(" & fy_yyyy & ")(" & Right(cur_yyyymm, 2) & "��).xlsx"
'''    Else
'''        xls_name = format(cur_day, "mmdd") & "���������G���A��(" & fy_yyyy & ").xlsx"
'''    End If
'''    file_destination = DESKTOP_PATH & "create\" & xls_name
'''
'''    FileCopy file_source, file_destination
'''
'''    Call write_log("�O�������t�@�C���̃R�s�[����")
'''    prc12_copy_preday_report = True
'''Exit Function
'''
'''err_handler:
'''    Call write_log("error")
'''    prc12_copy_preday_report = False
'''End Function

'��
Function prc21_exec_accessDB() As Boolean
'Function step111_Exec_Acc() As Boolean
'--- �ύX����   ---
'   ver.1.1     2018/10/31  katouk48    �ŏI����m�F����
'   ver.1.4     2021/12/14  katok21�F�����e
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
        .OpenCurrentDatabase (ThisWorkbook.path & "\" & DB_NAME)    'AccessDB���J��
        .DoCmd.SetWarnings False    '�A���[�g��\�������Ȃ�
        .Visible = True

        prc_sub_no = prc_no & "-03"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl10_�[����sys"
        .DoCmd.DeleteObject acTable, "tbl11_�[����sys_copy"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        .DoCmd.OpenQuery "qry10_�[����sys�f�[�^�e�[�u���쐬", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry11_�R�s�[�e�[�u���쐬", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry12_�̔����z�X�V", acViewNormal, acEdit

        prc_sub_no = prc_no & "-04"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl20_���{_���㍂"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        'cur_yyyymm = "202112"   '���f�o�b�O�p
        str_sql = "SELECT a.�i��R�[�h, h.���{�ꖼ�P AS �i�햼, m1.�n��ID, m1.�G���A, " & _
               "a.�{���R�[�h, s1.���{�ꖼ�P AS �{����, a.�x�ЃR�[�h, s2.���{�ꖼ�P AS �x�Ж�, a.�x�X�R�[�h, s3.���{�ꖼ�P AS �x�X��," & _
               "a.�c�Ə��R�[�h, s4.���{�ꖼ�P AS �c�Ə���, a.���[�g�R�[�h, s5.���{�ꖼ�P AS ���[�g��, " & _
               "a.���Ə��R�[�h, s6.���Ə���, " & _
               "a.����N��, Sum(a.���㍂�Q�O�N) AS ���㍂�Q�O�N, Sum(a.���㍂�Q����) AS ���㍂�Q���� " & _
        "INTO tbl20_���{_���㍂ " & _
        "FROM ((((((( SOUHON_�S�\�����Ə���FACT�Q���� a " & _
        "INNER JOIN SOUHON_�i�햼�� h ON a.�i��R�[�h = h.�R�[�h ) " & _
        "INNER JOIN SOUHON_�g�D���� s1 ON a.�{���R�[�h = s1.�R�[�h ) " & _
        "INNER JOIN SOUHON_�g�D���� s2 ON a.�x�ЃR�[�h = s2.�R�[�h ) " & _
        "INNER JOIN SOUHON_�g�D���� s3 ON a.�x�X�R�[�h = s3.�R�[�h ) " & _
        "INNER JOIN SOUHON_�g�D���� s4 ON a.�c�Ə��R�[�h = s4.�R�[�h ) " & _
        "INNER JOIN SOUHON_���Ə��}�X�^�Q���� s6 ON a.���Ə��R�[�h = s6.���Ə��R�[�h ) " & _
        "LEFT JOIN SOUHON_�g�D���� s5 ON a.[���[�g�R�[�h] = s5.�R�[�h ) " & _
        "LEFT JOIN mst_�x��_�V���Δ� m1 ON a.�x�ЃR�[�h = m1.�����x�X�R�[�h " & _
        "GROUP BY a.�i��R�[�h, h.���{�ꖼ�P, m1.�n��ID, m1.�G���A, " & _
                 "a.�{���R�[�h, s1.���{�ꖼ�P, a.�x�ЃR�[�h, s2.���{�ꖼ�P, a.�x�X�R�[�h, s3.���{�ꖼ�P, " & _
                 "a.�c�Ə��R�[�h, s4.���{�ꖼ�P, a.���Ə��R�[�h, s6.���Ə���, a.���[�g�R�[�h, s5.���{�ꖼ�P, " & _
                 "a.����N�� " & _
        "HAVING (a.�i��R�[�h='T41334')" ' AND a.����N��='" & cur_yyyymm & "');"
        Set qd = .CurrentDb.QueryDefs("qry21_���{���㍂_�擾")
        qd.Sql = str_sql
        .DoCmd.OpenQuery "qry21_���{���㍂_�擾", acViewNormal, acEdit


        prc_sub_no = prc_no & "-05"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl30_���{_�󒍎c"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        str_sql = "SELECT j.�{���R�[�h, j.�����x�X�R�[�h, j.�����x�X��, j.�c�Ə��R�[�h, j.�c�Ə���, a.���Ə��R�[�h, a.���Ə���, " & _
                       "a.�i��R�[�h, a.�i�햼, a.�N��, Sum(a.����v��) AS ����v��, Sum(a.���㍂) AS ���㍂, Sum(a.�󒍎c) AS �󒍎c " & _
                "INTO tbl30_���{_�󒍎c " & _
                "FROM SOUHON_�󒍕��͎󒍎c�Q���Ə� a " & _
                "INNER JOIN SOUHON_���Ə��}�X�^�Q���� j ON a.���Ə��R�[�h = j.���Ə��R�[�h " & _
                "WHERE a.V���x���R�[�h = 'V00100' " & _
                "GROUP BY j.�{���R�[�h, j.�����x�X�R�[�h, j.�����x�X��, j.�c�Ə��R�[�h, j.�c�Ə���, a.���Ə��R�[�h, a.���Ə���, a.�i��R�[�h, a.�i�햼, a.�N�� " & _
                "HAVING a.�i��R�[�h = 'T41334' AND a.�N�� = '" & cur_yyyymm & "';"
        Set qd = .CurrentDb.QueryDefs("qry31_���{�󒍎c_�擾")
        qd.Sql = str_sql
        .DoCmd.OpenQuery "qry31_���{�󒍎c_�擾", acViewNormal, acEdit

        .DoCmd.SetWarnings True

        prc_sub_no = prc_no & "-06"
        .Quit
    End With


prc_sub_no = prc_no & "-99"
    Set obj_access = Nothing

    Call write_log("AccessDB�X�V����")
    prc21_exec_accessDB = True
Exit Function

err_handler:
    Call write_log("error")
    prc21_exec_accessDB = False
End Function


'���ғ��P���ځE�R���ڗp
Function prc22_update_past_record() As Boolean
'--- �ύX����   ---
'   ver.1.4     2021/12/20  katok21�F�V�K�ݒ�
'------------------

    On Error GoTo err_handler

prc_sub_no = prc_no & "-01"
    Dim obj_access As Object: Set obj_access = CreateObject("Access.Application")

    Const acTable = 0
    Const acViewNormal = 0
    Const acEdit = 1

prc_sub_no = prc_no & "-02"
    With obj_access
        .OpenCurrentDatabase (ThisWorkbook.path & "\" & DB_NAME)    'AccessDB���J��
        .DoCmd.SetWarnings False    '�A���[�g��\�������Ȃ�
        .Visible = True

        prc_sub_no = prc_no & "-03"
        On Error Resume Next
        .DoCmd.DeleteObject acTable, "tbl40_(�����p)�ߋ�����"
        .DoCmd.DeleteObject acTable, "tbl41_tbl40+�x�ЃG���A�t�^"
        Application.Wait Now + TimeValue("00:00:03")
        On Error GoTo 0: On Error GoTo err_handler

        prc_sub_no = prc_no & "-04"
        .DoCmd.OpenQuery "qry41_�g�։ߋ�����_����", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry42_�g�։ߋ�����_�O�P", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry43_�g�։ߋ�����_�O�Q", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry44_�g�։ߋ�����_�O�R", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry45_�g�։ߋ�����_�O�S", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry46_�g�։ߋ�����_�O�T", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry47_�g�։ߋ�����_�O�U", acViewNormal, acEdit
        .DoCmd.OpenQuery "qry48_tbl40+�x�ЃG���A�t�^", acViewNormal, acEdit

        prc_sub_no = prc_no & "-05"
        .DoCmd.SetWarnings True
        .Quit
    End With


prc_sub_no = prc_no & "-99"
    Set obj_access = Nothing

    Call write_log("�ғ�" & format(working_day_cnt, "00") & "���ځF�ߋ����� �X�V����")
    prc22_update_past_record = True
Exit Function

err_handler:
    Call write_log("error")
    prc22_update_past_record = False
End Function



'���[����Sys�u��f�[�^�v�V�[�g�\��t���A�u�̔����z�v�̒l�ύX
Function prc31_input_rawdata() As Boolean
'Function step121_Paste_RawData() As Boolean
'Function step122_Update_Price() As Boolean
'--- �ύX����   ---
'   ver.1.0     2018/10/25  katouk48    ���ւ���ŏI�ł����킹��A�ǉ��v�]
'       �u��f�[�^�v�V�[�g�ւ̓]�L�͕s�v���Ǝv�������ǂ��A�m�F�p�ɕK�v������
'   ver.1.1     2018/10/31  katouk48    �ŏI����m�F����
'   ver.1.4     2021/12/14  katok21�F�����e�A�Q�̃v���V�[�W������
'------------------

    On Error GoTo err_handler

'    xls_name = "1213���������G���A��(2021).xlsx"    '���f�o�b�O�p

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
    DoEvents    '�Ȃ�ƂȂ�

prc_sub_no = prc_no & "-02"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")

    Dim str_sql As String
    str_sql = "SELECT * FROM [tbl10_�[����sys] WHERE �L�����Z���� IS NULL "

    Call ref_access.open_rs(str_sql)
    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo not_connect_db

prc_sub_no = prc_no & "-03"
    Dim r_num As Long

    Workbooks(xls_name).Worksheets("��f�[�^").Activate
    r_num = Range("A2").End(xlDown).Row
    Rows("2:" & r_num).ClearContents

    Range("A2").CopyFromRecordset ref_access.rs
    Set ref_access = Nothing

prc_sub_no = prc_no & "-04"
    Call set_datapage_configuration

prc_sub_no = prc_no & "-05"
    Const c_num_E As Integer = 5     '�Ǘ�
    Const c_num_H As Integer = 8     '�������e�敪
    Const c_num_AD As Integer = 30   '�̔����z

    r_num = 2
    Do Until Cells(r_num, 1).Value = ""
        If Left(Cells(r_num, c_num_E).Value, 1) = "F" And _
            (Cells(r_num, c_num_H).Value = "2" Or Cells(r_num, c_num_H).Value = "3") Then

            Cells(r_num, c_num_AD).Formula = "=CK" & r_num & "*1.43"
        End If
        r_num = r_num + 1
    Loop

prc_sub_no = prc_no & "-99"
    Call write_log("[��f�[�^]�V�[�g �f�[�^�\�t��OK")
    prc31_input_rawdata = True
Exit Function

not_connect_db:
    Set ref_access = Nothing

    Call write_log("��Ɨp�c�a�ڑ��G���[")
    GoTo err_handler

err_handler:
    Call write_log("error")
    prc31_input_rawdata = False
End Function



'���U�V�[�g�i�����SFAX�����`�O�������j�փf�[�^�\��t��
Function prc32_input_toku_sys_data() As Boolean
'Function step131_Input_TokuSysData() As Boolean
'--- �ύX����   ---
'   ver.1.1.0   2018/10/30  katouk48�F�m�F��
'   ver.1.1.1   2018/11/01  katouk48�F
'   ver.1.4     2021/12/14  katok21�F�����e
'------------------

    On Error GoTo err_handler

'    cur_day = Date
'    cur_yyyymm = "202112"    '���f�o�b�O�p
'    xls_name = "1213���������G���A��(2021).xlsx"    '���f�o�b�O�p

    Workbooks(xls_name).Activate

'���o�����Ɏg�p������t��ϐ��ɃZ�b�g
prc_sub_no = prc_no & "-01"
    Dim r_num As Long, c_num As Integer
    Dim fld As Object
    Dim str_sql As String

    '�P�����O�̍ŏI�����Z�b�g
    Dim pre_month_lastday As Date
    pre_month_lastday = DateSerial(Left(cur_yyyymm, 4), Right(cur_yyyymm, 2), 1) - 1
    '�����̍ŏI�����Z�b�g
    Dim cur_month_lastday As Date
    cur_month_lastday = DateSerial(Left(cur_yyyymm, 4), Right(cur_yyyymm, 2), 1)
    cur_month_lastday = DateAdd("m", 1, cur_month_lastday) - 1
    '�����A���X���̍ŏI�����Z�b�g
    Dim nxt1_month_lastday As Date, nxt2_month_lastday As Date
    nxt1_month_lastday = DateAdd("m", 1, cur_month_lastday) - 1
    nxt2_month_lastday = DateAdd("m", 1, nxt1_month_lastday) - 1

'���{�T�[�o�[�ڑ�
prc_sub_no = prc_no & "-02"
    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
    Call ref_access.connect("access")


'�@�u�����SFAX�����v�V�[�g����
prc_sub_no = prc_no & "-10"
    Worksheets("�����SFAX����").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-11"
    str_sql = "SELECT [*������], ������, �������e�敪, ���x�ЃR�[�h, [*�̔����z], �m��o�ד�, EOC�`�[������, �����x�X�R�[�h, �����x�X�� " & _
            "FROM tbl11_�[����sys_copy " & _
            "WHERE (�m��o�ד� >= #" & pre_month_lastday & "# AND �m��o�ד� < #" & cur_month_lastday & "#) "
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


'�A�u�������c�v�V�[�g����
prc_sub_no = prc_no & "-20"
    Worksheets("�������c").Select
    Cells.ClearContents

    If format(cur_day, "yyyymm") = cur_yyyymm And working_day_cnt < 3 Then
        prc_sub_no = prc_no & "-21"
        Worksheets("�����SFAX����").Cells.Copy Destination:=Worksheets("�������c").Range("A1")

    Else
        prc_sub_no = prc_no & "-22"
        Call ref_access.open_rs(str_sql)    '�\��t����f�[�^��"prc_no.32-11"�̂܂ܕύX�Ȃ�
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
            If Cells(r_num, 6).Value <= (cur_day - 1) Then  '�u�m��o�ד��v����Ɠ��O�����O�Ȃ�
                If Cells(r_num, 7).Value <> cur_day Then    '�uEOC�`�[�������v����Ɠ������Ŗ������
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


'�B�u�����[���m��v�V�[�g����
prc_sub_no = prc_no & "-30"
    Worksheets("�����[���m��").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-31"
    str_sql = "SELECT [*������], ������, �������e�敪, ���x�ЃR�[�h, [*�̔����z], �m��o�ד�, EOC�`�[������, �����x�X�R�[�h, �����x�X�� " & _
            "FROM tbl11_�[����sys_copy " & _
            "WHERE (�m��o�ד� >= #" & cur_month_lastday & "# AND �m��o�ד� < #" & nxt1_month_lastday & "#) "
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


'�C�u���X���v�V�[�g����
prc_sub_no = prc_no & "-40"
    Worksheets("���X��").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-41"
    str_sql = "SELECT [*������], ������, �������e�敪, ���x�ЃR�[�h, [*�̔����z], �m��o�ד�, EOC�`�[������, �����x�X�R�[�h, �����x�X�� " & _
            "FROM tbl11_�[����sys_copy " & _
            "WHERE (�m��o�ד� >= #" & nxt1_month_lastday & "# AND �m��o�ד� < #" & nxt2_month_lastday & "#) "
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


'�D�u�[�����m��v�V�[�g����
prc_sub_no = prc_no & "-50"
    Worksheets("�[�����m��").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-51"
    str_sql = "SELECT [*������], ������, �������e�敪, ���x�ЃR�[�h, [*�̔����z], �m��o�ד�, EOC�`�[������, �����x�X�R�[�h, �����x�X�� " & _
            "FROM tbl11_�[����sys_copy " & _
            "WHERE �m��o�ד� IS NULL "
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


'�E�u�O�������v�V�[�g����
prc_sub_no = prc_no & "-60"
    Worksheets("�O������").Select
    Cells.ClearContents

prc_sub_no = prc_no & "-61"
    str_sql = "SELECT [*������], ������, �������e�敪, ���x�ЃR�[�h, [*�̔����z], �m��o�ד�, EOC�`�[������, �����x�X�R�[�h, �����x�X�� " & _
            "FROM tbl11_�[����sys_copy " & _
            "WHERE [*������] = #" & (cur_day - 1) & "# "
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


'��Еt��
prc_sub_no = prc_no & "-99"
    Set ref_access = Nothing

    Call write_log("�[����sys �f�[�^�\�t������")
    prc32_input_toku_sys_data = True
Exit Function

not_connect_db:
    Set ref_access = Nothing

    Call write_log("��Ɨp�c�a�ڑ��G���[")
    GoTo err_handler

err_handler:
    Set ref_access = Nothing

    Call write_log("error")
    prc32_input_toku_sys_data = False
End Function

