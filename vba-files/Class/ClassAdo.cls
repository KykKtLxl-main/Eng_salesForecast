VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cn As Object
Public is_cn As Boolean

Public rs As Object
Public rec_cnt As Long      '���R�[�h�Z�b�g�irs�j�̃��R�[�h�����i�[
Public is_rs As Boolean

Private Const cnStr_souhon As String = _
    "Driver={Microsoft ODBC for Oracle}; ConnectString=RAH234; UID=SOUHON; PWD=kanri;"
Private Const cnStr_access As String = _
    "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="
'Private cnStr_excel As String

Public is_workingday As Boolean
Public ac_month As String
Public days_cnt_all As Integer, days_cnt As Integer

'�������g�p�������̂ŃR�����g�u���b�N
'Private Sub Class_Initialize()
'    Call connect
'End Sub

Public Sub connect(ByVal tgt As String, Optional path As String)
'Private Sub connect()
    On Error Resume Next
    Set cn = CreateObject("ADODB.Connection")
        
    Select Case tgt
    Case "souhon"
        cn.Open cnStr_souhon
    Case "access"
        cn.Open cnStr_access & ThisWorkbook.path & "\" & DB_NAME
    Case "excel"
        With cn
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Properties("Extended Properties") = "Excel 12.0"
            .Open path
        End With
        
    End Select
    On Error GoTo 0

    If cn.State = 1 Then
        is_cn = True
    Else
        is_cn = False
        Call write_log("error")
    End If
End Sub

Public Sub open_rs(ByVal mySql_ As String)
    If is_cn = False Then Exit Sub

    On Error Resume Next
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open mySql_, cn, 3  '3:adOpenStatic
    On Error GoTo 0

    Select Case rs.State
    Case 0  'adStateClosed
        is_rs = False
    Case 1  'adStateOpen
        is_rs = True
        rec_cnt = rs.RecordCount
    End Select
End Sub


Sub workingday_check(str_date As String)
    If is_cn = False Then Exit Sub
    
    Dim mySql As String: mySql = "SELECT * FROM �c�Ɨp�ғ������"
    Call open_rs(mySql)

    If is_rs Then

        '����N�����擾����
        rs.movefirst
        ac_month = Left(rs.fields("�ғ���"), 6)

        '��Ɠ����_�̌����o�ߓ������擾����
        rs.Filter = "�ғ��� = '" & str_date & "'"
        If rs.RecordCount = 0 Then
            is_workingday = False
        Else
            is_workingday = True
            days_cnt_all = rs.fields("�S�ғ�����")

            If rs.fields("�S�ғ�����") < (rs.fields("�ғ�����") + 2) Then     '�ғ��P��Q����
                days_cnt = (rs.fields("�ғ�����") + 2) - rs.fields("�S�ғ�����")
            Else    '�ғ��P���ځA�Q���ڈȊO
                days_cnt = rs.fields("�ғ�����") + 2
            End If
        End If
    End If

    rs.Close
End Sub

Private Sub Class_Terminate()
    Call dis_connect
End Sub

Private Sub dis_connect()
    On Error Resume Next
    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    If cn.State = 1 Then cn.Close
    Set cn = Nothing
    On Error GoTo 0
End Sub


'    Dim ref_access As ClassAdo: Set ref_access = New ClassAdo
'    Call ref_access.connect("access")
'    str_sql = "SELECT * FROM [Tbl] "
'    Call ref_access.open_rs(str_sql)
'    If ref_access.is_cn = False Or ref_access.is_rs = False Then GoTo err_handler
