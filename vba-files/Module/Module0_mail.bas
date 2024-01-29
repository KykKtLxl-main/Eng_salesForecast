Attribute VB_Name = "Module0_mail"
Option Explicit

'���m�Œ�n���[���iOutlook�j����p�̕ϐ�
'Public mail_to As String
'Public mail_cc As String
Public mail_sub As String
Public mail_body As String
Public importance As Boolean    '�d�v�x�ATrue�Ȃ獂
Public varAttach As Variant     '���Y�t�t�@�C���ivarAttach�j�̂�Varinat�^


Public Const mail_to As String = "kyoko1.kato@lixil.com"
'Public Const mail_to As String = "satoshi.anzai@lixil.com; kyoko1.kato@lixil.com"
Public Const mail_cc As String = ""
'Public Const mail_cc As String = "takehiko.nawa@lixil.com"
'Public Const mail_address_admi As String = "kyoko1.kato@lixil.com"

Sub send_mail(ByVal m_to As String, _
              ByVal m_cc As String, _
              ByVal m_sub As String, _
              ByVal m_body As String, _
              Optional ByVal m_imp As Boolean = False, _
              Optional ByRef m_attach As Variant = Empty)

    On Error GoTo finish

    Dim obj_ol  As Object: Set obj_ol = CreateObject("Outlook.Application")
    Dim obj_ns  As Object: Set obj_ns = obj_ol.GetNamespace("MAPI") '���O��Ԃ̎w��
    Dim obj_fld     As Object: Set obj_fld = obj_ns.GetDefaultFolder(6)    '��M�{�b�N�X
    Dim obj_item    As Object: Set obj_item = obj_ol.CreateItem(0)             '0�FolMailItem
    Dim i As Integer

    '�����[���̍쐬
    With obj_item
        .To = m_to                       '����
        .CC = m_cc                       'CC
        .Subject = m_sub                 '����
        .BodyFormat = 2                 'HTML�`��
        .HTMLBody = "<style type=""text/css"">" & _
                    "p {" & _
                        "font-family: Meiryo UI;" & _
                        "font-size: 10px;" & _
                    "}" & _
                    "</style>" & _
                    "<p>" & m_body & "</p>"             '�{��

        If m_imp Then .importance = 2   '�d�v�x

        '�Y�t�t�@�C��
        If IsEmpty(m_attach) = False Then
            If IsArray(m_attach) = True Then
                For i = LBound(m_attach) To UBound(m_attach)
                    .Attachments.Add m_attach(i)
                Next i
            Else
                .Attachments.Add m_attach
            End If
        End If


        Application.Wait Now() + TimeValue("00:00:03")  '�~���b
'        .Display    '�f�B�X�v���C�\��
        .Send      '���[�����M
        'Sleep 3000
    End With

finish:
    'objOutlook.Quit
    Set obj_item = Nothing: Set obj_ns = Nothing: Set obj_fld = Nothing: Set obj_ol = Nothing
End Sub







