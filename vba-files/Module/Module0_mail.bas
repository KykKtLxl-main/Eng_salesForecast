Attribute VB_Name = "Module0_mail"
Option Explicit

'■［固定］メール（Outlook）操作用の変数
'Public mail_to As String
'Public mail_cc As String
Public mail_sub As String
Public mail_body As String
Public importance As Boolean    '重要度、Trueなら高
Public varAttach As Variant     '※添付ファイル（varAttach）のみVarinat型


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
    Dim obj_ns  As Object: Set obj_ns = obj_ol.GetNamespace("MAPI") '名前空間の指定
    Dim obj_fld     As Object: Set obj_fld = obj_ns.GetDefaultFolder(6)    '受信ボックス
    Dim obj_item    As Object: Set obj_item = obj_ol.CreateItem(0)             '0：olMailItem
    Dim i As Integer

    '■メールの作成
    With obj_item
        .To = m_to                       '宛先
        .CC = m_cc                       'CC
        .Subject = m_sub                 '件名
        .BodyFormat = 2                 'HTML形式
        .HTMLBody = "<style type=""text/css"">" & _
                    "p {" & _
                        "font-family: Meiryo UI;" & _
                        "font-size: 10px;" & _
                    "}" & _
                    "</style>" & _
                    "<p>" & m_body & "</p>"             '本文

        If m_imp Then .importance = 2   '重要度

        '添付ファイル
        If IsEmpty(m_attach) = False Then
            If IsArray(m_attach) = True Then
                For i = LBound(m_attach) To UBound(m_attach)
                    .Attachments.Add m_attach(i)
                Next i
            Else
                .Attachments.Add m_attach
            End If
        End If


        Application.Wait Now() + TimeValue("00:00:03")  'ミリ秒
'        .Display    'ディスプレイ表示
        .Send      'メール送信
        'Sleep 3000
    End With

finish:
    'objOutlook.Quit
    Set obj_item = Nothing: Set obj_ns = Nothing: Set obj_fld = Nothing: Set obj_ol = Nothing
End Sub







