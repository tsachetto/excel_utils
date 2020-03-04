Sub envio()

'Criar uma planilha cujas colunas sejam:
'A inteira em Branco.
'B2 - Destinatário, B3 em diante adicionar e-mails dos destinatários.
'C2 - Assunto, C3 em diante adicionar os respectivos assuntos.
'D2 - Corpo do E-mail. D3 em diante escrever o conteúdo do email e assinatura.
'E2 - Anexo, E3 em diante o diretório\arquivo a ser anexado.
'F2 - Protocolo, que será preenchido automaticamente.
'Configurar o smtp com email e senha de evio etc..

Dim CDO_Mail As Object
Dim CDO_Config As Object
Dim SMTP_Config As Variant
Dim strSubject As String
Dim strFrom As String
Dim strTo As String
Dim strCc As String
Dim strBcc As String
Dim strBody As String
Dim anexo As String

For i = 3 To 5
    
    If Range("F" & i) > 0 Then
    GoTo prox
    Else: End If
    
    strSubject = Range("C" & i)
    strFrom = "xxx@xxx.xxx"
    strTo = Range("B" & i)
    strCc = ""
    strBcc = ""
    strBody = Range("D" & i)
    anexo = Range("E" & i)
    
    Set CDO_Mail = CreateObject("CDO.Message")
    On Error GoTo Error_Handling
    
    Set CDO_Config = CreateObject("CDO.Configuration")
    CDO_Config.Load -1
    
    Set SMTP_Config = CDO_Config.Fields
    
    With SMTP_Config
     .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
     .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "xxx@xxx.xxx"
     .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "xxxxxxx"
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
     .Update
    End With
    
    With CDO_Mail
     Set .Configuration = CDO_Config
    End With
    
    CDO_Mail.Subject = strSubject
    CDO_Mail.From = strFrom
    CDO_Mail.To = strTo
    CDO_Mail.TextBody = strBody
    CDO_Mail.CC = strCc
    CDO_Mail.BCC = strBcc
    CDO_Mail.AddAttachment anexo
    CDO_Mail.Send
    
Error_Handling:
    If Err.Description <> "" Then MsgBox Err.Description
    
    Range("F" & i) = Now()

prox:
Next i

End Sub
