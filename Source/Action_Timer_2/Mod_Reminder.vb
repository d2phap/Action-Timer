

Imports System.IO
Imports System.Net.Mail

Module Mod_Reminder

    Public Function ReadText(ByVal FileName As String) As String
        Dim r As StreamReader
        r = New StreamReader(Trim(FileName))

        ReadText = r.ReadToEnd()

        r.Close()
    End Function

    Public Sub WriteText(ByVal Text As String, ByVal FileName As String)
        Dim w As StreamWriter = New StreamWriter(Trim(FileName))

        w.Write(Text)

        w.Close()

    End Sub

    Public Sub CreateNewFolder(ByVal path As String, Optional ByVal FolderName As String = "New Folder")
        Dim fso, f, fc, nf
        fso = CreateObject("Scripting.FileSystemObject")
        f = fso.GetFolder(path)
        fc = f.SubFolders

        If LCase$(Trim$(FolderName)) = "new folder" Then
            FolderName = "New Folder"
        End If

        nf = fc.Add(Trim$(FolderName))

        fso = Nothing
    End Sub

    Public Function FolderExists(ByVal FolderName As String) As Boolean
        Dim fso
        fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FolderExists(FolderName)) Then
            FolderExists = True
        Else
            FolderExists = False
        End If

        fso = Nothing
    End Function


    'Biên soạn nội dung email theo mẫu :
    '[EMail]
    'To=...
    'From=...
    'Account=...
    'Password=...
    'Smtp=...

    '[Subject]...[Subject]
    '[Body]...[Body]


    Public Function Send_Email(ByVal sFilename As String) As Integer
        Dim INI As New ClassIniFile(sFilename)

        Dim sTo As String = INI.GetString("EMail", "To", "")
        Dim sFrom As String = INI.GetString("EMail", "From", "")
        Dim sAccount As String = INI.GetString("EMail", "Account", "")
        Dim sPassword As String = INI.GetString("EMail", "Password", "")

        Dim sSmtp As String = INI.GetString("EMail", "Smtp", "")

        Dim str() As String = ReadText(sFilename).Replace("[Subject]", "⅝").Split("⅝")
        Dim sSubject As String = str(1)

        str = ReadText(sFilename).Replace("[Body]", "⅝").Split("⅝")
        Dim sBody As String = str(1)

        Erase str   'xoá mảng

        'Start by creating a mail message object
        Dim MyMailMessage As New MailMessage()

        'From requires an instance of the MailAddress type
        MyMailMessage.From = New MailAddress(sFrom)

        'To is a collection of MailAddress types
        MyMailMessage.To.Add(sTo)

        MyMailMessage.Subject = sSubject
        MyMailMessage.Body = sBody

        'Create the SMTPClient object and specify the SMTP GMail server
        Dim SMTPServer As New SmtpClient(sSmtp)
        SMTPServer.Port = 587
        SMTPServer.Credentials = New System.Net.NetworkCredential(sAccount, sPassword)
        SMTPServer.EnableSsl = True

        Send_Email = 1

        Try
            SMTPServer.Send(MyMailMessage)
        Catch ex As SmtpException
            Send_Email = 0
        End Try

    End Function




End Module
