Imports System
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mail.MailMessage
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq

Imports CustomerSurvey.XMLer

Public Class Logger
    Public Property _logPath As String = "C:\emerge\survey\logs\"
    Public Property _logDir As String = _logPath & Date.Now.Year
    Public Property _logFile As String = _logDir & "\" & Date.Now.Month & "-" & MonthName(Date.Now.Month) & ".log"
    Public Property _settingsFile As String = "C:\emerge\survey\settings.xml"
    Public Property _logEmail As String = "info@emerge-it.co.uk"

    Public Sub Log(ByVal type As String, _
                   ByVal msg As String, _
                   Optional ByVal alsoEmail As Boolean = False, _
                   Optional ByVal logRecipients As List(Of String) = Nothing)
        CreateLog()
        Dim sb As New StringBuilder()
        With sb
            .Append(vbCrLf)
            .Append("[" & Date.Now.ToShortDateString)
            .Append(" " & Date.Now.ToShortTimeString & "] ")
            .Append(type & ": ")
            .Append(msg)
        End With
        Using sw As New StreamWriter(_logFile, True)
            sw.Write(sb.ToString())
        End Using
        If alsoEmail Then
            SendLogMail(logRecipients, type, msg)
        End If
    End Sub

    Public Sub CreateLog()
        If Not Directory.Exists(_logDir) Then
            Directory.CreateDirectory(_logDir)
        End If
        If Not File.Exists(_logFile) Then
            File.Create(_logFile).Close()
        End If
    End Sub

    Public Sub ReadLogSettings()
        Dim x As New XMLer
        Dim doc As New XDocument
        Dim settingsDoc As New XDocument
        If Not File.Exists(_settingsFile) Then
            x.createSettings()
        End If

        Dim settings As New List(Of String)
        settings = x.readSettings()

        If settings(4) IsNot Nothing Then
            _logEmail = settings(4)
        End If
    End Sub

    Public Sub SendLogMail(ByVal recipients As List(Of String),
                           ByVal details As String, _
                           ByVal msg As String, _
                           Optional ByVal fromAddress As String = "surveys@emerge-it.co.uk", _
                           Optional ByVal userName As String = "Info", _
                        Optional ByVal password As String = "", _
                        Optional attachments As List(Of String) = Nothing)

        Dim server As SmtpClient = New SmtpClient("blofeld", 25)
        With server
            .DeliveryMethod = SmtpDeliveryMethod.Network
            .PickupDirectoryLocation = "\\blofeld\Pickup"
        End With

        Dim email As New MailMessage()
        Try
            If (attachments IsNot Nothing) Then
                For Each attachment As String In attachments
                    email.Attachments.Add(New Attachment(attachment))
                Next
            End If

            For Each r As String In recipients
                email.To.Add(r)
            Next

            email.From = New MailAddress(fromAddress)
            email.Subject = "Emerge IT Customer Support Survey - " & details
            email.Body = msg
            email.IsBodyHtml = True

            server.Credentials = New System.Net.NetworkCredential(userName, password)
            server.Send(email)

            Log("Email sent", _
              "Survey logging email sent to: " & _
               recipients.ToString & _
               ". Reference: " & details)

            email.Dispose()
        Catch ex As Exception
            email.Dispose()

            Log("Email sending failed", _
                    "Failed to send survey logging email to: " & _
                    recipients.ToString & ". Error details: " & _
                    ex.ToString())
        End Try

    End Sub
End Class
