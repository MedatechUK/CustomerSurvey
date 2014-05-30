Imports System
Imports System.IO
Imports System.Net.Mail
Imports System.Text

Public Class Logger
    Public Property LogPath As String = "C:\emerge\survey\logs\"
    Public Property LogDir As String = LogPath & Date.Now.Year
    Public Property LogFile As String = LogDir & "\" & Date.Now.Month & "-" & MonthName(Date.Now.Month) & ".log"
    Public Property LogEmail As String = "info@emerge-it.co.uk"

    Public Sub New()
        CreateLog()
    End Sub

    Public Sub Log(ByVal type As String, _
                   ByVal msg As String, _
                   Optional ByVal alsoEmail As Boolean = False, _
                   Optional ByVal logRecipients As List(Of String) = Nothing)
        Dim sb As New StringBuilder()
        With sb
            .Append(vbCrLf)
            .Append("[" & Date.Now.ToShortDateString)
            .Append(" " & Date.Now.ToShortTimeString & "] ")
            .Append(type & ": ")
            .Append(msg)
        End With
        Using sw As New StreamWriter(LogFile, True)
            sw.Write(sb.ToString())
        End Using
        If alsoEmail Then
            SendLogMail(logRecipients, type, msg)
        End If
    End Sub

    Public Sub CreateLog()
        If Not Directory.Exists(LogDir) Then
            Directory.CreateDirectory(LogDir)
        End If
        If Not File.Exists(LogFile) Then
            File.Create(LogFile).Close()
        End If
    End Sub

    Public Sub SendLogMail(ByVal recipients As List(Of String),
                           ByVal SCdetails As String, _
                           ByVal emailBody As String, _
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

            With email
                .From = New MailAddress(fromAddress)
                .Subject = "Emerge IT Customer Support Survey - " & SCdetails
                .Body = emailBody
                .IsBodyHtml = True
            End With

            server.Credentials = New Net.NetworkCredential(userName, password)
            server.Send(email)

            Log("Email sent", _
              "Survey logging email sent to: " & _
               recipients.ToString & _
               ". Reference: " & SCdetails)

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
