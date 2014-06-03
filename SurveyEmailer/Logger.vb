Imports System
Imports System.IO
Imports System.Net.Mail
Imports System.Text

Public Class Logger
    Public Shared Property LogDir As String = "logs\" & Date.Now.Year
    Public Shared Property LogFile As String = LogDir & "\" & Date.Now.Month & "-" & MonthName(Date.Now.Month) & ".log"

    Public Shared Sub Log(ByVal type As String, _
                   ByVal msg As String, _
                   Optional ByVal alsoEmail As Boolean = False, _
                   Optional ByVal logRecipient As String = "")
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
            SendLogMail(logRecipient, type, msg)
        End If
    End Sub

    Public Shared Sub CreateLog()
        If Not Directory.Exists(LogDir) Then
            Directory.CreateDirectory(LogDir)
        End If
        If Not File.Exists(LogFile) Then
            File.Create(LogFile).Close()
        End If
    End Sub

    Public Shared Sub SendLogMail(ByVal recipient As String,
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

            email.To.Add(recipient)

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
               recipient.ToString & _
               ". Reference: " & SCdetails)

            email.Dispose()

        Catch ex As Exception
            email.Dispose()

            Log("Email sending failed", _
                    "Failed to send survey logging email to: " & _
                    recipient.ToString & ". Error details: " & _
                    ex.ToString())
        End Try

    End Sub
End Class
