Imports System.Net.Mail

Module Emailer
    Public curDir As String = My.Computer.FileSystem.CurrentDirectory()

    Public Class Emailer
        Public Sub New()
            
        End Sub
        Public Sub SendSurvey(ByVal recipients As List(Of String), _
                            ByVal fromAddress As String, _
                            ByVal subject As String, _
                            ByVal body As String, _
                            ByVal userName As String, _
                            ByVal password As String, _
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


                email.From = New MailAddress(fromAddress)

                For Each recipient As String In recipients
                    email.To.Add(recipient)
                Next

                email.Subject = subject
                email.Body = body

                server.Credentials = New System.Net.NetworkCredential(userName, password)
                server.Send(email)
                email.Dispose()
            Catch ex As SmtpException
                email.Dispose()
                Debug.WriteLine("Sending email failed: " & ex.ToString())
            Catch ex As ArgumentOutOfRangeException
                email.Dispose()
                Debug.WriteLine("Sending email failed. Check port number. " & ex.ToString())
            Catch ex As InvalidOperationException
                email.Dispose()
                Debug.WriteLine("Sending email failed. Check port number. " & ex.ToString())
            Catch ex As Exception
                email.Dispose()
                Debug.WriteLine("Sending email failed: " & ex.ToString())
            End Try

        End Sub

    End Class

End Module
