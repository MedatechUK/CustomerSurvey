Imports System.Net.Mail
Imports System.Net.Mail.MailMessage

Public Class Emailer
    Public curDir As String = My.Computer.FileSystem.CurrentDirectory()
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
            ' Email
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
            email.IsBodyHtml = True

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

    Public Function CreateEmail(ByVal name As String, _
                                ByVal docno As String, _
                                ByVal details As String)
        Dim email As String = "<html  xmlns=""http://www.w3.org/1999/xhtml"">" & _
                              "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8""/>" & _
                              "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"" /></head>" & _
                              "<title>Emerge IT - Customer Survey</title>" & _
                              "<style type=""text/css"">" & _
                              "#outlook a {padding:0;} body{width:100% !important; -webkit-text-size-adjust:100%; -ms-text-size-adjust:100%; margin:0; padding:0;} " & _
                              "#backgroundTable {margin:0; padding:0; width:100% !important; line-height: 100% !important;}" & _
                              "img {outline:none; text-decoration:none; -ms-interpolation-mode: bicubic;}a img {border:none;} .image_fix {display:block;} " & _
                              "table td {border-collapse: collapse;}table { border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; }" & _
                              "#choose{font-size: 1.5em; color: #6600FF;text-decoration:none;}#choose a{text-decoration:none;}" & _
                              ".center{text-align: center;} .orange{background-color: #F8B96C;}</style>" & _
                              "<body><table cellpadding=""0"" cellspacing=""0"" border=""0"" id=""backgroundTable"">" & _
                              "<tr class=""center""><td valign=""top""><img src=""http://www.emerge-it.co.uk/support/surveylogo.gif""/></td></tr>" & _
                              "<tr><td class=""center orange"" valign=""top""><h2>Support Desk Survey</h2></td></tr><br/>" & _
                              "<tr><td valign=""top"">Hello " & name & ",</td></tr><br/>" & _
                              "<tr><td valign=""top"">You recently raised service call: " & docno & ", regarding: " & details & "</td></tr>" & _
                              "<tr><td valign=""top"">We would love to hear your thoughts about our support desk.<br/> Please take a moment to answer the short question below.</td></tr><br/><br/>" & _
                              "<tr><td valign=""top"">How satisfied were you with our level of service? (1 is dissatisfied, 5 is delighted)</td></tr><br/>" & _
                              "<table align=""center"" id=""choose""><td width=""200""><a href=""http://www.emerge-it.co.uk/"">1</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/"">2</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/"">3</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/"">4</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/"">5</a></td>" & _
                              "</table><br/>" & _
                              "Thank you for taking the time,<br/>" & _
                              "The Emerge IT support team" & _
                              "</body></html>"

        Return email
    End Function

End Class
