Imports System.Net.Mail
Imports System.Text
Imports System.Xml.Linq


Public Class Emailer
    Dim l As New Logger

    Public Sub SendSurvey(ByVal contacts As Dictionary(Of Integer, List(Of String)), _
                        Optional ByVal fromAddress As String = "info@emerge-it.co.uk", _
                        Optional ByVal userName As String = "Info", _
                        Optional ByVal password As String = "", _
                        Optional attachments As List(Of String) = Nothing)

        Dim server As SmtpClient = New SmtpClient("blofeld", 25)
        With server
            .DeliveryMethod = SmtpDeliveryMethod.Network
            .PickupDirectoryLocation = "\\blofeld\Pickup"
        End With

        For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts

            Dim name As String = kvp.Value(0)
            Dim contactEmail As String = kvp.Value(1)
            Dim SCdetails As String = kvp.Value(2)
            Dim SCdocno As String = kvp.Value(3)
            Dim SCdoc As String = kvp.Value(4)

            l.Log("Email sending", _
                  "Sending email to: " & _
                  name & " , " & contactEmail & ". Reference: " & _
                  SCdocno & ": " & SCdetails)

            Dim body As String = CreateEmail(name, SCdocno, SCdetails, SCdoc).ToString()

            Dim email As New MailMessage()
            Try
                If (attachments IsNot Nothing) Then
                    For Each attachment As String In attachments
                        email.Attachments.Add(New Attachment(attachment))
                    Next
                End If

                email.To.Add(contactEmail)
                email.From = New MailAddress(fromAddress)

                email.Subject = "Emerge IT Support Customer Survey - Service Call " & SCdocno
                email.Body = body
                email.IsBodyHtml = True

                server.Credentials = New System.Net.NetworkCredential(userName, password)
                server.Send(email)

                l.Log("Email sent", _
                      "Sent email to: " & _
                       name & " , " & contactEmail & ". Reference: " & _
                       SCdocno & ": " & SCdetails, _
                        True, _
                        logRecipients)

                updateEmailedSetting()
                email.Dispose()

            Catch ex As Exception
                email.Dispose()

                l.Log("Email sending failed", _
                      "Failed to send email to: " & _
                       name & " , " & contactEmail & ". Reference: " & _
                       SCdocno & ": " & SCdetails & vbCrLf & ". Error details: " & _
                       ex.ToString(), _
                        True, _
                        logRecipients)
            End Try
        Next

    End Sub

    Public Function CreateEmail(ByVal name As String, _
                                ByVal SCdocno As String, _
                                ByVal details As String,
                                ByVal SCdoc As String)
        Dim emailBody As New StringBuilder
        With emailBody
            .Append("<html  xmlns=""http://www.w3.org/1999/xhtml"">" & _
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
                    "<tr><td valign=""top"">You recently raised service call: " & SCdocno & ", regarding: " & details & "</td></tr>" & _
                    "<tr><td valign=""top"">We would love to hear your thoughts about our support desk.<br/> Please take a moment to answer the short question below.</td></tr><br/><br/>" & _
                    "<tr><td valign=""top"">How satisfied were you with our level of service? (1 is dissatisfied, 5 is delighted)</td></tr><br/>")

            .Append(String.Format("<table align=""center"" id=""choose""><td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=1"">1</a></td>" & _
                    "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=2&d={0}"">2</a></td>" & _
                    "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=3&d={0}"">3</a></td>" & _
                    "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=4&d={0}"">4</a></td>" & _
                    "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=5&d={0}"">5</a></td>" & _
                    "</table><br/>", SCdoc))

            .Append("Thank you for taking the time,<br/>" & _
                    "The Emerge IT support team" & _
                    "</body></html>")
        End With

        Return emailBody
    End Function

    Private Sub updateEmailedSetting()
        Dim settingsFile As String = "C:\emerge\survey\settings.xml"
        Dim doc As New XDocument
        doc = XDocument.Load(settingsFile)
        doc.Root.Element("DateLastEmailed").Value = Date.Now
        doc.Save(settingsFile)
    End Sub

End Class
