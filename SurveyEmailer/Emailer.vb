﻿Imports System.Net.Mail
Imports System.Net.Mail.MailMessage
Imports System.Xml
Imports System.Xml.Linq

Public Class Emailer
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

            Dim body As String = CreateEmail(name, SCdocno, SCdetails)

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
                ' Update emailed date in settings
                emailedSetting()
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

        Next


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
                              "<table align=""center"" id=""choose""><td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=1"">1</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=2"">2</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=3"">3</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=4"">4</a></td>" & _
                              "<td width=""200""><a href=""http://www.emerge-it.co.uk/support/survey.aspx?q=5"">5</a></td>" & _
                              "</table><br/>" & _
                              "Thank you for taking the time,<br/>" & _
                              "The Emerge IT support team" & _
                              "</body></html>"

        Return email
    End Function

    Private Sub emailedSetting()
        Dim settingsFile As String = "C:\emerge\survey\settings.xml"
        Dim doc As New XDocument
        doc = XDocument.Load(settingsFile)
        Dim x As XElement = doc.Root.Element("DateLastEmailed")
        doc.Root.Element("DateLastEmailed").Value = Date.Now
        doc.Save(Settingsfile)
    End Sub

End Class
