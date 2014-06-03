Imports System.IO
Imports System.Text
Imports System.Xml.Linq

Public Class XMLer
    Public Shared ContactsFile As String = "contacts.xml"

    Public Shared Function ProcessContacts(ByVal contacts As Dictionary(Of Integer, List(Of String)))
        ' Takes a list of 10 random contacts as dictionary(phone autounique, ("name", "email"))

        Logger.Log("XML Contacts", _
              "Processing contacts recieved from database with the following settings: " & vbCrLf & _
              "DaysBetweenSurveyRuns: " & Settings.DaysBetweenSurveys & vbCrLf & _
              "TimeBetweenEmailingCustomerInMonths: " & Settings.MonthsBetweenEmails & vbCrLf & _
              "NumberCustomersToEmail: " & Settings.NumCustToEmail & vbCrLf & _
              "DateLastEmailed: " & Settings.DateLastEmailed & vbCrLf & _
              "AdminEmailAddress: " & Settings.AdminEmailAddress & vbCrLf & _
            "IgnoreSurveyFlag: " & Settings.IgnoreSurveyFlag & vbCrLf & _
            "NumberDaysInPastToCheckForServiceCalls: " & Settings.NumDaysInPast & vbCrLf & _
            "LogEmailAddress: " & Settings.LogEmailAddress)

        ' First, if the program has run in fewer than the DaysBetweenSurveyRuns setting, break out
        If Settings.DateLastEmailed <> "" Then
            Dim diff As TimeSpan = Date.Now.Subtract(Settings.DateLastEmailed)
            If diff.Days < Settings.DaysBetweenSurveys Then
                contacts.Clear()
                Logger.Log("XML Contacts", _
                      "It has been " & diff.Days & " days since the last run, which is less than " & _
                      "the " & Settings.DaysBetweenSurveys & " days defined in the settings." & vbCrLf & _
                      "Clearing list of contacts and closing program.")
                Return contacts
            End If
        End If

        Dim phonesToDelete As New List(Of Integer)

        If File.Exists(ContactsFile) Then
            Logger.Log("XML Contacts", _
                  "Contacts file exists. Processing...")

            Dim doc As XDocument
            doc = XDocument.Load(ContactsFile)
            Dim x As Integer = Settings.NumCustToEmail
            ' Make sure the contact has not recently been sent a survey
            For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                Dim emailDate = From result In doc.Root.Elements
                       Where result.Attribute("id").Value = kvp.Key
                       Select result.Element("date")

                ' If the contact does not exist in the XML contacts file then skip over this 
                If emailDate IsNot Nothing And x > 0 Then
                    ' If they are there but have been emailed x or more months ago, skip
                    If Not FormatDateTime(emailDate.Value, DateFormat.GeneralDate) < Date.Now.AddMonths(-Settings.MonthsBetweenEmails) Then
                        ' If they are there are have been emailed less than x months ago, remove from list to email
                        phonesToDelete.Add(kvp.Key)
                        Logger.Log("XML Contacts", _
                              kvp.Value(0) & " REMOVED, because they have been emailed recently.")
                    Else
                        Logger.Log("XML Contacts", _
                              kvp.Value(0) & " TO BE EMAILED... last emailed " & emailDate.Value)
                        emailDate.Value = Date.Now
                        x -= 1
                    End If
                Else
                    Logger.Log("XML Contacts", _
                             kvp.Value(0) & " TO BE EMAILED... never previously been emailed")
                    ' If their email address is blank then add them to the xml document
                    Dim id As Integer = kvp.Key
                    Dim name As String = kvp.Value(0)
                    Dim contactEmail As String = kvp.Value(1)
                    Dim SCdetails As String = kvp.Value(2)
                    Dim SCdocno As String = kvp.Value(3)
                    Dim SCdoc As String = kvp.Value(4)
                    Dim emailedDate As String = Date.Now

                    doc.Root.Add(New XElement("contact",
                          New XAttribute("id", id),
                          New XElement("name", name),
                          New XElement("email", contactEmail),
                          New XElement("SCdetails", SCdetails),
                          New XElement("SCdocno", SCdocno),
                          New XElement("SCdoc", SCdoc),
                          New XElement("date", emailedDate)
                          ))
                    x -= 1
                End If
            Next

            For Each i As Integer In phonesToDelete
                contacts.Remove(i)
            Next

            doc.Save(ContactsFile)
        Else
            Dim contactNames As New StringBuilder
            If contacts.Count > Settings.NumCustToEmail Then
                Dim x As Integer = Settings.NumCustToEmail
                For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                    If x < 1 Then
                        phonesToDelete.Add(kvp.Key)
                        contactNames.Append(vbCrLf)
                        contactNames.Append(kvp.Value(0))
                    End If
                    x -= 1
                Next
            End If

            For Each i As Integer In phonesToDelete
                contacts.Remove(i)
            Next

            Logger.Log("XML Contacts", _
                  "First email run. Using first six from batch:" & _
                  contactNames.ToString)

            WriteNewContactsFile(contacts)
        End If

        Return contacts
    End Function

    Public Shared Sub WriteNewContactsFile(ByVal contacts As Dictionary(Of Integer, List(Of String)))
        Dim doc As New XDocument
        Dim root As XElement = _
            <contacts>
            </contacts>
        doc.Add(root)

        For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
            Dim phone As Integer = kvp.Key
            Dim name As String = kvp.Value(0)
            Dim email As String = kvp.Value(1)
            Dim SCdetails As String = kvp.Value(2)
            Dim SCdocno As String = kvp.Value(3)
            Dim emailDate As Date = Date.Now

            doc.Root.Add(New XElement("contact",
                                      New XAttribute("id", phone),
                                      New XElement("name", name),
                                      New XElement("email", email),
                                      New XElement("SCdetails", SCdetails),
                                      New XElement("SCdocno", SCdocno),
                                      New XElement("date", emailDate)
                                      ))
        Next
        doc.Save(ContactsFile)
    End Sub
End Class
