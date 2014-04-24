Imports System.IO
Imports System.Security.AccessControl
Imports System.Xml
Imports System.Xml.Linq

Imports CustomerSurvey.Logger


Public Class XMLer
    Dim l As New Logger

    Public contactsFile As String = "C:\emerge\survey\contacts.xml"
    Public settingsFile As String = "C:\emerge\survey\settings.xml"

    Public Sub New(Optional ByVal dir As String = "C:\emerge\survey")
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
    End Sub

    Public Function Read(ByVal contacts As Dictionary(Of Integer, List(Of String)))

        ' Takes a list of 10 random contacts as dictionary(phone autounique, ("name", "email"))
        Dim doc As New XDocument
        Dim settingsDoc As New XDocument
        If Not File.Exists(settingsFile) Then
            createSettings()
        End If

        settingsDoc = XDocument.Load(settingsFile)

        Dim settings = From x In settingsDoc.Root.Elements
                       Select x.Value

        Dim DaysBetweenSurveyRuns As Integer = Convert.ToInt32(settings(0))
        Dim TimeBetweenEmailingCustomerInMonths As Integer = Convert.ToInt32(settings(1))
        Dim NumberCustomersToEmail As Integer = Convert.ToInt32(settings(2))
        Dim DateLastEmailed As String = settings(3).ToString

        l.Log("XML Contacts", _
              "Processing contacts recieved from database with the following settings: " & vbCrLf & _
              "DaysBetweenSurveyRuns: " & DaysBetweenSurveyRuns & vbCrLf & _
              "TimeBetweenEmailingCustomerInMonths: " & TimeBetweenEmailingCustomerInMonths & vbCrLf & _
              "NumberCustomersToEmail: " & NumberCustomersToEmail & vbCrLf & _
              "DateLastEmailed: " & DateLastEmailed & ".")

        ' First, if the program has run in fewer than the DaysBetweenSurveyRuns setting, break out
        If DateLastEmailed <> "" Then
            Dim diff As System.TimeSpan = Date.Now.Subtract(DateLastEmailed)
            If diff.Days < DaysBetweenSurveyRuns Then
                contacts.Clear()
                Return contacts
            End If
        End If

        Dim phonesToDelete As New List(Of Integer)

        If File.Exists(contactsFile) Then
            l.Log("XML Contacts", _
                  "Contacts file exists. Processing...")

            doc = XDocument.Load(contactsFile)

            ' Make sure the contact has not recently been sent a survey
            For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                Dim emailDate = From result In doc.Root.Elements
                       Where result.Attribute("id").Value = kvp.Key
                       Select result.Element("date")

                Dim x As Integer = NumberCustomersToEmail
                ' If the contact does not exist in the XML contacts file then skip over this 
                If emailDate IsNot Nothing And x > 0 Then
                    ' If they are there but have been emailed x or more months ago, skip
                    If Not FormatDateTime(emailDate.Value, DateFormat.GeneralDate) < Date.Now.AddMonths(-TimeBetweenEmailingCustomerInMonths) Then
                        ' If they are there are have been emailed less than x months ago, remove from list to email
                        phonesToDelete.Add(kvp.Key)
                        l.Log("XML Contacts", _
                              kvp.Value(0) & " REMOVED, because they have been emailed recently.")
                    Else
                        l.Log("XML Contacts", _
                              kvp.Value(0) & " TO BE EMAILED... last emailed " & emailDate.Value)
                        emailDate.Value = Date.Now
                        x -= 1
                    End If
                Else
                    l.Log("XML Contacts", _
                             kvp.Value(0) & " TO BE EMAILED... never previously been emailed")
                    ' If their email address is blank then add them to the xml document
                    Dim id As Integer = kvp.Key
                    Dim name As String = kvp.Value(0)
                    Dim contactEmail As String = kvp.Value(1)
                    Dim SCdetails As String = kvp.Value(2)
                    Dim SCdocno As String = kvp.Value(3)
                    Dim emailedDate As String = Date.Now

                    doc.Root.Add(New XElement("contact",
                          New XAttribute("id", id),
                          New XElement("name", name),
                          New XElement("email", contactEmail),
                          New XElement("SCdetails", SCdetails),
                          New XElement("SCdocno", SCdocno),
                          New XElement("date", emailedDate)
                          ))
                    x -= 1
                End If
            Next

            For Each x As Integer In phonesToDelete
                contacts.Remove(x)
            Next

            doc.Save(contactsFile)
        Else
            If contacts.Count > NumberCustomersToEmail Then
                Dim x As Integer = NumberCustomersToEmail
                For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                    If x < 1 Then
                        phonesToDelete.Add(kvp.Key)
                    End If
                    x -= 1
                Next
            End If

            For Each x As Integer In phonesToDelete
                contacts.Remove(x)
            Next

            l.Log("XML Contacts", _
                  "First email run. Using first six from batch.")

            writeNew(contacts)
        End If

        Return contacts
    End Function

    Public Sub writeNew(ByVal contacts As Dictionary(Of Integer, List(Of String)))
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
        doc.Save(contactsFile)
    End Sub

    Private Sub createSettings()
        Dim doc As New XDocument
        If Not File.Exists(settingsFile) Then
            Dim root As XElement = _
                <settings>
                </settings>
            doc.Add(root)

            Dim freq = <DaysBetweenSurveyRuns>14</DaysBetweenSurveyRuns>
            Dim count = <TimeBetweenEmailingCustomerInMonths>10</TimeBetweenEmailingCustomerInMonths>
            Dim month = <NumberCustomersToEmail>6</NumberCustomersToEmail>
            Dim lastEmail = <DateLastEmailed></DateLastEmailed>

            doc.Root.Add(freq)
            doc.Root.Add(count)
            doc.Root.Add(month)
            doc.Root.Add(lastEmail)

            doc.Save(settingsFile)
        End If
    End Sub

End Class
