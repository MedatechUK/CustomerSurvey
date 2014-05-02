Imports System.IO
Imports System.Text
Imports System.Xml.Linq

Public Class XMLer
    Dim l As New Logger

    Public ContactsFile As String = "C:\emerge\survey\contacts.xml"
    Public SettingsFile As String = "C:\emerge\survey\settings.xml"

    Public Sub New(Optional ByVal dir As String = "C:\emerge\survey")
        If Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
    End Sub

    Public Function Read(ByVal contacts As Dictionary(Of Integer, List(Of String)))

        ' Takes a list of 10 random contacts as dictionary(phone autounique, ("name", "email"))
        Dim doc As XDocument
        Dim settings = New List(Of String)
        settings = readSettings()

        Dim daysBetweenSurveyRuns As Integer = Convert.ToInt32(settings(0))
        Dim timeBetweenEmailingCustomerInMonths As Integer = Convert.ToInt32(settings(1))
        Dim numberCustomersToEmail As Integer = Convert.ToInt32(settings(2))
        Dim dateLastEmailed As String = settings(3).ToString
        Dim adminEmailAddress As String = settings(4).ToString

        _adminEmailAddress = adminEmailAddress
        'logRecipients.Add(_adminEmailAddress)
        'TODO remove this cheat
        logRecipients.Add("paul@emerge-it.co.uk")

        l.Log("XML Contacts", _
              "Processing contacts recieved from database with the following settings: " & vbCrLf & _
              "DaysBetweenSurveyRuns: " & daysBetweenSurveyRuns & vbCrLf & _
              "TimeBetweenEmailingCustomerInMonths: " & timeBetweenEmailingCustomerInMonths & vbCrLf & _
              "NumberCustomersToEmail: " & numberCustomersToEmail & vbCrLf & _
              "DateLastEmailed: " & dateLastEmailed & vbCrLf & _
              "AdminEmailAddress: " & adminEmailAddress & ".")

        ' First, if the program has run in fewer than the DaysBetweenSurveyRuns setting, break out
        If dateLastEmailed <> "" Then
            Dim diff As TimeSpan = Date.Now.Subtract(dateLastEmailed)
            If diff.Days < daysBetweenSurveyRuns Then
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

                Dim x As Integer = numberCustomersToEmail
                ' If the contact does not exist in the XML contacts file then skip over this 
                If emailDate IsNot Nothing And x > 0 Then
                    ' If they are there but have been emailed x or more months ago, skip
                    If Not FormatDateTime(emailDate.Value, DateFormat.GeneralDate) < Date.Now.AddMonths(-timeBetweenEmailingCustomerInMonths) Then
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

            For Each x As Integer In phonesToDelete
                contacts.Remove(x)
            Next

            doc.Save(contactsFile)
        Else
            Dim contactNames As New Stringbuilder
            If contacts.Count > numberCustomersToEmail Then
                Dim x As Integer = numberCustomersToEmail
                For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                    If x < 1 Then
                        phonesToDelete.Add(kvp.Key)
                        contactNames.Append(vbCrLf)
                        contactNames.Append(kvp.Value(0))
                    End If
                    x -= 1
                Next
            End If

            For Each x As Integer In phonesToDelete
                contacts.Remove(x)
            Next

            l.Log("XML Contacts", _
                  "First email run. Using first six from batch:" & _
                  contactNames.ToString)

            writeNew(contacts)
        End If

        Return contacts
    End Function

    Public Sub WriteNew(ByVal contacts As Dictionary(Of Integer, List(Of String)))
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

    Public Function ReadSettings()

        Dim settingsDoc As New XDocument
        If Not File.Exists(SettingsFile) Then
            createSettings()
        End If

        settingsDoc = XDocument.Load(SettingsFile)

        Dim settings = From x In settingsDoc.Root.Elements
                       Select x.Value

        Dim settingsList As New List(Of String)

        settingsList.Add(settings(0).ToString)
        settingsList.Add(settings(1).ToString)
        settingsList.Add(settings(2).ToString)
        settingsList.Add(settings(3).ToString)
        settingsList.Add(settings(4).ToString)

        Return settingsList
    End Function

    Public Sub CreateSettings()
        Dim doc As New XDocument
        If Not File.Exists(SettingsFile) Then
            Dim root As XElement = _
                <settings>
                </settings>
            doc.Add(root)

            Dim frequency = <DaysBetweenSurveyRuns>14</DaysBetweenSurveyRuns>
            Dim monthsBetweenEmails = <TimeBetweenEmailingCustomerInMonths>10</TimeBetweenEmailingCustomerInMonths>
            Dim numCustomers = <NumberCustomersToEmail>6</NumberCustomersToEmail>
            Dim lastEmail = <DateLastEmailed></DateLastEmailed>
            Dim adminEmailAddress = <AdminEmailAddress>info@emerge-it.co.uk</AdminEmailAddress>

            doc.Root.Add(frequency)
            doc.Root.Add(monthsBetweenEmails)
            doc.Root.Add(numCustomers)
            doc.Root.Add(lastEmail)
            doc.Root.Add(adminEmailAddress)

            doc.Save(SettingsFile)
        End If
    End Sub

End Class
