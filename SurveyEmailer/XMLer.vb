Imports System.IO
Imports System.Security.AccessControl
Imports System.Xml
Imports System.Xml.Linq


Public Class XMLer
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

        ' settings(0) = DaysBetweenSurveyRuns
        ' settings(1) = TimeBetweenEmailingCustomerInMonths
        ' settings(2) = NumberCustomersToEmail
        ' settings(3) = DateLastEmailed
        Dim settings = From x In settingsDoc.Root.Elements
                       Select x.Value

        Dim days As Integer = Convert.ToInt32(settings(0))
        Dim mons As Integer = Convert.ToInt32(settings(1))
        Dim numCust As Integer = Convert.ToInt32(settings(2))
        Dim dateEmailed As String = settings(3).ToString

        ' First, if the program has run in fewer than the DaysBetweenSurveyRuns setting, break out
        If dateEmailed <> "" Then
            Dim diff As System.TimeSpan = Date.Now.Subtract(dateEmailed)
            If diff.Days < days Then
                contacts.Clear()
                Return contacts
            End If
        End If

        If File.Exists(contactsFile) Then
            Dim phonesToDelete As New List(Of Integer)
            doc = XDocument.Load(contactsFile)

            ' Make sure the contact has not recently been sent a survey
            For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                Dim emailDate = From result In doc.Root.Elements
                       Where result.Attribute("id").Value = kvp.Key
                       Select result.Element("date")

                ' If the contact does not exist in the XML contacts file then skip over this 
                If emailDate IsNot Nothing Then
                    ' If they are there but have been emailed x or more months ago, skip
                    If Not FormatDateTime(emailDate.Value, DateFormat.GeneralDate) < Date.Now.AddMonths(-mons) Then
                        ' If they are there are have been emailed less than x months ago, remove from list to email
                        phonesToDelete.Add(kvp.Key)
                    Else
                        emailDate.Value = Date.Now
                    End If
                Else
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
                End If
            Next

            For Each x As Integer In phonesToDelete
                contacts.Remove(x)
            Next

            ' Remove any contacts that would leave more than the desired number
            If contacts.Count > numCust Then
                Dim x As Integer = 0
                For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
                    If x > (numCust - 1) Then
                        contacts.Remove(kvp.Key)
                    End If
                    x += 1
                Next
            End If
        Else
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
