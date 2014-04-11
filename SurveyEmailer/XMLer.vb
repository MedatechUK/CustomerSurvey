Imports System.Xml
Imports System.Xml.Linq
Imports System.IO

Public Class XMLer
    Dim contactsFile As String = curDir & "contacts.xml"

    Public Function Read(ByVal contacts As Dictionary(Of Integer, List(Of String)))

        ' Takes a list of 10 random contacts as dictionary(phone autounique, ("name", "email"))
        Dim doc As New XDocument
        doc = XDocument.Load(contactsFile)

        Dim phonesToDelete As New List(Of Integer)

        For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
            Dim emailDate = From result In doc.Root.Elements
                   Where result.Attribute("id").Value = kvp.Key
                   Select result.Element("date")

            ' If the random contact does not exist in the XML contacts file then skip over this 
            If emailDate IsNot Nothing Then
                ' If they are there but have been emailed 6 or more months ago, skip
                If Not FormatDateTime(emailDate.Value, DateFormat.GeneralDate) < Date.Now.AddMonths(-6) Then
                    ' If they are there are have been emailed less than 6 months ago, remove from list to email
                    phonesToDelete.Add(kvp.Key)
                End If
            End If
        Next

        For Each x As Integer In phonesToDelete
            contacts.Remove(x)
        Next

        Return contacts
    End Function

    Public Sub Write(ByVal contacts As Dictionary(Of Integer, List(Of String)))
        Dim doc As New XDocument

        Dim root As XElement = _
            <contacts>
            </contacts>
        doc.Add(root)

        For Each kvp As KeyValuePair(Of Integer, List(Of String)) In contacts
            Dim phone As Integer = kvp.Key
            Dim name As String = kvp.Value(0)
            Dim email As String = kvp.Value(1)
            Dim emailDate As Date = kvp.Value(2)

            doc.Root.Add(New XElement("contact",
                                      New XAttribute("id", phone),
                                      New XElement("name", name),
                                      New XElement("email", email),
                                      New XElement("date", emailDate)))
        Next

        doc.Save(contactsFile)
    End Sub

End Class
