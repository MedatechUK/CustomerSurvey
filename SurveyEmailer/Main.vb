Imports CustomerSurvey.Databaser
Imports CustomerSurvey.Emailer
Imports CustomerSurvey.Settings
Imports CustomerSurvey.XMLer

Module Main
    Public Sub Main()
        Dim doc As New XMLer

        Dim list As New List(Of String)(New String() {"paul", "two", Date.Now})
        Dim con As New Dictionary(Of Integer, List(Of String))
        con.Add(1, list)
        doc.Write(con)
        doc.Read(con)
        doc.Write(con)
    End Sub

End Module
