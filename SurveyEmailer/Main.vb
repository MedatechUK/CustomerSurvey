Imports CustomerSurvey.Databaser
Imports CustomerSurvey.Emailer
Imports CustomerSurvey.XMLer

Module Main
    Public Sub Main()
        Dim d As New Databaser, e As New Emailer, x As New XMLer

        Dim contacts As New Dictionary(Of Integer, List(Of String))
        Dim emailList As New List(Of String)

        ' Return dictionary of contacts of format:
        ' {AUTOUNIQUE, (NAME, EMAIL, DATELASTEMAILED)}
        contacts = d.Read()

        ' TODO: swap this with contacts for release version
        Dim con As New Dictionary(Of Integer, List(Of String))
        con = cheat()

        con = x.Read(con)
        If con IsNot Nothing Then
            e.SendSurvey(con)
        End If

    End Sub

    Public Function cheat()
        Dim dict As New Dictionary(Of Integer, List(Of String))
        Dim list As New List(Of String)
        list.Add("Paul Wilson")
        list.Add("paul@emerge-it.co.uk")
        list.Add("Service Call Details")
        list.Add("SC1400001")
        dict.Add(1, list)
        Return dict
    End Function
End Module
