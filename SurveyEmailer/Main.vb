Module Main
    Public Property _adminEmailAddress As String
    Public Property logRecipients As New List(Of String)

    Public Sub Main()
        Dim d As New Databaser, e As New Emailer, x As New XMLer, l As New Logger

        'l.Log("Email", "test msg")

        'Dim contacts As New Dictionary(Of Integer, List(Of String))
        'Dim emailList As New List(Of String)

        ' Return dictionary of contacts of format:
        ' {AUTOUNIQUE, (NAME, EMAIL, DATELASTEMAILED)}
        'contacts = d.ListServiceCallContacts()

        ' TODO: swap this with contacts for release version
        Dim con As Dictionary(Of Integer, List(Of String))

        con = Cheat()

        con = x.Read(con)
        If con.Count > 0 Then
            e.SendSurvey(con)
        ElseIf con.Count > 6 Then
            l.Log("Email sending", "Did not send any surveys: more than 6 ", True, )
        Else
            l.Log("Email sending", "Did not send any surveys: list of contacts was empty.")
        End If


    End Sub

    Public Function Cheat()
        Dim dict As New Dictionary(Of Integer, List(Of String))
        Dim list As New List(Of String)
        Dim l2 As New List(Of String)
        list.Add("Paul Wilson")
        list.Add("paul@emerge-it.co.uk")
        list.Add("Service Call Details")
        list.Add("SC1400001")
        list.Add("199")
        l2.Add("hoams")
        l2.Add("thomas@emerge-it.co.uk")
        l2.Add("Service call details")
        l2.Add("SC140001")
        dict.Add(1, list)
        'dict.Add(2, l2)
        Return dict
    End Function
End Module
