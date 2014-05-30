Module Main
    Public Property _AdminEmailAddress As String
    Public Property LogRecipients As New List(Of String)

    Public Sub Main()
        Dim d As New Databaser, e As New Emailer, _
            l As New Logger, s As New Settings, x As New XMLer

        Dim contacts As Dictionary(Of Integer, List(Of String))
        Dim emailList As New List(Of String)

        If Not s.ReadSettings() Then
            l.Log("Settings", "Unable to read settings from database." & vbCrLf & _
                  "Shutting down...", True, LogRecipients)
        End If

        ' Return dictionary of contacts of format:
        ' {AUTOUNIQUE, (NAME, EMAIL, DATELASTEMAILED)}
        contacts = d.ListServiceCallContacts()
        contacts = x.ProcessContacts(contacts)

        If contacts.Count = x.NumberCustomersToEmail Then
            e.SendSurvey(contacts)
        ElseIf contacts.Count > x.NumberCustomersToEmail Then

        Else

        End If
        ' TODO: swap this with contacts for release version
        'Dim con As Dictionary(Of Integer, List(Of String))
        'con = Cheat()

        'con = x.ProcessContacts(con)
        'If con.Count > 0 Then
        '    e.SendSurvey(con)
        'ElseIf con.Count > 6 Then
        '    l.Log("Email sending", "Did not send any surveys: more than 6 ", True, )
        'Else
        '    l.Log("Email sending", "Did not send any surveys: list of contacts was empty.")
        'End If
        s.ReadSettings()
        Dim y As DateTime = New DateTime(1988, 1, 1).AddMinutes(13885920)
        Dim yz As TimeSpan = (Date.Now.Subtract(y))
        Dim lal As Integer = yz.Days
        Dim z As Int64
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
