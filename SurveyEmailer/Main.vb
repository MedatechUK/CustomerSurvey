Module Main
    Public Sub Main()
        Dim contacts As Dictionary(Of Integer, List(Of String))
        Logger.CreateLog()

        If Not Settings.ReadSettings() Then
            Logger.Log("Settings", "Unable to read settings from database." & vbCrLf & _
                  "Shutting down...", True, Settings.LogEmailAddress)
            Exit Sub
        End If


        ' Return dictionary of contacts of format:
        ' {AUTOUNIQUE, (NAME, EMAIL, DATELASTEMAILED)}
        contacts = Databaser.ListServiceCallContacts()
        contacts = XMLer.ProcessContacts(contacts)

        If contacts.Count = Settings.NumCustToEmail Then
            Emailer.SendSurvey(contacts)
        ElseIf contacts.Count > Settings.NumCustToEmail Then
            Logger.Log("Email sending failed", _
                       "Did not send any emails as more than " & _
                       Settings.NumCustToEmail & " (as defined in settings) contacts " & _
                       "were passed to the emailer. Please contact Emerge support.",
                       True, Settings.LogEmailAddress)
        Else
            Logger.Log("Email sending failed", _
                       "Did not send any emails as less than " & _
                       Settings.NumCustToEmail & " (as defined in settings) contacts " & _
                       "were passed to the emailer. Please contact Emerge support.",
                       True, Settings.LogEmailAddress)
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
