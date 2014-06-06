Module Main
    Public Sub Main()
        Dim contacts As Dictionary(Of Integer, List(Of String))
        Logger.CreateLog()

        If Not Settings.ReadSettings() Then
            Logger.Log("Settings", "Unable to read settings from database." & vbCrLf & _
                  "Shutting down.", True, Settings.LogEmailAddress)
            Exit Sub
        ElseIf Settings.SurveysSentRecently() = True Then
            Logger.Log("Settings", "Surveys sent less than " & Settings.DaysBetweenSurveys & " days ago. " & _
                       vbCrLf & "Shutting down.")
            Exit Sub
        End If

        contacts = GetContacts.ListServiceCallContacts()
        contacts = ProcessContacts.Process(contacts)

        If contacts.Count = Settings.NumCustToEmail Then
            Emailer.SendSurvey(contacts)
        ElseIf contacts.Count = 0 Then
            Logger.Log("Surveys not sent", _
                       "No surveys sent as no contacts were supplied." & _
                       " See logs for details (&lt;Install directory&gt;\logs\&lt;month&gt;.log).",
                       True, Settings.LogEmailAddress)
        ElseIf contacts.Count < Settings.NumCustToEmail Then
            Emailer.SendSurvey(contacts)
            Logger.Log("Surveys sent", _
                       "Sent surveys, although less than " & _
                       Settings.NumCustToEmail & " (as defined in settings) contacts " & _
                       "were passed to the emailer." & _
                       " See logs for details (&lt;Install directory&gt;\logs\&lt;month&gt;.log).",
                       True, Settings.LogEmailAddress)
        Else
            Logger.Log("Survey sending failed", _
                       "Did not send any emails as more than " & _
                       Settings.NumCustToEmail & " (as defined in settings) contacts " & _
                       "were passed to the emailer. Please contact Emerge support." & _
                       " See logs for details (&lt;Install directory&gt;\logs\&lt;month&gt;.log).",
                       True, Settings.LogEmailAddress)
        End If

        Settings.UpdateEmailedDate()
    End Sub
End Module
