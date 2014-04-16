Imports CustomerSurvey.Databaser
Imports CustomerSurvey.Emailer
Imports CustomerSurvey.XMLer

Module Main
    Public Sub Main()
        Dim e As New Emailer
        Dim body As String = e.CreateEmail("andy", "sc140005", "DAMNED PRIORITYYYYY!!")
        Dim list As New List(Of String)
        'list.Add("andymackintosh@emerge-it.co.uk")
        list.Add("paul@emerge-it.co.uk")
        e.SendSurvey(list, "survey@emerge-it.co.uk", "Customer Survey", body, "Info", "")
    End Sub
End Module
