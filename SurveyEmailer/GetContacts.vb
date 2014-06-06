Imports System.Data.SqlClient

Public Class GetContacts
    Public Shared Property DbContacts As New Dictionary(Of Integer, List(Of String))

    Public Shared Function ListServiceCallContacts()
        ' Selects all contacts who have had calls in the last month via the SurveyCount view
        Dim startDate As New DateTime(1988, 1, 1)
        Dim minsTodaySince1988 As Integer = (Date.Now - startDate).TotalMinutes
        Dim numMinsInPast As Integer = minsTodaySince1988 - (Settings.NumDaysInPast * 1440)

        Dim query As String = "SELECT TOP(100) PERCENT * " & _
                                "FROM (SELECT TOP(100) PERCENT dbo.PHONEBOOK.PHONE, dbo.PHONEBOOK.NAME, " & _
                                "dbo.PHONEBOOK.EMAIL, dbo.DOCUMENTS.DETAILS, dbo.DOCUMENTS.DOCNO, " & _
                                "dbo.DOCUMENTS.DOC, " & _
                                "ROW_NUMBER() OVER (PARTITION BY dbo.PHONEBOOK.PHONE " & _
                                "ORDER BY PHONEBOOK.PHONE DESC) rn " & _
                                "FROM dbo.PHONEBOOK JOIN " & _
                                "dbo.SERVCALLS ON dbo.PHONEBOOK.PHONE = dbo.SERVCALLS.PHONE JOIN " & _
                                "dbo.DOCUMENTS ON dbo.SERVCALLS.DOC = dbo.DOCUMENTS.DOC JOIN " & _
                                "dbo.CALLSTATUSES ON dbo.SERVCALLS.CALLSTATUS = dbo.CALLSTATUSES.CALLSTATUS " & _
                                "WHERE (dbo.DOCUMENTS.TYPE = 'Q') " & _
                                "AND (dbo.DOCUMENTS.CURDATE > " & numMinsInPast & ") " & _
                                "AND (dbo.CALLSTATUSES.CANCEL <> 'Y') AND (dbo.DOCUMENTS.DOC IS NOT NULL) " & _
                                "AND (dbo.PHONEBOOK.EMAIL <> '') AND (dbo.PHONEBOOK.INACTIVE <> 'Y') "

        If Settings.IgnoreSurveyFlag = False Then
            query &= "AND (dbo.CALLSTATUSES.ZEMG_SURVEYALLOWED = 'Y') "
        End If

        query &= "ORDER BY dbo.PHONEBOOK.NAME) a " & _
                 "WHERE rn = 1 " & _
                 "ORDER BY NAME"

        Try
            Dim con As SqlConnection = New SqlConnection(DbSettings.ConString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                If reader.HasRows Then
                    Do While reader.Read()
                        Dim list As New List(Of String)

                        Dim phone As Integer = reader.GetInt64(0)
                        Dim name As String = reader.GetString(1)
                        Dim email As String = reader.GetString(2)
                        Dim details As String = reader.GetString(3)
                        Dim docno As String = reader.GetString(4)
                        Dim doc As String = reader.GetInt64(5)

                        list.Add(name)
                        list.Add(email)
                        list.Add(details)
                        list.Add(docno)
                        list.Add(doc)

                        DbContacts.Add(phone, list)
                    Loop

                End If
                Logger.Log("Database connection succeeded", _
            "Retrieved a list of " & DbContacts.Count.ToString & _
            " contacts from the database who have had service calls in the last month.")
            Catch ex As Exception
                Logger.Log("Database query failed", _
                "Retrieved no data from the database." & vbCrLf & _
                "Error details: " & ex.ToString())
            End Try
        Catch ex As Exception
            Logger.Log("Database connection failed", _
            "Failed to connect to the database. " & _
           ". Error details: " & vbCrLf & _
            ex.ToString(), _
            True, _
            Settings.LogEmailAddress)
        End Try
        Return DbContacts
    End Function
End Class

