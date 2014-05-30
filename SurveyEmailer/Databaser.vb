Imports System.Data.SqlClient

Public Class Databaser
    Public Property DbContacts As New Dictionary(Of Integer, List(Of String))

    Dim l As New Logger

    Public Function ListServiceCallContacts()
        ' Selects all contacts who have had calls in the last month via the SurveyCount view
        Dim conString As String = DbSettings.ConString

        Dim query As String = "SELECT TOP(100) PERCENT * " & _
                                "FROM (SELECT TOP(100) PERCENT dbo.PHONEBOOK.PHONE, dbo.PHONEBOOK.NAME, " & _
                                "dbo.PHONEBOOK.EMAIL, dbo.DOCUMENTS.DETAILS, dbo.DOCUMENTS.DOCNO " & _
                                "ORDER BY PHONEBOOK.PHONE DESC) rn" & _
                                "FROM dbo.PHONEBOOK JOIN " & _
                                "dbo.SERVCALLS ON dbo.PHONEBOOK.PHONE = dbo.SERVCALLS.PHONE JOIN " & _
                                "dbo.DOCUMENTS ON dbo.SERVCALLS.DOC = dbo.DOCUMENTS.DOC JOIN " & _
                                "dbo.CALLSTATUSES ON dbo.SERVCALLS.CALLSTATUS = dbo.CALLSTATUSES.CALLSTATUS " & _
                                "WHERE (dbo.DOCUMENTS.TYPE = 'Q') AND (dbo.CALLSTATUSES.ZEMG_SURVEYALLOWED = 'Y') " & _
                                "AND (dbo.MINTODATE(dbo.DOCUMENTS.CURDATE) > GETDATE() - @daysBetween) " & _
                                "AND (dbo.CALLSTATUSES.CANCEL <> 'Y') AND (dbo.DOCUMENTS.DOC IS NOT NULL)" & _
                                "AND (dbo.PHONEBOOK.EMAIL <> '') AND (dbo.PHONEBOOK.INACTIVE <> 'Y') " & _
                                "ORDER BY dbo.PHONEBOOK.NAME) a " & _
                                "WHERE rn = 1 " & _
                                "ORDER BY NAME"

        Try
            Dim con As SqlConnection = New SqlConnection(conString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            'TODO here
            cmd.Parameters.Add()

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                If reader.HasRows Then
                    Do While reader.Read()
                        Dim list As New List(Of String)

                        Dim phone As Integer = reader.GetInt32(0)
                        Dim name As String = reader.GetString(1)
                        Dim email As String = reader.GetString(2)
                        Dim details As String = reader.GetString(3)
                        Dim docno As String = reader.GetString(4)
                        Dim doc As String = reader.GetString(5)

                        list.Add(name)
                        list.Add(email)
                        list.Add(details)
                        list.Add(docno)
                        list.Add(doc)

                        DbContacts.Add(phone, list)
                    Loop

                End If
                l.Log("Database connection succeeded", _
            "Retrieved a list of " & DbContacts.Count.ToString & _
            " contacts from the database who have had service calls in the last month.")
            Catch ex As Exception
                l.Log("Database query failed", _
                "Retrieved no data from the database." & vbCrLf & _
                "Error details: " & ex.ToString())
            End Try
        Catch ex As Exception
            l.Log("Database connection failed", _
            "Failed to connect to the database. " & _
           ". Error details: " & vbCrLf & _
            ex.ToString(), _
            True, _
            logRecipients)
        End Try
        Return dbContacts
    End Function
End Class

