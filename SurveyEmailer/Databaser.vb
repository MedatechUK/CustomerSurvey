Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO

Imports CustomerSurvey.Settings
Imports CustomerSurvey.Logger

Public Class Databaser
    Public Property dbContacts As New Dictionary(Of Integer, List(Of String))

    Dim l As New Logger

    Public Function ListServiceCallContacts()
        ' Selects all contacts who have had calls in the last month via the SurveyCount view
        Dim findDB As New Settings
        Dim query As String = "select * from dbo.v_SurveyCount;"
        Try
            Dim con As SqlConnection = New SqlConnection(findDB.conString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                If reader.HasRows Then
                    Do While reader.Read()
                        Dim list As New List(Of String)

                        Dim PHONE As Integer = reader.GetInt32(0)
                        Dim NAME As String = reader.GetString(1)
                        Dim EMAIL As String = reader.GetString(2)
                        Dim DETAILS As String = reader.GetString(3)
                        Dim DOCNO As String = reader.GetString(4)
                        Dim DOC As Integer = reader.GetInt32(5)

                        list.Add(NAME)
                        list.Add(EMAIL)
                        list.Add(DETAILS)
                        list.Add(DOCNO)
                        list.Add(DOC)

                        dbContacts.Add(PHONE, list)
                    Loop

                End If
                l.Log("Database connection succeeded", _
            "Retrieved a list of " & dbContacts.Count.ToString & _
            " contacts from the database who have had service calls in the last month.")
            Catch ex As Exception

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

