Imports System.Configuration
Imports System.Data.SqlClient

Public Class Databaser
    Public Property contacts As List(Of Dictionary(Of String, String()))

    Public Sub New()
        Dim query As String = "select * from dbo.v_CustomerSurvey;"
        Try
            Dim constring As String = ConfigurationManager.ConnectionStrings(0).ConnectionString
            Dim con As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("joker").ConnectionString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                While reader.Read()
                    contacts.Add(New Dictionary(Of String, String())() From { _
                                 {"PHONE", reader.Item("PHONE")} _
                    })
                    contacts.Add(New Dictionary(Of String, String())() From { _
                                 {"NAME", reader.Item("NAME")}})
                End While
            Catch ex As Exception

            End Try
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
        


    End Sub
End Class

