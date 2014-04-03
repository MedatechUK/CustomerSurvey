Imports System.Configuration
Imports System.Data.SqlClient

Public Class Databaser
    Public Sub New()
        Using con As New SqlConnection(ConfigurationManager.ConnectionStrings("joker").ConnectionString)
            Dim cmd As New SqlCommand = 
        End Using
    End Sub
End Class
