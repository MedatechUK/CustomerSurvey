Imports System.Data.SqlClient

Public Class Settings
    ReadOnly _l As New Logger
    ReadOnly _findDb As New DbSettings

    Public SurveySettings As New List(Of String)

    Public Function ReadSettings()
        Const query As String = "select DAYSBETWEENSURVEYS, MONSBETWEENEMAILS, " & _
                                "NUMCUSTTOEMAIL, ADMINEMAILADDRESS, " & _
                                "DATELASTEMAILED " & _
                                "from dbo.ZEMG_SURVEYSETTINGS " & _
                                "where SETTING = 1"

        Try
            Dim con As SqlConnection = New SqlConnection(_findDb.ConString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                If reader.HasRows Then
                    Do While reader.Read()
                        SurveySettings.Add(reader.GetInt64(0))  ' 0 DAYSBETWEENSURVEYS
                        SurveySettings.Add(reader.GetInt64(1))  ' 1 MONSBETWEENEMAILS
                        SurveySettings.Add(reader.GetInt64(2))  ' 2 NUMCUSTTOEMAIL
                        SurveySettings.Add(reader.GetString(3)) ' 3 ADMINEMAILADDRESS

                        Dim tempDate As DateTime = _
                            New DateTime(1988, 1, 1).AddMinutes(reader.GetInt64(4))
                        SurveySettings.Add(tempDate.ToString()) ' 4 DATELASTEMAILED
                    Loop
                End If
                con.Close()
                _l.Log("Database connection succeeded", _
           "Retrieved a list of survey settings: " & vbCrLf & _
           "Days between surveys: " & SurveySettings(0).ToString() & vbCrLf & _
           "Months between emailing the same customer: " & SurveySettings(1).ToString() & vbCrLf & _
           "Number of customers to email: " & SurveySettings(2).ToString() & vbCrLf & _
           "Admin email address: " & SurveySettings(3).ToString() & vbCrLf & _
           "Date the last survey was sent: " & SurveySettings(4).ToString())
            Catch ex As Exception
                con.Close()
                _l.Log("Database query failed", _
                "Retrieved no data from the database." & vbCrLf & _
                "Error details: " & ex.ToString())
            End Try
            Return True
        Catch ex As Exception
            _l.Log("Database connection failed", _
           "Failed to connect to the database to retrieve survey settings. " & _
          ". Error details: " & vbCrLf & _
           ex.ToString(), _
           True, _
           LogRecipients)
            Return False
        End Try
    End Function

    Public Sub UpdateEmailedDate()
        ' Priority SQL.DATE
        Dim dateNow As Int64 = _
            (DateTime.UtcNow.Date - New DateTime(1988, 1, 1)).TotalMinutes

        Dim query As String = "UPDATE dbo.ZEMG_SURVEYSETTINGS " & _
                                "SET DATELASTEMAILED = " & dateNow.ToString() & _
                                " where SETTING = 1;"

        Try
            Dim con As SqlConnection = New SqlConnection(_findDb.ConString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            cmd.ExecuteNonQuery()
            _l.Log("Database connection succeeded", _
                   "Updated date last emailed to today.")
        Catch ex As Exception
            _l.Log("Database connection failed", _
           "Failed to connect to the database to retrieve survey settings. " & _
          ". Error details: " & vbCrLf & _
           ex.ToString(), _
           True, _
           LogRecipients)
        End Try
    End Sub
End Class
