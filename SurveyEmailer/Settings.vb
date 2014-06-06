Imports System.Data.SqlClient

Public Class Settings
    Public Shared DaysBetweenSurveys As Int64
    Public Shared MonthsBetweenEmails As Int64
    Public Shared NumCustToEmail As Int64
    Public Shared AdminEmailAddress As String
    Public Shared DateLastEmailed As String
    Public Shared IgnoreSurveyFlag As Boolean
    Public Shared NumDaysInPast As Int64
    Public Shared LogEmailAddress As String

    Public Shared Function ReadSettings()
        Const query As String = "select DAYSBETWEENSURVEYS, MONSBETWEENEMAILS, " & _
                                "NUMCUSTTOEMAIL, ADMINEMAILADDRESS, " & _
                                "DATELASTEMAILED, IGNORESURVEYFLAG, " & _
                                "NUMDAYSINPAST, LOGEMAILADDRESS " & _
                                "from dbo.ZEMG_SURVEYSETTINGS " & _
                                "where SETTING = 1"

        Try
            Dim con As SqlConnection = New SqlConnection(DbSettings.ConString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Try
                If reader.HasRows Then
                    Do While reader.Read()
                        LogEmailAddress = reader.GetString(7)
                        DaysBetweenSurveys = reader.GetInt64(0)
                        MonthsBetweenEmails = reader.GetInt64(1)
                        NumCustToEmail = reader.GetInt64(2)
                        AdminEmailAddress = reader.GetString(3)

                        Dim tempDate As DateTime = _
                            New DateTime(1988, 1, 1).AddMinutes(reader.GetInt64(4))
                        DateLastEmailed = tempDate.ToString()

                        If reader.GetString(5) <> "Y" Then
                            IgnoreSurveyFlag = False
                        Else
                            IgnoreSurveyFlag = True
                        End If

                        NumDaysInPast = reader.GetInt64(6)
                    Loop
                End If
                con.Close()
                Logger.Log("Database connection succeeded", _
           "Retrieved a list of survey settings: " & vbCrLf & _
           "Days between surveys: " & DaysBetweenSurveys.ToString() & vbCrLf & _
           "Months between emailing the same customer: " & MonthsBetweenEmails.ToString() & vbCrLf & _
           "Number of customers to email: " & NumCustToEmail.ToString() & vbCrLf & _
           "Admin email address: " & AdminEmailAddress & vbCrLf & _
           "Date the last survey was sent: " & DateLastEmailed)
            Catch ex As Exception
                con.Close()
                Logger.Log("Database query failed", _
                "Retrieved no data from the database." & vbCrLf & _
                "Error details: " & ex.ToString())
                Return False
            End Try
            Return True
        Catch ex As Exception
            Logger.Log("Database connection failed", _
           "Failed to connect to the database to retrieve survey settings. " & _
          ". Error details: " & vbCrLf & _
           ex.ToString(), _
           True, _
           LogEmailAddress)
            Return False
        End Try
    End Function

    Public Shared Sub UpdateEmailedDate()
        ' Priority SQL.DATE
        Dim dateNow As Int64 = _
            (DateTime.UtcNow.Date - New DateTime(1988, 1, 1)).TotalMinutes

        Dim query As String = "UPDATE dbo.ZEMG_SURVEYSETTINGS " & _
                                "SET DATELASTEMAILED = " & dateNow.ToString() & _
                                " where SETTING = 1;"

        Try
            Dim con As SqlConnection = New SqlConnection(DbSettings.ConString)
            con.Open()
            Dim cmd As New SqlCommand(query, con)

            cmd.ExecuteNonQuery()
            Logger.Log("Database connection succeeded", _
                   "Updated date last emailed to today.")
        Catch ex As Exception
            Logger.Log("Database connection failed", _
           "Failed to connect to the database to retrieve survey settings. " & _
          ". Error details: " & vbCrLf & _
           ex.ToString(), _
           True, _
           LogEmailAddress)
        End Try
    End Sub

    Public Shared Function SurveysSentRecently()
        If DateLastEmailed <> "" Then
            Dim diff As TimeSpan = Date.Now.Subtract(DateLastEmailed)
            If diff.Days < DaysBetweenSurveys Then
                Logger.Log("Settings", _
                      "It has been " & diff.Days & " days since the last run, which is less than " & _
                      "the " & DaysBetweenSurveys & " days defined in the settings." & vbCrLf & _
                      "Closing the program.")
                Return True
            Else : Return False
            End If
        Else : Return False
        End If
    End Function
End Class
