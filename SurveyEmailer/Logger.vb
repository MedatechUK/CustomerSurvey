Imports System
Imports System.IO
Imports System.Text

Public Class Logger
    Public Property _logPath As String = "C:\emerge\survey\logs\"
    Public Property _logDir As String = _logPath & Date.Now.Year
    Public Property _logFile As String = _logDir & "\" & Date.Now.Month & "-" & MonthName(Date.Now.Month) & ".log"

    Public Sub Log(ByVal type As String, _
                   ByVal msg As String)
        CreateLog()
        Dim sb As New StringBuilder()
        With sb
            .Append(vbCrLf)
            .Append("[" & Date.Now.ToShortDateString)
            .Append(" " & Date.Now.ToShortTimeString & "] ")
            .Append(type & ": ")
            .Append(msg)
        End With
        Using sw As New StreamWriter(_logFile, True)
            sw.Write(sb.ToString())
        End Using
    End Sub

    Public Sub CreateLog()
        If Not Directory.Exists(_logDir) Then
            Directory.CreateDirectory(_logDir)
        End If
        If Not File.Exists(_logFile) Then
            File.Create(_logFile).Close()
        End If
    End Sub
End Class
