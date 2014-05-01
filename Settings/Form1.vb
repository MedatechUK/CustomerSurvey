Imports System.Xml
Imports System.Xml.Linq
Imports System.IO

Public Class Form1
    'TODO: Make prettier
    Dim settingsFile As String = "C:\emerge\survey\settings.xml"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim doc As New XDocument
        grdSettings.AutoGenerateColumns = False

        If Not File.Exists(settingsFile) Then
            Dim root As XElement = _
                <settings>
                </settings>
            doc.Add(root)

            Dim frequency As Integer = 14
            Dim count As Integer = 10
            Dim months As Integer = 6
            Dim email As String = "info@emerge-it.co.uk"

            doc.Root.Add(New XElement("DaysBetweenSurveyRuns"), frequency)
            doc.Root.Add(New XElement("TimeBetweenEmailingCustomerInMonths"), months)
            doc.Root.Add(New XElement("NumberCustomersToEmail"), count)
            doc.Root.Add(New XElement("AdminEmailAddress"), email)
            doc.Save(settingsFile)

            grdSettings.Rows.Add("DaysBetweenSurveyRuns", frequency)
            grdSettings.Rows.Add("NumberCustomersToEmail", count)
        Else
            doc = XDocument.Load(settingsFile)

            Dim query = From x In doc.Root.Elements
                        Select x.Name, x.Value

            Dim y = query.ToList
            Dim dictSettings As New List(Of String())
            For Each thing In y
                dictSettings.Add({thing.Name.ToString, thing.Value})
            Next

            For Each s As String() In dictSettings
                grdSettings.Rows.Add(s(0), s(1))
            Next
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim doc As New XDocument

        Dim root As XElement = _
            <settings>
            </settings>
        doc.Add(root)

        For Each row As DataGridViewRow In grdSettings.Rows
            If Not row.IsNewRow Then
                doc.Root.Add(New XElement(row.Cells(0).Value.ToString, _
                                          row.Cells(1).Value.ToString))
            End If
        Next

        doc.Save(settingsFile)
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

End Class
