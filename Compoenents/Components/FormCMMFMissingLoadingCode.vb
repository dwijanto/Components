Imports System.Threading
Imports Microsoft.Office.Interop
Imports System.Text
Public Class FormCMMFMissingLoadingCode

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty


        Dim sqlstr As String = "select cmmf from cmmf where loadingcode isnull"

        Dim SaveFileDialog1 As New SaveFileDialog


        SaveFileDialog1.FileName = "CMMFMissingLoadingCode.xlsx"
        If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim mypath As String = System.IO.Path.GetDirectoryName(SaveFileDialog1.FileName)

            Dim reportname = "CMMFMissingLoadingCode"
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim datasheet As Integer = 3
            Dim myreport As New ExportToExcelFile(Me, sqlstr, mypath, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException       
    End Sub


End Class