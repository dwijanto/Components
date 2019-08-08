Imports System.Threading
Imports System.Text

Public Class FormImportOPLTComments
    Dim mythread As New Thread(AddressOf doWork)
    Dim errmsg As New StringBuilder
    Dim selectedfile As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not mythread.IsAlive Then
            'Get file
            errmsg = New StringBuilder
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                selectedfile = openfiledialog1.FileName
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub doWork()
        Dim ImportOPLTComments = New ImportOPLTComments(Me, OpenFileDialog1.FileName)
        ProgressReport(1, "Processing. Please wait..")
        ProgressReport(2, "Marque")
        If ImportOPLTComments.ValidateFile Then
            ProgressReport(1, "Do Import file..")
            If ImportOPLTComments.DoImportFile Then
                'Thread.Sleep(5000)
                ProgressReport(1, "Done.")
            Else
                ProgressReport(1, String.Format("Error::{0}", ImportOPLTComments.ErrorMsg))
            End If
        Else
            ProgressReport(1, String.Format("Error::{0}", ImportOPLTComments.ErrorMsg))

        End If
        ProgressReport(3, "Continuous")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
            End Select

        End If

    End Sub
End Class

