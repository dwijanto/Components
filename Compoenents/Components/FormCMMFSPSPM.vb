Imports System.Threading
Public Class FormCMMFSPSPM
    Dim mySelectedPath As String

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myController As CMMFSPSPMController
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim OpenFileDialog1 As New OpenFileDialog

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        'AddHandler FormCategoryIssue.myRefresh, AddressOf loadData
        'AddHandler FormManagementMessage.myRefresh, AddressOf loadData
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myController = New CMMFSPSPMController(Me)
        Try
            ProgressReport(1, "Loading..")
            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If

            ProgressReport(1, String.Format("Done. Records Count({0}).", myController.BS.Count))
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select
        End If
    End Sub

    Private Sub FormCommentMapping_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Public Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = myController.BS.AddNew()
        drv.row.item("pcmmf") = False
        drv.row.item("sp") = False
        drv.row.item("lg") = False
        drv.row.item("bu") = False
        drv.row.item("cp") = False
        drv.row.item("act") = False
        drv.EndEdit()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.loadData()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Me.Validate() Then
            If myController.Validate Then
                myController.save()
            End If

        End If

    End Sub

    'Public Overloads Function validate() As Boolean
    '    Dim drv As DataRowView = myController.BS.Current

    'End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(myController.BS.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
                For Each dsrow In DataGridView1.SelectedRows
                    myController.BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myform As New FormCategoryIssue
        myform.Show()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myform As New FormManagementMessage
        myform.Show()
    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        myController.ApplyFilter = ToolStripTextBox1.Text
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)

    End Sub

    Private Sub ToolStripButton6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        myController.Model.ExportToExcel()
    End Sub

    Private Sub ToolStripButton7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        ImportData()
    End Sub

    Private Sub ImportData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            Dim OpenFileDialog1 = New OpenFileDialog
            If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                myController.ImportFileName = OpenFileDialog1.FileName
                myThread = New Thread(AddressOf DoImport)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoImport()

        If myController.ImportData() Then
            DoWork()
        End If

    End Sub
End Class