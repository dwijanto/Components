Imports System.Threading
Public Class FormActivity
    Dim mySelectedPath As String

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)

    Dim myController As ActivityController
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim OpenFileDialog1 As New OpenFileDialog





    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        AddHandler DialogActivity.RefreshData, AddressOf RefreshDataGrid
        'AddHandler FormManagementMessage.myRefresh, AddressOf loadData
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myController = New ActivityController

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
        loadData()
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
        showTx(TXRecord.AddRecord)
    End Sub

    Public Sub showTx(ByVal tx As TXRecord)
        If Not myThread.IsAlive Then
            Dim drv As DataRowView = Nothing
            Select Case tx
                Case TXRecord.AddRecord
                    drv = myController.BS.AddNew
                Case TXRecord.UpdateRecord
                    drv = myController.BS.Current
            End Select

            Dim myform = New DialogActivity(drv)
            If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                DataGridView1.Invalidate()
            End If
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.loadData()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        myController.save()
    End Sub

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

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        showTx(TXRecord.UpdateRecord)
    End Sub


    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        myController.ApplyFilter = ToolStripTextBox1.Text
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs)

    End Sub

    Private Sub ToolStripButton6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        myController.Model.ExportToExcel()
    End Sub


    Public Sub RefreshDataGrid()
        DataGridView1.Invalidate()
    End Sub



    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        showTx(TXRecord.UpdateRecord)
    End Sub

End Class