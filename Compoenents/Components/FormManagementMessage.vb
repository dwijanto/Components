Imports System.Threading
Public Class FormManagementMessage
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myadapter As ManagementMessageAdapter
    Public Shared Event myRefresh()
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myadapter = New ManagementMessageAdapter
        Try
            ProgressReport(1, "Loading..")
            If myadapter.loaddata() Then
                ProgressReport(4, "Init Data")
            End If


            ProgressReport(1, "Done.")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4

                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myadapter.BS

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
                    drv = myadapter.BS.AddNew
                Case TXRecord.UpdateRecord
                    drv = myadapter.BS.Current
            End Select

            Dim myform = New DialogAddManagementMessage(drv)
            If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
                DataGridView1.Invalidate()
            End If
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.loadData()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If myadapter.save() Then
            RaiseEvent myRefresh()
        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(myadapter.BS.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
                For Each dsrow In DataGridView1.SelectedRows
                    myadapter.BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click, DataGridView1.CellDoubleClick
        showTx(TXRecord.UpdateRecord)
    End Sub
End Class