Imports System.Threading
Imports System.Text
Public Enum TXRecord
    AddRecord = 0
    UpdateRecord = 1
    CancelRecord = 2
    DeleteRecord = 3
    ViewRecord = 4
End Enum
Public Class FormSAOAllocation
    Dim myAdapter As SAOAllocationAdapter
    Dim myThread As New Thread(New ThreadStart(AddressOf DoWork))
    Dim myFields() = {"customercode", "customername", "pol", "userid", "customertype"}

    Private Sub FormSAOAllocation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton5.Click
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf DoWork)
            myThread.SetApartmentState(ApartmentState.MTA)
            myThread.Start()
        End If
    End Sub

    Sub DoWork()
        myAdapter = New SAOAllocationAdapter
        Try
            ProgressReport(5, "Marque")
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
        Catch ex As Exception

        End Try
        ProgressReport(6, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1

                Case 4
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myAdapter.BS
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
            End Select
        End If
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ToolStripComboBox1.ComboBox.SelectedIndex = 3
        myAdapter = New SAOAllocationAdapter
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        myAdapter.Save()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        ShowTX(TXRecord.AddRecord)
    End Sub

    Private Sub ShowTX(ByVal StatusTx As TXRecord)
        Dim drv As DataRowView = Nothing
        Select Case StatusTx
            Case tXRecord.AddRecord
                drv = myAdapter.BS.AddNew
            Case tXRecord.UpdateRecord
                drv = myAdapter.BS.Current
        End Select
        Dim customerbs As New BindingSource
        Dim dt As DataTable
        dt = myAdapter.DS.Tables(1).Copy
        customerbs.DataSource = dt
        Dim myform As New DialogSAOAllocationInput(drv, customerbs)

        If myform.ShowDialog = DialogResult.OK Then
            DataGridView1.Invalidate()
        End If

    End Sub

    'Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick, ToolStripButton6.Click
    '    ShowTX(TXRecord.UpdateRecord)
    'End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.CellDoubleClick, ToolStripButton6.Click
        ShowTX(TXRecord.UpdateRecord)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myAdapter.BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

 

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj As ToolStripTextBox
        obj = CType(sender, ToolStripTextBox)
        Dim myfilter As String = String.Empty
        If obj.Text <> "" Then
            myfilter = String.Format("[{0}] like '*{1}*'", myFields(ToolStripComboBox1.ComboBox.SelectedIndex), obj.Text)
        End If
        myAdapter.BS.Filter = myfilter
    End Sub



    Private Sub ToolStripButton3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        myAdapter.BS.Filter = ""
    End Sub


End Class

Public Class SAOAllocationAdapter
    Public Property BS As New BindingSource
    Public Property DS As DataSet
    Public Property dbAdapter1 As DbAdapter
    Private SB As New StringBuilder

    Public Sub New()
        dbAdapter1 = DbAdapter.getInstance
    End Sub

    Public Function LoadData()
        DS = New DataSet
        BS = New BindingSource
        SB.Append("select sa.id,sa.customercode::character varying,sa.ctype,sa.userid,sa.pol,c.customername,getcustomertypename(sa.ctype) as customertype from saoallocation sa left join customer c on c.customercode = sa.customercode;")
        SB.Append("select customercode,customername from customer order by customercode;")
        Try
            dbAdapter1.TbgetDataSet(SB.ToString, DS)
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk


            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
            DS.Tables(0).Columns("customercode").AllowDBNull = False
            BS.DataSource = DS.Tables(0)
        Catch ex As Exception
            MessageBox.Show(String.Format("Error Found : {0}"), ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function Save() As Boolean
        Dim myret As Boolean = False
        Try
            BS.EndEdit()
            Dim ra As Integer
            Dim ds2 As DataSet = DS.GetChanges
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                Try
                    If dbAdapter1.SAOAllocation(Me, mye) Then
                        DS.Merge(ds2)
                        DS.AcceptChanges()
                        MessageBox.Show("Saved.")
                        myret = True
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    DS.Merge(ds2)
                End Try
            End If

        Catch ex As Exception
            myret = False
            MessageBox.Show(ex.Message)
        End Try
        

        Return myret
    End Function

End Class