Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Imports DJLib
Public Class FormSPManager
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)

    Protected CM As CurrencyManager

    Dim Dataset As DataSet
    Dim WithEvents Dataset2 As New DataSet
    Dim WithEvents DataTable2 As DataTable
    Dim sqlstr As String = String.Empty

    Dim WithEvents smbs As New BindingSource
    Dim DS As DataSet
    Dim bs As BindingSource

    Private Sub FormBrowseFolder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        Me.Validate()
        Dim ds2 = DS.GetChanges()
        If Not IsNothing(ds2) Then
            If MessageBox.Show("There is unsaved data in a row. Do you want to store to the database?", "Save Changes", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                ToolStripButton4.PerformClick()
            End If
        End If

    End Sub

    Private Sub FormBrowseFolder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton5.Click

        LoadData()

    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf doLoad)           
            myThread.Start()
        End If
    End Sub

    Private Sub doLoad()
        ProgressReport(6, "Marquee")

        sqlstr = "select null as smid,''::character varying as sm union all select 0 as smid,''::character as sm union all select o.ofsebid::bigint as smid,o.officersebname::character varying as sm from officerseb o" & _
                 " left join teamtitle tt on tt.teamtitleid = o.teamtitleid" & _
                 " where levelid = 3 and isactive and (teamtitleshortname = 'SM' or teamtitleshortname = 'SCM') order by sm"
        Dataset2 = New DataSet
        BS = New BindingSource

        
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, Dataset2, mymessage) Then
            DataTable2 = Dataset2.Tables(0)
            smbs.DataSource = DataTable2
            ProgressReport(1, "Assign DataGridView DataSource")
        Else
            ProgressReport(2, mymessage)
        End If
        Dataset = New DataSet
        FillData()


        ProgressReport(5, "Continuous")
    End Sub
    Public Overridable Sub FillData()
        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.spmanager(Dataset, message, ra) Then
            Dataset.Tables(0).TableName = "spmanager"
            ProgressReport(4, "Fill Datagrid")
        Else
            MessageBox.Show(message)
        End If


    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                   

                    Dim ColSM As DataGridViewComboBoxColumn = DataGridView1.Columns("CSupplyManager")
                    With ColSM                       
                        .DataSource = smbs 'Dataset2.Tables(0)
                        .DisplayMember = "sm"
                        .ValueMember = "smid"                       
                    End With

                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 4
                    bs.DataSource = Dataset.Tables("spmanager")
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = bs
                    CM = CType(BindingContext(bs), CurrencyManager)
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



    Sub doQuery()

    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

      
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        BS.EndEdit()
        Dim ds2 = Dataset.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            'Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.spmanager(ds2, mymessage, ra) Then
                'DS.Merge(ds2)
                'DS.AcceptChanges()
                'Else
                MessageBox.Show(mymessage)
                Exit Sub
            End If
            Dataset.Merge(ds2)
            Dataset.AcceptChanges()
            MessageBox.Show("Saved.")
        End If

    End Sub

    'Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
    '    Dim drv As DataRowView = BS.AddNew()
    '    Dim dr As DataRow = drv.Row
    '    dr.Item("vendorcode") = DBNull.Value
    '    dr.Item("buid") = DBNull.Value
    '    dr.Item("spid") = DBNull.Value
    'End Sub

    'Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

    '    If Not IsNothing(BS.Current) Then
    '        If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
    '            'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
    '            For Each dsrow In DataGridView1.SelectedRows
    '                BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
    '            Next
    '        End If
    '    Else
    '        MessageBox.Show("No record to delete.")
    '    End If
    'End Sub

    'Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
    '    'MessageBox.Show("endedit")
    '    If e.ColumnIndex = 0 Then
    '        Dim bs As New BindingSource
    '        'MessageBox.Show(DirectCast(DataGridView1.Rows(0).Cells(0), DataGridViewComboBoxCell).Value)
    '        bs = DirectCast(DataGridView1.Rows(0).Cells(0), DataGridViewComboBoxCell).DataSource
    '        Dim dr As DataRowView = bs.Current
    '        'MessageBox.Show(dr.Row.Item("customername"))
    '        DataGridView1.Rows(e.RowIndex).Cells(1).Value = dr.Row.Item("vendorname")
    '    End If

    'End Sub

    'Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated

    'End Sub

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Me.Validate()
        If e.ColumnIndex = 1 Then
            If CType(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell).EditedFormattedValue = "" Then
                Dim myrow = CType(bs.Current, DataRowView).Row
                myrow.Item("smid") = DBNull.Value
            End If
        ElseIf e.ColumnIndex = 3 Then
            'If CType(DataGridView1.Rows(e.RowIndex).Cells(3), DataGridViewComboBoxCell).EditedFormattedValue = "" Then
            '    Dim myrow = CType(BS.Current, DataRowView).Row
            '    myrow.Item("spid") = DBNull.Value
            'End If

        End If
    End Sub




    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        bs.CancelEdit()
    End Sub


    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim myform As New FormImportVendorBUSP
        myform.Show()
        Me.LoadData()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    'Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
    '    Dim filename As String = "VendorSP-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
    '    Dim dbtools = New Dbtools
    '    dbtools.Userid = "admin"
    '    dbtools.Password = "admin"
    '    Dim sqlstr = "select v.vendorcode,v.vendorname,s.sbuname as bu, o.officersebname as sp from vp.vendorbusp vp" & _
    '             " left join vendor v on v.vendorcode = vp.vendorcode" & _
    '             " left join sbu s on s.sbuid = vp.buid" & _
    '             " left join officerseb o on o.ofsebid = vp.spid" & _
    '             " order by vendorname,bu,sp "
    '    ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools)
    'End Sub

End Class