Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Public Class FormMarketContacts
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)


    Dim ds As DataSet

    Dim BS As BindingSource
    Dim Customerbs As BindingSource
    Dim Brandbs As BindingSource
    Dim startdate As DateTime
    Dim enddate As DateTime
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim myfilter As String = String.Empty
    Dim mybasefolder As String
    Private Sub FormBrowseFolder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        Me.Validate()
        Dim ds2 = ds.GetChanges()
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
            startdate = startdateDTP.Value.Date
            enddate = enddateDTP.Value.Date.AddDays(1)
            myThread.Start()
        End If
    End Sub

    Private Sub doLoad()
        ProgressReport(6, "Marquee")

        Dim sqlstr As String = "select m.*,c.customername::character varying,b.brandname::character varying from marketemail m" &
                            " left join customer c on c.customercode = m.customercode" &
                            " left join brand b on b.brandid = m.brandid order by customercode;" &
                            " select null as brandid,''::character varying as brandname union all (select brandid,brandname::character varying from brand order by brandname);" &
                            " select customercode,customername::character varying from customer order by customercode;"
        ds = New DataSet
        BS = New BindingSource

        customerbs = New BindingSource
        Brandbs = New BindingSource

        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            ds.Tables(0).TableName = "MarketEmail"
            Dim idx0(0) As DataColumn
            idx0(0) = ds.Tables(0).Columns("marketemailid")
            ds.Tables(0).PrimaryKey = idx0

            ds.Tables(0).Columns(0).AutoIncrement = True
            ds.Tables(0).Columns(0).AutoIncrementSeed = 0
            ds.Tables(0).Columns(0).AutoIncrementStep = -1

            'Dim docemailhdidxU As UniqueConstraint = New UniqueConstraint(New DataColumn() {ds.Tables(0).Columns("docemailname")})
            'ds.Tables(0).Constraints.Add(docemailhdidxU)

            BS.DataSource = ds.Tables(0)

            Brandbs.DataSource = ds.Tables(1)
            Customerbs.DataSource = ds.Tables(2)

            ProgressReport(1, "Assign DataGridView DataSource")
        Else
            ProgressReport(2, mymessage)
        End If
        ProgressReport(5, "Continuous")
    End Sub


    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    BS.DataSource = ds.Tables(0)
                    DataGridView1.AutoGenerateColumns = False


                    With DirectCast(DataGridView1.Columns("ColumnCustomerCode"), DataGridViewComboBoxColumn)
                        .DataSource = Customerbs
                        .DisplayMember = "customercode"
                        .ValueMember = "customercode"
                    End With

                    With DirectCast(DataGridView1.Columns("ColumnBrandName"), DataGridViewComboBoxColumn)
                        .DataSource = Brandbs
                        .DisplayMember = "brandname"
                        .ValueMember = "brandid"
                    End With

                    DataGridView1.DataSource = BS
                    'ToolStripComboBox1.ComboBox.DisplayMember = "foldername"
                    'ToolStripComboBox1.ComboBox.DataSource = cbbs
                    'ToolStripComboBox1.ComboBox.Text = myfilter
                Case 2
                    ToolStripStatusLabel1.Text = message
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

        ' Add any initialization after the InitializeComponent() call.
        With startdateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        With enddateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        Dim host1 = New ToolStripControlHost(startdateDTP)
        Dim host2 = New ToolStripControlHost(enddateDTP)
        ' ToolStrip1.Items.Insert(9, host1)
        ' ToolStrip1.Items.Insert(11, host2)
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        BS.EndEdit()
        Dim ds2 = ds.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.MarketEmailTx(Me, mye) Then
                'delete the modfied row for Merged
                Dim modifiedRows = From row In ds.Tables(0)
                   Where row.RowState = DataRowState.Added
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
            Else
                MessageBox.Show(mye.message)
                Exit Sub
            End If
            ds.Merge(ds2)
            ds.AcceptChanges()
            MessageBox.Show("Saved.")
        End If

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = BS.AddNew()
        Dim dr As DataRow = drv.Row
        'dr.Item("docemailtype") = 0
        'dr.Item("receiveddate") = Today
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click

        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
                For Each dsrow In DataGridView1.SelectedRows
                    BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub DataGridView1_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        'MessageBox.Show("endedit")
        If e.ColumnIndex = 0 Then
            Dim bs As New BindingSource
            'MessageBox.Show(DirectCast(DataGridView1.Rows(0).Cells(0), DataGridViewComboBoxCell).Value)
            bs = DirectCast(DataGridView1.Rows(0).Cells(0), DataGridViewComboBoxCell).DataSource
            Dim dr As DataRowView = bs.Current
            'MessageBox.Show(dr.Row.Item("customername"))
            DataGridView1.Rows(e.RowIndex).Cells(1).Value = dr.Row.Item("customername")
        End If
        
    End Sub

    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated

    End Sub

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Me.Validate()
        If e.ColumnIndex = 2 Then
            If CType(DataGridView1.Rows(e.RowIndex).Cells(2), DataGridViewComboBoxCell).EditedFormattedValue = "" Then
                Dim myrow = CType(BS.Current, DataRowView).Row
                myrow.Item("brandid") = DBNull.Value
            End If

        End If
    End Sub


   

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        BS.CancelEdit()
    End Sub


    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

    End Sub
End Class