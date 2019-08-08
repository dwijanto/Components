Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Imports DJLib
Public Class FormCustomerFlow
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)


    Dim ds As DataSet

    Dim BS As BindingSource
    Dim SOBS As BindingSource
    Dim SIBS As BindingSource

    Dim startdate As DateTime
    Dim enddate As DateTime
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim myfilter As String = String.Empty
    Dim mybasefolder As String
    Public myFlow As New List(Of FlowType)


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
            
            myThread.Start()
        End If
    End Sub

    Private Sub doLoad()
        ProgressReport(6, "Marquee")

        'Dim sqlstr = "select vp.vendorcode,v.vendorname::character varying,vp.buid, vp.spid,vendorbuspid from vp.vendorbusp vp" & _
        '        " left join vendor v on v.vendorcode = vp.vendorcode" &
        '        " order by vendorname; " &
        '        "select vendorcode::bigint ,vendorname::character varying from  vendor order by vendorname;" & _
        '        "select null as sbuid,''::character varying as sbuname union all (select sbuid::bigint as buid, sbuname::character varying from sbu where bu or sp or lg or pcmmf order by sbuname);" & _
        '        "select null as ofsebid,''::character varying as  officersebname union all (select ofsebid::bigint as spid,officersebname::character varying from officerseb where levelid = 3 and  parent <> ofsebid and isactive order by officersebname);"
        Dim sqlstr = "select cf.*,c1.customername::text as soldtopartyname,c2.customername::text  as shiptopartyname from customerflow cf" &
            " left join customer c1 on c1.customercode = cf.soldtoparty" &
            " left join customer c2 on c2.customercode = cf.shiptoparty order by soldtopartyname,shiptopartyname;" &
                "select customercode ,customername::character varying from  customer order by customercode;" &
                "select customercode ,customername::character varying from  customer order by customercode;"

        ds = New DataSet
        BS = New BindingSource

        SOBS = New BindingSource
        SIBS = New BindingSource

        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            BS.DataSource = ds.Tables(0)
            ds.Tables(0).TableName = "CustomerFlow"
            Dim idx0(0) As DataColumn
            idx0(0) = ds.Tables(0).Columns("id")
            ds.Tables(0).PrimaryKey = idx0

            ds.Tables(0).Columns("id").AutoIncrement = True
            ds.Tables(0).Columns("id").AutoIncrementSeed = 0
            ds.Tables(0).Columns("id").AutoIncrementStep = -1

            'Dim docemailhdidxU As UniqueConstraint = New UniqueConstraint(New DataColumn() {ds.Tables(0).Columns("docemailname")})
            'ds.Tables(0).Constraints.Add(docemailhdidxU)



            SOBS.DataSource = ds.Tables(1)
            SIBS.DataSource = ds.Tables(2)            

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


                    With DirectCast(DataGridView1.Columns("ColumnSoldToParty"), DataGridViewComboBoxColumn)
                        .DataSource = SOBS
                        .DisplayMember = "customercode"
                        .ValueMember = "customercode"
                    End With

                    With DirectCast(DataGridView1.Columns("ColumnShipToParty"), DataGridViewComboBoxColumn)
                        .DataSource = SIBS
                        .DisplayMember = "customercode"
                        .ValueMember = "customercode"
                    End With

                    With DirectCast(DataGridView1.Columns("ColumnFlow"), DataGridViewComboBoxColumn)
                        .DataSource = MyFlow
                        .DisplayMember = "flowid"
                        .ValueMember = "flowname"
                    End With
                    DataGridView1.DataSource = BS

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

        myFlow.Add(New FlowType With {.FlowId = "DI", .FlowName = "DI"})
        myFlow.Add(New FlowType With {.FlowId = "SIS", .FlowName = "SIS"})
        myFlow.Add(New FlowType With {.FlowId = "BRAZIL", .FlowName = "BRAZIL"})
        myFlow.Add(New FlowType With {.FlowId = "DOMESTIC", .FlowName = "DOMESTIC"})
        myFlow.Add(New FlowType With {.FlowId = "GSE", .FlowName = "GSE"})

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        BS.EndEdit()
        Dim ds2 = ds.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.CustomerFlowTx(Me, mye) Then
                MessageBox.Show(mye.message)
                Exit Sub
            End If
            ds.Merge(ds2)
            ds.AcceptChanges()
            DataGridView1.Invalidate()
            MessageBox.Show("Saved.")
        End If

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = BS.AddNew()
        Dim dr As DataRow = drv.Row
        dr.Item("shiptoparty") = DBNull.Value
        dr.Item("soldtoparty") = DBNull.Value
        dr.Item("flow") = DBNull.Value
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
            DataGridView1.Rows(e.RowIndex).Cells(1).Value = dr.Row.Item("customername")
        ElseIf e.ColumnIndex = 2 Then
            Dim bs As New BindingSource
            'MessageBox.Show(DirectCast(DataGridView1.Rows(0).Cells(0), DataGridViewComboBoxCell).Value)
            bs = DirectCast(DataGridView1.Rows(0).Cells(2), DataGridViewComboBoxCell).DataSource
            Dim dr As DataRowView = bs.Current
            DataGridView1.Rows(e.RowIndex).Cells(3).Value = dr.Row.Item("customername")
        End If

    End Sub

    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated

    End Sub

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Me.Validate()
        'If e.ColumnIndex = 2 Then
        '    If CType(DataGridView1.Rows(e.RowIndex).Cells(2), DataGridViewComboBoxCell).EditedFormattedValue = "" Then
        '        Dim myrow = CType(BS.Current, DataRowView).Row
        '        myrow.Item("buid") = DBNull.Value
        '    End If
        'ElseIf e.ColumnIndex = 3 Then
        '    If CType(DataGridView1.Rows(e.RowIndex).Cells(3), DataGridViewComboBoxCell).EditedFormattedValue = "" Then
        '        Dim myrow = CType(BS.Current, DataRowView).Row
        '        myrow.Item("spid") = DBNull.Value
        '    End If

        'End If
    End Sub




    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        BS.CancelEdit()
    End Sub


    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim myform As New FormImportCustomerFlow
        myform.ShowDialog()
        Me.LoadData()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dim filename As String = "CustomerFlow-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim dbtools = New Dbtools
        dbtools.Userid = "admin"
        dbtools.Password = "admin"
        Dim sqlstr = "select cf.soldtoparty,c1.customername::text as soldtopartyname,cf.shiptoparty,c2.customername::text  as shiptopartyname,cf.flow,cf.dicustomer,cf.continent,cf.continent_group,cf.continent_group_emea from customerflow cf" &
            " left join customer c1 on c1.customercode = cf.soldtoparty" &
            " left join customer c2 on c2.customercode = cf.shiptoparty" &
            " order by soldtopartyname,shiptopartyname"
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools)
    End Sub


End Class

Public Class FlowType
    Public Property FlowId As String
    Public Property FlowName As String
    Public Overrides Function ToString() As String
        Return FlowName
    End Function
End Class