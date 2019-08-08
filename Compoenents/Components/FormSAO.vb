﻿Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Imports DJLib
Public Class FormSAO
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)


    Dim ds As DataSet

    Dim BS As BindingSource
    Dim SoldToPartyBS As BindingSource
    Dim ShipToPartyBS As BindingSource
    Dim VendorBS As BindingSource

    Dim startdate As DateTime
    Dim enddate As DateTime
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim myfilter As String = String.Empty
    Dim mybasefolder As String
    Private Sub FormBrowseFolder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

        Dim sqlstr = "select soldtoparty,shiptoparty,saoname,saost,saofg,saocp,saofgstartdate,saocpstartdate,c.customername::character varying as soldtopartyname,c1.customername::character varying as shiptopartyname,saoopltid from saooplt left join customer c on c.customercode = soldtoparty left join customer c1 on c1.customercode = shiptoparty order by soldtoparty;" &
                     "select customercode::bigint as soldtoparty ,customername::character varying as soldtopartyname from  customer where not customername isnull order by customercode;" & _
                     "select customercode::bigint as shiptoparty ,customername::character varying as shiptopartyname from  customer where not customername isnull order by customercode;"


        ds = New DataSet
        BS = New BindingSource
        SoldToPartyBS = New BindingSource
        ShipToPartyBS = New BindingSource


        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            ds.Tables(0).TableName = "SAO"
            Dim idx0(0) As DataColumn
            idx0(0) = ds.Tables(0).Columns("saoopltid")
            ds.Tables(0).PrimaryKey = idx0

            ds.Tables(0).Columns("saoopltid").AutoIncrement = True
            ds.Tables(0).Columns("saoopltid").AutoIncrementSeed = 0
            ds.Tables(0).Columns("saoopltid").AutoIncrementStep = -1

            BS.DataSource = ds.Tables(0)
            SoldToPartyBS.DataSource = ds.Tables(1)
            ShipToPartyBS.DataSource = ds.Tables(2)

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
                    'BS.DataSource = ds.Tables(0)
                    DataGridView1.AutoGenerateColumns = False
                    With DirectCast(DataGridView1.Columns("Columnsoldtoparty"), DataGridViewComboBoxColumn)
                        .DataSource = SoldToPartyBS
                        .DisplayMember = "soldtoparty"
                        .ValueMember = "soldtoparty"
                    End With
                    With DirectCast(DataGridView1.Columns("columnShiptoparty"), DataGridViewComboBoxColumn)
                        .DataSource = ShipToPartyBS
                        .DisplayMember = "shiptoparty"
                        .ValueMember = "shiptoparty"
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
            If DbAdapter1.SAOOPLT(Me, mye) Then
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
        dr.Item("soldtoparty") = DBNull.Value
        dr.Item("shiptoparty") = DBNull.Value

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
        Select Case e.ColumnIndex
            Case 0
                Dim bs As New BindingSource
                bs = DirectCast(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell).DataSource
                Dim dr As DataRowView = bs.Current
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = dr.Row.Item("soldtopartyname")
            Case 2
                Dim bs As New BindingSource
                bs = DirectCast(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell).DataSource
                Dim dr As DataRowView = bs.Current
                DataGridView1.Rows(e.RowIndex).Cells(3).Value = dr.Row.Item("shiptopartyname")
        End Select
    End Sub

    Private Sub DataGridView1_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValidated

    End Sub

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        Me.Validate()       
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
        Dim myform As New FormImportSAO
        myform.Show()
        Me.LoadData()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dim filename As String = "SAOOPLT-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim dbtools = New Dbtools
        dbtools.Userid = "admin"
        dbtools.Password = "admin"
        Dim sqlstr = "select soldtoparty,c.customername::character varying as soldtopartyname,shiptoparty,c1.customername::character varying as shiptopartyname,saoname,saost,saofg,saofgstartdate,saocp,saocpstartdate from saooplt left join customer c on c.customercode = soldtoparty left join customer c1 on c1.customercode = shiptoparty order by soldtoparty; "
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools)
    End Sub
End Class