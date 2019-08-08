Imports Components.HelperClass
Imports Components.PublicClass
Imports DJLib
Public Class FormVendorSP1
    Protected CM As CurrencyManager
    Protected mypanel1 As UCSortTx
    Protected mypanel As UCFilterTx
    Dim Dataset As DataSet
    Dim Dataset2 As New DataSet
    Dim sqlstr As String = String.Empty

    Dim vendorbs As New BindingSource
    Dim bubs As New BindingSource
    Dim spbs As New BindingSource
    Dim col2 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim col3 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim col4 As New System.Windows.Forms.DataGridViewTextBoxColumn
    Dim col5 As New System.Windows.Forms.DataGridViewTextBoxColumn

    Dim sortcol As New Dictionary(Of Integer, System.Windows.Forms.DataGridViewTextBoxColumn)

    Private Sub FormVendorSP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()

        sortcol.Add(0, col3)
        sortcol.Add(1, col4)
        sortcol.Add(2, col5)

        Dim message As String = String.Empty
        'AddDateTimePickerToBindingNavigator()
        sqlstr = "select suppliercode::bigint as vendorcode,vendorname::character varying from vp.supplier vp left join vendor v on v.vendorcode = vp.suppliercode::bigint order by v.vendorname;" & _
                "select sbuid::bigint as buid, sbuname::character varying from sbu where bu or sp or lg or pcmmf order by sbuname;" & _
                "select ofsebid::bigint as spid,officersebname::character varying from officerseb where levelid = 3 and  parent <> ofsebid and isactive order by officersebname;"
        '"select ofsebid::bigint as spid,officersebname::character varying from officerseb where levelid = 3 and  not parent is null and isactive order by officersebname;"
        If Not DbAdapter1.TbgetDataSet(sqlstr, Dataset2, message) Then
            MessageBox.Show(message)
        End If
        vendorbs.DataSource = Dataset2.Tables(0)
        bubs.DataSource = Dataset2.Tables(1)
        spbs.DataSource = Dataset2.Tables(2)

        LoadData()
        LoadToolstrip()
        ToolStrip1.Visible = False


    End Sub
    Public Sub LoadData()

        InitObject()
        FillData()
        BindDataSource()
        BindingObject()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub
    Public Overridable Sub InitObject()
        InitDataGrid()
        BindingSource1 = New BindingSource

        Dataset = New DataSet
        With DataGridView1
            .DataSource = BindingSource1
            .RowsDefaultCellStyle.BackColor = System.Drawing.Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke
        End With
        'DataGridView1.AutoGenerateColumns = False
        BindingNavigator1.BindingSource = BindingSource1


    End Sub
    Public Overridable Sub FillData()
        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.VendorBUSP1(Dataset, message, ra) Then
            Dataset.Tables(0).TableName = "vendorbusp"
        Else
            MessageBox.Show(message)
        End If


    End Sub
    Public Overridable Sub BindDataSource()
        BindingSource1.DataSource = Dataset.Tables("vendorbusp")
        DataGridView1.DataSource = BindingSource1
    End Sub

    Public Overridable Sub BindingObject()
        DataGridView1.Columns.Clear()
        'DataGridView1.AutoGenerateColumns = False
        'DataGridView1.DataSource = BindingSource1

        Dim ColVendor = New DataGridViewComboBoxColumn
        With ColVendor
            .HeaderText = "Vendor Name"
            .DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
            .FlatStyle = FlatStyle.Flat
            .MaxDropDownItems = 10
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            '.DropDownWidth = 300
            .Width = 300
            .DataPropertyName = "vendorcode"
            .DataSource = vendorbs 'Dataset2.Tables(0)
            .DisplayMember = "vendorname"
            .ValueMember = "vendorcode"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            .Tag = "vendorname"
        End With

        Dim Colbu As New System.Windows.Forms.DataGridViewComboBoxColumn
        With Colbu
            .HeaderText = "BU"
            .DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
            .FlatStyle = FlatStyle.Flat
            .MaxDropDownItems = 10
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 300
            .DataPropertyName = "buid"
            .DataSource = bubs 'Dataset2.Tables(1)
            .ValueMember = "buid"
            .DisplayMember = "sbuname"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .Tag = "bu"
        End With

        Dim Colsp As New System.Windows.Forms.DataGridViewComboBoxColumn
        With Colsp
            .HeaderText = "Supply Planner"
            .DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.[Nothing]
            .FlatStyle = FlatStyle.Flat
            .MaxDropDownItems = 10
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
            .SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 300
            .DataPropertyName = "spid"
            .DataSource = spbs 'Dataset2.Tables(2)
            .ValueMember = "spid"
            .DisplayMember = "officersebname"
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            .Tag = "sp"
        End With

        With col2
            .HeaderText = "Vendor Code"
            .DataPropertyName = "vendorcode"
            .Name = "DataGridViewTextBoxColumn1"
            .Visible = True
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 82
        End With

        With col3
            .HeaderText = "Vendor Name"
            .DataPropertyName = "vendorname"
            .Name = "DataGridViewTextBoxColumn1"
            .Visible = True
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 82
        End With
        With col4
            .HeaderText = "bu"
            .DataPropertyName = "bu"
            .Name = "DataGridViewTextBoxColumn1"
            .Visible = True
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 82
        End With
        With col5
            .HeaderText = "sp"
            .DataPropertyName = "sp"
            .Name = "DataGridViewTextBoxColumn1"
            .Visible = True
            '.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
            .Width = 82
        End With
        With DataGridView1
            '.Columns.Insert(0, ColVendor)
            '.Columns.Insert(1, Colbu)
            '.Columns.Insert(2, Colsp)
            .Columns.Insert(0, col2)
            .Columns.Insert(1, col3)
            .Columns.Insert(2, col4)
            .Columns.Insert(3, col5)
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader 'DataGridViewAutoSizeColumnMode.AllCellsExceptHeader            
        End With

    End Sub
    Private Sub LoadToolstrip()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, DataGridView1, mypanel1, ToolStrip1, mypanel)
    End Sub

    Private Sub toolstripvisible(ByVal toolstripvisible As Boolean)
        ToolStrip1.Visible = Not (toolstripvisible)
        'Button3.Visible = toolstripvisible
    End Sub

    Public Overridable Sub InitDataGrid()

    End Sub


    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadData()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Call toolstripvisible(ToolStrip1.Visible)
    End Sub


    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        BindingSource1.AddNew()
    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                If DataGridView1.SelectedRows.Count = 0 Then
                    BindingSource1.RemoveAt(CM.Position)
                Else
                    For Each a As DataGridViewRow In DataGridView1.SelectedRows
                        BindingSource1.RemoveAt(a.Index)
                    Next
                End If
                UpdateRecord()
            Catch ex As Exception
            End Try
        End If
    End Sub
    Public Sub UpdateRecord()

        Dim ra As Integer
        Dim message As String = String.Empty
        Dim sb As New System.Text.StringBuilder
        Try
            BindingSource1.EndEdit()
            Dim ds2 = Dataset.GetChanges
            If Not IsNothing(ds2) Then

                If DbAdapter1.VendorBUSP(ds2, message, ra) Then
                    'sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    Dataset.Merge(ds2)
                    Dataset.AcceptChanges()
                End If
                If Dataset.HasErrors Then
                    sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh shortly.")
                    'sb.Append(message)
                    MessageBox.Show(sb.ToString)
                    LoadData()
                Else
                    If sb.ToString <> "" Then
                        MessageBox.Show(sb.ToString)
                    End If
                    LoadData()
                    sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    MessageBox.Show(sb.ToString)
                End If

            Else
                MessageBox.Show("Nothing to save.")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub SaveToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        Me.Validate()
        UpdateRecord()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        refreshrecord()
    End Sub

    Private Sub refreshrecord()
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                    LoadData()
                Case Windows.Forms.DialogResult.Cancel

                Case Windows.Forms.DialogResult.No
                    LoadData()
            End Select
        Else
            LoadData()
        End If
    End Sub
    Private Sub FormUser_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Dataset.HasChanges Then
            Dim datasetchanges As DataSet
            datasetchanges = Dataset.GetChanges()
            Dim response As Windows.Forms.DialogResult
            response = MessageBox.Show(datasetchanges.Tables(0).Rows.Count & " unsaved data. Do you want to store to the database?", "Unsaved data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Select Case response
                Case Windows.Forms.DialogResult.Yes
                    UpdateRecord()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
                Case Windows.Forms.DialogResult.No
            End Select
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        Dim sortDirection As SortOrder
        If TypeOf (DataGridView1.Columns(e.ColumnIndex)) Is DataGridViewComboBoxColumn Then
            'get column sorted
            Dim currentsortcolumn As DataGridViewColumn = DataGridView1.SortedColumn
            Dim column As DataGridViewComboBoxColumn = DataGridView1.Columns(e.ColumnIndex)


            If Not IsNothing(currentsortcolumn) Then
                If Not currentsortcolumn.Equals(sortcol(e.ColumnIndex)) Then
                    sortDirection = SortOrder.Ascending
                Else
                    If DataGridView1.SortOrder <> SortOrder.Ascending Then
                        sortDirection = SortOrder.Ascending
                    Else
                        sortDirection = SortOrder.Descending
                    End If
                End If
            Else
                If DataGridView1.SortOrder <> SortOrder.Ascending Then
                    sortDirection = SortOrder.Ascending
                Else
                    sortDirection = SortOrder.Descending
                End If

            End If
            Me.DataGridView1.Sort(sortcol(e.ColumnIndex), IIf(sortDirection = SortOrder.Ascending, System.ComponentModel.ListSortDirection.Ascending, System.ComponentModel.ListSortDirection.Descending))
            column.HeaderCell.SortGlyphDirection = sortDirection

        End If
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim filename As String = "VendorSP-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim dbtools = New Dbtools
        dbtools.Userid = "admin"
        dbtools.Password = "admin"
        sqlstr = "select v.vendorcode,v.vendorname,s.sbuname as bu, o.officersebname as sp from vp.vendorbusp vp" & _
                 " left join vendor v on v.vendorcode = vp.vendorcode" & _
                 " left join sbu s on s.sbuid = vp.buid" & _
                 " left join officerseb o on o.ofsebid = vp.spid" & _
                 " order by vendorname,bu,sp "
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim myform As New FormImportVendorBUSP
        myform.Show()
    End Sub
End Class