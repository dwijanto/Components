Imports Components.HelperClass
Imports Components.PublicClass
Public Class FormWeeklyTx
    Protected CM As CurrencyManager
    Protected mypanel1 As UCSortTx
    Protected mypanel As UCFilterTx
    Dim Dataset As DataSet
    Dim WithEvents Dataset2 As New DataSet
    Dim WithEvents DataTable2 As DataTable
    Dim sqlstr As String = String.Empty


    Private Sub FormWeeklyTx_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        Dim message As String = String.Empty
        LoadData()
        LoadToolstrip()
        ToolStripCustom1.Visible = False


    End Sub
    Public Sub LoadData()

        InitObject()
        FillData()
        BindDataSource()
        'BindingObject()
        CM = CType(BindingContext(BindingSource1), CurrencyManager)
    End Sub
    Public Overridable Sub InitObject()
        InitDataGrid()
        BindingSource1 = New BindingSource
        Dataset = New DataSet
        With DataGridView1
            .AutoGenerateColumns = False            
            .RowsDefaultCellStyle.BackColor = System.Drawing.Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke
        End With
    End Sub
    Public Overridable Sub FillData()
        Dim message As String = String.Empty
        Dim ra As Integer = 0

        If DbAdapter1.FillWeeklyTx(Dataset, message, ra) Then
            Dataset.Tables(0).TableName = "cxweeklyevolution"
            Dim idx0(0) As DataColumn
            idx0(0) = Dataset.Tables(0).Columns(0)
            Dataset.Tables(0).PrimaryKey = idx0
            Dataset.Tables(0).Columns(0).AutoIncrement = True
            Dataset.Tables(0).Columns(0).AutoIncrementSeed = -1
            Dataset.Tables(0).Columns(0).AutoIncrementStep = 1
        Else
            MessageBox.Show(message)
        End If
    End Sub
    Public Overridable Sub BindDataSource()
        BindingSource1.DataSource = Dataset.Tables("cxweeklyevolution")
        BindingNavigator1.BindingSource = BindingSource1
        DataGridView1.DataSource = BindingSource1
    End Sub

    'Public Overridable Sub BindingObject()
    '    'DataGridView1.Columns.Clear()
    '    'DataGridView1.AutoGenerateColumns = False
    '    'DataGridView1.DataSource = BindingSource1
    'End Sub


    Private Sub LoadToolstrip()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, DataGridView1, mypanel1, ToolStripCustom1, mypanel)
    End Sub

    Private Sub toolstripvisible(ByVal toolstripvisible As Boolean)
        ToolStripCustom1.Visible = Not (toolstripvisible)
        'Button3.Visible = toolstripvisible
    End Sub

    Public Overridable Sub InitDataGrid()
        With DataGridView1
            .Columns(0).DataPropertyName = "myyear"
            .Columns(1).DataPropertyName = "myweek"
            .Columns(2).DataPropertyName = "sasl"
            .Columns(3).DataPropertyName = "pctsasl"
            .Columns(4).DataPropertyName = "targetsasl"
            .Columns(5).DataPropertyName = "pctssl"
            .Columns(6).DataPropertyName = "countordertype"
            .Columns(7).DataPropertyName = "id"
            .Columns(8).DataPropertyName = "idori"
            .Columns(9).DataPropertyName = "yearweek"
            .Columns(10).DataPropertyName = "sp_updateweeklyevolution"
        End With
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoadData()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Call toolstripvisible(ToolStripCustom1.Visible)
    End Sub

    Public Sub UpdateRecord()
        Dim ra As Integer
        Dim message As String = String.Empty
        Dim sb As New System.Text.StringBuilder
        Try
            BindingSource1.EndEdit()
            Dim ds2 = Dataset.GetChanges
            If Not IsNothing(ds2) Then

                If DbAdapter1.AdapterWeeklyTx(ds2, message, ra) Then
                    'sb.Append(ra & " Record" & IIf(ra > 1, "s", "") & " Affected.")
                    Dataset.Merge(ds2)
                    Dataset.AcceptChanges()
                End If
                If Dataset.HasErrors Then
                    sb.Append("Some Record(s) has been modified/deleted by other user. Records will refresh.")
                    'sb.Append(message)
                    MessageBox.Show(sb.ToString)
                    LoadData()
                Else
                    If sb.ToString <> "" Then
                        MessageBox.Show(sb.ToString)
                    End If
                    'LoadData()
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
        If Me.Validate() Then
            UpdateRecord()
        End If

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

    Private Sub DataGridView1_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridView1.CellValidating
        If e.ColumnIndex = 1 Then
           
        End If
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub



    Private Sub BindingNavigatorAddNewItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorAddNewItem.Click
        Dim dr As DataRowView = BindingSource1.AddNew()
        dr.Item("myyear") = Year(Date.Today)
        Dataset.Tables(0).Rows.Add(dr.Row)
    End Sub

    Private Sub BindingNavigatorDeleteItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BindingNavigatorDeleteItem.Click
        If MessageBox.Show("Delete this record?", "Delete Record.", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            BindingSource1.RemoveAt(CM.Position)
        End If
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim myform As New FormGenerateChart
        myform.ShowDialog()
    End Sub
End Class