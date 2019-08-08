Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.ComponentModel
Public Class FormMarketContactCP
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents MEBS As BindingSource

    Dim DS As DataSet
    Dim sb As New StringBuilder

    Dim bsVendorName As BindingSource
    Dim bsVendorNameHelper As BindingSource
    Dim bsShipToPartyName As BindingSource
    Dim bsShipToPartyNameHelper As BindingSource

    Dim myDict As Dictionary(Of String, Integer)
    Dim myFields As String() = {"shiptopartycode", "shiptopartyname", "vendorcode", "vendorname", "name", "email"}

    Public Property groupid As Long = 0
    Public Property groupname As String = String.Empty

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load      
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()
        sb.Append("select shiptopartycode::text,customername::text as shiptopartyname,me.vendorcode::text,v.vendorname::text,me.name,me.email,me.id from marketemailcp me" &
                  " left join vendor v on v.vendorcode = me.vendorcode left join customer c on c.customercode = me.shiptopartycode;")
        sb.Append("select null as vendorcode,''::text as description,''::text as vendorname union all (select vendorcode, vendorcode::text || ' - ' || vendorname::text as description,vendorname::text from vendor order by vendorname);")
        sb.Append("select null as shiptopartycode,''::text as description,''::text as shiptopartyname union all (select customercode, customercode::text || ' - ' || customername::text as description,customername::text from customer where not customername isnull order by customername,customercode);")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "MarketEmail"
                DS.Tables(1).TableName = "Vendor"
                DS.Tables(2).TableName = "ShipTo"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try

                            MEBS = New BindingSource
                            bsVendorName = New BindingSource
                            bsVendorNameHelper = New BindingSource
                            bsShipToPartyName = New BindingSource
                            bsShipToPartyNameHelper = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1


                            MEBS.DataSource = DS.Tables(0)

                            bsVendorName.DataSource = New DataView(DS.Tables(1))
                            bsVendorNameHelper.DataSource = New DataView(DS.Tables(1))

                            bsShipToPartyName.DataSource = New DataView(DS.Tables(2))
                            bsShipToPartyNameHelper.DataSource = New DataView(DS.Tables(2))

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = MEBS
                            DataGridView1.RowTemplate.Height = 23

                            ComboBox1.DataBindings.Clear()
                            ComboBox2.DataBindings.Clear()
                            TextBox1.DataBindings.Clear()
                            TextBox2.DataBindings.Clear()

                            TextBox1.DataBindings.Add(New Binding("Text", MEBS, "name", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            TextBox2.DataBindings.Add(New Binding("Text", MEBS, "email", True, DataSourceUpdateMode.OnPropertyChanged, ""))

                            ComboBox1.DataSource = bsVendorName
                            ComboBox1.DisplayMember = "description"
                            ComboBox1.ValueMember = "vendorcode"
                            ComboBox1.DataBindings.Add("SelectedValue", MEBS, "vendorcode", True, DataSourceUpdateMode.OnPropertyChanged)

                            ComboBox2.DataSource = bsShipToPartyName
                            ComboBox2.DisplayMember = "description"
                            ComboBox2.ValueMember = "shiptopartycode"
                            ComboBox2.DataBindings.Add("SelectedValue", MEBS, "shiptopartycode", True, DataSourceUpdateMode.OnPropertyChanged)

                            If IsNothing(MEBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                                ComboBox2.SelectedIndex = -1
                            End If

                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles MEBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(MEBS.Current)
        ComboBox2.Enabled = Not IsNothing(MEBS.Current)
        TextBox1.Enabled = Not IsNothing(MEBS.Current)
        TextBox2.Enabled = Not IsNothing(MEBS.Current)
        Button8.Enabled = Not IsNothing(MEBS.Current)
        Button1.Enabled = Not IsNothing(MEBS.Current)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = MEBS.AddNew()

        drv.Row.BeginEdit()
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Try
            MEBS.EndEdit()
            If Me.validate Then
                Try
                    'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                    Dim ds2 As DataSet
                    ds2 = DS.GetChanges

                    If Not IsNothing(ds2) Then
                        Dim mymessage As String = String.Empty
                        Dim ra As Integer
                        Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                        If Not DbAdapter1.MarketContactCPTx(Me, mye) Then
                            MessageBox.Show(mye.message)
                            DS.Merge(ds2)
                            Exit Sub
                        End If
                        DS.Merge(ds2)
                        DS.AcceptChanges()
                        DataGridView1.Invalidate()
                        MessageBox.Show("Saved.")
                    End If
                Catch ex As Exception
                    MessageBox.Show(" Error:: " & ex.Message)
                End Try
            End If
            DataGridView1.Invalidate()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In MEBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next
        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        'If IsDBNull(drv.Row.Item("vendorcode")) Then
        '    myret = False
        '    sb.Append("Vendor Code cannot be blank")
        'End If
        If IsDBNull(drv.Row.Item("shiptopartycode")) Then
            myret = False
            sb.Append("Ship To Party cannot be blank")
        End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(MEBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    MEBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub


    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click, Button1.Click
        Dim myobj As Button = CType(sender, Button)
        Try
            Select Case myobj.Name
                Case "Button8"
                    Dim myform = New FormHelper(bsVendorNameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = bsVendorNameHelper.Current
                        Dim mydrv As DataRowView = MEBS.Current
                        mydrv.BeginEdit()
                        mydrv.Row.Item("vendorcode") = drv.Row.Item("vendorcode")
                        mydrv.Row.Item("vendorname") = drv.Row.Item("vendorname")
                        mydrv.EndEdit()
                        'Need bellow code to sync with combobox
                        Dim myposition = bsVendorName.Find("vendorcode", drv.Row.Item("vendorcode"))
                        bsVendorName.Position = myposition
                    End If
                Case "Button1"
                    Dim myform = New FormHelper(bsShipToPartyNameHelper)
                    myform.DataGridView1.Columns(0).DataPropertyName = "description"
                    If myform.ShowDialog = DialogResult.OK Then
                        Dim drv As DataRowView = bsShipToPartyNameHelper.Current
                        Dim mydrv As DataRowView = MEBS.Current
                        mydrv.BeginEdit()
                        mydrv.Row.Item("shiptopartycode") = drv.Row.Item("shiptopartycode")
                        mydrv.Row.Item("shiptopartyname") = drv.Row.Item("shiptopartyname")
                        mydrv.EndEdit()
                        'Need bellow code to sync with combobox
                        Dim myposition = bsShipToPartyName.Find("shiptopartycode", drv.Row.Item("shiptopartycode"))
                        bsShipToPartyName.Position = myposition
                    End If
            End Select
        Catch ex As Exception

            MessageBox.Show(ex.Message)
        End Try

        DataGridView1.Invalidate()
    End Sub


    Private Sub ComboBox2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        '1. Force Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(MEBS.Current) Then
            Dim myselected1 As DataRowView = myobj.SelectedItem

            Dim drv As DataRowView = MEBS.Current
            Try

                drv.Row.BeginEdit()
                Select Case myobj.Name
                    Case "ComboBox1"
                        drv.Row.Item("vendorname") = myselected1.Row.Item("vendorname")
                    Case "ComboBox2"
                        drv.Row.Item("shiptopartyname") = myselected1.Row.Item("shiptopartyname")
                End Select

                MEBS.EndEdit()
            Catch ex As Exception
                Select Case myobj.Name
                    Case "Combobox1"
                        ComboBox1.SelectedValue = drv.Row.Item("vendorcode", DataRowVersion.Original)
                    Case "Combobox2"
                        ComboBox1.SelectedValue = drv.Row.Item("shiptopartycode", DataRowVersion.Original)
                End Select

                drv.Row.CancelEdit()
                MessageBox.Show(ex.Message)
            End Try

        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        MEBS.Filter = ""
        ToolStripStatusLabel1.Text = ""
        If ToolStripTextBox1.Text <> "" And (ToolStripComboBox1.SelectedIndex <> -1) Then
            Select Case ToolStripComboBox1.SelectedIndex
                Case 0, 2
                    If Not IsNumeric(ToolStripTextBox1.Text) Then
                        ToolStripTextBox1.Select()
                        SendKeys.Send("{BACKSPACE}")
                        Exit Sub
                    End If
            End Select
            MEBS.Filter = myFields(ToolStripComboBox1.SelectedIndex).ToString & " like '%" & sender.ToString.Replace("'", "''") & "%'"
            ToolStripStatusLabel1.Text = "Record Count " & MEBS.Count
        End If
    End Sub


    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message.ToString)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim myform As New FormImportMasterContactCP
        myform.ShowDialog()
        Me.loaddata()
    End Sub

    Private Sub TextBox1_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        DataGridView1.Invalidate()
    End Sub
End Class