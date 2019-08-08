Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Public Class FormBrowseFolderCP
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)


    Dim ds As DataSet

    Dim BS As BindingSource
    Dim cbbs As BindingSource
    Dim startdate As DateTime
    Dim enddate As DateTime
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim myfilter As String = String.Empty
    Dim mybasefolder As String
    Dim mytextfilter As String = String.Empty
    Private Sub FormBrowseFolder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Validate()
        Dim ds2 = ds.GetChanges()
        If Not IsNothing(ds2) Then
            If MessageBox.Show("There is unsaved data in a row. Do you want to store to the database?", "Save Changes", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                ToolStripButton4.PerformClick()
            End If
        End If

    End Sub

    Private Sub FormBrowseFolder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton6.Click
        mytextfilter = ToolStripComboBox1.Text
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

        Dim sqlstr As String = "select * from docemailhdcp where docemailtype = 0 and receiveddate >= " & DateFormatyyyyMMdd(startdate) & " and receiveddate <=  " & DateFormatyyyyMMdd(enddate) & "  order by receiveddate desc;" &
                                "select ''::character varying as foldername union all (select distinct foldername from docemailhdcp where docemailtype = 0" &
                                " order by foldername);" &
                                " select * from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'logbookcp' order by pd.ivalue;" &
                                " select d.* from docemaildtcp d left join docemailhdcp h on h.docemailhdid = d.docemailhdid  where docemailtype = 0 and h.receiveddate >= " & DateFormatyyyyMMdd(startdate) & " and h.receiveddate <= " & DateFormatyyyyMMdd(enddate) & ";"

        'Dim sqlstr As String = "select * from docemailhd order by receiveddate desc"
        ds = New DataSet
        BS = New BindingSource
        cbbs = New BindingSource

        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            Dim idx0(0) As DataColumn
            idx0(0) = ds.Tables(0).Columns("docemailhdid")
            ds.Tables(0).PrimaryKey = idx0

            ds.Tables(0).Columns(0).AutoIncrement = True
            ds.Tables(0).Columns(0).AutoIncrementSeed = 0
            ds.Tables(0).Columns(0).AutoIncrementStep = -1

            Dim docemailhdidxU As UniqueConstraint = New UniqueConstraint(New DataColumn() {ds.Tables(0).Columns("docemailname")})
            ds.Tables(0).Constraints.Add(docemailhdidxU)

            ds.Tables(0).TableName = "DocEmailHD"
            BS.DataSource = ds.Tables(0)

            cbbs.DataSource = ds.Tables(1)

            Dim idx3(0) As DataColumn
            idx3(0) = ds.Tables(3).Columns("docemaildtid")
            ds.Tables(3).PrimaryKey = idx3

            ds.Tables(3).Columns(0).AutoIncrement = True
            ds.Tables(3).Columns(0).AutoIncrementSeed = 0
            ds.Tables(3).Columns(0).AutoIncrementStep = -1


            Dim rel As DataRelation
            Dim hcol As DataColumn
            Dim dcol As DataColumn

            hcol = ds.Tables(0).Columns("docemailhdid") 'docemailhdid in table header
            dcol = ds.Tables(3).Columns("docemailhdid") 'docemailhdid in table dtl
            rel = New DataRelation("hdrel", hcol, dcol)
            ds.Relations.Add(rel)


            ProgressReport(1, "Assign DataGridView DataSource")
            mybasefolder = ds.Tables(2).Rows(4).Item("cvalue") & "\"
        Else
            ProgressReport(2, mymessage)
        End If
        ProgressReport(5, "Continuous")
        ProgressReport(8, "Combobox text")
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
                    DataGridView1.DataSource = BS


                    ToolStripComboBox1.ComboBox.DisplayMember = "foldername"
                    ToolStripComboBox1.ComboBox.DataSource = cbbs
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
                Case 8
                    ToolStripComboBox1.Text = mytextfilter
            End Select

        End If

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Dim myrow As DataRowView = BS.Current
        Dim myfolder = myrow.Row.Item("docemailname")
        Dim p As New System.Diagnostics.Process
        p.StartInfo.FileName = "explorer.exe"
        p.StartInfo.Arguments = String.Format("{0},""{1}{2}""", "/select", mybasefolder, DbAdapter1.validfilename(myfolder))
        'Process.Start("explorer.exe", "/select," & "C:\temp\Documents\Forwarder\""" & DbAdapter1.validfilename(myfolder) & """")
        'Process.Start("explorer.exe", "/select," & mybasefolder & """" & DbAdapter1.validfilename(myfolder) & """")
        p.Start()
    End Sub

    Sub doQuery()

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        'If ToolStripComboBox1.ComboBox.Text = "" Then
        '    BS.Filter = ""
        'Else
        '    BS.Filter = "[foldername] = '" & ToolStripComboBox1.ComboBox.Text & "'"
        '    'myfilter = ToolStripComboBox1.ComboBox.Text
        'End If
        ToolStripTextBox1_TextChanged(Me, e)
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
        ToolStrip1.Items.Insert(10, host1)
        ToolStrip1.Items.Insert(12, host2)
    End Sub


    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        BS.EndEdit()
        Dim ds2 = ds.GetChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.BrowseFolderTx(Me, mye) Then
                'delete the modfied row for Merged
                Dim modifiedRows = From row In ds.Tables(0)
                   Where row.RowState = DataRowState.Added
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
                modifiedRows = From row In ds.Tables(3)
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
        dr.Item("docemailtype") = 0
        dr.Item("receiveddate") = Today
        ds.Tables(0).Rows.Add(dr)

        Dim myform As New FormDocEmail(BS, ds)
        If Not myform.ShowDialog = DialogResult.OK Then
            BS.RemoveCurrent()
        Else
            BS.EndEdit()
        End If

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click

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


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim sb As New StringBuilder
        Dim myfilter As String = String.Empty
        Dim userfolder As String = String.Empty
        BS.Filter = ""
        If Not ToolStripComboBox1.Text = "" Then
            userfolder = "[foldername] = '" & ToolStripComboBox1.Text & "'"
            sb.Append(userfolder)
        End If
        If ToolStripTextBox1.Text <> "" Then
            myfilter = "[docemailname] like '" & ToolStripTextBox1.Text & "'"
        End If

        If sb.Length > 0 And myfilter <> "" Then
            sb.Append(" and ")
        End If
        sb.Append(myfilter)
        BS.Filter = sb.ToString
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim myrow = BS.Current
        Dim myform = New FormDocEmail(BS, ds)

        If Not myform.ShowDialog = DialogResult.OK Then
            'MessageBox.Show("Add New One")
            BS.CancelEdit()
        Else
            BS.EndEdit()
        End If
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        BS.CancelEdit()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class