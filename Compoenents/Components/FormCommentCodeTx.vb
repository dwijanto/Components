Imports System.Threading
Imports Components.PublicClass
Delegate Sub ProgressReport(ByVal id As Integer, ByVal message As String)
Public Class FormCommentCodeTx
    Dim myThreadStart As New ThreadStart(AddressOf doLoad)
    Dim myThread As New Thread(myThreadStart)
    Dim bs As BindingSource
    Dim CategoryBS As BindingSource
    Dim GroupBS As BindingSource
    Dim DS As DataSet
    Dim myfilter As String() = {"", "Code", "Category", "Description", "Group", "Ranking"}
    Dim myfiltervalue As String() = {"", "cmnttxdtlname", "cmnttxhdname", "description", "cmnttxgrpname", "rank"}
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ToolStripComboBox1.Items.AddRange(myfilter)
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub FormCommentCodeTx_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim myobj = DS.GetChanges
        If Not IsNothing(myobj) Then
            Select Case MessageBox.Show("Save modified records?", "Question", System.Windows.Forms.MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    ToolStripButton5.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Sub FormCommentCodeTx_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf doLoad)
            myThread.Start()
        End If
    End Sub

    Sub doLoad()
        ProgressReport(6, "Marquee")
        Dim sqlstr As String = "select c.cmnttxdtlid, c.cmnttxdtlname::character varying,h.cmnttxhdname::character varying,c.description::character varying,g.cmnttxgrpname::character varying,c.rank from cmnttxdtl c " &
                               " left join cmnttxhd h on h.cmnttxhdid = c.cmnttxhdid" &
                               " left join cmnttxgrp g on g.cmnttxgrpid = c.cmnttxgrpid" &
                               " where h.cmnttxhdid <> 11 order by c.cmnttxdtlname;" &
                               " select cmnttxhdname::character varying ,cmnttxhdid from cmnttxhd order by cmnttxhdname;" &
                               " select cmnttxgrpname::character varying,cmnttxgrpid from cmnttxgrp order by cmnttxgrpname;"
        DS = New DataSet
        bs = New BindingSource
        CategoryBS = New BindingSource
        GroupBS = New BindingSource
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then

            Dim idx0(0) As DataColumn
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0
            DS.Tables(0).Columns(0).AutoIncrement = True
            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
            DS.Tables(0).Columns(0).AutoIncrementStep = -1
            DS.Tables(0).TableName = "CommentCode"

            Dim idx1(0) As DataColumn
            idx1(0) = DS.Tables(1).Columns(0)
            DS.Tables(1).PrimaryKey = idx1
            'DS.Tables(1).Columns(0).AutoIncrement = True
            'DS.Tables(1).Columns(0).AutoIncrementSeed = 0
            'DS.Tables(1).Columns(0).AutoIncrementStep = -1

            Dim idx2(0) As DataColumn
            idx2(0) = DS.Tables(2).Columns(0)
            DS.Tables(2).PrimaryKey = idx2
            'DS.Tables(2).Columns(0).AutoIncrement = True
            'DS.Tables(2).Columns(0).AutoIncrementSeed = 0
            'DS.Tables(2).Columns(0).AutoIncrementStep = -1

            bs.DataSource = DS.Tables(0)
            CategoryBS.DataSource = DS.Tables(1)
            GroupBS.DataSource = DS.Tables(2)
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
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = bs
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

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripComboBox1.SelectedIndexChanged
        bs.Filter = ""
        If (ToolStripTextBox1.Text <> "" AndAlso ToolStripComboBox1.Text <> "") Then
            bs.Filter = "[" & myfiltervalue(ToolStripComboBox1.SelectedIndex) & "] like '" & ToolStripTextBox1.Text & "'"
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim myrow = bs.AddNew
        Dim myform As New FormInputCommentCodeTx(bs, CategoryBS, GroupBS, DS)
        If myform.ShowDialog = DialogResult.OK Then

            DataGridView1.Invalidate()
        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        'Dim myrow = bs.Current
        Dim myform As New FormInputCommentCodeTx(bs, CategoryBS, GroupBS, DS)
        If myform.ShowDialog = DialogResult.OK Then
            DataGridView1.Invalidate()

        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton3.PerformClick()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(bs.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
                For Each dsrow In DataGridView1.SelectedRows
                    bs.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        'save
        Me.Validate()
        Me.bs.EndEdit()
        Me.CategoryBS.EndEdit()
        Me.GroupBS.EndEdit()

        Dim Ds2 = DS.GetChanges
        If Not IsNothing(Ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(Ds2, True, mymessage, ra, True)
            If DbAdapter1.SaveComment(Me, mye) Then
                Dim addedrows1 = From row As DataRow In DS.Tables(0).Rows
                                Where row.RowState = DataRowState.Added
                For Each row In addedrows1.ToArray
                    row.Delete()
                Next
                Dim addedrows2 = From row As DataRow In DS.Tables(1).Rows
                Where row.RowState = DataRowState.Added
                For Each row In addedrows2.ToArray
                    row.Delete()
                Next
                Dim addedrows3 = From row As DataRow In DS.Tables(2).Rows
                Where row.RowState = DataRowState.Added
                For Each row In addedrows3.ToArray
                    row.Delete()
                Next
                DS.Merge(Ds2)
                DS.AcceptChanges()
                MessageBox.Show("Saved.")
            Else
                MessageBox.Show("Not Saved.")
            End If
        End If
    End Sub
End Class