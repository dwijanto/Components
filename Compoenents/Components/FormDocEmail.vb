Public Class FormDocEmail
    Dim BS As New BindingSource
    Dim BSHeader As BindingSource
    Dim dr As DataRow
    Dim DS As DataSet
    Dim drdetail As DataRow
    Dim drvdetail As DataRowView
    Public Sub New(ByRef bsheader As BindingSource, ByRef DS As DataSet)

        ' This call is required by the designer.
        InitializeComponent()
        Me.BSHeader = bsheader
        Me.DS = DS
        ClearBindingObject()
        BindingObject()
        dr = CType(bsheader.Current, DataRowView).Row
        BS.DataSource = Nothing
        BS.DataSource = Me.DS.Tables(3)
        BS.Filter = "docemailhdid =" & dr.Item("docemailhdid")
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ClearBindingObject()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        DataGridView1.DataSource = Nothing
        DateTimePicker1.DataBindings.Clear()

    End Sub

    Private Sub BindingObject()
        TextBox1.DataBindings.Add(New Binding("Text", BSHeader, "docemailname"))
        TextBox2.DataBindings.Add(New Binding("Text", BSHeader, "foldername"))
        TextBox3.DataBindings.Add(New Binding("Text", BSHeader, "sendername"))
        TextBox4.DataBindings.Add(New Binding("Text", BSHeader, "sender"))
        DateTimePicker1.DataBindings.Add(New Binding("Text", BSHeader, "receiveddate"))
        DataGridView1.DataSource = BS
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        drvdetail = BS.AddNew
        drdetail = drvdetail.Row
        drdetail.Item("docemailhdid") = Me.dr.Item("docemailhdid")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        BS.CancelEdit()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(3).Rows.Remove(CType(BS.Current, DataRowView).Row)
                For Each dsrow In DataGridView1.SelectedRows
                    BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)

                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub
End Class