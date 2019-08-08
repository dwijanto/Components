Imports System.Windows.Forms

Public Class DialogAddCommentMapping
    Private DRV As DataRowView
    Private MMBS As ManagementMessageAdapter
    Private CIBS As CategoryIssuesAdapter
    Private CGBS As CommentGroupAdapter
    Private EP1 As New ErrorProvider

    Public Sub New(ByVal drv As DataRowView, ByVal MMBS As ManagementMessageAdapter, ByVal CIBS As CategoryIssuesAdapter)
        InitializeComponent()

        Me.DRV = drv
        Me.MMBS = MMBS
        Me.CIBS = CIBS

        initData()

    End Sub
    Public Sub New(ByVal drv As DataRowView, ByVal MMBS As ManagementMessageAdapter, ByVal CIBS As CategoryIssuesAdapter, ByVal CGBS As CommentGroupAdapter)
        InitializeComponent()

        Me.DRV = drv
        Me.MMBS = MMBS
        Me.CIBS = CIBS
        Me.CGBS = CGBS
        initData()

    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        DRV.EndEdit()
        If Me.validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub
    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        EP1.SetError(ComboBox1, "")
        EP1.SetError(ComboBox2, "")
        If IsNothing(ComboBox1.SelectedItem) Then
            EP1.SetError(ComboBox1, "Please select from list.")
            myret = False
        Else
            DRV.Row.Item("catissues") = DirectCast(ComboBox1.SelectedItem, DataRowView).Row.Item("catissues")
        End If
        If IsNothing(ComboBox2.SelectedItem) Then
            EP1.SetError(ComboBox2, "Please select from list.")
            myret = False
        Else
            DRV.Row.Item("mgtmsg") = DirectCast(ComboBox2.SelectedItem, DataRowView).Row.Item("mgtmsg")
        End If
        If Not IsNothing(ComboBox3.SelectedItem) Then
            DRV.Row.Item("cmnttxgrpname") = DirectCast(ComboBox3.SelectedItem, DataRowView).Row.Item("cmnttxgrpname")
        End If

        Return myret
    End Function
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initData()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()

        TextBox1.DataBindings.Clear()

        ComboBox1.ValueMember = "id"
        ComboBox1.DisplayMember = "catissues"
        ComboBox1.DataSource = CIBS.BS

        ComboBox2.ValueMember = "id"
        ComboBox2.DisplayMember = "mgtmsg"
        ComboBox2.DataSource = MMBS.BS

        ComboBox3.ValueMember = "cmnttxgrpid"
        ComboBox3.DisplayMember = "cmnttxgrpname"
        ComboBox3.DataSource = CGBS.BS

        TextBox1.DataBindings.Add(New Binding("text", DRV, "comment", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("selectedvalue", DRV, "catid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("selectedvalue", DRV, "mgtmsgid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("selectedvalue", DRV, "groupid", True, DataSourceUpdateMode.OnPropertyChanged))


    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim myform As New DialogAddCategoryIssue(CIBS)
        myform.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim myform As New DialogAddManagementMessage(MMBS)
        myform.Show()
    End Sub
End Class
