Imports System.Windows.Forms

Public Class DialogActivity

    Private DRV As DataRowView
    Dim SBUController1 As New SBUController
    Dim EP1 As New ErrorProvider

    Public Shared Event RefreshData()

    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
        initData()
    End Sub
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        DRV.EndEdit()
        If Me.validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initData()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()
        ComboBox3.DataBindings.Clear()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.ValueMember = "sbuid"
        ComboBox1.DisplayMember = "sbuname"
        ComboBox1.DataSource = SBUController1.GetSBUBS("where act")

        ComboBox2.ValueMember = "sbuid"
        ComboBox2.DisplayMember = "sbuname"
        ComboBox2.DataSource = SBUController1.GetSBUBS("where act")

        ComboBox3.ValueMember = "sbuid"
        ComboBox3.DisplayMember = "sbuname"
        ComboBox3.DataSource = SBUController1.GetSBUBS("where act")

        ComboBox4.ValueMember = "sbuid"
        ComboBox4.DisplayMember = "sbuname"
        ComboBox4.DataSource = SBUController1.GetSBUBS("where act")

        TextBox1.DataBindings.Add(New Binding("text", DRV, "activitycode", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "activityname", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("selectedvalue", DRV, "sbuid", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox2.DataBindings.Add(New Binding("selectedvalue", DRV, "sbuidsp", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox3.DataBindings.Add(New Binding("selectedvalue", DRV, "sbuidlg", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox4.DataBindings.Add(New Binding("selectedvalue", DRV, "sbuidvpi", True, DataSourceUpdateMode.OnPropertyChanged))


    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted, ComboBox2.SelectionChangeCommitted, ComboBox3.SelectionChangeCommitted, ComboBox4.SelectionChangeCommitted
        Dim mycb As ComboBox = DirectCast(sender, ComboBox)
        Dim cb1drv As DataRowView = mycb.SelectedItem
        Select Case mycb.Name
            Case "ComboBox1"
                DRV.Item("sbuname") = cb1drv.Item("sbuname")
            Case "ComboBox2"
                DRV.Item("sbunamesp") = cb1drv.Item("sbuname")
            Case "ComboBox3"
                DRV.Item("sbunamelg") = cb1drv.Item("sbuname")
            Case "ComboBox4"
                DRV.Item("bu") = cb1drv.Item("sbuname")
        End Select
        RaiseEvent RefreshData()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        RaiseEvent RefreshData()
    End Sub

    Private Sub DialogActivity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


    End Sub
End Class
