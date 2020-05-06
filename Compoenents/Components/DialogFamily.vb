Imports System.Windows.Forms

Public Class DialogFamily

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
       
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()

        ComboBox1.ValueMember = "sbuid"
        ComboBox1.DisplayMember = "sbuname"
        ComboBox1.DataSource = SBUController1.GetSBUBS("where act")


        TextBox1.DataBindings.Add(New Binding("text", DRV, "familyid", True, DataSourceUpdateMode.OnPropertyChanged))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "familyname", True, DataSourceUpdateMode.OnPropertyChanged))
        ComboBox1.DataBindings.Add(New Binding("selectedvalue", DRV, "sbuid", True, DataSourceUpdateMode.OnPropertyChanged))
        
    End Sub



    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim mycb As ComboBox = DirectCast(sender, ComboBox)
        Dim cb1drv As DataRowView = mycb.SelectedItem
        DRV.Item("sbuname") = cb1drv.Item("sbuname")        
        RaiseEvent RefreshData()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        Dim myobj As TextBox = DirectCast(sender, TextBox)
        If myobj.Name = "TextBox1" Then
            DRV.Item("familyidtext") = TextBox1.Text
        End If
        RaiseEvent RefreshData()
    End Sub


End Class
