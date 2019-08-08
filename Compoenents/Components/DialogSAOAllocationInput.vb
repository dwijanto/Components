Imports System.Windows.Forms

Public Class DialogSAOAllocationInput
    Private DRV As DataRowView
    Private CustomerBS As New BindingSource
    Private CustomerTypes As New List(Of CustomerType)
    Public Sub New(ByVal drv As DataRowView, ByVal customerbs As BindingSource)
        InitializeComponent()
        Me.DRV = drv
        Me.CustomerBS = customerbs
        CustomerTypes.Add(New CustomerType(1, "SoldToParty (FG)"))
        CustomerTypes.Add(New CustomerType(2, "ShipToParty (COMP)"))
        CustomerTypes.Add(New CustomerType(3, "USA - POL"))
        CustomerTypes.Add(New CustomerType(4, "USA - ShipToParty"))
        InitDataDrv()

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Me.Validate Then
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        Dim custtype As CustomerType = ComboBox2.SelectedItem
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
        ErrorProvider1.SetError(ComboBox2, "")
        If IsNothing(drv) Then
            ErrorProvider1.SetError(ComboBox2, "Value cannot be blank.")
            myret = False
        Else
            Me.DRV.Row.Item("customertype") = custtype.name
            If custtype.id = 3 Then
                If TextBox2.Text = "" Then
                    ErrorProvider1.SetError(TextBox2, "Value cannot be blank.")
                    myret = False
                End If
            End If
        End If
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank.")
            myret = False
        End If


        Return myret
    End Function
    Private Sub InitDataDrv()
        ComboBox1.DataBindings.Clear()

        ComboBox1.DisplayMember = "customercode"
        ComboBox1.ValueMember = "customercode"
        ComboBox1.DataSource = CustomerBS

        ComboBox2.DisplayMember = "name"
        ComboBox2.ValueMember = "id"
        ComboBox2.DataSource = CustomerTypes

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox1.DataBindings.Add(New Binding("text", DRV, "userid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "pol", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox1.DataBindings.Add(New Binding("selectedvalue", DRV, "customercode", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        ComboBox2.DataBindings.Add(New Binding("selectedvalue", DRV, "ctype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        Label2.Text = "" & DRV.Row.Item("customername")
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim drv As DataRowView = ComboBox1.SelectedItem
        Label2.Text = drv.Row.Item("customername")
        Me.DRV.Row.Item("customername") = Label2.Text
    End Sub

End Class

Public Class CustomerType
    Public Property id As Integer
    Public Property name As String
    Public Sub New(ByVal id As Integer, ByVal name As String)
        Me.id = id
        Me.name = name
    End Sub
    Public Overrides Function ToString() As String
        Return name
    End Function
End Class