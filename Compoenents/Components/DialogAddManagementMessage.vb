Public Class DialogAddManagementMessage
    Private NewRecord As Boolean = False
    Private drv As DataRowView
    Private EP1 As New ErrorProvider
    Dim myadapter As ManagementMessageAdapter
    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.drv = drv
        'Me.myadapter = myadapter
        initdata()
    End Sub

    Public Sub New(ByVal myadapter As ManagementMessageAdapter)
        InitializeComponent()
        NewRecord = True
        Me.myadapter = myadapter
        drv = myadapter.BS.AddNew
        initdata()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        drv.EndEdit()
        If Me.validate Then
            If NewRecord Then
                If myadapter.save() Then
                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            Else
                Me.DialogResult = System.Windows.Forms.DialogResult.OK
                Me.Close()
            End If


        End If

    End Sub
    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        EP1.SetError(TextBox1, "")
        If TextBox1.Text = "" Then
            EP1.SetError(TextBox1, "Value cannot be blank.")
            myret = False
        Else

        End If
        Return myret
    End Function
    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initdata()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Add(New Binding("text", drv, "mgtmsg", True, DataSourceUpdateMode.OnPropertyChanged))
    End Sub
End Class