Public Class FormInputCommentCodeTx

    Private BS As BindingSource
    Private CategoryBS As BindingSource
    Private GroupBS As BindingSource
    Dim WithEvents oBindingNumeric1 As Binding
    Private myrow As DataRowView
    Private DS As DataSet
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByRef BS As BindingSource, ByRef CategoryBS As BindingSource, ByRef GroupBS As BindingSource, ByRef ds As DataSet)

        ' This call is required by the designer.
        InitializeComponent()
        Me.BS = BS
        Me.CategoryBS = CategoryBS
        Me.GroupBS = GroupBS
        ' Add any initialization after the InitializeComponent() call.
        oBindingNumeric1 = New Binding("Text", BS, "rank")

        TextBox1.DataBindings.Add(New Binding("Text", BS, "cmnttxdtlname"))
        TextBox2.DataBindings.Add(New Binding("Text", BS, "description"))
        TextBox3.DataBindings.Add(oBindingNumeric1)

        ComboBox1.DataSource = CategoryBS
        ComboBox1.DisplayMember = "cmnttxhdname"
        ComboBox1.ValueMember = "cmnttxhdid"
        ComboBox1.SelectedIndex = -1
        ComboBox1.DataBindings.Add("Text", BS, "cmnttxhdname")

        ComboBox2.DataSource = GroupBS
        ComboBox2.DisplayMember = "cmnttxgrpname"
        ComboBox2.ValueMember = "cmnttxgrpid"
        ComboBox2.SelectedIndex = -1
        ComboBox2.DataBindings.Add("text", BS, "cmnttxgrpname")
        Me.DS = ds
        myrow = BS.Current
    End Sub
    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim e1 = New System.ComponentModel.CancelEventArgs
    '    Dim e2 = New System.ComponentModel.CancelEventArgs
    '    Dim e3 = New System.ComponentModel.CancelEventArgs
    '    Dim e4 = New System.ComponentModel.CancelEventArgs
    '    Dim e5 = New System.ComponentModel.CancelEventArgs
    '    TextBox1_Validating(TextBox1, e1)
    '    TextBox1_Validating(TextBox2, e2)
    '    TextBox1_Validating(TextBox3, e4)
    '    ComboBox1_Validating(ComboBox1, e5)
    '    If e1.Cancel Or e2.Cancel Or e3.Cancel Or e4.Cancel Or e5.Cancel Then
    '        DialogResult = Windows.Forms.DialogResult.None
    '        Exit Sub
    '    End If
    '    myrow.Item("cmnttxhdname") = ComboBox1.Text
    '    myrow.Item("cmnttxgrpname") = ComboBox2.Text
    '    If ComboBox1.SelectedIndex = -1 Then
    '        'assign to CategoryBS

    '        Dim dr As DataRow = DS.Tables(1).NewRow()
    '        dr.Item("cmnttxhdname") = ComboBox1.Text
    '        dr.Item("cmnttxhdid") = DS.Tables(1).Rows.Count
    '        DS.Tables(1).Rows.Add(dr)
    '    End If
    '    If ComboBox2.SelectedIndex = -1 Then
    '        'assign to GroupBS
    '        Dim dr As DataRow = DS.Tables(2).NewRow()
    '        dr.Item("cmnttxgrpname") = ComboBox2.Text
    '        dr.Item("cmnttxgrpid") = DS.Tables(2).Rows.Count
    '        DS.Tables(2).Rows.Add(dr)
    '    End If
    '    Me.Validate()

    '    BS.EndEdit()
    'End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            Dim obj = CType(sender, TextBox)

            If obj.Tag = "String" Then
                If obj.Text = "" Then
                    ErrorProvider1.SetError(obj, "Value cannot be empty.")
                    Button1.DialogResult = Windows.Forms.DialogResult.None
                    e.Cancel = True
                End If
            ElseIf obj.Tag = "Number" Then
                If Not IsNumeric(obj.Text) Then
                    ErrorProvider1.SetError(obj, "Please enter numeric value.")
                    Button1.DialogResult = Windows.Forms.DialogResult.None
                    e.Cancel = True
                    'ElseIf obj.Text = 0 Then
                    '    ErrorProvider1.SetError(obj, "Value cannot be 0.")
                    '    Button1.DialogResult = Windows.Forms.DialogResult.None
                    '    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TextBox_validate(ByVal sender As Object, ByVal e As System.EventArgs)
        ErrorProvider1.SetError(sender, "")
        Button1.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
    Private Sub ComboBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim obj = CType(sender, ComboBox)
        If obj.SelectedIndex = -1 Then
            ErrorProvider1.SetError(sender, "Please select value from the list, Are you creating a new item?")
            'e.Cancel = True
            'Button1.DialogResult = Windows.Forms.DialogResult.None
        Else
            ErrorProvider1.SetError(sender, "")
        End If
    End Sub

    Private Sub ComboBox1_Validated(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox1.Validated, ComboBox2.Validated
        'ErrorProvider1.SetError(sender, "")
        'Button1.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        BS.CancelEdit()
    End Sub

    Private Sub obinding_Format(ByVal sender As Object, ByVal e As System.Windows.Forms.ConvertEventArgs) Handles oBindingNumeric1.Format
        If Not IsDBNull(e.Value) Then
            Select Case CType(sender, System.Windows.Forms.Binding).BindingMemberInfo.BindingField
                Case "rank"
                    e.Value = Format(e.Value, "#,##0")
                Case Else
                    e.Value = Format(e.Value, "#,##0.00")
            End Select
        End If
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim e1 = New System.ComponentModel.CancelEventArgs
        Dim e2 = New System.ComponentModel.CancelEventArgs
        Dim e3 = New System.ComponentModel.CancelEventArgs
        Dim e4 = New System.ComponentModel.CancelEventArgs
        Dim e5 = New System.ComponentModel.CancelEventArgs
        TextBox1_Validating(TextBox1, e1)
        TextBox1_Validating(TextBox2, e2)
        TextBox1_Validating(TextBox3, e4)
        ComboBox1_Validating(ComboBox1, e5)
        If e1.Cancel Or e2.Cancel Or e3.Cancel Or e4.Cancel Or e5.Cancel Then
            DialogResult = Windows.Forms.DialogResult.None
            Exit Sub
        End If
        myrow.Item("cmnttxhdname") = ComboBox1.Text
        myrow.Item("cmnttxgrpname") = ComboBox2.Text
        If ComboBox1.SelectedIndex = -1 Then
            'assign to CategoryBS
            Dim pkey(0) As Object
            pkey(0) = ComboBox1.Text
            Dim result As DataRow = DS.Tables(1).Rows.Find(pkey)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(1).NewRow()
                dr.Item("cmnttxhdname") = ComboBox1.Text
                'dr.Item("cmnttxhdid") = DS.Tables(1).Rows.Count
                DS.Tables(1).Rows.Add(dr)
            End If
            
        End If
        If ComboBox2.SelectedIndex = -1 Then
            'assign to GroupBS
            'find first if not avail then create
            Dim pkey(0) As Object
            pkey(0) = ComboBox2.Text
            Dim result As DataRow = DS.Tables(2).Rows.Find(pkey)
            If IsNothing(result) Then
                Dim dr As DataRow = DS.Tables(2).NewRow()
                dr.Item("cmnttxgrpname") = ComboBox2.Text
                ' dr.Item("cmnttxgrpid") = DS.Tables(2).Rows.Count
                DS.Tables(2).Rows.Add(dr)
            End If


            
        End If
        Me.Validate()

        BS.EndEdit()
    End Sub
End Class