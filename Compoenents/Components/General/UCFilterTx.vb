Imports System.ComponentModel

Public Class UCFilterTx
    Inherits UCToolstrip
    Private bs As New BindingSource
    Private DG As DataGridView
    Private FilterOperatorHash As New Hashtable
    Private ComboList As New Dictionary(Of String, String)
    Private isDateTime As Boolean
    Dim _myfilter As String
    Private _ColumnIndexFiltered As New ArrayList

    Public ReadOnly Property ColumnIndexFiltered As ArrayList
        Get
            Return _ColumnIndexFiltered
        End Get
    End Property
    Public ReadOnly Property myFilter As String
        Get
            Return _myfilter
        End Get
    End Property

    Public Sub New(ByVal DG As DataGridView)
        ' This call is required by the designer.
        InitializeComponent()
        ToolStripLabel1.Text = "Filter"
        ' Add any initialization after the InitializeComponent() call.        
        bs = CType(DG.DataSource, BindingSource)
        Me.DG = DG
        InitDataLayout()
        MyBase.setPanelSize = Me.Panel1.Size

    End Sub

    Private Sub InitDataLayout()
        BindFieldCombobox()
        BuildAutoCompleteString()
        OperatorComboBox.DataSource = System.Enum.GetNames(GetType(FilterOperator))
        InitFilterOperatorHash()
    End Sub

    Public Sub BuildAutoCompleteString()

        bs = CType(DG.DataSource, BindingSource)
        Dim myfilter As String
        Dim myFieldClass = CType(FieldComboBox.SelectedItem, FieldClass)
        isDateTime = False
        'clear first
        FilterTextBox.AutoCompleteCustomSource.Clear()

        If bs.List.Count <= 0 OrElse FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 Then Return

        'Get Column Name
        myfilter = bs.Filter
        If RadioButton2.Checked Then
            bs.Filter = ""
        End If

        Dim FilterField As String = myFieldClass.id 'CType(FieldComboBox.SelectedItem, FieldClass).id.ToString
        Dim filterVals As AutoCompleteStringCollection = New AutoCompleteStringCollection


        If myFieldClass.ColumnType = "DataGridViewComboBoxColumn" Then
            Dim bs2 As New BindingSource
            Dim dgcombo = CType(DG.Columns(myFieldClass.ColumnIndex), DataGridViewComboBoxColumn)

            bs2 = CType(dgcombo.DataSource, BindingSource)

            Try
                ComboList.Clear()
                For Each dataitem As Object In bs2.List
                    Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(dataitem)
                    Dim propdesc As PropertyDescriptor = props.Find(dgcombo.DisplayMember, True)
                    Dim propdesc2 As PropertyDescriptor = props.Find(dgcombo.DataPropertyName, True)
                    'Dim propdesc2 As PropertyDescriptor = props.Find(dgcombo.V, True)
                    Dim mykey As String = propdesc.GetValue(dataitem).ToString
                    Dim myvalue As String = propdesc2.GetValue(dataitem).ToString
                    Try
                        ComboList.Add(mykey.ToLower, myvalue)
                    Catch ex As Exception

                    End Try
                    filterVals.Add(mykey)
                Next
            Catch ex As Exception
                MessageBox.Show(ex.Message)

            End Try
        Else
            For Each dataitem As Object In bs.List
                Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(dataitem)
                Dim propdesc As PropertyDescriptor = props.Find(FilterField, True)
                Try
                    Dim fieldval As String = propdesc.GetValue(dataitem).ToString
                    If propdesc.PropertyType.Name = "DateTime" Then
                        isDateTime = True
                    End If
                    filterVals.Add(fieldval)
                Catch ex As Exception

                End Try

            Next
        End If
        bs.Filter = myfilter
        FilterTextBox.AutoCompleteCustomSource = filterVals
    End Sub
    Private Sub FieldComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FieldComboBox.SelectedIndexChanged
        BuildAutoCompleteString()
    End Sub
    Private Sub InitFilterOperatorHash()
        FilterOperatorHash.Add(0, "None")
        FilterOperatorHash.Add(1, "=")
        FilterOperatorHash.Add(2, "Like")
        FilterOperatorHash.Add(3, "<")
        FilterOperatorHash.Add(4, "<=")
        FilterOperatorHash.Add(5, ">")
        FilterOperatorHash.Add(6, ">=")
        FilterOperatorHash.Add(7, "<>")
    End Sub
    Private Sub BindFieldCombobox()
        UCHelper.BindFieldCombobox(FieldComboBox, DG)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
        FilterOperatorHash.Clear()
        ComboList.Clear()
    End Sub

    Private Sub executefilter()
        bs = CType(DG.DataSource, BindingSource)
        If bs.List.Count <= 0 OrElse
            FieldComboBox.Items.Count <= 0 OrElse
            FieldComboBox.SelectedIndex <= 0 OrElse
            OperatorComboBox.SelectedIndex <= 0 Then
            Return
        End If

        If String.IsNullOrEmpty(FilterTextBox.Text) Then Return

        'inFilterMode = True

        '##getpropertyname##
        '1.get columnname from combo
        Dim myFieldClass = CType(FieldComboBox.SelectedItem, FieldClass)
        Dim filterMember As String = myFieldClass.id.ToString 'CType(FieldComboBox.SelectedItem, FieldClass).id.ToString

        If myFieldClass.ColumnType = "DataGridViewComboBoxColumn" Then
            If myFieldClass.columnLookup <> "" Then
                filterMember = myFieldClass.columnLookup.ToString
            End If
        End If

        '1.b Check for ComboboxColumn
        Dim filterValue As String = Nothing
        Dim SearchValue As String = Nothing
        SearchValue = FilterTextBox.Text
        'If myFieldClass.ColumnType = "DataGridViewComboBoxColumn" Then
        '    Try
        '        SearchValue = ComboList(SearchValue.ToLower)
        '    Catch ex As Exception
        '        Err.Raise("1", Description:=String.Format("Operator ""{0}"" is not applicable for column {1}", OperatorComboBox.SelectedValue.ToString, myFieldClass.name))
        '    End Try
        'End If
        '2.Get dataitem from bindinglist.list(0)
        Dim DataItem As Object = bs.List(0)
        '3.Get Propertiescollection from dataitem
        Dim props As PropertyDescriptorCollection = TypeDescriptor.GetProperties(DataItem)
        '4.Get Selected PropertyDescriptor based on filtermember
        Dim propDesc As PropertyDescriptor = props.Find(filterMember, True)

        'getoperator
        Dim stringoperator As String = FilterOperatorHash(OperatorComboBox.SelectedIndex).ToString
        'putbindingfilter
        'Check for different format

        Dim JoinFilter As String = "AND "
        If Not CheckBox1.Checked Then
            _myfilter = ""
            'ColumnIndexFiltered.Clear()
        Else
            _myfilter = bs.Filter
        End If
        'If bs.Filter <> "" AndAlso bs.Filter IsNot Nothing Then
        If _myfilter <> "" AndAlso _myfilter IsNot Nothing Then
            If RadioButton2.Checked Then JoinFilter = "OR "

        Else
            JoinFilter = ""
        End If

        Select Case OperatorComboBox.SelectedIndex
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
                If isDateTime Then
                    filterValue = String.Format("{0}[{1}] {2} '#{3}#'", JoinFilter, propDesc.Name, stringoperator, SearchValue)
                Else
                    filterValue = String.Format("{0}[{1}] {2} '{3}'", JoinFilter, propDesc.Name, stringoperator, SearchValue)
                End If
        End Select

        Try
            If Not (ColumnIndexFiltered.Contains(myFieldClass.ColumnIndex)) Then
                ColumnIndexFiltered.Add(myFieldClass.ColumnIndex)
            End If
            _myfilter = _myfilter & filterValue
            bs.Filter = myFilter
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            _myfilter = ""
            bs.Filter = ""
        End Try

    End Sub


    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        RadioButton1.Enabled = CheckBox1.Checked
        RadioButton2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            executefilter()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        BS.Filter = ""
        _myfilter = ""
        ColumnIndexFiltered.Clear()
    End Sub
End Class
