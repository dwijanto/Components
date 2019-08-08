
Public Class UCHelper
    Public Shared Sub BindFieldCombobox(ByRef CB As ComboBox, ByVal DG As DataGridView)
        Dim cols As List(Of FieldClass) = New List(Of FieldClass)
        If DG.Columns.Count > 0 Then

            For i = 0 To DG.Columns.Count - 1
                If DG.Columns(i).Visible Then
                    cols.Add(New FieldClass With {.id = DG.Columns(i).DataPropertyName,
                                                  .name = DG.Columns(i).HeaderText,
                                                  .ColumnType = DG.Columns(i).GetType.Name,
                                                  .ColumnIndex = i,
                                                  .columnLookup = DG.Columns(i).Tag})
                End If
            Next
            cols.Insert(0, New FieldClass With {.id = "None", .name = "None", .ColumnType = "None"})
        End If
        CB.DataSource = cols
        CB.DisplayMember = "Name"
        CB.SelectedItem = "id"
    End Sub
End Class
Public Enum ExpandedState
    Expanded
    Collapsed
End Enum

Public Enum FilterOperator
    None
    IsEqualTo
    IsLike
    IsLessThan
    IsLessThanOrEqualTo
    IsGreaterThan
    IsGreaterThanOrEqualTo
    IsNotEqualTo
End Enum
Public Class FieldClass
    Public Property name
    Public Property id
    Public Property ColumnType
    Public Property ColumnIndex
    Public Property columnLookup


    Public Overrides Function ToString() As String
        'Return MyBase.ToString()
        Return name.ToString
    End Function
End Class

Public Enum SortDirection
    Asc
    Desc
End Enum

Public Delegate Sub HideToolbarDelegate(ByVal ToolstripVisible As Boolean)