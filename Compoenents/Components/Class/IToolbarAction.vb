Public Interface IToolbarAction
    Property ApplyFilter As String
    Function GetNewRecord() As DataRowView
    Function GetCurrentRecord() As DataRowView
    Sub RemoveAt(ByVal value As Integer)
    Function Save(ByVal mye As ContentBaseEventArgs) As Boolean
End Interface
