Public Delegate Function SaveDelegate(ByRef sender As Object, ByRef e As EventArgs) As Boolean
Public Class ModelAdapter


    Dim myAdapter As DbAdapter
    Public DS As DataSet
    Public BS As BindingSource
    Dim Sqlstr As String = String.Empty
    'Public Sub New()
    '    myAdapter = DbAdapter.getInstance       
    'End Sub

    Public Sub New(ByVal sqlstr)
        myAdapter = DbAdapter.getInstance
        Me.Sqlstr = sqlstr
    End Sub

    Public Overloads Function loaddata() As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        BS = New BindingSource

        If myAdapter.TbgetDataSet(Sqlstr, DS) Then
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret
    End Function

    Public Overloads Function save(ByVal SaveDelegate1 As SaveDelegate) As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If SaveDelegate1.Invoke(mye, New EventArgs) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

    'Public Overloads Function save(ByVal mye As ContentBaseEventArgs) As Boolean

    'End Function

End Class
