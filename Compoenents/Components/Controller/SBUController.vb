Public Class SBUController
    Implements IController
    Implements IToolbarAction
    Public Model As New SBUModel
    Public BS As BindingSource
    Public DS As DataSet

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property

    Public Function GetSBUBS(ByVal criteria As String) As BindingSource
        Return Model.GetBindingSource(criteria)
    End Function


    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New SBUModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("sbuid")
            DS.Tables(0).Columns("sbuid").AutoIncrement = True
            DS.Tables(0).Columns("sbuid").AutoIncrementSeed = 0
            DS.Tables(0).Columns("sbuid").AutoIncrementStep = -1
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret

    End Function

    Public Function Validate() As Boolean
        Dim myret = True
        BS.EndEdit()
        Dim ds2 As DataSet = DS.GetChanges
        For Each dr As DataRow In ds2.Tables(0).Rows
            dr.RowError = ""
            If dr.RowState <> DataRowState.Deleted Then
                If (Not dr.Item("pcmmf")) And (Not dr.Item("sp")) And (Not dr.Item("lg")) And (Not dr.Item("bu")) And (Not dr.Item("cp")) And (Not dr.Item("act")) Then
                    dr.RowError = "At least 1 item checkbox selected."
                    myret = False
                End If
            End If
        Next

        DS.Merge(ds2)

        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False

        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
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

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property

    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(ByVal value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub
End Class
