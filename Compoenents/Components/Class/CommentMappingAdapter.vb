Imports Npgsql
Public Class CommentMappingAdapter
    Inherits ModelAdapter
    Dim myAdapter As DbAdapter
    'Dim Sqlstr As String = String.Empty
    Dim savedelegate As SaveDelegate = AddressOf saverecord
    Public Sub New()
        MyBase.new("select * from commentmapping cm left join ;")
        myAdapter = DbAdapter.getInstance
    End Sub

    Public Sub New(ByVal sqlstr)
        MyBase.New(sqlstr)
        myAdapter = DbAdapter.getInstance
        'Me.Sqlstr = sqlstr
    End Sub
    Public Overloads Function loaddata() As Boolean
        Dim myret As Boolean = False
        If MyBase.loaddata Then
            Dim idx0(0) As DataColumn
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0
            DS.Tables(0).Columns(0).AutoIncrement = True
            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
            DS.Tables(0).Columns(0).AutoIncrementStep = -1
            DS.Tables(0).TableName = "CommentCode"
            myret = True
        End If
        Return myret
    End Function

    Public Overloads Function save() As Boolean
        Dim myret As Boolean = False
        If MyBase.save(savedelegate) Then
            myret = True
        End If
        Return myret
    End Function


    Public Overloads Function saverecord(ByRef sender As Object, ByRef e As EventArgs) As Boolean
        Dim mye As ContentBaseEventArgs = DirectCast(sender, ContentBaseEventArgs)
        Dim dataadapter As NpgsqlDataAdapter = New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection

            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sp_updatcommentmapping"
            dataadapter.UpdateCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comment").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "mgtmsgid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "catid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_insertcommentmapping"
            dataadapter.InsertCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "comment").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "mgtmsgid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "catid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "groupid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletecommentmapping"
            dataadapter.DeleteCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

End Class
