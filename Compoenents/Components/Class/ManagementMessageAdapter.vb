Imports Npgsql
Public Class ManagementMessageAdapter
    Inherits ModelAdapter


    Dim myAdapter As DbAdapter
    Dim Sqlstr As String = String.Empty
    Dim savedelegate As SaveDelegate = AddressOf saverecord

    Public Sub New()
        MyBase.New("select * from managementmessage;")        
        myAdapter = DbAdapter.getInstance
    End Sub

    Public Sub New(ByVal sqlstr)
        MyBase.New(sqlstr)
        myAdapter = DbAdapter.getInstance
        Me.Sqlstr = sqlstr
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
            DS.Tables(0).TableName = "Management Message"
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
            Dim sqlstr = "sp_updatemanagementmessage"
            dataadapter.UpdateCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mgtmsg").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_insertmanagementmessage"
            dataadapter.InsertCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "mgtmsg").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletemanagementmessage"
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
