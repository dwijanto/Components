Imports Npgsql
Imports Components.PublicClass
Imports System.Text
Public Class CMMFSPSPMModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Property cmmf As String    
    Public Property sp As String
    Public Property spm As String

    Private _ErrorMessage As String
    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cmmfspspm"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[cmmftext] like '*{0}*'  or [sp] like '*{0}*' or [spm] like '*{0}*'"
        End Get
    End Property


    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.*,cmmf::text as cmmftext from {0} u  order by {1};", TableName, SortField)
          
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        'AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        'Dim mytransaction As Npgsql.NpgsqlTransaction
        'Using conn As Object = myadapter.getConnection
        '    conn.Open()
        '    mytransaction = conn.BeginTransaction
        '    'Update
        '    Dim sqlstr = "sp_updatesbu"
        '    dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").SourceVersion = DataRowVersion.Original
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "pcmmf").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sp").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "lg").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "bu").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "act").SourceVersion = DataRowVersion.Current
        '    dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

        '    sqlstr = "sp_insertsbu"
        '    dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "pcmmf").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sp").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "lg").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "bu").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "act").SourceVersion = DataRowVersion.Current
        '    dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").Direction = ParameterDirection.InputOutput
        '    dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

        '    sqlstr = "sp_deletesbu"
        '    dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
        '    dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").Direction = ParameterDirection.Input
        '    dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

        '    dataadapter.InsertCommand.Transaction = mytransaction
        '    dataadapter.UpdateCommand.Transaction = mytransaction
        '    dataadapter.DeleteCommand.Transaction = mytransaction

        '    mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))
        '    mytransaction.Commit()
        '    myret = True
        'End Using
        Return myret
    End Function

    Public Sub ExportToExcel()
        Dim filename As String = "CMMFSPSPM-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim sqlstr = "select * from cmmfspspm order by cmmf,sp,spm;"
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools1)
    End Sub

    Public Function copyToDb(ByVal sqlstr As String, ByVal data As StringBuilder, ByRef result As Boolean) As Boolean
        _ErrorMessage = DbAdapter1.copy(sqlstr, data.ToString, result)
        Return result
    End Function
End Class
