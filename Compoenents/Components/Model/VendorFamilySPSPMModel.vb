Imports Npgsql
Imports Components.PublicClass
Imports System.Text
Public Class VendorFamilySPSPMModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Public Property familyid As String
    Public Property vendorcode As String
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
            Return "familyspspm"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[familyname] like '*{0}*' or [vendorname] like '*{0}*' or [sp] like '*{0}*' or [spm] like '*{0}*'"
        End Get
    End Property

    'Public Function GetBindingSource(ByVal criteria As String) As BindingSource
    '    Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
    '    Dim myret As Boolean = False
    '    Dim DS As New DataSet
    '    Using conn As Object = myadapter.getConnection
    '        conn.Open()
    '        'Dim sqlstr = String.Format("select null::integer as sbuid,null::text as sbuname, null::boolean as pcmmf,null::boolean as sp,null::boolean as lg,null::boolean as bu,null::boolean as cp,null::boolean as act  union all ( select u.* from {0} u {2} order by {1})", TableName, SortField, criteria)
    '        Dim sqlstr = String.Format("select null::integer as sbuid,null::text as sbuname, null::boolean as pcmmf,null::boolean as sp,null::boolean as lg,null::boolean as bu,null::boolean as cp,null::boolean as act  union all ( " &
    '                                   "with u as (select * from {0}  where(pcmmf Or sp Or lg Or bu Or cp Or act) order by sbuid) select * from u {2} order by {1})", TableName, SortField, criteria)
    '        dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
    '        dataadapter.SelectCommand.CommandType = CommandType.Text
    '        dataadapter.Fill(DS, TableName)
    '        myret = True
    '    End Using
    '    Dim BS As New BindingSource
    '    BS.DataSource = DS.Tables(0)
    '    Return BS
    'End Function

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()            
            'Dim sqlstr = String.Format("select u.* from {0} u  order by {1})", TableName, SortField)
            Dim sqlstr = String.Format("select fv.*,f.familyname::text,v.vendorname::text,fv.vendorcode::text || ' - '  || v.vendorname::text as vendorcodename,fv.familyid::text || ' - ' || f.familyname::text as familyidname from {0} fv" &
                                       " left join family f on f.familyid = fv.familyid" &
                                       " left join vendor v on v.vendorcode = fv.vendorcode  order by {1};", TableName, SortField)
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
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sp_updatesbu"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "pcmmf").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sp").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "lg").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "bu").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "act").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_insertsbu"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "pcmmf").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "sp").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "lg").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "bu").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cp").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "act").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletesbu"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sbuid").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Public Sub ExportToExcel()
        Dim filename As String = "FamilySPSPM-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim sqlstr = "select fss.familyid,fss.vendorcode,v.vendorname::text,fss.sp,fss.spm from familyspspm fss left join vendor v on v.vendorcode = fss.vendorcode left join family f on f.familyid = fss.familyid order by v.vendorcode,f.familyid;"
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools1)
    End Sub

    Public Function copyToDb(ByVal sqlstr As String, ByVal data As StringBuilder, ByRef result As Boolean) As Boolean
        _ErrorMessage = DbAdapter1.copy(sqlstr, data.ToString, result)
        Return result
    End Function


End Class
