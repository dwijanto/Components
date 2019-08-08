Imports Npgsql
Partial Public Class Dbtools

#Region "SaveChanges"
    Public Function SaveChanges(ByRef DataSet As DataSet, ByVal sqlstr As String, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                Dim cmd = New NpgsqlCommand(sqlstr)
                Dim cmdbuilder = New NpgsqlCommandBuilder(DataAdapter)
                DataAdapter.SelectCommand = cmd
                DataAdapter.SelectCommand.Connection = conn
                DataAdapter.Update(DataSet.Tables(0))
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "GetDataSet"
    Public Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "TBProgram"
    
    Public Function TBProgramSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim cmd As NpgsqlCommand
        Dim param As NpgsqlParameter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                Try
                    conn.Open()
                    'Select
                    sqlstr = "select * from auth.tbprogram"
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from auth.tbprogram where programid = @programid"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").SourceVersion = DataRowVersion.Original

                    'Update
                    sqlstr = "Update auth.tbprogram set parentid = @parentid,myorder = @myorder,description = @description, programname = @programname,isactive = @isactive,icon = @icon, iconindex = @iconindex,members = @members" &
                             " where programid = @programid and latestupdate=@latestupdate"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@parentid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@myorder", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@description", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@programname", NpgsqlTypes.NpgsqlDbType.Text, 0, "programname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@isactive", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@icon", NpgsqlTypes.NpgsqlDbType.Text, 0, "icon").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@iconindex", NpgsqlTypes.NpgsqlDbType.Integer, 0, "iconindex").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@members", NpgsqlTypes.NpgsqlDbType.Text, 0, "members").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@latestupdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ, 0, "latestupdate").SourceVersion = DataRowVersion.Original

                    'insert
                    sqlstr = "insert into auth.tbprogram(parentid,myorder,description,programname,isactive,icon,iconindex,members) values " & _
                             "(@parentid,@myorder,@description,@programname,@isactive,@icon,@iconindex,@members);select currval('auth.tbprogram_programid_seq');"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("@parentid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "parentid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@myorder", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@description", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@programname", NpgsqlTypes.NpgsqlDbType.Text, 0, "programname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@isactive", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "isactive").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@icon", NpgsqlTypes.NpgsqlDbType.Text, 0, "icon").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@iconindex", NpgsqlTypes.NpgsqlDbType.Integer, 0, "iconindex").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@members", NpgsqlTypes.NpgsqlDbType.Text, 0, "members").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@programid", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "programid").Direction = ParameterDirection.Output

                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))
                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region

#Region "TBRoles"
    Public Function TBRolesSaveChanges(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim cmd As NpgsqlCommand
        Dim param As NpgsqlParameter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                Try
                    conn.Open()
                    'Select
                    sqlstr = "select * from auth.roles"
                    cmd = New NpgsqlCommand(sqlstr)
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn) 'cmd
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                    'Delete
                    sqlstr = "Delete from auth.roles where rolename = @rolename and applicationname=@applicationname"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").SourceVersion = DataRowVersion.Original

                    'Update
                    sqlstr = "Update auth.roles set rolename = @rolename,applicationname = @applicationname" &
                             " where rolename = @rolenameori and applicationname=@applicationnameori"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("@rolenameori", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("@applicationnameori", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").SourceVersion = DataRowVersion.Original
                    
                    'insert
                    sqlstr = "insert into auth.roles(rolename,applicationname) values " & _
                             "(@rolename,@applicationname)"

                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Text, 0, "rolename").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Text, 0, "applicationname").SourceVersion = DataRowVersion.Current                    

                    RecordAffected = DataAdapter.Update(DataSet.Tables(0))
                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
#End Region
End Class
