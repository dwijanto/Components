Imports System.Web.Security
Imports System.Configuration.Provider
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Diagnostics
Imports System.Web
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports Npgsql
Imports DJLib.NpgsqlMembershipProvider
Imports DJLib


Public Class NpgsqlRoleProvider
    Inherits RoleProvider

    Private conn As NpgsqlConnection
    Private eventSource As String = "NpgsqlRoleProvider"
    Private eventLog As String = "Application"
    Private exceptionMessage As String = "An exception occured. Please check the Event Log."
    Private pConnectionStringSetting As ConnectionStringSettings
    Private connectionString As String
    Private pWriteExceptionsToEventLog As Boolean = False
    Dim membership As New NpgsqlMembershipProvider

    Public Property WriteExceptionsToEventLog As Boolean
        Get
            Return pWriteExceptionsToEventLog
        End Get
        Set(ByVal value As Boolean)
            pWriteExceptionsToEventLog = value
        End Set
    End Property

    Public Sub New(ByVal name As String, ByVal config As System.Collections.Specialized.NameValueCollection)
        Initialize(name, config)
    End Sub

    Public Overrides Sub Initialize(ByVal name As String, ByVal config As System.Collections.Specialized.NameValueCollection)
        If config Is Nothing Then Throw New ArgumentNullException("config")
        If name Is Nothing OrElse name.Length = 0 Then name = "NpgsqlRoleProvider"
        If String.IsNullOrEmpty(config("description")) Then
            config.Remove("description")
            config.Add("description", "Npgsql Role provider")
        End If
        'Initialize the abstract base class.
        MyBase.Initialize(name, config)
        If config("applicationName") Is Nothing OrElse config("applicationName").Trim = "" Then

            pApplicationName = "\" 'System.Web.Hosting.HostingEnvironment.ApplicationVirtualPath
        Else
            pApplicationName = config("applicationName")
        End If

        If Not config("writeExceptionsToEventLog") Is Nothing Then
            If config("writeExceptionsToEventLog").ToUpper() = "TRUE" Then
                pWriteExceptionsToEventLog = True
            End If
        End If

        'Initialize NpgsqlConnection
        'pConnectionStringSetting = ConfigurationManager.ConnectionStrings(config("connectionStringName"))
        'If pConnectionStringSetting Is Nothing OrElse pConnectionStringSetting.ConnectionString.Trim() = "" Then
        '    Throw New ProviderException("Connection string cannot be blank.")
        'End If
        connectionString = membership.GetConfigValue(config("connectionStringName"), "")
        'Dim ConnectionStringSettings As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(config("connectionStringName"))
        'If ConnectionStringSettings Is Nothing OrElse ConnectionStringSettings.ConnectionString.Trim() = "" Then
        '    Throw New ProviderException("Connection string cannot be blank")
        'End If
        'connectionString = ConnectionStringSettings.ConnectionString
        'connectionString = pConnectionStringSetting.ConnectionString


    End Sub


    Public Overrides Sub AddUsersToRoles(ByVal usernames() As String, ByVal roleNames() As String)
        For Each rolename As String In roleNames
            If Not RoleExists(rolename) Then
                Throw New ProviderException("Role name not found.")
            End If
        Next

        For Each username As String In usernames
            If username.Contains(",") Then
                Throw New ArgumentException("User name cannot contain commas.")
            End If
            For Each rolename As String In roleNames
                If IsUserInRole(username, rolename) Then
                    Throw New ProviderException("User is already in role")
                End If
            Next
        Next

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "insert into usersinroles(username,rolename,applicationname) values(@username,@rolename,@applicationname)"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        Dim userParm As NpgsqlParameter = cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar)
        Dim roleParm As NpgsqlParameter = cmd.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Varchar)
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Dim trans As NpgsqlTransaction = Nothing
        Try
            conn.Open()
            trans = conn.BeginTransaction
            cmd.Transaction = trans
            For Each username As String In usernames
                For Each rolename As String In roleNames
                    userParm.Value = username
                    roleParm.Value = rolename
                    cmd.ExecuteNonQuery()
                Next
            Next
            trans.Commit()
        Catch ex As Exception
            Try
                trans.Rollback()
            Catch
            End Try
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "AddUsersToRoles")
            End If
        End Try


    End Sub


    Private pApplicationName As String
    Public Overrides Property ApplicationName As String
        Get
            Return pApplicationName
        End Get
        Set(ByVal value As String)
            pApplicationName = value
        End Set
    End Property

    Public Overrides Sub CreateRole(ByVal roleName As String)
        If roleName.Contains(",") Then
            Throw New ArgumentException("Role names cannot contain commas.")
        End If

        If RoleExists(roleName) Then
            Throw New ProviderException("Role name already exists.")
        End If

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "insert into roles(rolename,applicationname) values(@rolename,@applicationname)"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "CreateRole")
            Else
                Throw ex
            End If
        Finally
            conn.Close()
        End Try

    End Sub

    Public Overrides Function DeleteRole(ByVal roleName As String, ByVal throwOnPopulatedRole As Boolean) As Boolean
        If Not RoleExists(roleName) Then
            Throw New ProviderException("Role does not exist.")
        End If
        If throwOnPopulatedRole AndAlso GetUsersInRole(roleName).Length > 0 Then
            Throw New ProviderException("Cannot delete a populated role.")
        End If
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "delete from usersinroles where rolename = @rolename and applicationname = @applicationname;delete from roles where rolename = @rolename and applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@rolename", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName
        Dim tran As NpgsqlTransaction = Nothing
        Try
            conn.Open()
            tran = conn.BeginTransaction
            cmd.Transaction = tran
            cmd.ExecuteNonQuery()
            tran.Commit()
        Catch ex As Exception
            Try
                tran.Rollback()
            Catch
            End Try
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "DeleteRole")
                Return False
            Else
                Throw ex
            End If
        Finally
            conn.Close()
        End Try
        Return True
    End Function

    Public Overrides Function FindUsersInRole(ByVal roleName As String, ByVal usernameToMatch As String) As String()
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT username FROM usersinroles  " & _
                  "WHERE upper(username) LIKE @UsernameSearch AND upper(roleName) = @Rolename AND applicationname = @ApplicationName", conn)
        cmd.Parameters.Add("@UsernameSearch", NpgsqlTypes.NpgsqlDbType.Varchar).Value = usernameToMatch
        cmd.Parameters.Add("@RoleName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = pApplicationName

        Dim tmpUserNames As String = ""
        Dim reader As NpgsqlDataReader = Nothing
        Dim tmp(0) As String
        Try
            conn.Open()

            reader = cmd.ExecuteReader()
            Dim i As Integer = 0
            Do While reader.Read()
                ReDim Preserve tmp(i)
                tmp(i) = reader.GetString(0)
                'tmpUserNames &= reader.GetString(0) & ","
                i = i + 1
            Loop
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "FindUsersInRole")
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        'If tmpUserNames.Length > 0 Then
        '    ' Remove trailing comma.
        '    tmpUserNames = tmpUserNames.Substring(0, tmpUserNames.Length - 1)
        '    Return tmpUserNames.Split(CChar(","))
        'End If
        If tmp.Length > 0 Then
            Return tmp
        End If
        Return New String() {}

    End Function

    Public Overrides Function GetAllRoles() As String()
        Dim tmpRoleNames(0) As String

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT rolename FROM roles " & _
                  " WHERE ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader()
            Dim i As Integer = 0
            Do While reader.Read()
                ReDim Preserve tmpRoleNames(i)
                tmpRoleNames(i) = reader.GetString(0)
                i = i + 1
            Loop
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetAllRoles")
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        If tmpRoleNames.Length > 0 Then
            Return tmpRoleNames
        End If

        Return New String() {}

    End Function

    Public Overrides Function GetRolesForUser(ByVal username As String) As String()
        Dim tmpRoleNames(0) As String

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT rolename FROM usersinroles " & _
                " WHERE username = @Username AND applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()
            cmd.ExecuteScalar()
            reader = cmd.ExecuteReader()
            Dim i As Integer = 0
            Do While reader.Read()
                ReDim Preserve tmpRoleNames(i)
                tmpRoleNames(i) = reader.GetString(0)
                i = i + 1
            Loop
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetRolesForUser")
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        If tmpRoleNames.Length > 0 Then
            Return tmpRoleNames
        End If

        Return New String() {}

    End Function

    Public Overrides Function GetUsersInRole(ByVal roleName As String) As String()
        Dim tmpUserNames(0) As String

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT username FROM usersinroles " & _
                  " WHERE rolename = @Rolename AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Rolename", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader()
            Dim i As Integer = 0
            Do While reader.Read()
                ReDim Preserve tmpUserNames(i)
                tmpUserNames(i) = reader.GetString(0)
                i = i + 1
            Loop
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUsersInRole")
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        If tmpUserNames.Length > 0 Then
            Return tmpUserNames
        End If

        Return New String() {}

    End Function



    Public Overrides Function IsUserInRole(ByVal username As String, ByVal roleName As String) As Boolean
        Dim userIsInRole As Boolean = False

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT COUNT(*) FROM usersinroles " & _
                " WHERE username = @Username AND rolename = @Rolename AND ApplicationName = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@Rolename", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Try
            conn.Open()

            Dim numRecs As Integer = CType(cmd.ExecuteScalar(), Integer)

            If numRecs > 0 Then
                userIsInRole = True
            End If
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "IsUserInRole")
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        Return userIsInRole

    End Function

    Public Overrides Sub RemoveUsersFromRoles(ByVal usernames() As String, ByVal roleNames() As String)
        For Each rolename As String In roleNames
            If Not RoleExists(rolename) Then
                Throw New ProviderException("Role name not found.")
            End If
        Next

        For Each username As String In usernames
            For Each rolename As String In roleNames
                If Not IsUserInRole(username, rolename) Then
                    Throw New ProviderException("User is not in role.")
                End If
            Next
        Next

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("DELETE FROM usersinroles " & _
                " WHERE username = @Username AND rolename = @Rolename AND applicationname = @ApplicationName", conn)

        Dim userParm As NpgsqlParameter = cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar)
        Dim roleParm As NpgsqlParameter = cmd.Parameters.Add("@Rolename", NpgsqlTypes.NpgsqlDbType.Varchar)
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Dim tran As NpgsqlTransaction = Nothing

        Try
            conn.Open()
            tran = conn.BeginTransaction
            cmd.Transaction = tran

            For Each username As String In usernames
                For Each rolename As String In roleNames
                    userParm.Value = username
                    roleParm.Value = rolename
                    cmd.ExecuteNonQuery()
                Next
            Next

            tran.Commit()
        Catch e As NpgsqlException
            Try
                tran.Rollback()
            Catch
            End Try


            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "RemoveUsersFromRoles")
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

    End Sub

    Public Overrides Function RoleExists(ByVal roleName As String) As Boolean
        Dim exists As Boolean = False

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT COUNT(*) FROM roles " & _
                  " WHERE rolename = @Rolename AND applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@Rolename", NpgsqlTypes.NpgsqlDbType.Varchar).Value = roleName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName

        Try
            conn.Open()

            Dim numRecs As Integer = CType(cmd.ExecuteScalar(), Integer)

            If numRecs > 0 Then
                exists = True
            End If
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "RoleExists")
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        Return exists

    End Function



    Private Sub WriteToEventLog(ByVal e As NpgsqlException, ByVal action As String)
        Dim log As EventLog = New EventLog()
        log.Source = eventSource
        log.Log = eventLog

        Dim message As String = exceptionMessage & vbCrLf & vbCrLf
        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
        sb.Append("Action: " & action & vbCrLf & vbCrLf)
        sb.Append("Exception: " & e.ToString)
        log.WriteEntry(sb.ToString)
    End Sub

    Public Sub New()
        connectionString = My.Settings.connectionstring
    End Sub


End Class
