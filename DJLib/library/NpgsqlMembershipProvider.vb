Imports System
Imports System.Configuration
Imports System.Configuration.Provider
Imports System.Collections.Specialized
Imports System.Data
Imports System.Diagnostics
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports System.Web.Security
Imports System.Web.Configuration
Imports Npgsql
Imports DJLib.NpgsqlRoleProvider
'Imports System.Web.Mvc

'
' This provider works with the following schema for the table of user data.
' 
' CREATE TABLE Users
' (
'   PKID Guid NOT NULL PRIMARY KEY,
'   Username Text (255) NOT NULL,
'   ApplicationName Text (255) NOT NULL,
'   Email Text (128) NOT NULL,
'   Comment Text (255),
'   Password Text (128) NOT NULL,
'   PasswordQuestion Text (255),
'   PasswordAnswer Text (255),
'   IsApproved YesNo, 
'   LastActivityDate DateTime,
'   LastLoginDate DateTime,
'   LastPasswordChangedDate DateTime,
'   CreationDate DateTime, 
'   IsOnLine YesNo,
'   IsLockedOut YesNo,
'   LastLockedOutDate DateTime,
'   FailedPasswordAttemptCount Integer,
'   FailedPasswordAttemptWindowStart DateTime,
'   FailedPasswordAnswerAttemptCount Integer,
'   FailedPasswordAnswerAttemptWindowStart DateTime
' )






Public Class NpgsqlMembershipProvider
    Inherits MembershipProvider

    '
    'Global generated password length, generic exception message, event log info.
    '

    Private newPasswordLength As Integer = 8
    Private eventSource As String = "NpgsqlMembershipProvider"
    Private eventLog As String = "Application"
    Private exceptionMessage As String = "An exception occured. Please check the Event Log."
    Private connectionString As String

    '
    'Used When determining encryption key values
    '
    'Private machineKey As MachineKeySection

    '
    'if false, exception are thrown to the caller. if true, exceptions are written to the event log.
    '

    Private _UserIsOnlineTimeWindow As Integer
    Public Property UserIsOnlineTimeWindow() As Integer
        Get
            Return _UserIsOnlineTimeWindow
        End Get
        Set(ByVal value As Integer)
            _UserIsOnlineTimeWindow = value
        End Set
    End Property

    Private _WriteExceptionsToEventLog As Boolean
    Public Property WriteExceptionsToEventLog() As Boolean
        Get
            Return _WriteExceptionsToEventLog
        End Get
        Set(ByVal value As Boolean)
            _WriteExceptionsToEventLog = value
        End Set
    End Property
    Public Sub New()
        Me.New("NpgsqlMembershipProvider", AppConfig.CreateConfig)
    End Sub

    Public Sub New(ByVal name As String, ByVal config As System.Collections.Specialized.NameValueCollection)
        Initialize(name, config)
    End Sub

    'System.Configuration.Provider.ProviderBase.Initialize Method
    Public Overrides Sub Initialize(ByVal name As String, ByVal config As System.Collections.Specialized.NameValueCollection)
        '
        'initialize values from web.config
        '
        If config Is Nothing Then Throw New ArgumentNullException("config")
        If name Is Nothing OrElse name.Length = 0 Then name = "NpgsqlMembershipProvider"
        If String.IsNullOrEmpty(config("description")) Then
            config.Remove("description")
            config.Add("description", "Npgsql Membership Provider")
        End If
        MyBase.Initialize(name, config)

        _ApplicationName = "/" 'GetConfigValue(config("applicationName"), System.Web.Hosting.HostingEnvironment.ApplicationVirtualPath)
        _MaxInvalidPasswordAttemps = Convert.ToInt32(GetConfigValue(config("MaxInvalidPasswordAttempts"), "5"))
        _PasswordAttemptWindow = Convert.ToInt32(GetConfigValue(config("PasswordAttemptWindow"), "10"))
        _MinRequiredNonAlphanumericCharacters = Convert.ToInt32(GetConfigValue(config("MinRequiredNonAlphanumericCharacters"), "1"))
        _MinRequiredPasswordLength = Convert.ToInt32(GetConfigValue(config("MinRequiredPasswordLength"), "6"))
        _PasswordStrengthRegularExpression = Convert.ToString(GetConfigValue(config("PasswordStrengthRegularExpression"), ""))
        _EnablePasswordReset = Convert.ToBoolean(GetConfigValue(config("EnablePasswordReset"), "True"))
        _EnablePasswordRetrieval = Convert.ToBoolean(GetConfigValue(config("EnablePasswordRetrieval"), "True"))
        _RequiresQuestionAndAnswer = Convert.ToBoolean(GetConfigValue(config("RequiresQuestionAndAnswer"), "False"))
        _RequiresUniqueEmail = Convert.ToBoolean(GetConfigValue(config("RequiresUniqueEmail"), "False"))
        _WriteExceptionsToEventLog = Convert.ToBoolean(GetConfigValue(config("WriteExceptionsToEventLog"), "True"))
        _UserIsOnlineTimeWindow = Convert.ToInt32(GetConfigValue(config("UserIsOnlineTimeWindow"), "20"))
        Dim temp_format As String = config("passwordFormat")
        If temp_format Is Nothing Then
            temp_format = "Hashed"
        End If

        Select Case temp_format
            Case "Hashed"
                _PasswordFormat = MembershipPasswordFormat.Hashed
            Case "Encrypted"
                _PasswordFormat = MembershipPasswordFormat.Encrypted
            Case "Clear"
                _PasswordFormat = MembershipPasswordFormat.Clear
        End Select


        connectionString = GetConfigValue(config("connectionStringName"), "")
        'Dim ConnectionStringSettings As ConnectionStringSettings = ConfigurationManager.ConnectionStrings(config("connectionStringName"))
        'If ConnectionStringSettings Is Nothing OrElse ConnectionStringSettings.ConnectionString.Trim() = "" Then
        If connectionString = "" Then
            Throw New ProviderException("Connection string cannot be blank")
        End If
        'connectionString = ConnectionStringSettings.ConnectionString

        'Dim cfg As System.Configuration.Configuration = WebConfigurationManager.OpenWebConfiguration(System.Web.Hosting.HostingEnvironment.ApplicationVirtualPath)
        'machineKey = CType(cfg.GetSection("system.web/machineKey"), MachineKeySection)
        'If Not machineKey Is Nothing Then
        '    If machineKey.ValidationKey.Contains("AutoGenerate") Then
        '        If PasswordFormat <> MembershipPasswordFormat.Clear Then
        '            Throw New ProviderException("Hashed or Encrypted passwords are not supported with auto-generated keys")
        '        End If
        '    End If
        'End If
    End Sub



    Public Function GetConfigValue(ByVal configValue As String, ByVal defaultValue As String) As String
        If String.IsNullOrEmpty(configValue) Then Return defaultValue
        Return configValue
    End Function


    Private _ApplicationName As String
    Public Overrides Property ApplicationName As String
        Get
            Return _ApplicationName
        End Get
        Set(ByVal value As String)
            _ApplicationName = value
        End Set
    End Property

    Public Overrides Function ChangePassword(ByVal username As String, ByVal oldPassword As String, ByVal newPassword As String) As Boolean
        Dim myreturn As Boolean = False
        If Not ValidateUser(username, oldPassword) Then Return False

        Dim args As ValidatePasswordEventArgs = New ValidatePasswordEventArgs(username, newPassword, True)

        OnValidatingPassword(args)

        If args.Cancel Then
            If Not args.FailureInformation Is Nothing Then
                Throw args.FailureInformation
            Else
                Throw New ProviderException("Change password canceled due to New password validation failure.")
            End If
        End If

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "Update users set passwd = @passwd,lastpasswordchangeddate=@lastpasswordchangeddate,lastactivitydate=@lastactivitydate " & _
                               " where username=@username and applicationname=@applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        Dim lasttime As DateTime = DateTime.Now
        cmd.Parameters.Add("@passwd", NpgsqlTypes.NpgsqlDbType.Varchar).Value = EncodePassword(newPassword)
        cmd.Parameters.Add("@lastpasswordchangeddate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = lasttime
        cmd.Parameters.Add("@lastactivitydate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = lasttime
        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim rowsaffected As Integer = 0
        Try
            conn.Open()
            rowsaffected = cmd.ExecuteNonQuery()
            If rowsaffected > 0 Then myreturn = True
        Catch ex As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "ChangePassword")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            conn.Close()
        End Try

        Return myreturn

    End Function

    Public Overrides Function ChangePasswordQuestionAndAnswer(ByVal username As String, ByVal password As String, ByVal newPasswordQuestion As String, ByVal newPasswordAnswer As String) As Boolean
        Dim myReturn As Boolean = False

        If ValidateUser(username, password) Then
            Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
            Dim sqlstr = "Update users set passwordquestion = @passwordquestion,passwordanswer=@passwordanswer" & _
                               " where username=@username and applicationname=@applicationname"
            Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("@passwordquestion", NpgsqlTypes.NpgsqlDbType.Varchar).Value = newPasswordQuestion
            cmd.Parameters.Add("@passwordanswer", NpgsqlTypes.NpgsqlDbType.Varchar).Value = newPasswordAnswer
            cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
            cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

            Dim rowAffected As Integer = 0
            Try
                conn.Open()
                rowAffected = cmd.ExecuteNonQuery
                myReturn = True
            Catch ex As NpgsqlException
                If WriteExceptionsToEventLog Then
                    WriteToEventLog(ex, "ChangePasswordQuestionAndAnswer")
                Else
                    Throw New ProviderException(exceptionMessage)
                End If
            Finally
                conn.Close()
            End Try
        End If
        Return myReturn
    End Function

    Public Overrides Function CreateUser(ByVal username As String, ByVal password As String, ByVal email As String, ByVal passwordQuestion As String, ByVal passwordAnswer As String, ByVal isApproved As Boolean, ByVal providerUserKey As Object, ByRef status As System.Web.Security.MembershipCreateStatus) As System.Web.Security.MembershipUser

        If passwordAnswer Is Nothing Then passwordAnswer = ""
        Dim Args As ValidatePasswordEventArgs = New ValidatePasswordEventArgs(username, password, True)
        OnValidatingPassword(Args)
        If Args.Cancel Then
            status = MembershipCreateStatus.InvalidPassword
            Return Nothing
        End If


        If RequiresUniqueEmail AndAlso GetUserNameByEmail(email) <> "" Then
            status = MembershipCreateStatus.DuplicateEmail
            Return Nothing
        End If

        Dim u As MembershipUser = GetUser(username, False)
        If u Is Nothing Then
            Dim createdate As DateTime = DateTime.Now
            If providerUserKey Is Nothing Then
                providerUserKey = Guid.NewGuid()
            Else
                If Not TypeOf providerUserKey Is Guid Then
                    status = MembershipCreateStatus.InvalidProviderUserKey
                    Return Nothing
                End If
            End If
            Dim con As NpgsqlConnection = New NpgsqlConnection(connectionString)
            Dim sqlstr As String = "insert into users (pkid,username,passwd,email,passwordquestion,passwordanswer,isapproved,comments,creationdate,lastpasswordchangeddate,lastactivitydate," & _
                                   "applicationname,islockedout,lastlockedoutdate,failedpasswordattemptcount,failedpasswordattemptwindowstart,failedpasswordanswerattemptcount," & _
                                   "failedpasswordanswerattemptwindowstart) values(@pkid,@username,@password,@email,@passwordquestion,@passwordanswer,@isapproved,@comments," & _
                                   "@creationdate,@lastpasswordchangeddate,@lastactivitydate,@applicationname,@islockedout,@lastlockedoutdate,@failedpasswordattemptcount," & _
                                   "@failedpasswordattemptwindowstart,@failedpasswordanswerattemptcount,@failedpasswordanswerattemptwindowstart)"
            Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, con)
            cmd.Parameters.Add("@pkid", NpgsqlTypes.NpgsqlDbType.Uuid).Value = providerUserKey
            cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
            cmd.Parameters.Add("@password", NpgsqlTypes.NpgsqlDbType.Varchar).Value = EncodePassword(password)
            cmd.Parameters.Add("@email", NpgsqlTypes.NpgsqlDbType.Varchar).Value = email
            cmd.Parameters.Add("@passwordquestion", NpgsqlTypes.NpgsqlDbType.Varchar).Value = passwordQuestion
            cmd.Parameters.Add("@passwordanswer", NpgsqlTypes.NpgsqlDbType.Varchar).Value = EncodePassword(passwordAnswer)
            cmd.Parameters.Add("@isapproved", NpgsqlTypes.NpgsqlDbType.Boolean).Value = isApproved
            cmd.Parameters.Add("@comments", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ""
            cmd.Parameters.Add("@creationdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate
            cmd.Parameters.Add("@lastpasswordchangeddate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate
            cmd.Parameters.Add("@lastactivitydate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate
            cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
            cmd.Parameters.Add("@islockedout", NpgsqlTypes.NpgsqlDbType.Boolean).Value = False
            cmd.Parameters.Add("@lastlockedoutdate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate
            cmd.Parameters.Add("@failedpasswordattemptcount", NpgsqlTypes.NpgsqlDbType.Integer).Value = 0
            cmd.Parameters.Add("@failedpasswordattemptwindowstart", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate
            cmd.Parameters.Add("@failedpasswordanswerattemptcount", NpgsqlTypes.NpgsqlDbType.Integer).Value = 0
            cmd.Parameters.Add("@failedpasswordanswerattemptwindowstart", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = createdate

            Try
                con.Open()

                Dim recadded As Integer = cmd.ExecuteNonQuery
                If recadded > 0 Then
                    status = MembershipCreateStatus.Success
                Else
                    status = MembershipCreateStatus.UserRejected
                End If
            Catch ex As NpgsqlException
                If WriteExceptionsToEventLog Then
                    WriteToEventLog(ex, "CreateUser")
                End If
                status = MembershipCreateStatus.ProviderError
            Finally
                con.Close()
            End Try
        Else
            status = MembershipCreateStatus.DuplicateUserName

        End If
        Return Nothing
    End Function

    Public Overrides Function DeleteUser(ByVal username As String, ByVal deleteAllRelatedData As Boolean) As Boolean
        Dim myReturn As Boolean = False

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "Delete from users where username = @username and applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
        Dim rowsAffected As Integer = 0

        Try
            conn.Open()
            rowsAffected = cmd.ExecuteNonQuery()
            If rowsAffected > 0 Then
                If deleteAllRelatedData Then
                    'Process command to delete all related data for the user in the database
                End If
                myReturn = True
            End If
        Catch ex As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "DeleteUser")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            conn.Close()
        End Try

        Return myReturn
    End Function
    Private _EnablePasswordReset As Boolean
    Public Overrides ReadOnly Property EnablePasswordReset As Boolean
        Get
            Return _EnablePasswordReset
        End Get
    End Property

    Private _EnablePasswordRetrieval As Boolean
    Public Overrides ReadOnly Property EnablePasswordRetrieval As Boolean
        Get
            Return _EnablePasswordRetrieval
        End Get
    End Property

    Public Overrides Function FindUsersByEmail(ByVal emailToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "select count(*) from users where email like @emailtomatch and applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@emailtomatch", NpgsqlTypes.NpgsqlDbType.Varchar).Value = emailToMatch
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
        Dim users As MembershipUserCollection = New MembershipUserCollection()
        Dim reader As NpgsqlDataReader = Nothing
        totalRecords = 0
        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar())
            cmd.CommandText = "select pkid,username,email,passwordquestion,comments,isapproved,islockedout,creationdate,lastlogindate,lastactivitydate,lastpasswordchangeddate,lastlockedoutdate from users" & _
                              " where email like @emailtomatch and applicationname = @applicationname order by username asc"
            reader = cmd.ExecuteReader
            Dim counter As Integer = 0
            Dim startIndex As Integer = pageIndex * pageIndex
            Dim endIndex As Integer = startIndex + pageSize - 1

            Do While reader.Read()
                If counter >= startIndex Then
                    Dim u As MembershipUser = GetUserFromReader(reader)
                    users.Add(u)
                End If
                If counter >= endIndex Then cmd.Cancel()
                counter += 1
            Loop

        Catch ex As NpgsqlException
            If WriteExceptionsToEventLog() Then
                WriteToEventLog(ex, "FindUserByEmail")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            If Not reader Is Nothing Then
                reader.Close()
            End If
            conn.Close()
        End Try
        Return users
    End Function

    Public Overrides Function FindUsersByName(ByVal usernameToMatch As String, ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "select count(*) from users where upper(username) like @username and applicationname=@applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = UCase(usernameToMatch)
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
        Dim users As MembershipUserCollection = New MembershipUserCollection
        Dim reader As NpgsqlDataReader = Nothing
        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar)
            If totalRecords > 0 Then
                cmd.CommandText = "select pkid,username,email,passwordquestion,comments,isapproved,islockedout,creationdate,lastlogindate,lastactivitydate,lastpasswordchangeddate,lastlockedoutdate from users" & _
                              " where upper(username) like @username and applicationname = @applicationname order by username asc"

                reader = cmd.ExecuteReader()
                Dim counter As Integer = 0
                Dim startindex As Integer = pageSize * pageIndex
                Dim endIndex As Integer = startindex + pageSize - 1
                Do While reader.Read()
                    If counter >= startindex Then
                        Dim u As MembershipUser = GetUserFromReader(reader)
                        users.Add(u)
                    End If
                    If counter >= endIndex Then cmd.Cancel()
                    counter += 1
                Loop
            End If
        Catch ex As NpgsqlException
            If WriteExceptionsToEventLog() Then
                WriteToEventLog(ex, "FindUserByName")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try
        Return users
    End Function

    Public Overrides Function GetAllUsers(ByVal pageIndex As Integer, ByVal pageSize As Integer, ByRef totalRecords As Integer) As System.Web.Security.MembershipUserCollection
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "SELECT Count(*) FROM users WHERE applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName
        Dim users As MembershipUserCollection = New MembershipUserCollection()
        Dim reader As NpgsqlDataReader = Nothing
        totalRecords = 0

        Try
            conn.Open()
            totalRecords = CInt(cmd.ExecuteScalar())

            If totalRecords <= 0 Then Return users

            cmd.CommandText = "SELECT pkid, username, email, passwordquestion, comments, isapproved, islockedout, creationdate, lastlogindate,lastactivitydate, lastpasswordchangeddate, lastlockedoutdate " & _
                     " FROM users  " & _
                     " WHERE applicationname = @applicationname ORDER BY username asc"

            reader = cmd.ExecuteReader()

            Dim counter As Integer = 0
            Dim startIndex As Integer = pageSize * pageIndex
            Dim endIndex As Integer = startIndex + pageSize - 1

            Do While reader.Read()
                If counter >= startIndex Then
                    Dim u As MembershipUser = GetUserFromReader(reader)
                    users.Add(u)
                End If

                If counter >= endIndex Then cmd.Cancel()

                counter += 1
            Loop
        Catch ex As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "GetAllUsers")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        Return users

    End Function

    Public Overrides Function GetNumberOfUsersOnline() As Integer

        Dim onlineSpan As TimeSpan = New TimeSpan(0, UserIsOnlineTimeWindow, 0)
        Dim compareTime As DateTime = DateTime.Now.Subtract(onlineSpan)

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "SELECT Count(*) FROM Users WHERE lastactivitydate > @comparedate AND applicationname = @applicationname "
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)

        cmd.Parameters.Add("@comparedate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = compareTime
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim numOnline As Integer = 0
        Try
            conn.Open()
            numOnline = CInt(cmd.ExecuteScalar())
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetNumberOfUsersOnline")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        Return numOnline

    End Function

    Public Overrides Function GetPassword(ByVal username As String, ByVal answer As String) As String
        If Not EnablePasswordRetrieval Then
            Throw New ProviderException("Password Retrieval Not Enabled.")
        End If

        If PasswordFormat = MembershipPasswordFormat.Hashed Then
            Throw New ProviderException("Cannot retrieve Hashed passwords.")
        End If

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "SELECT passwd, passwordanswer, islockedout FROM users WHERE username = @username AND applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)

        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim password As String = ""
        Dim passwordAnswer As String = ""
        Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()

                If reader.GetBoolean(2) Then _
                  Throw New MembershipPasswordException("The supplied user is locked out.")

                password = reader.GetString(0)
                passwordAnswer = reader.GetString(1)
            Else
                Throw New MembershipPasswordException("The supplied user name is not found.")
            End If
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetPassword")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try


        If RequiresQuestionAndAnswer AndAlso Not CheckPassword(answer, passwordAnswer) Then
            UpdateFailureCount(username, "passwordAnswer")
            Throw New MembershipPasswordException("Incorrect password answer.")
        End If


        If PasswordFormat = MembershipPasswordFormat.Encrypted Then
            password = UnEncodePassword(password)
        End If

        Return password

    End Function

    Private Sub UpdateFailureCount(ByVal username As String, ByVal failureType As String)

        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "SELECT failedpasswordattemptcount,   failedpasswordattemptwindowstart,  failedpasswordanswerattemptcount,   failedpasswordanswerattemptwindowstart   FROM users  " & _
                                "  WHERE username = @username AND applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)

        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@applicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim reader As NpgsqlDataReader = Nothing
        Dim windowStart As DateTime = New DateTime()
        Dim failureCount As Integer = 0

        Try
            conn.Open()

            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)

            If reader.HasRows Then
                reader.Read()

                If failureType = "password" Then
                    failureCount = reader.GetInt32(0)
                    windowStart = reader.GetDateTime(1)
                End If

                If failureType = "passwordAnswer" Then
                    failureCount = reader.GetInt32(2)
                    windowStart = reader.GetDateTime(3)
                End If
            End If

            reader.Close()

            Dim windowEnd As DateTime = windowStart.AddMinutes(PasswordAttemptWindow)

            If failureCount = 0 OrElse DateTime.Now > windowEnd Then
                ' First password failure or outside of PasswordAttemptWindow. 
                ' Start a New password failure count from 1 and a New window starting now.

                If failureType = "password" Then _
                  cmd.CommandText = "UPDATE users   SET failedpasswordattemptcount = @Count,   failedpasswordattemptwindowstart = @WindowStart " & _
                                    "  WHERE username = @Username AND applicationname = @ApplicationName"

                If failureType = "passwordAnswer" Then _
                  cmd.CommandText = "UPDATE users   SET failedpasswordanswerattemptcount = @Count, failedpasswordanswerattemptwindowstart = @WindowStart " & _
                                    "  WHERE username = @Username AND applicationame = @ApplicationName"

                cmd.Parameters.Clear()

                cmd.Parameters.Add("@Count", NpgsqlTypes.NpgsqlDbType.Integer).Value = 1
                cmd.Parameters.Add("@WindowStart", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTime.Now
                cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
                cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

                If cmd.ExecuteNonQuery() < 0 Then _
                  Throw New ProviderException("Unable to update failure count and window start.")
            Else
                failureCount += 1

                If failureCount >= MaxInvalidPasswordAttempts Then
                    ' Password attempts have exceeded the failure threshold. Lock out
                    ' the user.

                    cmd.CommandText = "UPDATE users SET islockedout = @IsLockedOut, lastlockedoutdate = @LastLockedOutDate " & _
                                      "  WHERE username = @Username AND applicationName = @ApplicationName"

                    cmd.Parameters.Clear()

                    cmd.Parameters.Add("@IsLockedOut", NpgsqlTypes.NpgsqlDbType.Bit).Value = True
                    cmd.Parameters.Add("@LastLockedOutDate", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTime.Now
                    cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
                    cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

                    If cmd.ExecuteNonQuery() < 0 Then _
                      Throw New ProviderException("Unable to lock out user.")
                Else
                    ' Password attempts have not exceeded the failure threshold. Update
                    ' the failure counts. Leave the window the same.

                    If failureType = "password" Then _
                      cmd.CommandText = "UPDATE users  " & _
                                        "  SET failedpasswordattemptcount = @Count" & _
                                        "  WHERE username = @Username AND ApplicationName = @ApplicationName"

                    If failureType = "passwordAnswer" Then _
                      cmd.CommandText = "UPDATE users  " & _
                                        "  SET failedpasswordanswerattemptcount = @Count" & _
                                        "  WHERE username = @username AND applicationname = @ApplicationName"

                    cmd.Parameters.Clear()

                    cmd.Parameters.Add("@Count", NpgsqlTypes.NpgsqlDbType.Integer).Value = failureCount
                    cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
                    cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

                    If cmd.ExecuteNonQuery() < 0 Then _
                      Throw New ProviderException("Unable to update failure count.")
                End If
            End If
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UpdateFailureCount")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try
    End Sub



    Public Overloads Overrides Function GetUser(ByVal providerUserKey As Object, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT pkid, username, email, passwordquestion," & _
              " comments, isapproved, islockedout, creationdate, lastlogindate," & _
              " lastactivitydate, lastpasswordchangeddate, lastlockedoutdate" & _
              " FROM users  WHERE pkid = @pkid", conn)

        cmd.Parameters.Add("@pkid", NpgsqlTypes.NpgsqlDbType.Uuid).Value = providerUserKey

        Dim u As MembershipUser = Nothing
        Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                reader.Read()
                u = GetUserFromReader(reader)

                If userIsOnline Then
                    Dim updateCmd As NpgsqlCommand = New NpgsqlCommand("UPDATE users  " & _
                              "SET lastactivityaate = @lastactivitydate " & _
                              "WHERE pkid = @pkid", conn)

                    updateCmd.Parameters.Add("@lastactivitydate", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTime.Now
                    updateCmd.Parameters.Add("@PKID", NpgsqlTypes.NpgsqlDbType.Uuid).Value = providerUserKey

                    updateCmd.ExecuteNonQuery()
                End If
            End If
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUser(Object, Boolean)")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()

            conn.Close()
        End Try

        Return u

    End Function

    Public Overloads Overrides Function GetUser(ByVal username As String, ByVal userIsOnline As Boolean) As System.Web.Security.MembershipUser
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr As String = "SELECT pkid, username, email, passwordquestion," & _
                  " comments, isapproved, islockedout, creationdate, lastlogindate," & _
                  " lastactivitydate, lastpasswordchangeddate, lastlockedoutdate" & _
                  " FROM users  WHERE username = @username AND applicationname = @applicationname"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
        Dim u As MembershipUser = Nothing
        Dim reader As NpgsqlDataReader = Nothing
        Try
            conn.Open()
            reader = cmd.ExecuteReader
            If reader.HasRows Then
                reader.Read()
                u = GetUserFromReader(reader)
                reader.Close()
                If userIsOnline Then
                    sqlstr = "UPDATE users  " & _
                                  "SET lastactivitydate = @lastactivitydate " & _
                                  "WHERE username = @username AND applicationname = @applicationname"
                    Dim updatecmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
                    updatecmd.Parameters.Add("@lastactivitydate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = DateTime.Now
                    updatecmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
                    updatecmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName
                    updatecmd.ExecuteNonQuery()
                End If
            End If
        Catch ex As Exception
            If WriteExceptionsToEventLog Then
                WriteToEventLog(ex, "GetUser(string,Boolean)")
                Throw New ProviderException(exceptionMessage)
            Else
                Throw ex
            End If
        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try
        Return u
    End Function

    Private Function GetUserFromReader(ByVal reader As NpgsqlDataReader) As MembershipUser
        Dim providerUserKey As Object = reader.GetValue(0)
        Dim username As String = reader.GetString(1)
        Dim email As String = reader.GetString(2)

        Dim passwordQuestion As String = ""
        If Not reader.GetValue(3) Is DBNull.Value Then
            passwordQuestion = reader.GetValue(3)
        End If

        Dim comments As String = ""
        If Not reader.GetValue(4) Is DBNull.Value Then
            comments = reader.GetString(4)
        End If

        Dim isapproved As Boolean = reader.GetBoolean(5)
        Dim islockedout As Boolean = reader.GetBoolean(6)
        Dim creationDate As DateTime = reader.GetDateTime(7)

        Dim lastlogindate As DateTime = New DateTime
        If Not reader.GetValue(8) Is DBNull.Value Then
            lastlogindate = reader.GetDateTime(8)
        End If

        Dim lastactivitydate As DateTime = reader.GetDateTime(9)
        Dim lastpasswordChangedDate As DateTime = reader.GetDateTime(10)
        Dim lastlockedoutdate As DateTime = New DateTime()
        If Not reader.GetValue(11) Is DBNull.Value Then
            lastlockedoutdate = reader.GetDateTime(11)
        End If
        Dim u As MembershipUser = New MembershipUser(Me.Name, username, providerUserKey, email, passwordQuestion, comments, isapproved, islockedout, creationDate, lastlogindate, lastactivitydate, lastpasswordChangedDate, lastlockedoutdate)
        'Dim u As MembershipUser = New MembershipUser(Me.Name, username, providerUserKey, email, passwordQuestion, comments, isapproved, islockedout, creationDate, lastlogindate, lastactivitydate, lastpasswordChangedDate, lastlockedoutdate)
        Return u
    End Function

    Public Overrides Function GetUserNameByEmail(ByVal email As String) As String
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim sqlstr = "SELECT username FROM users  WHERE email = @Email AND applicationname = @ApplicationName"
        Dim cmd As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)

        cmd.Parameters.Add("@Email", NpgsqlTypes.NpgsqlDbType.Varchar).Value = email
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar, 255).Value = _ApplicationName

        Dim username As String = ""

        Try
            conn.Open()

            username = cmd.ExecuteScalar().ToString()
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "GetUserNameByEmail")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Catch
        Finally
            conn.Close()
        End Try

        If username Is Nothing Then _
          username = ""

        Return username

    End Function

    Private _MaxInvalidPasswordAttemps As Integer
    Public Overrides ReadOnly Property MaxInvalidPasswordAttempts As Integer
        Get
            Return _MaxInvalidPasswordAttemps
        End Get
    End Property

    Private _MinRequiredNonAlphanumericCharacters As Integer
    Public Overrides ReadOnly Property MinRequiredNonAlphanumericCharacters As Integer
        Get
            Return _MinRequiredNonAlphanumericCharacters
        End Get
    End Property

    Private _MinRequiredPasswordLength As Integer
    Public Overrides ReadOnly Property MinRequiredPasswordLength As Integer
        Get
            Return _MinRequiredPasswordLength
        End Get
    End Property

    Private _PasswordAttemptWindow As Integer
    Public Overrides ReadOnly Property PasswordAttemptWindow As Integer
        Get
            Return _PasswordAttemptWindow
        End Get
    End Property

    Private _PasswordFormat As System.Web.Security.MembershipPasswordFormat
    Public Overrides ReadOnly Property PasswordFormat As System.Web.Security.MembershipPasswordFormat
        Get
            Return _PasswordFormat
        End Get
    End Property

    Private _PasswordStrengthRegularExpression As String
    Public Overrides ReadOnly Property PasswordStrengthRegularExpression As String
        Get
            Return _PasswordStrengthRegularExpression
        End Get
    End Property

    Private _RequiresQuestionAndAnswer As Boolean
    Public Overrides ReadOnly Property RequiresQuestionAndAnswer As Boolean
        Get
            Return _RequiresQuestionAndAnswer
        End Get
    End Property

    Private _RequiresUniqueEmail As Boolean
    Public Overrides ReadOnly Property RequiresUniqueEmail As Boolean
        Get
            Return _RequiresUniqueEmail
        End Get
    End Property


    Public Overrides Function ResetPassword(ByVal username As String, ByVal answer As String) As String
        If Not EnablePasswordReset Then
            Throw New NotSupportedException("Password Reset is not enabled.")
        End If

        If answer Is Nothing AndAlso RequiresQuestionAndAnswer Then
            UpdateFailureCount(username, "passwordAnswer")

            Throw New ProviderException("Password answer required for password Reset.")
        End If

        Dim newPassword As String = answer
        'System.Web.Security.Membership.GeneratePassword(newPasswordLength, MinRequiredNonAlphanumericCharacters)


        Dim Args As ValidatePasswordEventArgs = _
          New ValidatePasswordEventArgs(username, newPassword, True)

        OnValidatingPassword(Args)

        If Args.Cancel Then
            If Not Args.FailureInformation Is Nothing Then
                Throw Args.FailureInformation
            Else
                Throw New MembershipPasswordException("Reset password canceled due to password validation failure.")
            End If
        End If


        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("SELECT passwordanswer, islockedout FROM users  WHERE username = @UserName AND applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim rowsAffected As Integer = 0
        Dim passwordAnswer As String = ""
        'Dim reader As NpgsqlDataReader = Nothing

        Try
            conn.Open()
            Using reader As NpgsqlDataReader = cmd.ExecuteReader(CommandBehavior.SingleRow)
                If reader.HasRows Then
                    reader.Read()

                    If reader.GetBoolean(1) Then _
                      Throw New MembershipPasswordException("The supplied user is locked out.")

                    passwordAnswer = reader.GetString(0)
                Else
                    Throw New MembershipPasswordException("The supplied user name is not found.")
                End If
            End Using

            If RequiresQuestionAndAnswer AndAlso Not CheckPassword(answer, passwordAnswer) Then
                UpdateFailureCount(username, "passwordAnswer")

                Throw New MembershipPasswordException("Incorrect password answer.")
            End If

            Dim updateCmd As NpgsqlCommand = New NpgsqlCommand("UPDATE Users  SET passwd = @Password, lastpasswordchangeddate = @LastPasswordChangedDate" & _
                " WHERE username = @Username AND applicationname = @ApplicationName AND islockedout = False", conn)

            updateCmd.Parameters.Add("@Password", NpgsqlTypes.NpgsqlDbType.Varchar).Value = EncodePassword(newPassword)
            updateCmd.Parameters.Add("@LastPasswordChangedDate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = DateTime.Now
            updateCmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
            updateCmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

            rowsAffected = updateCmd.ExecuteNonQuery()
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "ResetPassword")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            'If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try

        If rowsAffected > 0 Then
            Return newPassword
        Else
            Throw New MembershipPasswordException("User not found, or user is locked out. Password not Reset.")
        End If

    End Function

    Public Overrides Function UnlockUser(ByVal userName As String) As Boolean
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("UPDATE users SET islockedout = False, lastlockedoutdate = @LastLockedOutDate " & _
                                          " WHERE username = @Username AND applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@LastLockedOutDate", NpgsqlTypes.NpgsqlDbType.Date).Value = DateTime.Now
        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = userName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName

        Dim rowsAffected As Integer = 0

        Try
            conn.Open()

            rowsAffected = cmd.ExecuteNonQuery()
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UnlockUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

        If rowsAffected > 0 Then _
          Return True

        Return False

    End Function

    Public Overrides Sub UpdateUser(ByVal user As System.Web.Security.MembershipUser)
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("UPDATE Users SET email = @Email, comments = @Comment,isapproved = @IsApproved" & _
                " WHERE username = @Username AND applicationname = @ApplicationName", conn)

        cmd.Parameters.Add("@Email", NpgsqlTypes.NpgsqlDbType.Varchar).Value = user.Email
        cmd.Parameters.Add("@Comment", NpgsqlTypes.NpgsqlDbType.Varchar).Value = user.Comment
        cmd.Parameters.Add("@IsApproved", NpgsqlTypes.NpgsqlDbType.Bit).Value = user.IsApproved
        cmd.Parameters.Add("@Username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = user.UserName
        cmd.Parameters.Add("@ApplicationName", NpgsqlTypes.NpgsqlDbType.Varchar).Value = _ApplicationName


        Try
            conn.Open()

            cmd.ExecuteNonQuery()
        Catch e As NpgsqlException
            If WriteExceptionsToEventLog Then
                WriteToEventLog(e, "UpdateUser")

                Throw New ProviderException(exceptionMessage)
            Else
                Throw e
            End If
        Finally
            conn.Close()
        End Try

    End Sub

    Public Overrides Function ValidateUser(ByVal username As String, ByVal password As String) As Boolean

        Dim isValid As Boolean = False
        Dim conn As NpgsqlConnection = New NpgsqlConnection(connectionString)
        Dim cmd As NpgsqlCommand = New NpgsqlCommand("select passwd,isApproved from users where username = @username and applicationname=@applicationname and islockedout = false", conn)
        cmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
        cmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName
        Dim reader As NpgsqlDataReader = Nothing
        Dim isApproved As Boolean = False
        Dim pwd As String = ""
        Try
            conn.Open()
            reader = cmd.ExecuteReader(CommandBehavior.SingleRow)
            If reader.HasRows Then
                reader.Read()
                pwd = reader.GetString(0)
                isApproved = reader.GetBoolean(1)
            Else
                Err.Raise(1)
            End If
            reader.Close()

            If CheckPassword(password, pwd) Then
                If isApproved Then isValid = True
                Dim lasttime = DateTime.Now
                Dim updateCmd As NpgsqlCommand = New NpgsqlCommand("update users set lastlogindate = @lastlogindate,lastactivitydate = @lastactivitydate where username = @username and applicationname=@applicationname ", conn)
                updateCmd.Parameters.Add("@lastlogindate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = lasttime
                updateCmd.Parameters.Add("@lastactivitydate", NpgsqlTypes.NpgsqlDbType.TimestampTZ).Value = lasttime
                updateCmd.Parameters.Add("@username", NpgsqlTypes.NpgsqlDbType.Varchar).Value = username
                updateCmd.Parameters.Add("@applicationname", NpgsqlTypes.NpgsqlDbType.Varchar).Value = ApplicationName
                updateCmd.ExecuteNonQuery()
                'get role
                'Dim sampleprincipal As New SampleIPrincipal(username, pwd)
                'Dim instance As New AuthorizeAttribute
                'instance.Users = username
            End If
        Catch ex As Exception

        Finally
            If Not reader Is Nothing Then reader.Close()
            conn.Close()
        End Try
        Return isValid
    End Function


    Private Function CheckPassword(ByVal password As String, ByVal dbPassword As String) As Boolean
        Dim myReturn As Boolean = False

        Dim pass1 As String = password
        Dim pass2 As String = dbPassword

        Select Case PasswordFormat
            Case MembershipPasswordFormat.Encrypted
                pass2 = UnEncodePassword(dbPassword)
            Case MembershipPasswordFormat.Hashed
                pass1 = EncodePassword(password)
        End Select
        If pass1 = pass2 Then
            myReturn = True
        End If
        Return myReturn
    End Function

    Private Function EncodePassword(ByRef password As String) As String
        Dim encodedPassword As String = password
        Select Case PasswordFormat
            Case MembershipPasswordFormat.Clear
            Case MembershipPasswordFormat.Encrypted
                encodedPassword = Convert.ToBase64String(EncryptPassword(Encoding.Unicode.GetBytes(password)))
            Case MembershipPasswordFormat.Hashed
                Dim hash As HMACSHA1 = New HMACSHA1()
                hash.Key = HexToByte(My.Settings.validationkey) 'HexToByte(machineKey.ValidationKey)
                encodedPassword = Convert.ToBase64String(hash.ComputeHash(Encoding.Unicode.GetBytes(password)))
            Case Else
                Throw New ProviderException("Unsupported password format")
        End Select
        Return encodedPassword
    End Function

    Private Function UnEncodePassword(ByRef encodedPassword As String) As String
        Dim password As String = encodedPassword
        Select Case PasswordFormat
            Case MembershipPasswordFormat.Clear
            Case MembershipPasswordFormat.Encrypted
                password = Encoding.Unicode.GetString(DecryptPassword(Convert.FromBase64String(password)))
            Case MembershipPasswordFormat.Hashed
                Throw New ProviderException("Cannot unEncode a hashed password")
            Case Else
                Throw New ProviderException("Unsupported password format")
        End Select
        Return password
    End Function

    Private Function HexToByte(ByRef hexstring As String) As Byte()
        Dim ReturnBytes((hexstring.Length \ 2) - 1) As Byte
        For i As Integer = 0 To ReturnBytes.Length - 1
            ReturnBytes(i) = Convert.ToByte(hexstring.Substring(i * 2, 2), 16)
        Next
        Return ReturnBytes
    End Function

    Private Sub WriteToEventLog(ByRef e As Exception, ByRef action As String)
        Dim log As EventLog = New EventLog()
        log.Source = eventSource
        log.Log = eventLog
        Dim message As String = "An exception occured communicating with the data source." & vbCrLf & vbCrLf
        message &= "Action:" & action & vbCrLf & vbCrLf
        message &= "Exception:" & e.Message.ToString()
        log.WriteEntry(message)
    End Sub

    

End Class
