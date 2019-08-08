Imports System.Collections.Specialized

Public Class AppConfig
    Public Shared MembershipService As AccountMembershipService
    Public Shared RoleAttribute As NpgsqlRoleProvider
    Public Shared Identity As MyIdentity
    Public Shared Principal As MyPrincipal
    Public Shared dbTools As Dbtools
    Public Shared ConnectionStringCollections As Collection

    Public Shared Function CreateConfig() As NameValueCollection
        Dim config As NameValueCollection = New NameValueCollection
        config.Add("connectionStringName", dbTools.getConnectionString) '"securityTest.My.MySettings.ConnectionString1"
        config.Add("applicationName", "/")
        config.Add("enablePasswordReset", "True")
        config.Add("enablePasswordRetrieval", "true")
        config.Add("maxInvalidPasswordAttempts", "5")
        config.Add("minRequiredNonalphanumericCharacters", "2")
        config.Add("minRequiredPasswordLength", "6")
        config.Add("requiresQuestionAndAnswer", "False")
        config.Add("requiresUniqueEmail", "true")
        config.Add("passwordAttemptWindow", "10")
        config.Add("UserIsOnlineTimeWindow", "20")
        Return config
    End Function
End Class
