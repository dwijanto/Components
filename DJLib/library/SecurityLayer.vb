Imports System.Security.Principal
Imports System.Threading
Imports System.Globalization

Public Class SecurityPrincipal
    Implements IPrincipal
    Private SecurityGroups As Hashtable
    Private SecurityRights As Hashtable
    Private UserIdentity As UserIdentity

    Private Sub New(ByVal SecurityGroups As Hashtable, ByVal SecurityRights As Hashtable, ByVal UserInfo As Hashtable)
        Me.SecurityGroups = SecurityGroups
        Me.SecurityRights = SecurityRights
        Me.UserIdentity = UserIdentity.CreateUserIdentity(UserInfo)
    End Sub

    Public ReadOnly Property Identity As System.Security.Principal.IIdentity Implements System.Security.Principal.IPrincipal.Identity
        Get
            Return UserIdentity
        End Get
    End Property

    Public Function IsInRole(ByVal role As String) As Boolean Implements System.Security.Principal.IPrincipal.IsInRole
        Return SecurityGroups.ContainsValue(role)
    End Function

    Public Shared Function CreatePrincipal(ByVal SecurityGroups As Hashtable, ByVal SecurityRights As Hashtable, ByVal UserInfo As Hashtable) As SecurityPrincipal
        Return New SecurityPrincipal(SecurityGroups, SecurityRights, UserInfo)
    End Function

    Public Shared Function setSecurityPrincipal(ByVal SecurityGroups As Hashtable, ByVal SecurityRights As Hashtable, ByVal UserInfo As Hashtable) As SecurityPrincipal
        AppDomain.CurrentDomain.SetPrincipalPolicy(PrincipalPolicy.WindowsPrincipal)
        If Not (TypeOf Thread.CurrentPrincipal Is SecurityPrincipal) Then
            Dim SecurityPrincipal = New SecurityPrincipal(SecurityGroups, SecurityRights, UserInfo)
            Dim currentSecurityPrincipal = Thread.CurrentPrincipal
            Thread.CurrentPrincipal = SecurityPrincipal
            Return currentSecurityPrincipal
        Else
            Return Nothing
        End If
    End Function

    Public Shared Sub RestoreSecurityPrincipal(ByVal OriginalPrincipal As IPrincipal)
        Thread.CurrentPrincipal = OriginalPrincipal
    End Sub
End Class


Public Class UserIdentity
    Implements IIdentity

    Private _AuthenticationType As String = "Database"
    Private _isAuthenticated As Boolean
    Private _Name As String
    Private UserInfo As New Hashtable
    Private Shared UserNameKey As String = "UserName"

    Private Sub New(ByVal UserInfo As Hashtable)
        Me.UserInfo = UserInfo
    End Sub

    Public ReadOnly Property AuthenticationType As String Implements System.Security.Principal.IIdentity.AuthenticationType
        Get
            Return _AuthenticationType
        End Get
    End Property

    Public ReadOnly Property IsAuthenticated As Boolean Implements System.Security.Principal.IIdentity.IsAuthenticated
        Get
            Return _isAuthenticated
        End Get
    End Property

    Public ReadOnly Property Name As String Implements System.Security.Principal.IIdentity.Name
        Get
            Return Convert.ToString(UserInfo(UserNameKey), CultureInfo.InvariantCulture).Trim()
        End Get
    End Property

    Public Shared Function CreateUserIdentity(ByVal UserInfo As Hashtable) As UserIdentity
        Return New UserIdentity(UserInfo)
    End Function

    Public Function GetPropertyNames() As String()
        Dim PropertyNames() As String = New String(UserInfo.Count - 1) {}
        Dim Count As Integer = 0
        For Each key As Object In UserInfo.Keys
            PropertyNames(Count) = DirectCast(key, String)
            Count += 1
        Next
        Return PropertyNames

    End Function
    Public Function GetProperty(ByVal PropertyName As String) As String
        Return Convert.ToString(UserInfo(PropertyName), CultureInfo.InvariantCulture)
    End Function
    Public Sub SetProperty(ByVal PropertyName As String, ByVal PropertyValue As String)
        UserInfo(PropertyName) = PropertyValue
    End Sub

End Class
