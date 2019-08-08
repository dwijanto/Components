Imports System.Security.Principal

Public Class MyPrincipal
    Implements IPrincipal

    Private _Identity As IIdentity

    Public ReadOnly Property Identity As System.Security.Principal.IIdentity Implements System.Security.Principal.IPrincipal.Identity
        Get
            Return _Identity
        End Get
    End Property

    Public Function IsInRole(ByVal role As String) As Boolean Implements System.Security.Principal.IPrincipal.IsInRole
        Return AppConfig.RoleAttribute.IsUserInRole(_Identity.Name, role)
    End Function

    'Hide Constructor
    Private Sub New()
    End Sub

    Public Shared Function CreatePrincipal(ByVal Identity As IIdentity) As IPrincipal
        Dim principal As IPrincipal = New MyPrincipal With {._Identity = Identity}
        Return principal
    End Function
End Class

Public Class MyIdentity
    Implements IIdentity
    Private _name As String
    Private _IsAuthenticated As Boolean
    Private _AuthenticationType As String
    Public ReadOnly Property AuthenticationType As String Implements System.Security.Principal.IIdentity.AuthenticationType
        Get
            Return _AuthenticationType
        End Get
    End Property

    Public ReadOnly Property IsAuthenticated As Boolean Implements System.Security.Principal.IIdentity.IsAuthenticated
        Get
            Return _IsAuthenticated
        End Get
    End Property

    Public ReadOnly Property Name As String Implements System.Security.Principal.IIdentity.Name
        Get
            Return _name
        End Get
    End Property

    'Hide Constructor
    Private Sub New()
    End Sub
    
    Public Shared Function CreateIdentity(ByVal Name As String) As IIdentity
        Dim identity As IIdentity = New MyIdentity With {._name = Name,
                                                         ._AuthenticationType = "DataBase",
                                                         ._IsAuthenticated = True}
        Return identity
    End Function
End Class