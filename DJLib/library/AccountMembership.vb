Imports System.Web.Security
Imports System


Public Interface IMembershipService
    ReadOnly Property MinPasswordLength() As Integer
    ReadOnly Property EnablePasswordReset As Boolean
    ReadOnly Property RequiresQuestionAndAnswer As Boolean

    Function ValidateUser(ByVal userName As String, ByVal password As String) As Boolean
    Function CreateUser(ByVal userName As String, ByVal password As String, ByVal email As String) As MembershipCreateStatus
    Function ChangePassword(ByVal userName As String, ByVal oldPassword As String, ByVal newPassword As String) As Boolean
    Function ResetPassword(ByVal username As String, ByVal loginUrl As String)
    Function GetUserQuestion(ByVal username As String) As String
    Function GetNumberOFUsersOnline() As Integer


End Interface

Public Class AccountMembershipService
    Implements IMembershipService

    Private ReadOnly _provider As MembershipProvider

    Public Sub New()
        Me.New(New NpgsqlMembershipProvider)
    End Sub

    Public Sub New(ByVal provider As MembershipProvider)
        _provider = provider 'If(provider, Membership.Provider)
    End Sub

    Public ReadOnly Property MinPasswordLength() As Integer Implements IMembershipService.MinPasswordLength
        Get
            Return _provider.MinRequiredPasswordLength
        End Get
    End Property

    Public Function ValidateUser(ByVal userName As String, ByVal password As String) As Boolean Implements IMembershipService.ValidateUser
        If String.IsNullOrEmpty(userName) Then Throw New ArgumentException("Value cannot be null or empty.", "userName")
        If String.IsNullOrEmpty(password) Then Throw New ArgumentException("Value cannot be null or empty.", "password")

        Return _provider.ValidateUser(userName, password)
    End Function

    Public Function CreateUser(ByVal userName As String, ByVal password As String, ByVal email As String) As MembershipCreateStatus Implements IMembershipService.CreateUser
        If String.IsNullOrEmpty(userName) Then Throw New ArgumentException("Value cannot be null or empty.", "userName")
        If String.IsNullOrEmpty(password) Then Throw New ArgumentException("Value cannot be null or empty.", "password")
        If String.IsNullOrEmpty(email) Then Throw New ArgumentException("Value cannot be null or empty.", "email")

        Dim status As MembershipCreateStatus
        _provider.CreateUser(userName, password, email, Nothing, Nothing, True, Nothing, status)
        Return status
    End Function

    Public Function ChangePassword(ByVal userName As String, ByVal oldPassword As String, ByVal newPassword As String) As Boolean Implements IMembershipService.ChangePassword
        If String.IsNullOrEmpty(userName) Then Throw New ArgumentException("Value cannot be null or empty.", "username")
        If String.IsNullOrEmpty(oldPassword) Then Throw New ArgumentException("Value cannot be null or empty.", "oldPassword")
        If String.IsNullOrEmpty(newPassword) Then Throw New ArgumentException("Value cannot be null or empty.", "newPassword")

        ' The underlying ChangePassword() will throw an exception rather
        ' than return false in certain failure scenarios.
        Try
            'Dim currentUser As MembershipUser = _provider.GetUser(userName, True)
            'Return currentUser.ChangePassword(oldPassword, newPassword)
            Return _provider.ChangePassword(userName, oldPassword, newPassword)

        Catch ex As ArgumentException
            Return False
        Catch ex As MembershipPasswordException
            Return False
        End Try
    End Function
    Public Function ResetPassword(ByVal username As String, ByVal loginUrl As String) As Object Implements IMembershipService.ResetPassword
        Return _provider.ResetPassword(username, loginUrl)
    End Function
    Public ReadOnly Property EnablePasswordReset As Boolean Implements IMembershipService.EnablePasswordReset
        Get
            Return _provider.EnablePasswordReset
        End Get
    End Property

    Public ReadOnly Property RequiresQuestionAndAnswer As Boolean Implements IMembershipService.RequiresQuestionAndAnswer
        Get
            Return _provider.RequiresQuestionAndAnswer
        End Get
    End Property

    Public Function GetUserQuestion1(ByVal username As String) As String Implements IMembershipService.GetUserQuestion
        Dim user As MembershipUser = _provider.GetUser(username, False)
        If IsNothing(user) Then
            Throw New Exception("User name not found")
        Else
            Return user.PasswordQuestion
        End If
    End Function

    Public Function GetNumberOFUsersOnline() As Integer Implements IMembershipService.GetNumberOFUsersOnline
        Return _provider.GetNumberOfUsersOnline
    End Function
End Class

Public NotInheritable Class AccountValidation
    Public Shared Function ErrorCodeToString(ByVal createStatus As MembershipCreateStatus) As String
        ' See http://go.microsoft.com/fwlink/?LinkID=177550 for
        ' a full list of status codes.
        Select Case createStatus
            Case MembershipCreateStatus.DuplicateUserName
                Return "Username already exists. Please enter a different user name."

            Case MembershipCreateStatus.DuplicateEmail
                Return "A username for that e-mail address already exists. Please enter a different e-mail address."

            Case MembershipCreateStatus.InvalidPassword
                Return "The password provided is invalid. Please enter a valid password value."

            Case MembershipCreateStatus.InvalidEmail
                Return "The e-mail address provided is invalid. Please check the value and try again."

            Case MembershipCreateStatus.InvalidAnswer
                Return "The password retrieval answer provided is invalid. Please check the value and try again."

            Case MembershipCreateStatus.InvalidQuestion
                Return "The password retrieval question provided is invalid. Please check the value and try again."

            Case MembershipCreateStatus.InvalidUserName
                Return "The user name provided is invalid. Please check the value and try again."

            Case MembershipCreateStatus.ProviderError
                Return "The authentication provider returned an error. Please verify your entry and try again. If the problem persists, please contact your system administrator."

            Case MembershipCreateStatus.UserRejected
                Return "The user creation request has been canceled. Please verify your entry and try again. If the problem persists, please contact your system administrator."

            Case Else
                Return "An unknown error occurred. Please verify your entry and try again. If the problem persists, please contact your system administrator."
        End Select
    End Function
End Class


