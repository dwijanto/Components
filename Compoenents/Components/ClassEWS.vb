Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Microsoft.Exchange.WebServices.Data
Imports System.Xml
Public Class ClassEWS
    Inherits base
    Dim url As String
    Dim username As String
    Dim domain As String
    Dim password As String
    Dim service As ExchangeService
    Dim networkcredential As Boolean

    Public Sub New(ByVal url As String, ByVal username As String, ByVal password As String, ByVal domain As String, ByVal networkcredential As Boolean)
        Me.url = url
        Me.username = username
        Me.password = password
        Me.domain = domain
        Me.networkcredential = networkcredential
    End Sub

    Public Function CreateConnection() As ExchangeService
        'Using URL
        ServicePointManager.ServerCertificateValidationCallback = AddressOf mycallback

        service = New ExchangeService(ExchangeVersion.Exchange2007_SP1)
        service.Url = New Uri(url)
        If networkcredential Then
            service.Credentials = New NetworkCredential(username, password, domain)
        Else
            'service.Credentials = New WebCredentials(username, password, domain)
            service.Credentials = New WebCredentials(username, password)
        End If
        Return service
    End Function

    Public Function CreateConnectionAutoDiscover() As ExchangeService
        'Using Autodiscover
        ServicePointManager.ServerCertificateValidationCallback = AddressOf mycallback
        service = New ExchangeService
        'service.Credentials = New WebCredentials("dlie@groupeseb.com", "dwi12345")
        service.Credentials = New WebCredentials(username, password)
        'if UseDefaultCredentials = true, the value of service.credentials property is ignored
        'service.UseDefaultCredentials = True 
        service.TraceListener = New TraceListener
        service.TraceEnabled = True
        service.TraceFlags = TraceFlags.All
        service.AutodiscoverUrl("dlie@groupeseb.com", AddressOf RedirectionUrlValidationCallback)
        Return service
    End Function

    Private Function mycallback(ByVal obj As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal SslPolicyErrors As SslPolicyErrors) As Boolean
        'Added for server to pass the request. this will bypass the checking above
        Return True

        If SslPolicyErrors = System.Net.Security.SslPolicyErrors.None Then
            Return True
        End If

        ' If there are errors in the certificate chain, look at each error to determine the cause.
        If (SslPolicyErrors And System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) <> 0 Then
            If chain IsNot Nothing AndAlso chain.ChainStatus IsNot Nothing Then
                For Each status As System.Security.Cryptography.X509Certificates.X509ChainStatus In chain.ChainStatus
                    If (certificate.Subject = certificate.Issuer) AndAlso (status.Status = System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot) Then
                        ' Self-signed certificates with an untrusted root are valid. 
                        Continue For
                    Else
                        If status.Status <> System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError Then
                            ' If there are any other errors in the certificate chain, the certificate is invalid,
                            ' so the method returns false.
                            Return False
                        End If
                    End If
                Next
            End If

            ' When processing reaches this line, the only errors in the certificate chain are 
            ' untrusted root errors for self-signed certificates. These certificates are valid
            ' for default Exchange server installations, so return true.
            Return True
        Else
            ' In all other cases, return false.
            Return False
        End If
    End Function

    Private Function RedirectionUrlValidationCallback(ByVal redirectionUrl As String) As Boolean
        ' The default for the validation callback is to reject the URL.
        Dim result As Boolean = False

        Dim redirectionUri As New Uri(redirectionUrl)

        ' Validate the contents of the redirection URL. In this simple validation
        ' callback, the redirection URL is considered valid if it is using HTTPS
        ' to encrypt the authentication credentials. 
        If redirectionUri.Scheme = "https" Then
            result = True
        End If
        Return result
    End Function

End Class

Public Class base
    Implements IDisposable

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

Class TraceListener
    Implements ITraceListener


    Public Sub Trace(ByVal traceType As String, ByVal traceMessage As String) Implements Microsoft.Exchange.WebServices.Data.ITraceListener.Trace
        CreateXMLTextFile(traceType, traceMessage.ToString())
    End Sub

    Private Sub CreateXMLTextFile(ByVal filename As String, ByVal traceContent As String)
        Try
            ' If the trace data is valid XML, create an XmlDocument object and save.
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(traceContent)
            xmlDoc.Save(filename + ".xml")
        Catch
            ' If the trace data is not valid XML, save it as a text document.
            System.IO.File.AppendAllText(filename + ".txt", traceContent)
        End Try
    End Sub

End Class