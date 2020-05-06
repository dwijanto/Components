Imports Npgsql
Imports NpgsqlTypes
Imports System.IO
Imports System.Text
Public Class DbAdapterExeption
    Inherits ApplicationException
    Public Sub New(ByVal errormessage As String)
        MyBase.New(errormessage)
    End Sub
End Class

Public Class DbAdapter
    Implements IDisposable


    Dim _ConnectionStringDict As Dictionary(Of String, String)
    Dim _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Dim _userid As String
    Dim _password As String

    Dim myTransaction As NpgsqlTransaction

    Public Shared myInstance As DbAdapter

    Public Shared Function getInstance() As DbAdapter
        If myInstance Is Nothing Then
            myInstance = New DbAdapter
        End If
        Return myInstance
    End Function

    Public ReadOnly Property userid As String
        Get
            Return _userid
        End Get
    End Property
    Public ReadOnly Property password As String
        Get
            Return _password
        End Get
    End Property

    Public Property Connectionstring As String
        Get
            Return _connectionstring

        End Get
        Set(ByVal value As String)
            _connectionstring = value
        End Set
    End Property

    Public Sub New()
        InitConnectionStringDict()
        _connectionstring = getConnectionString()
    End Sub

    Public ReadOnly Property ConnectionStringDict As Dictionary(Of String, String)
        Get
            Return _ConnectionStringDict
        End Get
    End Property

    Private Sub InitConnectionStringDict()
        _ConnectionStringDict = New Dictionary(Of String, String)
        Dim connectionstring = getConnectionString()
        Dim connectionstrings() As String = connectionstring.Split(";")
        For i = 0 To (connectionstrings.Length - 1)
            Dim mystrs() As String = connectionstrings(i).Split("=")
            _ConnectionStringDict.Add(mystrs(0), mystrs(1))
        Next i

    End Sub

    Private Function getConnectionString() As String
        _userid = "admin"
        _password = "admin"
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.Connectionstring1
        builder.Add("User Id", _userid)
        builder.Add("password", _password)
        'builder.Add("CommandTimeout", "300")
        'builder.Add("TimeOut", "300")
        Return builder.ConnectionString
    End Function

    Public Function getConnection() As NpgsqlConnection
        If IsNothing(_userid) Or IsNothing(_password) Then
            Throw New DbAdapterExeption("User Id or Password is blank.")
        End If
        Return New NpgsqlConnection(_connectionstring)
    End Function

    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function


    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function

#Region "GetDataSet"
    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey

                DataAdapter.Fill(DataSet)
            End Using
            myret = True

        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            If Not IsNothing(ex.Where) Then
                myerror = myerror & " " & ex.Where
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function
#End Region

    Function TBScorecardDataAdapter(ByRef DataSet As DataSet, Optional ByRef message As String = "", Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getscorecard() as tb(scorecardid bigint,supplierid bigint,mydate date,deptid integer,category integer,myvalue numeric)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey

                    'DataAdapter.SelectCommand.Parameters.Add("col1", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = dbtools1.Region
                    DataAdapter.Fill(DataSet)

                    'Delete
                    sqlstr = "sp_deletescorecard"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updatescorecard"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "scorecardid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertscorecard"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "mydate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "category").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "myvalue").SourceVersion = DataRowVersion.Current
                    'DataAdapter.InsertCommand.Parameters.Add("paramhdid", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = paramhdid
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


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

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(_ConnectionStringDict) Then
                    _ConnectionStringDict.Clear()
                    _ConnectionStringDict = Nothing
                End If
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

    Function getQualityYTDM(ByVal mydate As Date, ByVal vendorcode As Long, ByVal cvalue As String, ByRef message As String, ByRef myresult As Double) As Boolean
        Dim myret As Boolean
        Dim result As Object
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_getytdm", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = vendorcode
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(cvalue)
                result = cmd.ExecuteScalar
                myret = True
                If IsDBNull(result) Then
                    myret = False
                Else
                    myresult = result
                End If
            End Using

        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString.Replace("\", "\\")))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & vbCrLf & ex.Detail & vbCrLf & ex.Where
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ex.Where
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function

    Function getSalesdocid(ByVal salesdoc As String, ByRef message As String, ByRef myresult As Long) As Boolean
        Dim myret As Boolean
        Dim result As Object
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getsalesdocid", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(salesdoc)
                result = cmd.ExecuteScalar
                myret = True
                If IsDBNull(result) Then
                    myret = False
                Else
                    myresult = result
                End If
            End Using

        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Function getSalesdocid(ByVal salesdoc As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getsalesdocid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(salesdoc)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Sub deletevp()
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.deletefive", conn) 'Scheduledline
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()
            'End Using
            'Using conn As New NpgsqlConnection(Connectionstring)
            '    conn.Open()
            cmd = New NpgsqlCommand("vp.deletesix", conn) 'POHeader
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()
            'End Using
            'Using conn As New NpgsqlConnection(Connectionstring)
            '    conn.Open()
            cmd = New NpgsqlCommand("vp.deleteone", conn) 'Salesdoccustomerclient
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()
            'End Using

            'Using conn As New NpgsqlConnection(Connectionstring)
            '    conn.Open()
            cmd = New NpgsqlCommand("vp.deletetwo", conn) 'salesdoc
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()
            'End Using

            'Using conn As New NpgsqlConnection(Connectionstring)
            '    conn.Open()
            cmd = New NpgsqlCommand("vp.deletethree", conn) 'customerclientpodetail
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()

            cmd = New NpgsqlCommand("vp.deletefour", conn) 'customerclientpoheader
            cmd.CommandType = CommandType.StoredProcedure
            cmd.ExecuteScalar()

            'cmd = New NpgsqlCommand("vp.deleteseven", conn)
            'cmd.CommandType = CommandType.StoredProcedure
            'cmd.ExecuteScalar()

            'cmd = New NpgsqlCommand("vp.deleteeight", conn)
            'cmd.CommandType = CommandType.StoredProcedure
            'cmd.ExecuteScalar()
        End Using
    End Sub

    Function getCustomerClientpoheaderid(ByVal customerclient As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcustomerclientpoheaderid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(customerclient)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getCustomerClientpodetailid(ByVal customerclientpoheaderid As Long, ByVal customerorderlineno As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcustomerclientpodetailid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = Trim(customerclientpoheaderid)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(customerorderlineno)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getcustomerid(ByVal customercode As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcustomerid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(customercode)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getCustomerDescId(ByVal customerid As Long, ByVal customerdesc As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcustomerdescid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = Trim(customerid)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(customerdesc)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getcmmfid(ByVal cmmfcode As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcmmfid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(cmmfcode)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getcmmfDescId(ByVal cmmfid As Long, ByVal cmmfdesc As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcmmfdescid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = Trim(cmmfid)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(cmmfdesc)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getsupplierid(ByVal suppliercode As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getsupplierid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = Trim(suppliercode)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getSalesdocCustomerClientId(ByVal salesdoc As String, ByVal customerpo As String, ByVal customerlineno As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getsalesdoccustomerclientid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(salesdoc) ' Trim(salesdoc)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(customerpo) 'Trim(customerpo)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(customerlineno) 'IIf(Trim(customerlineno) = "", DBNull.Value, Trim(customerlineno))
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function

    Function getPoDetailid(ByVal sebpono As String, ByVal updatedate As String, ByVal updateby As String, ByVal typeofitem As String,
                           ByVal ordercreationdate As String, ByVal incoterm As String, ByVal port As String, ByVal sao As String,
                           ByVal sp As String, ByVal suppliercode As String, ByVal suppliername As String, ByVal billtopartycode As String,
                           ByVal billtopartyname As String, ByVal shiptopartycode As String, ByVal shiptopartyname As String, ByVal currency As String, ByVal customerremark As String,
                           ByVal cmmf As String, ByVal cmmfdesc As String, ByVal vendormatnumber As String, ByVal customscode As String, ByVal productfamily As String, ByVal brand As String,
                           ByVal unitprice As String, ByVal perunit As String, ByVal newproject As String, ByVal pricedifferent As String, ByVal rri As String, ByVal status As String,
                           ByVal inqqty As String, ByVal inqetd As String, ByVal supplierreasoncode As String, ByVal sebasialineno As String, ByVal reasoncode As String) As Long
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getpodetailid", conn)
            cmd.CommandType = CommandType.StoredProcedure
            'Poheader
            Dim sunitprice As String = unitprice.Replace(",", "").Replace("""", "")
            Dim sperunit As String = perunit.Replace(",", "").Replace("""", "")
            Dim sinqqty As String = inqqty.Replace(",", "").Replace("""", "")

            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(sebpono) 'Trim(sebpono)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(updatedate) 'IIf(Trim(updatedate) = "", DBNull.Value, CDateddMMyyyy(updatedate)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(updateby) 'IIf(Trim(updateby) = "", DBNull.Value, Trim(updateby))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(typeofitem) 'Trim(typeofitem)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(ordercreationdate)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(incoterm) 'Trim(incoterm)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(port) 'IIf(Trim(port) = "", DBNull.Value, Trim(port))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(sao) 'IIf(Trim(sao) = "", DBNull.Value, Trim(sao))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(sp) 'IIf(Trim(sp) = "", DBNull.Value, Trim(sp))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(suppliercode) 'Trim(suppliercode)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(suppliername) 'Trim(suppliername)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(billtopartycode) 'Trim(billtopartycode)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(billtopartyname) 'Trim(billtopartyname)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(shiptopartycode) 'Trim(shiptopartycode)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(shiptopartyname) 'Trim(shiptopartyname.Replace("'", "''"))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(currency) 'Trim(currency)
            'Podetail
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(customerremark) 'IIf(Trim(customerremark) = "", DBNull.Value, Trim(customerremark))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(cmmf) 'Trim(cmmf)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(cmmfdesc) 'Trim(cmmfdesc)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(vendormatnumber) 'IIf(Trim(vendormatnumber) = "", DBNull.Value, Trim(vendormatnumber))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(customscode) 'IIf(Trim(customscode) = "", DBNull.Value, Trim(customscode))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(productfamily) 'IIf(Trim(productfamily) = "", DBNull.Value, Trim(productfamily))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(brand) 'IIf(Trim(brand) = "", DBNull.Value, Trim(brand))

            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = validdec(unitprice) 'IIf(sunitprice = "", DBNull.Value, CDec(sunitprice))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0).Value = validdec(perunit) 'IIf(sperunit = "", DBNull.Value, CDec(sperunit))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = IIf(newproject = "Y", True, False)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = IIf(pricedifferent = "Y", True, False)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(rri) 'IIf(Trim(rri) = "", DBNull.Value, Trim(rri))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(status) 'Trim(status)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = validint(inqqty) 'IIf(sinqqty = "", DBNull.Value, CInt(sinqqty))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(inqetd) 'IIf(Trim(inqetd) = "", DBNull.Value, CDateddMMyyyy(inqetd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(supplierreasoncode) 'IIf(Trim(supplierreasoncode) = "", DBNull.Value, Trim(supplierreasoncode))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(sebasialineno) 'Trim(sebasialineno)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(reasoncode)
            result = cmd.ExecuteScalar
        End Using
        Return result
    End Function



    Sub setscheduledline(ByVal salesdocCustomerClientId As Long, ByVal podetailid As Long, ByVal scheduledlineno As String, ByVal confirmedqty As String, ByVal confirmedetd As String,
                         ByVal consolidateflag As String, ByVal ncstatus As String, ByVal remarks As String, ByVal psed As String, ByVal psrd As String, ByVal pscd As String,
                         ByVal pbtfed As String, ByVal pbtfrd As String, ByVal pbtfcd As String, ByVal peed As String, ByVal perd As String, ByVal pecd As String,
                         ByVal qced As String, ByVal qcrd As String, ByVal qccd As String, ByVal efed As String, ByVal efrd As String, ByVal efcd As String, ByVal qd As String)
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_setscheduledline", conn)
            cmd.CommandType = CommandType.StoredProcedure
            'Poheader

            'Dim scqty As Integer = confirmedqty.Replace(",", "").Replace("""", "")
            'Dim sqdv As Integer = qd.Replace(",", "").Replace("""", "")
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = Trim(salesdocCustomerClientId)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = Trim(podetailid)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(scheduledlineno) 'Trim(scheduledlineno)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = validint(confirmedqty) 'IIf(Trim(scqty) = "", DBNull.Value, Trim(scqty))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(confirmedetd) 'IIf(Trim(confirmedetd) = "", DBNull.Value, CDateddMMyyyy(confirmedetd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = IIf(consolidateflag = "Y", True, False)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0).Value = IIf(ncstatus = "Y", True, False)
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = validchar(remarks) 'IIf(Trim(remarks) = "", DBNull.Value, Trim(remarks))
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(psed) 'IIf(Trim(psed) = "", DBNull.Value, CDateddMMyyyy(psed)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(psrd) 'IIf(Trim(psrd) = "", DBNull.Value, CDateddMMyyyy(psrd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(pscd) 'IIf(Trim(pscd) = "", DBNull.Value, CDateddMMyyyy(pscd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(pbtfed) 'IIf(Trim(pbtfed) = "", DBNull.Value, CDateddMMyyyy(pbtfed)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(pbtfrd) 'IIf(Trim(pbtfrd) = "", DBNull.Value, CDateddMMyyyy(pbtfrd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(pbtfcd) 'IIf(Trim(pbtfcd) = "", DBNull.Value, CDateddMMyyyy(pbtfcd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(peed) 'IIf(Trim(peed) = "", DBNull.Value, CDateddMMyyyy(peed)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(perd) 'IIf(Trim(perd) = "", DBNull.Value, CDateddMMyyyy(perd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(pecd) 'IIf(Trim(pecd) = "", DBNull.Value, CDateddMMyyyy(pecd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(qced) 'IIf(Trim(qced) = "", DBNull.Value, CDateddMMyyyy(qced)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(qcrd) 'IIf(Trim(qcrd) = "", DBNull.Value, CDateddMMyyyy(qcrd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(qccd) 'IIf(Trim(qccd) = "", DBNull.Value, CDateddMMyyyy(qccd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(efed) 'IIf(Trim(efed) = "", DBNull.Value, CDateddMMyyyy(efed)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(efrd) 'IIf(Trim(efrd) = "", DBNull.Value, CDateddMMyyyy(efrd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = CDateddMMyyyy(efcd) 'IIf(Trim(efcd) = "", DBNull.Value, CDateddMMyyyy(efcd)) ' convert date to cdate
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = validint(qd) 'IIf(Trim(sqdv) = "", DBNull.Value, Trim(sqdv))
            result = cmd.ExecuteScalar
        End Using
    End Sub

    Public Function validint(ByVal sinqqty As String) As Object
        If sinqqty = "" Then
            Return DBNull.Value
        Else
            Return CInt(sinqqty.Replace(",", "").Replace("""", ""))
        End If
    End Function
    Public Function validbool(ByVal mybool As String) As Object
        If mybool = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function
    Public Function validdec(ByVal sunitprice As String) As Object
        If sunitprice = "" Then
            Return DBNull.Value
        Else
            Return CDec(sunitprice.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Public Function validchar(ByVal updateby As String) As Object
        If updateby = "" Then
            'Return DBNull.Value
            Return ""
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", "").Replace("\", "\\"))
        End If
    End Function
    Public Function validstr(ByVal updateby As String) As Object
        If updateby = "" Then
            Return DBNull.Value
        Else
            Return Trim(updateby.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Public Function CDateddMMyyyy(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return CDate(mydata(2) & "-" & mydata(1) & "-" & mydata(0))
        End If
        Return DBNull.Value
    End Function
    Public Function ddMMyyyytoyyyyMMdd(ByVal updatedate As String) As Object
        Dim mydata() As String
        If updatedate.Contains(".") Then
            mydata = updatedate.Split(".")
        Else
            mydata = updatedate.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return DBNull.Value
    End Function

    Function VendorBUSP(ByVal Dataset As DataSet, ByRef message As String, Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "vp.sp_getvendorbusp() as tb(vendorcode bigint,buid bigint,spid bigint,vendorbuspid bigint)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(Dataset)

                    'Delete
                    sqlstr = "vp.sp_deletevendorbusp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "vp.sp_updatevendorbusp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "vp.sp_insertvendorbusp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(Dataset.Tables(0))

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


    Function CustomerFlowTx(ByVal myObj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    myTransaction = conn.BeginTransaction
                    'Delete
                    sqlstr = "sp_deletecustomerflow"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updatecustomerflow"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "soldtoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "flow").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dicustomer").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent_group").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent_group_emea").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertcustomerflow"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "soldtoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "flow").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "dicustomer").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent_group").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "continent_group_mea").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True

                Catch ex As Exception
                    mye.message = ex.Message
                    Return False
                End Try
            End Using
            myret = True
        Catch ex As NpgsqlException
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function spmanager(ByVal Dataset As DataSet, ByRef message As String, Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "vp.sp_getspmanager() as tb(supplyplanner character varying ,smid bigint,ofsebid bigint)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(Dataset)


                    'Update
                    sqlstr = "vp.sp_updatespmanager"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "ofsebid").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "smid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "smid").Value = IIf(DataRowVersion.Current = 0, DataRowVersion.Current, DBNull.Value)
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    RecordAffected = DataAdapter.Update(Dataset.Tables(0))

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

    Function VendorBUSP1(ByVal Dataset As DataSet, ByRef message As String, Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        'Dim cmd As NpgsqlCommand
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "vp.sp_getvendorbusp1() as tb(vendorcode bigint,buid bigint,spid bigint,vendorbuspid bigint,vendorname character varying,bu character varying,sp character varying)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(Dataset)

                    'Delete
                    sqlstr = "vp.sp_deletevendorbusp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "vp.sp_updatevendorbusp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "vp.sp_insertvendorbusp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    RecordAffected = DataAdapter.Update(Dataset.Tables(0))

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
    Function FillWeeklyTx(ByVal Dataset As DataSet, ByRef message As String, Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_getweeklyevolution() as tb(id bigint, myyear integer,myweek integer,sasl numeric,pctsasl numeric,targetsasl numeric,pctssl numeric,yearweek character varying,countordertype bigint,idori bigint)"
                    DataAdapter.ContinueUpdateOnError = continueupdateonerror
                    DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.SelectCommand.Connection = conn
                    DataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                    'DataAdapter.MissingMappingAction = MissingSchemaAction.AddWithKey
                    DataAdapter.Fill(Dataset)

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
    Function AdapterWeeklyTx(ByVal Dataset As DataSet, ByRef message As String, Optional ByRef RecordAffected As Integer = 0, Optional ByVal continueupdateonerror As Boolean = True) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'Delete
                    sqlstr = "sp_deleteWeeklyevolution"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "idori").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updateweeklyevolution"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "idori").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myweek").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "pctsasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "targetsasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "pctssl").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "yearweek").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "countordertype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertweeklyevolution"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myyear").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myweek").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "sasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "pctsasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "targetsasl").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "pctssl").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "yearweek").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "countordertype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "idori").Direction = ParameterDirection.Output

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    RecordAffected = DataAdapter.Update(Dataset.Tables(0))

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
    Public Function ExNonQuery(ByVal sqlstr As String) As Long
        Dim myRet As Long
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                myRet = command.ExecuteNonQuery
            End Using
        End Using
        Return myRet
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery
                    'recordAffected = command.ExecuteNonQuery
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteNonQueryAsync(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery

                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, ByRef recordresult As Object, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False       
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordresult = command.ExecuteScalar
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function getcbmosqty(ByVal rowid As Long) As String
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_getcbmosqty", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0).Value = rowid
            result = cmd.ExecuteScalar
        End Using
        If IsDBNull(result) Then
            result = ""
        End If
        Return result
    End Function

    Function validselection(ByVal sebpono As String, ByVal sebpolineno As String, ByVal scheduledlineid As String) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("vp.sp_validselection", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sebpono
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = sebpolineno
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = scheduledlineid
            result = cmd.ExecuteScalar
        End Using
        If IsDBNull(result) Then
            result = True
        Else
            result = False
        End If
        Return result
    End Function

    Sub deleteWOR()
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deletewor", conn)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Function deleteWOR(ByRef message As String) As Boolean
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deletewor", conn)
                cmd.CommandType = CommandType.StoredProcedure

                cmd.ExecuteScalar()
                myret = True
            Catch ex As Exception
                message = ex.Message
            End Try
        End Using
        Return myret
    End Function

    'Sub deleteWOR(ByVal mydate As Date)
    '    Using conn As New NpgsqlConnection(Connectionstring)
    '        Try
    '            conn.Open()
    '            Dim cmd As NpgsqlCommand = New NpgsqlCommand("deletewor", conn)
    '            cmd.CommandType = CommandType.StoredProcedure
    '            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate.Date
    '            cmd.ExecuteScalar()
    '        Catch ex As Exception
    '            MessageBox.Show(ex.Message)
    '        End Try
    '    End Using
    'End Sub

    Function deleteWOR(ByVal mydate As Date) As Boolean
        Dim myret = False
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deletewor", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate.Date
                cmd.ExecuteScalar()
                myret = True
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using
        Return myret
    End Function
    Function deleteWOR(ByVal mydate As Date, ByVal mydate2 As Date, Optional ByRef message As String = "") As Boolean
        Dim myret = False
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deletewor", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate.Date
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = mydate2.Date
                cmd.ExecuteScalar()
                myret = True
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using
        Return myret
    End Function


    Sub deleteIPLT()
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deleteiplt", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Sub deleteOPLT(ByVal startdate As Date, ByVal enddate As Date)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deleteoplt", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startdate.Date
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = enddate.Date
                cmd.ExecuteScalar()
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

        End Using
    End Sub

    Sub deleteOPLT(ByVal startdate As Date)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deleteoplt", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startdate.Date
                cmd.ExecuteScalar()
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

        End Using
    End Sub
    Sub deleteOPLTSIS(ByVal startdate As Date)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_deleteopltsis", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = startdate.Date
                cmd.ExecuteScalar()
            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try

        End Using
    End Sub

    Sub deleteOPLT()
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand("deleteoplt", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Sub ExecuteStoreProcedure(ByVal storeprocedurename As String)
        Using conn As New NpgsqlConnection(Connectionstring)
            Try
                conn.Open()
                Dim cmd As NpgsqlCommand = New NpgsqlCommand(storeprocedurename, conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.ExecuteScalar()
            Catch ex As Exception

            End Try

        End Using
    End Sub

    Public Function getproglock(ByVal programname As String, ByVal userid As String, ByVal status As Integer) As Boolean
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("proglock", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = status
            result = cmd.ExecuteScalar
        End Using

        Return result
    End Function
    Public Function checkproglock(ByVal programname As String) As Boolean
        Dim result As Object = Nothing
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("checkproglock", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = programname
            result = cmd.ExecuteScalar
        End Using

        Return result
    End Function
    Public Function AdapterSASLTx(ByVal formSASLStatusComments As Form, ByVal ds2 As DataSet, Optional ByRef RecordAffected As Integer = 0, Optional ByRef message As String = "") As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    ''Delete
                    'sqlstr = "sp_deleteposasl"
                    'DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Current
                    'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Current

                    'DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    'Update
                    sqlstr = "sp_updateposasl"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cslstatus").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "shipdate").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "shipdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertposasl"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cslstatus").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "shipdate").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    RecordAffected = DataAdapter.Update(ds2.Tables(0))
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
    Public Function AdapterSASLShipment(ByVal formSASLStatusComments As FormSASLStatusComments, ByVal ds2 As DataSet, Optional ByRef RecordAffected As Integer = 0, Optional ByRef message As String = "") As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False

        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    'Update
                    sqlstr = "sp_updateposaslshipmentmodule"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cslstatus").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "shipdate").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'insert
                    sqlstr = "sp_insertposasl"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "cslstatus").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "comment").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "shipdate").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    RecordAffected = DataAdapter.Update(ds2.Tables(0))
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
    Function dateformatdotdate(ByVal myrecord As String) As Date
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = CDate(mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0))
        Return myreturn
    End Function
    Function dateformatdot(ByVal myrecord As String) As Object
        Dim myreturn = "Null"
        If myrecord = "" Then
            Return myreturn
        End If
        Dim mysplit = Split(myrecord, ".")
        myreturn = "'" & mysplit(2) & "-" & mysplit(1) & "-" & mysplit(0) & "'"
        Return myreturn
    End Function
    Public Function validlong(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return DBNull.Value
        Else
            Return CLng(myvalue)
        End If
    End Function
    Public Function validlongNull(ByVal myvalue As String) As Object
        If myvalue = "" Then
            Return "Null"
        Else
            Return CLng(myvalue)
        End If
    End Function

    Function logbook(ByVal formLogBook As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pohd").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 2
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'myTransaction = conn.BeginTransaction
                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pohd").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 2
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    sqlstr = "sp_updatebillingdocreversal"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 1
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    'sqlstr = "sp_updatebillingdocreversal"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    ''DataAdapter.UpdateCommand.Transaction = myTransaction
                    ''mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 1
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))


                    'myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    'myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function
    Function logbookreversal(ByVal formLogBook As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()

                    sqlstr = "sp_updatebillingdocreversal"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))


                    sqlstr = "sp_updatebillingdocreversal"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                    For Each dr As DataRow In mye.dataset.Tables(3).Rows
                        If dr.RowState = DataRowState.Unchanged Then
                            dr.SetModified()
                        End If
                    Next

                    'myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    'myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function
    Function logbook1(ByVal formLogBook As FormLogBook, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    'sqlstr = "sp_updatepackinglistdocument"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pohd").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 2
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    myTransaction = conn.BeginTransaction
                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


                    'sqlstr = "sp_updatepackinglistdocument"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pohd").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 2
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    sqlstr = "sp_updatebillingdocreversal"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 1
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    sqlstr = "sp_updatebillingdocreversal"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "salesdoc").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "salesdocitem").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "status").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                    sqlstr = "sp_updatepackinglistdocument"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "billingdocument").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sebasiapono").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "deliveryitem").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 1
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                    'new one

                    'sqlstr = "sp_updatepackinglistdocument"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "mironumber").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "pohd").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "polineno").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "item").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0).Value = 3
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    'mye.ra = DataAdapter.Update(mye.dataset.Tables(5))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function SaveComment(ByVal formLogBook As FormCommentCodeTx, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_deletecmnttxdtl"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cmnttxdtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                    '(icmnttxdtlid integer, icmnttxdtlname text, icmnttxhdname text, idescription text, icmnttxgrpname text, irank integer)
                    sqlstr = "sp_updatecmnttxdtl"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cmnttxdtlid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxdtlname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxhdname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxgrpname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    '(IN icmnttxdtlname text, IN icmnttxhdname text, IN idescription text, IN icmnttxgrpname text, IN irank integer, INOUT icmnttxdtlid integer)
                    sqlstr = "sp_insertcmnttxdtl"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxdtlname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxhdname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "description").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Text, 0, "cmnttxgrpname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "rank").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "cmnttxdtlid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    myTransaction = conn.BeginTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction


                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function MaterialMaster(ByVal formImportMaterialMaster As FormImportMaterialMaster, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    sqlstr = "sp_updatematerialmaster"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "materialdesc").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "commref").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "familylv1").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbu").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rri").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "range").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cmmftype").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "owner").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertmaterialmaster"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "sorg").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "plant").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "materialdesc").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "commref").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "familylv1").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbu").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rri").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "range").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "cmmftype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "owner").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    myTransaction = conn.BeginTransaction

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    sqlstr = "sp_updatefamily"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertfamily"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                    sqlstr = "sp_updatefamilylv2"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2id").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2name").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertfamilylv2"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2id").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2name").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    sqlstr = "sp_updatebrand"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "brandname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertbrand"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "brandname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))

                    sqlstr = "sp_updatesbusap"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertsbusap"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuid").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(4))


                    sqlstr = "sp_updaterange"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "range").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rangedesc").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertrange"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "range").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "rangedesc").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(5))

                    'sqlstr = "sp_updatevendor"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertvendor"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "vendorname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(6))

                    sqlstr = "sp_updateowner"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "owner").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ownerdescription").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertowner"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "owner").SourceVersion = DataRowVersion.Original
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "ownerdescription").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(7))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function

    Function ConvertFamilySBU(ByVal formConvertFamilySBU As FormConvertFamilySBU, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    sqlstr = "sp_updatefamilysbu"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "turnoverhistoryid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbuid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    ' myTransaction = conn.BeginTransaction
                    ' DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    ' myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    ' myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try
        Return myret
    End Function

   
    Function DocEmailTx(ByVal formGetEmailFromExServerCP As Object, ByRef mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_insertdocemailhdcp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "docemailtype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemailhdcp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                    sqlstr = "sp_insertdocemaildtcp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret

    End Function
    Function checkLockFile(ByVal path As String) As Boolean
        Dim myret As Boolean = False
        If File.Exists(path) Then
            myret = True
        Else
            'create file
            Using fs As FileStream = File.Create(path)
                Dim info As Byte() = New UTF8Encoding(True).GetBytes("0")
                ' Add some information to the file.
                fs.Write(info, 0, info.Length)
                fs.Close()
            End Using            
        End If
        Return myret
    End Function

    Public Function validfilename(ByVal strToReplace As String) As String
        Dim mychar() As String = {"\", "/", ":", "*", "?", "<", ">", "|", """"}
        For Each value As String In mychar
            strToReplace = strToReplace.Replace(value, " ")
        Next
        Return strToReplace
    End Function

    Function BrowseFolderTx(ByVal formBrowseFolder As FormBrowseFolder, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletedocemailhd"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertdocemailhd"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "docemailtype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemailhd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))



                    sqlstr = "sp_deletedocemaildt"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemaildtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertdocemaildt"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemaildt"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemaildtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function
    Function BrowseFolderTx(ByVal formBrowseFolder As FormBrowseFolderCP, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletedocemailhdcp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertdocemailhdcp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "docemailtype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemailhdcp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))



                    sqlstr = "sp_deletedocemaildtcp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemaildtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertdocemaildtcp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current

                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemaildtcp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemaildtid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(3))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function
    Function UploadInvoiceReceivedDateTx(ByVal myform As FormInvoiceHardCopy, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_updatehardcopyreceiveddate"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "invoicehardcopyreceiveddateid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "fcrnumber").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.UpdateCommand.Transaction = myTransaction

                    sqlstr = "sp_inserthardcopyreceiveddate"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "supplierinvoicenumber").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "fcrnumber").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remark").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    sqlstr = "sp_updatehousebilldoc"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "housebilldocid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "courier").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "trackingno").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "senddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.UpdateCommand.Transaction = myTransaction

                    sqlstr = "sp_inserthousebilldoc"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "housebill").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "courier").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "trackingno").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "senddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try

        Return myret
    End Function

    Function UpdatePackingListBillofLadingTx(ByVal formImportPackingListBillofLading As FormImportPackingListBillofLading, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                   
                    sqlstr = "sp_updatepackinglistbilloflading"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "housebill").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function DocEmailDraftTx(ByVal formEmailPreparation As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_updatedocemailtx"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "billoflading").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "draftcreateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertdocemailtx"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "billoflading").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "draftcreateddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function DocEmailTx(ByVal formGetEmailFromExServer As FormGetEmailFromExServer, ByRef mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_insertdocemailhd"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "docemailtype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemailhd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                    sqlstr = "sp_insertdocemaildt"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret

    End Function
    Function DocEmailTx2(ByVal formGetSelectedEmail As FormGetSelectedEmail, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_insertdocemailhd"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemailname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "docemailtype").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatedocemailhd"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sender").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sendername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "foldername").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Timestamp, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(1))


                    sqlstr = "sp_insertdocemaildt"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "docemailhdid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "docemaildtname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                    DataAdapter.InsertCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(2))


                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function MarketEmailTx(ByVal formMarketContacts As FormMarketContacts, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletemarketcontacts"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "marketemailid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertmarketcontacts"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "name").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "marketemailid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatemarketcontacts"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "brandid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "name").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "marketemailid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function VendorBUSP(ByVal formvendorbusp As FormVendorBuSp, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "vp.sp_deletevendorbusp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "vp.sp_insertvendorbusp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "vp.sp_updatevendorbusp"
                     DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorbuspid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "buid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "spid").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function VendorSPComp(ByVal formVendorSPComp As FormVendorSPComp, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletevendorspcomp"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertvendorspcomp"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current                    
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sp").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatevendorspcomp"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sp").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function SAOOPLT(ByVal formSAO As FormSAO, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletesaooplt"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "saoopltid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertsaooplt"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "soldtoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saoname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saost").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "saoopltid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatesaooplt"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "saoopltid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "soldtoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptoparty").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saoname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saost").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function vendorsaosis(ByVal formSAOSIS As FormSAOSIS, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_deletevendorsaosis"
                    DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorsaosisid").SourceVersion = DataRowVersion.Original
                    DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_insertvendorsaosis"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saoname").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorsaosisid").Direction = ParameterDirection.InputOutput
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    sqlstr = "sp_updatevendorsaosis"
                    DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorsaosisid").SourceVersion = DataRowVersion.Original
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "saoname").SourceVersion = DataRowVersion.Current
                    DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    DataAdapter.UpdateCommand.Transaction = myTransaction
                    DataAdapter.DeleteCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function InvoiceReceivedDateTx(ByVal formInvoiceHardCopy As FormInvoiceHardCopy, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction




                    sqlstr = "sp_insertupdateinvoicereceiveddate"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "supplierinvoicenumber").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "billoflading").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "fcrnumber").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "receiveddate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "remarks").SourceVersion = DataRowVersion.Current                    
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure



                    DataAdapter.InsertCommand.Transaction = myTransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function UploadHousebillsenddate(ByVal formInvoiceHardCopy As FormInvoiceHardCopy, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    sqlstr = "sp_insertupdatehousebilldoc"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "billoflading").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = formInvoiceHardCopy.TextBox9.Text
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = formInvoiceHardCopy.TextBox10.Text
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0).Value = formInvoiceHardCopy.DateTimePicker3.Value
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure



                    DataAdapter.InsertCommand.Transaction = myTransaction

                    mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Public Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If Not e.Status = UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If

        End If
    End Sub

    Function MarketContactCPTx(ByVal formMarketContactCP As FormMarketContactCP, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                myTransaction = conn.BeginTransaction
                'Update
                sqlstr = "sp_updatemarketemailcp"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptopartycode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "name").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_insertmarketemailcp"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "shiptopartycode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "vendorcode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "name").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_deletemarketemailcp"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = myTransaction
                DataAdapter.UpdateCommand.Transaction = myTransaction
                DataAdapter.DeleteCommand.Transaction = myTransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                myTransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Private Sub onRowUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Then
            If Not e.Status = UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub


    Function DocEmailDraftCPTx(ByVal formEmailPreparation As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                Try
                    'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster
                    conn.Open()
                    myTransaction = conn.BeginTransaction

                    'sqlstr = "sp_updatedocemailcptx"
                    'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "billoflading").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "containernumber").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "draftcreateddate").SourceVersion = DataRowVersion.Current
                    'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                    sqlstr = "sp_insertupdatedocemailcptx"
                    DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)

                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "delivery").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "draftdate").SourceVersion = DataRowVersion.Current
                    DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure


                    DataAdapter.InsertCommand.Transaction = myTransaction
                    'DataAdapter.UpdateCommand.Transaction = myTransaction
                    mye.ra = DataAdapter.Update(mye.dataset.Tables(4))

                    myTransaction.Commit()
                    myret = True
                Catch ex As Exception
                    myTransaction.Rollback()
                    mye.message = ex.Message
                    Return False
                End Try
            End Using

        Catch ex As Exception
            mye.message = ex.Message
        End Try


        Return myret
    End Function

    Function CustSASTX(ByVal formConvCustSAS As FormConvCustSAS, ByVal mye As ContentBaseEventArgs) As Boolean

        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                myTransaction = conn.BeginTransaction
                'Update
                sqlstr = "sp_updateconvcustsas"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sassebasia").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_insertconvcustsas"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "sassebasia").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_deleteconvcustsas"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = myTransaction
                DataAdapter.UpdateCommand.Transaction = myTransaction
                DataAdapter.DeleteCommand.Transaction = myTransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                myTransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret


    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(Connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

    Function SAOAllocation(ByVal sAOAllocationAdapter As SAOAllocationAdapter, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(Connectionstring)
                conn.Open()
                myTransaction = conn.BeginTransaction
                'Update
                sqlstr = "sp_updatesaoallocation"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pol").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ctype").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_insertsaoallocation"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "customercode").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "pol").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "userid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "ctype").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_deletesaoallocation"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = myTransaction
                DataAdapter.UpdateCommand.Transaction = myTransaction
                DataAdapter.DeleteCommand.Transaction = myTransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                myTransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

End Class
