Imports System.Text

Public Class ScoreboardModel
    Public DS As New DataSet
    Public BS As New BindingSource
    Public VendorBS As New BindingSource

    Dim myadapter As DbAdapter = DbAdapter.getInstance
    Dim _ErrorMessage As Object
    Dim _WMFList As Object
    Dim _SISList As Object
    Dim _VendorSASL As String

    Public ReadOnly Property ErrorMessage
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public ReadOnly Property WMFList
        Get
            Return _WMFList
        End Get
    End Property
    Public ReadOnly Property SISList
        Get
            Return _SISList
        End Get
    End Property

    Public ReadOnly Property VendorSASL As String
        Get
            Return _VendorSASL
        End Get
    End Property

    Public Function getTarget(ByVal VendorCode As Long, ByVal department As Integer, ByVal myField As String) As Double
        Dim result As Double
        Dim Params(2) As Npgsql.NpgsqlParameter
        Params(0) = New Npgsql.NpgsqlParameter("vendorcode", VendorCode)
        Params(1) = New Npgsql.NpgsqlParameter("type", department)
        Params(2) = New Npgsql.NpgsqlParameter("field", myField)
        myadapter.ExecuteStoreProcedure("getvendorprocessvalue", result, Params)
        Return result
    End Function

    Public Function getTargetVendor(ByVal vendorcode As Long) As Double
        Dim result As Double
        Dim params(0) As Npgsql.NpgsqlParameter
        params(0) = New Npgsql.NpgsqlParameter("vendorcode", vendorcode)
        myadapter.ExecuteStoreProcedure("getvendorprocessvalue", result, params)
        Return result
    End Function
    Public Function LoadInitialDataFG(ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        Dim sb As New StringBuilder
        sb.Append(String.Format(" select 'All' as vendorname,0 as vendorcode union all select 'All Supplier',1  " &
                               " union all" &
                               " (select distinct v.vendorname, v.vendorcode from vp.vendorbusp vp left join vendor v on v.vendorcode = vp.vendorcode {0} order by v.vendorname);", criteria))
        sb.Append("select pd.cvalue from paramdt pd where paramname = 'ScoreboardWMF';")
        sb.Append("select pd.cvalue from paramdt pd where paramname = 'ScoreboardSIS';")
        sb.Append("select vendorcode::character varying from vendorsasl order by vendorcode")
        Try
            myadapter.TbgetDataSet(sb.ToString, DS)
            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables(0)
            _WMFList = DS.Tables(1).Rows(0).Item(0)
            _SISList = DS.Tables(2).Rows(0).Item(0)
            sb.Clear()
            For Each dr In DS.Tables(3).Rows
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(dr.item(0))
            Next
            _VendorSASL = sb.ToString
            myret = True
        Catch ex As Exception
            _ErrorMessage = ex.Message
        End Try
        Return myret
    End Function

    Public Function LoadInitialDataCP(ByVal criteria As String) As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        Dim sb As New StringBuilder
        sb.Append(String.Format(" select 'All' as vendorname,0 as vendorcode union all select 'All Supplier',1  union all(select distinct v.vendorname, v.vendorcode from vendorspcomp vp left join vendor v on v.vendorcode = vp.vendorcode {0} order by v.vendorname);", criteria))
        sb.Append("select pd.cvalue from paramdt pd where paramname = 'ScoreboardWMF';")
        sb.Append("select pd.cvalue from paramdt pd where paramname = 'ScoreboardSIS';")

        Try
            myadapter.TbgetDataSet(sb.ToString, DS)
            VendorBS = New BindingSource
            VendorBS.DataSource = DS.Tables(0)
            _WMFList = DS.Tables(1).Rows(0).Item(0)
            _SISList = DS.Tables(2).Rows(0).Item(0)
            myret = True
        Catch ex As Exception
            _ErrorMessage = ex.Message
        End Try
        Return myret
    End Function
End Class
