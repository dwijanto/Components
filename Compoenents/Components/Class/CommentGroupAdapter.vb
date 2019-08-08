Imports Npgsql
Public Class CommentGroupAdapter
    Inherits ModelAdapter

    Dim myAdapter As DbAdapter
    Dim Sqlstr As String = String.Empty

    Public Sub New()
        MyBase.new("select cmnttxgrpid,cmnttxgrpname::character varying from cmnttxgrp order by cmnttxgrpname;")
        myAdapter = DbAdapter.getInstance

    End Sub

    Public Sub New(ByVal sqlstr)
        MyBase.New(sqlstr)
        myAdapter = DbAdapter.getInstance
        Me.Sqlstr = sqlstr
    End Sub

    Public Overloads Function loaddata() As Boolean
        Dim myret As Boolean = False
        If MyBase.loaddata Then
            Dim idx0(0) As DataColumn
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0
            DS.Tables(0).Columns(0).AutoIncrement = True
            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
            DS.Tables(0).Columns(0).AutoIncrementStep = -1
            DS.Tables(0).TableName = "CommentGroup"
            myret = True
        End If
        Return myret
    End Function

   
End Class
