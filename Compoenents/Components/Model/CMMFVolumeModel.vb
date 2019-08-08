Imports Npgsql
Imports Components.PublicClass
Imports System.Text

Public Class CMMFVolumeModel
    Implements IModel

    Dim myadapter As DbAdapter = DbAdapter.getInstance

    Private _ErrorMessage As String
    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cmmfvolume"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property

    Public ReadOnly Property FilterField
        Get
            Return "[cmmf] like '*{0}*'"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.cmmf::text, u.avgvolume from {0} u order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sp_updatecmmfvolume"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "avgvolume").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_insertcmmfvolume"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "avgvolume").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sp_deletecmmfvolume"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Public Sub ExportToExcel()
        Dim filename As String = "CMMFVolume-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim sqlstr = "select cmmf,avgvolume from cmmfvolume order by cmmf;"
        ExcelStuff.ExportToExcelAskDirectory(filename, sqlstr, dbtools1)
    End Sub

    Function ImportFromText(ByVal Myform As FormCMMFVolume, ByVal mySelectedPath As String) As Boolean

        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of String())

        Myform.ProgressReport(1, "Open Text File...")
        Dim i As Long
        Try
            'Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)

            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            'Dim arrFI As IO.FileInfo() = dir.GetFiles("*.txt")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            Myform.ProgressReport(6, "Set To Marque")
            'For Each fi As IO.FileInfo In arrFI
            Myform.ProgressReport(1, String.Format("Read Text File...{0}", mySelectedPath))
            Using objTFParser = New FileIO.TextFieldParser(mySelectedPath)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = False
                    Dim count As Long = 0

                    Do Until .EndOfData
                        'If count > 0 Then
                        myrecord = .ReadFields
                        If count > 1 Then
                            myList.Add(myrecord)
                        End If

                        tcount += 1
                        'End If
                        count += 1

                    Loop
                End With
            End Using
            'Next
            If myList.Count = 0 Then
                _ErrorMessage = "Nothing to process."
                Myform.ProgressReport(3, _ErrorMessage)
                Return myret
            End If

            Myform.ProgressReport(1, String.Format("Build Data row..........."))
            Myform.ProgressReport(5, "Set To Continuous")
            For i = 0 To myList.Count - 1
                Myform.ProgressReport(7, i + 1 & "," & myList.Count)
                sb.Append(myList(i)(0) & vbTab & myList(i)(1) & vbCrLf)
            Next
            Myform.ProgressReport(6, "Set To Marque")


            If sb.Length > 0 Then
                Myform.ProgressReport(1, String.Format("Copy CMMF Volume"))
                'Delete First
                Dim sqlstr = "delete from cmmfvolume;"
                DbAdapter1.ExecuteNonQuery(sqlstr)

                sqlstr = "copy cmmfvolume(cmmf,avgvolume)  from stdin with null as 'Null';"
                _ErrorMessage = DbAdapter1.copy(sqlstr, sb.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            myret = True

        Catch ex As Exception
            _ErrorMessage = String.Format("Error : {0} ", i) & ex.Message
            Myform.ProgressReport(1, String.Format(_ErrorMessage))
        End Try
        'copy


        'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(SW.Elapsed.Minutes, "00"), Format(SW.Elapsed.Seconds, "00"), SW.Elapsed.Milliseconds.ToString))
        Return myret
    End Function

End Class
