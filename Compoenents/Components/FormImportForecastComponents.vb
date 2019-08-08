Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
Imports System.IO

Public Class FormImportForecastComponents
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim ds As New DataSet
    Dim errmsg As New StringBuilder
    Dim readfilestatus As Boolean = False
    Dim selectedfile As String = String.Empty
    Dim ForecastComponents As StringBuilder
    Dim CMMFSB As StringBuilder

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            errmsg = New StringBuilder
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                selectedfile = openfiledialog1.FileName
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        ' Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Read Folder..")

        readfilestatus = ImportTextFile(selectedfile)
        If readfilestatus Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(5, "Set to continuous mode again")
        Else
            If Not errmsg.ToString.Contains(vbCrLf) Then
                ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errmsg.ToString))
            Else
                Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                    mystream.WriteLine(errmsg.ToString)
                End Using
                Process.Start(Application.StartupPath & "\error.txt")
                ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            End If
        End If
        sw.Stop()
    End Sub
    Private Function ImportTextFile(ByVal selectedfile As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        Dim list As New List(Of String())
        CMMFSB = New StringBuilder
        ForecastComponents = New StringBuilder
        Dim i As Long
        Try
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            ProgressReport(2, String.Format("Read Text File...{0}", selectedfile))
            Using objTFParser = New FileIO.TextFieldParser(selectedfile)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    '.SetDelimiters(vbTab)
                    .SetDelimiters(";")
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0
                    Do Until .EndOfData
                        'If count > 0 Then
                        myrecord = .ReadFields
                        If count > 0 Then
                            list.Add(myrecord)
                        End If
                        count += 1
                    Loop
                End With
            End Using
            If list.Count = 0 Then
                errmsg.Append("Text File Wrong Format")
                ProgressReport(5, "Set To Marque")
                Return myret
            End If
            'get dataset
            ds = New DataSet
            'get initial keys from Database fro related table

            If Not FillDataset(ds, errmsg) Then
                Return False
            End If

            'Create object for handleing row creation
            'Dim WOR As New WOR(DS)

            ProgressReport(2, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")


            For i = 0 To list.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & list.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                buildSB(list(i))
                'End If
            Next
            myret = True
            If errmsg.Length > 0 Then
                myret = False
                Return myret
            End If
            ProgressReport(6, "Set To Marque")
            ProgressReport(2, String.Format("Copy To Db"))
            If Not copyToDb() Then
                Return False
            End If


        Catch ex As Exception
            errmsg.Append(String.Format("Row : {0} ", i) & ex.Message)
        End Try
        Return myret
    End Function
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripStatusLabel1.Text = message
            End Select
        End If

    End Sub


    Private Sub buildSB(ByVal myrecord As String())
        'get vendorcode
        Dim customercode As String = String.Empty

        Dim result As DataRow
        Dim pkey0(0) As Object
        pkey0(0) = myrecord(6)
        result = ds.Tables(0).Rows.Find(pkey0)
        If IsNothing(result) Then
            customercode = "Null"
            errmsg.Append(String.Format("CustomerName '{0}' not avail.", myrecord(6)) & vbCrLf)
        Else
            customercode = result.Item(0).ToString
        End If

        Dim pkey1(0) As Object
        pkey1(0) = myrecord(4)
        result = ds.Tables(1).Rows.Find(pkey1)
        If IsNothing(result) Then
            Dim dr = ds.Tables(1).NewRow
            dr.Item(0) = myrecord(4)
            ds.Tables(1).Rows.Add(dr)
            CMMFSB.Append(myrecord(4) & vbTab &
                          validstr(myrecord(5)) & vbCrLf)
        End If

        'append ForecastComponents
        'cmmf bigint,  vendorcode bigint,  weeketa integer,  qty integer,        customercode(bigint,
        ForecastComponents.Append(myrecord(4) & vbTab &
                                  myrecord(1) & vbTab &
                                  myrecord(14) & vbTab &
                                  validint(myrecord(11)) & vbTab &
                                  customercode & vbCrLf)
    End Sub

    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As StringBuilder) As Boolean
        Dim progress As String = String.Empty
        Dim myret As Boolean = False
        Dim myerror As String = String.Empty
        Dim sqlstr As String = "select max(customercode),customername from customer where customername <> '' group by customername ;" &
                               " select cmmf from cmmf;"


        If DbAdapter1.TbgetDataSet(sqlstr, DS, myerror) Then
            Try
                DS.Tables(0).TableName = "customer"
                DS.CaseSensitive = True
                progress = "Table Customer"
                Dim idx0(0) As DataColumn               'vendor
                idx0(0) = DS.Tables(0).Columns(1)       'vendorcode
                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "CMMF"
                DS.CaseSensitive = True
                progress = "Table CMMF"
                Dim idx1(0) As DataColumn               'vendor
                idx1(0) = DS.Tables(1).Columns(0)       'vendorcode
                DS.Tables(1).PrimaryKey = idx1

                myret = True
            Catch ex As Exception
                myerror = ex.Message
                errmessage.Append(progress & " " & myerror & vbCrLf)
            End Try
        Else
            errmessage.Append(progress & " " & myerror & vbCrLf)
            Return False
        End If


        Return myret
    End Function

    Private Function copyToDb() As Boolean
        Dim myret As Boolean = False
        Dim mystr As New StringBuilder
        Dim errmessage As String

        ProgressReport(1, "Start Add New Records")
        mystr.Append("delete from forecastestimationcomp;")
        mystr.Append("select setval('forecastestimationcomp_feid_seq',1,false);")
        'cmmf bigint,  vendorcode bigint,  weeketa integer,  qty integer,        customercode(bigint,
        Dim sqlstr As String = String.Empty
        Dim ra As Long = 0
        Try

            If CMMFSB.Length > 0 Then
                sqlstr = "copy cmmf(cmmf,materialdesc) from stdin with null as 'Null';"
                errmessage = DbAdapter1.copy(sqlstr, CMMFSB.ToString, myret)
                If Not myret Then
                    errmsg.Append("Copy CMMF Error " & errmessage & vbCrLf)
                    Return False
                End If

            End If
            sqlstr = "copy forecastestimationcomp(cmmf,vendorcode,weeketa,qty,customercode) from stdin with null as 'Null';"
            ra = DbAdapter1.ExNonQuery(mystr.ToString)
            errmessage = DbAdapter1.copy(sqlstr, ForecastComponents.ToString, myret)
            If Not myret Then
                errmsg.Append("Copy Error " & errmessage & vbCrLf)
                Return myret
            End If


        Catch ex As Exception
            errmsg.Append(ex.Message & vbCrLf)
        End Try
        Return myret
    End Function
End Class