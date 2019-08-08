Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports System.IO
Public Class FormImportCustomerFlow

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim ds As New DataSet
    Dim errmsg As New StringBuilder
    Dim readfilestatus As Boolean = False
    Dim selectedfile As String = String.Empty
    Dim customerflowSB As New StringBuilder

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

        Dim i As Long
        Try
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            ProgressReport(2, String.Format("Read Text File...{0}", selectedfile))
            Using objTFParser = New FileIO.TextFieldParser(selectedfile)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(vbTab)
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


            'ProgressReport(2, String.Format("Delete Old Data ..........."))
            'DbAdapter1.deleteVendorSBUSP()
            'get initial keys from Database for related table
            If Not FillDataset(ds, errmsg) Then
                Return False
            End If

            'Create object for handleing row creation
            'Dim WOR As New WOR(DS)

            ProgressReport(2, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")
            customerflowSB.Clear()
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
        'get SoldToParty
        Dim soldtoparty As String = String.Empty
        Dim shiptoparty As String = String.Empty
        Dim ofsebid As String = String.Empty
        Dim result As DataRow
        Dim pkey0(0) As Object
        pkey0(0) = myrecord(0)
        result = ds.Tables(0).Rows.Find(pkey0)
        If IsNothing(result) Then
            soldtoparty = "Null"
            errmsg.Append(String.Format("Sold To Party Code '{0}' not avail.", myrecord(0)) & vbCrLf)
        Else
            soldtoparty = result.Item(0).ToString
        End If
        'get sbuid
        Dim pkey1(0) As Object
        pkey1(0) = myrecord(2)
        result = ds.Tables(0).Rows.Find(pkey1)
        If IsNothing(result) Then
            shiptoparty = "Null"
            errmsg.Append(String.Format("Shipt To Party '{0}' not avail.", myrecord(2)) & vbCrLf)
        Else
            shiptoparty = result.Item(0).ToString
        End If       

        'append vendorbuspSB
        customerflowSB.Append(soldtoparty & vbTab & shiptoparty & vbTab & validstr(myrecord(4)) & vbTab & validstr(myrecord(5)) & vbTab & validstr(myrecord(6)) & vbTab & validstr(myrecord(7)) & vbTab & validstr(myrecord(8)) & vbCrLf)
    End Sub
    Private Function validstr(ByVal data As Object) As Object
        If data = "" Then
            Return "Null"
        End If
        Return data
    End Function
    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As StringBuilder) As Boolean
        Dim progress As String = String.Empty
        Dim myret As Boolean = False
        Dim myerror As String = String.Empty
        Dim sqlstr As String = "select customercode,customername from customer;"


        If DbAdapter1.TbgetDataSet(sqlstr, DS, myerror) Then
            Try
                DS.Tables(0).TableName = "Customer"
              

                progress = "Table Customer"
                Dim idx0(0) As DataColumn               'customer
                idx0(0) = DS.Tables(0).Columns(0)       'customercode
                DS.Tables(0).PrimaryKey = idx0

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
        mystr.Append("delete from customerflow;")
        mystr.Append("select setval('customerflow_id_seq',1,false);")
        Dim sqlstr As String = "copy customerflow(soldtoparty,shiptoparty,flow,dicustomer,continent,continent_group,continent_group_emea) from stdin with null as 'Null';"
        Dim ra As Long = 0
        Try
            ra = DbAdapter1.ExNonQuery(mystr.ToString)
            errmessage = DbAdapter1.copy(sqlstr, customerflowSB.ToString, myret)
            If myret Then
                ProgressReport(1, "Add Records Done.")
            Else
                errmsg.Append("Copy Error " & errmessage & vbCrLf)
            End If
        Catch ex As Exception
            errmsg.Append(ex.Message & vbCrLf)
        End Try
        Return myret
    End Function

End Class