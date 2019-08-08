
Imports System.Threading
Imports Components.SharedClass
Imports Components.PublicClass
Imports System.Text

Public Class FormImportMissingBillOfLading

    Dim myImportDelegate As New ThreadStart(AddressOf DoImport)

    Dim myImport As New System.Threading.Thread(myImportDelegate)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)



    Sub DoImport()
        Dim sw As New Stopwatch

        Dim forwarderhousebillSB As New StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())

        Dim sqlstr As String = String.Empty
        'Dim myID As Long
        Dim DS As New DataSet
        sw.Start()


        Dim mymessage As String = String.Empty

        sqlstr = "select housebill,containerno from forwarderhousebill "
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, message:=mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        DS.Tables(0).TableName = "ForwarderHouseBill"
        Dim idx0(1) As DataColumn
        idx0(0) = DS.Tables(0).Columns(0)
        idx0(1) = DS.Tables(0).Columns(1)
        DS.Tables(0).PrimaryKey = idx0


        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0


                ProgressReport(2, "Read Text File...")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)




                    If i >= 0 And myrecord(0) <> "" And myrecord(1) <> "" Then
                        Dim pkey0(1) As Object
                        pkey0(0) = myrecord(0)
                        pkey0(1) = myrecord(1)
                        Dim result = DS.Tables(0).Rows.Find(pkey0)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(0)
                            dr.Item(1) = myrecord(1)
                            DS.Tables(0).Rows.Add(dr)
                            forwarderhousebillSB.Append(validstr(myrecord(0)) & vbTab &
                                                        validstr(myrecord(1)) & vbTab &
                                                        validstr(myrecord(2)) & vbCrLf)
                        End If
                       

                    End If
                Next


            End With
        End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            'If ImportHardCopyHdSB.Length > 0 Then
            ProgressReport(2, "Copy ForwarderHousebill")

            'sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "'," & DateFormatyyyyMMdd(startdate) & ");"
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'myret = DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmessage)
            'If Not myret Then
            '    ProgressReport(2, "Insert InvoiceHardCopy HD" & "::" & errmessage)
            '    Exit Sub
            'End If
            ' End If

            If forwarderhousebillSB.Length > 0 Then
                ProgressReport(2, "Copy ForwarderHousebill")
                'billingdocument bigint,item integer,billedqty numeric, salesunit text, requiredqty numeric,
                'netweight numeric, grossweight numeric, weightunit text, volume numeric,  vunit text,
                'pricingdate date,exrate numeric, netvalue numeric,curr,
                'salesdoc bigint, salesdocitem integer, material bigint, description text, shippingpoint integer,
                'plant integer, country text, cost numeric,curr2 subtotal(numeric)
                sqlstr = "copy forwarderhousebill(housebill,containerno,forwarder) from stdin with null as 'Null';"

                errmessage = DbAdapter1.copy(sqlstr, forwarderhousebillSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy ForwarderHouseBill" & "::" & errmessage)
                    ProgressReport(5, "Continue")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message

                Case (5)
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)


            End Select

        End If

    End Sub

    Private Function filldataset(ByVal DS As DataSet, ByVal mymessage As String) As Boolean
        Throw New NotImplementedException
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myImport.IsAlive Then
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myImport = New Thread(AddressOf DoImport)
                myImport.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub
End Class