Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Public Class FormImportBillingDocument
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    'Dim miroSeq As Long
    'Dim podtlseq As Long
    'Dim cmmfpriceseq As Long
    'Dim cmmfvendorpriceseq As Long

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            'startdate = getfirstdate(DateTimePicker1.Value)
            'enddate = getlastdate(DateTimePicker2.Value)
            startdate = DateTimePicker1.Value
            enddate = DateTimePicker2.Value
            'appendfile = RadioButton1.Checked

            If openfiledialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
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
                Case 4

                Case 5
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


    Sub DoWork()
        Dim sw As New Stopwatch
        'Dim AccountingHDSB As New System.Text.StringBuilder
        Dim BillingHDSB As New System.Text.StringBuilder
        Dim BillingDtlSB As New System.Text.StringBuilder
        Dim ReversalSB As New System.Text.StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        Dim billingdate As Date
        Dim DS As New DataSet
        sw.Start()
        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Delete Existing Record
                ProgressReport(2, "Delete ..")
                ProgressReport(6, "Marque")

                sqlstr = "delete from billingdtl bd where bd.billingdocument in (select billingdocument from billinghd bh where" &
                         " bh.billingdate >= " & DateFormatyyyyMMdd(startdate) & " and bh.billingdate <= " & DateFormatyyyyMMdd(enddate) & ");" &
                         " select setval('billingdtl_billingdtlid_seq',(select billingdtlid from billingdtl order by billingdtlid desc limit 1) + 1,false);" &
                         " delete from billinghd bh where bh.billingdate >= " & DateFormatyyyyMMdd(startdate) & " and bh.billingdate <= " & DateFormatyyyyMMdd(enddate) & ";"

                Dim mymessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                'Fill Header
                ProgressReport(2, "Initialize Table..")
                sqlstr = "select billingdocument from billinghd ph where ph.billingdocument= 0;" &
                         "select billingdoc,salesdoc,item from billingdocreversal;"


                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "BillingHD"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "Reversal"
                Dim idx1(2) As DataColumn
                idx1(0) = DS.Tables(1).Columns(0)
                idx1(1) = DS.Tables(1).Columns(1)
                idx1(2) = DS.Tables(1).Columns(2)
                DS.Tables(1).PrimaryKey = idx1


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
                Dim result As DataRow
                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If i >= 0 Then
                        billingdate = DbAdapter1.dateformatdotdate(myrecord(2))
                        'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
                        If billingdate >= startdate.Date AndAlso billingdate <= enddate.Date Then

                            Dim pkey0(0) As Object
                            pkey0(0) = myrecord(0)
                            result = DS.Tables(0).Rows.Find(pkey0)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item(0) = myrecord(0)
                                DS.Tables(0).Rows.Add(dr)
                                'billingdocument bigint ,  billingtype text,  salesorg text,  pricingprocedure text,
                                ' billingdate date,  incoterm text,  incoterm2 text, termofpayment text,  destcountry text,  companycode text,
                                ' netvalue numeric,  crcy text,  officer text,  createdon date,  payer bigint, soldtoparty bigint,

                                BillingHDSB.Append(validlong(myrecord(0)) & vbTab &
                                                   validstr(myrecord(7)) & vbTab &
                                                   validstr(myrecord(9)) & vbTab &
                                                   validstr(myrecord(10)) & vbTab &
                                                   dateformatdotyyyymmdd(myrecord(2)) & vbTab &
                                                   validstr(myrecord(11)) & vbTab &
                                                   validstr(myrecord(12)) & vbTab &
                                                   validstr(myrecord(13)) & vbTab &
                                                   validstr(myrecord(14)) & vbTab &
                                                   validstr(myrecord(6)) & vbTab &
                                                   validreal(myrecord(15)) & vbTab &
                                                   validstr(myrecord(16)) & vbTab &
                                                   validstr(myrecord(17)) & vbTab &
                                                   dateformatdotyyyymmdd(myrecord(3)) & vbTab &
                                                   validlong(myrecord(18)) & vbTab &
                                                   validlong(myrecord(19)) & vbCrLf)

                            End If
                            'billingdocument bigint,item integer,billedqty numeric, salesunit text, requiredqty numeric,
                            'netweight numeric, grossweight numeric, weightunit text, volume numeric,  vunit text,
                            'pricingdate date,exrate numeric, netvalue numeric,curr,
                            'salesdoc bigint, salesdocitem integer, material bigint, description text, shippingpoint integer,
                            'plant integer, country text, cost numeric,curr2 subtotal(numeric)


                            BillingDtlSB.Append(validlong(myrecord(0)) & vbTab &
                                                   validint(myrecord(1)) & vbTab &
                                                   validreal(myrecord(20)) & vbTab &
                                                   validstr(myrecord(21)) & vbTab &
                                                   validreal(myrecord(22)) & vbTab &
                                                   validreal(myrecord(24)) & vbTab &
                                                   validreal(myrecord(26)) & vbTab &
                                                   validstr(myrecord(27)) & vbTab &
                                                   validreal(myrecord(28)) & vbTab &
                                                   validstr(myrecord(29)) & vbTab &
                                                   dateformatdotyyyymmdd(myrecord(30)) & vbTab &
                                                   validreal(myrecord(31)) & vbTab &
                                                   validreal(myrecord(32)) & vbTab &
                                                   validstr(myrecord(33)) & vbTab &
                                                   validlong(myrecord(4)) & vbTab &
                                                   validint(myrecord(5)) & vbTab &
                                                   validlong(myrecord(34)) & vbTab &
                                                   validstr(myrecord(35)) & vbTab &
                                                   validstr(myrecord(36)) & vbTab &
                                                   validint(myrecord(37)) & vbTab &
                                                   validstr(myrecord(38)) & vbTab &
                                                   validreal(myrecord(39)) & vbTab &
                                                   validstr(myrecord(40)) & vbTab &
                                                   validreal(myrecord(41)) & vbCrLf)
                            'check reversal
                            'If myrecord(7) = "37S1" Then
                            '    'find table reversal, if not avail then create
                            '    Dim pkey1(2) As Object
                            '    pkey1(0) = myrecord(0)
                            '    pkey1(1) = myrecord(4)
                            '    pkey1(2) = myrecord(5)
                            '    result = DS.Tables(1).Rows.Find(pkey1)
                            '    If IsNothing(result) Then
                            '        Dim dr2 As DataRow = DS.Tables(1).NewRow
                            '        dr2.Item(0) = myrecord(0)
                            '        dr2.Item(1) = myrecord(4)
                            '        dr2.Item(2) = myrecord(5)
                            '        DS.Tables(1).Rows.Add(dr2)
                            '        ReversalSB.Append(myrecord(0) & vbTab &
                            '                          myrecord(4) & vbTab &
                            '                          myrecord(5) & vbCrLf)
                            '    End If
                            'End If

                        End If
                    End If
                Next


            End With
        End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If BillingHDSB.Length > 0 Then
                ProgressReport(2, "Copy BillingHD")

                sqlstr = "copy billinghd( billingdocument,billingtype,salesorg,pricingprocedure,billingdate,incoterm,incoterm2,termofpayment,destcountry,companycode,netvalue,crcy,officer,createdon,payer,soldtoparty) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, BillingHDSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Billing HD" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            'If ReversalSB.Length > 0 Then
            '    ProgressReport(2, "Copy Reversal")

            '    sqlstr = "copy billingdocreversal( billingdoc,salesdoc,item) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, ReversalSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy Reversal" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If
            If BillingDtlSB.Length > 0 Then
                ProgressReport(2, "Copy Billing DTL")
                'billingdocument bigint,item integer,billedqty numeric, salesunit text, requiredqty numeric,
                'netweight numeric, grossweight numeric, weightunit text, volume numeric,  vunit text,
                'pricingdate date,exrate numeric, netvalue numeric,curr,
                'salesdoc bigint, salesdocitem integer, material bigint, description text, shippingpoint integer,
                'plant integer, country text, cost numeric,curr2 subtotal(numeric)
                sqlstr = "copy billingdtl(billingdocument,item,billedqty,salesunit,requiredqty,netweight,grossweight,weightunit,volume,vunit,pricingdate,exrate,netvalue,curr,salesdoc,salesdocitem,material,description,shippingpoint,plant,country,cost,curr2,subtotal) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, BillingDtlSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy BillingDTL" & "::" & errmessage)
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

End Class