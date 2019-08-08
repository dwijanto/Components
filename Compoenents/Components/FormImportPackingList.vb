Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Public Class FormImportPackingList

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)


    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    Dim miroSeq As Long
    Dim podtlseq As Long
    Dim cmmfpriceseq As Long
    Dim cmmfvendorpriceseq As Long

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

    Sub DoWorkOld()
        Dim sw As New Stopwatch
        'Dim AccountingHDSB As New System.Text.StringBuilder
        Dim PackingListHDSB As New System.Text.StringBuilder
        Dim PackingListDtSB As New System.Text.StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        Dim createdon As Date
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

                'sqlstr = "delete from accountinghd where postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & ";" &
                '         " select setval('accountinghd_accountinghdid_seq',(select accountinghdid from accountinghd order by accountinghdid desc limit 1) + 1,false);"

                sqlstr = "delete from packinglistdt pd where pd.delivery in (select delivery from packinglisthd ph where" &
                         " ph.createdon >= " & DateFormatyyyyMMdd(startdate) & " and ph.createdon <= " & DateFormatyyyyMMdd(enddate) & ");" &
                         " select setval('packinglistdt_packinglistdtid_seq',(select packinglistdtid from packinglistdt order by packinglistdtid desc limit 1) + 1,false);" &
                         " delete from packinglisthd ph where ph.createdon >= " & DateFormatyyyyMMdd(startdate) & " and ph.createdon <= " & DateFormatyyyyMMdd(enddate) & ";"


                Dim mymessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                'Fill Header
                ProgressReport(2, "Initialize Table..")
                sqlstr = "select delivery from packinglisthd ph where ph.delivery= 0;"

                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "PackingListHD"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0


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
                    'If i = 68052 Then
                    'Debug.Print("hello")
                    'End If
                    myrecord = mylist(i)
                    If i >= 0 Then
                        createdon = DbAdapter1.dateformatdotdate(myrecord(15))
                        'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
                        If createdon >= startdate.Date AndAlso createdon <= enddate.Date Then

                            Dim pkey0(0) As Object
                            pkey0(0) = myrecord(5)
                            result = DS.Tables(0).Rows.Find(pkey0)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item(0) = myrecord(5)
                                DS.Tables(0).Rows.Add(dr)
                                'delivery bigint,createdby text, createdon date,  reference text, shippingpoint text,
                                'deliverydate date,  incoterm text,  incoterm2 text,  biloflading text,
                                'meansoftranstype text,  meansoftransid text,  documentdate date,vendorcode bigint,
                                'PackingListHDSB.Append(DbAdapter1.validlong(myrecord(5)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(14)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(15)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(3)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(26)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(2)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(27)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(28)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(11)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(12)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(13)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(30)) & vbTab &
                                '                       DbAdapter1.validlong(myrecord(9)) & vbCrLf)
                                PackingListHDSB.Append(validlong(myrecord(5)) & vbTab &
                                                       validstr(myrecord(14)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(15)) & vbTab &
                                                       validstr(myrecord(3)) & vbTab &
                                                       validstr(myrecord(26)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(2)) & vbTab &
                                                       validstr(myrecord(27)) & vbTab &
                                                       validstr(myrecord(28)) & vbTab &
                                                       validstr(myrecord(11)) & vbTab &
                                                       validstr(myrecord(12)) & vbTab &
                                                       validstr(myrecord(13)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(30)) & vbTab &
                                                       validlong(myrecord(9)) & vbTab &
                                                       validstr(GetBillofLading(myrecord(31), myrecord(32))) & vbCrLf)
                            End If
                            'plant integer,  delivery bigint,  deliveryitem integer,  pohd bigint,  poitem integer,  
                            'cmmf bigint,  description text,  deliveredqty numeric,  unit text,  netweight numeric,
                            'nunit text,  grossweight numeric,  gunit text,  volume numeric,  vunit text,

                            'PackingListDtSB.Append(DbAdapter1.validint(myrecord(4)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(5)) & vbTab &
                            '                       DbAdapter1.validint(myrecord(6)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(7)) & vbTab &
                            '                       DbAdapter1.validint(myrecord(8)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(16)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(17)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(18)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(19)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(20)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(21)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(22)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(23)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(24)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(25)) & vbCrLf)
                            PackingListDtSB.Append(validint(myrecord(4)) & vbTab &
                                                   validlong(myrecord(5)) & vbTab &
                                                   validint(myrecord(6)) & vbTab &
                                                   validlong(myrecord(7)) & vbTab &
                                                   validint(myrecord(8)) & vbTab &
                                                   validlong(myrecord(16)) & vbTab &
                                                   validstr(myrecord(17)) & vbTab &
                                                   validreal(myrecord(18)) & vbTab &
                                                   validstr(myrecord(19)) & vbTab &
                                                   validreal(myrecord(20)) & vbTab &
                                                   validstr(myrecord(21)) & vbTab &
                                                   validreal(myrecord(22)) & vbTab &
                                                   validstr(myrecord(23)) & vbTab &
                                                   validreal(myrecord(24)) & vbTab &
                                                   validstr(myrecord(25)) & vbCrLf)

                        End If
                    End If
                Next


            End With
        End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If PackingListHDSB.Length > 0 Then
                ProgressReport(2, "Copy PackingListHD")
                'delivery bigint,createdby text, createdon date,  reference text, shippingpoint text,
                'deliverydate date,  incoterm text,  incoterm2 text,  biloflading text,
                'meansoftranstype text,  meansoftransid text,  documentdate date,vendorcode bigint,
                sqlstr = "copy packinglisthd(delivery,createdby,createdon,reference,shippingpoint,deliverydate,incoterm,incoterm2,biloflading,meansoftranstype,meansoftransid,  documentdate ,vendorcode,housebill) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PackingListHDSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PackingListHD" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If PackingListDtSB.Length > 0 Then
                ProgressReport(2, "Copy PackingListDt")
                'plant integer,  delivery bigint,  deliveryitem integer,  pohd bigint,  poitem integer,  
                'cmmf bigint,  description text,  deliveredqty numeric,  unit text,  netweight numeric,
                'nunit text,  grossweight numeric,  gunit text,  volume numeric,  vunit text,
                sqlstr = "copy packinglistdt(plant,delivery,deliveryitem,pohd,poitem,cmmf,description,deliveredqty,unit,netweight,nunit,grossweight,gunit,volume,vunit) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PackingListDtSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PackingListdt" & "::" & errmessage)
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


    Sub DoWork()
        Dim sw As New Stopwatch
        'Dim AccountingHDSB As New System.Text.StringBuilder
        Dim PackingListHDSB As New System.Text.StringBuilder
        Dim PackingListDtSB As New System.Text.StringBuilder
        Dim PackingListHousebillSB As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        Dim createdon As Date
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

                'sqlstr = "delete from accountinghd where postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & ";" &
                '         " select setval('accountinghd_accountinghdid_seq',(select accountinghdid from accountinghd order by accountinghdid desc limit 1) + 1,false);"

                sqlstr = "delete from packinglistdt pd where pd.delivery in (select delivery from packinglisthd ph where" &
                         " ph.createdon >= " & DateFormatyyyyMMdd(startdate) & " and ph.createdon <= " & DateFormatyyyyMMdd(enddate) & ");" &
                         "delete from packinglisthousebill pd where pd.delivery in (select delivery from packinglisthd ph where" &
                         " ph.createdon >= " & DateFormatyyyyMMdd(startdate) & " and ph.createdon <= " & DateFormatyyyyMMdd(enddate) & ");" &
                         " select setval('packinglistdt_packinglistdtid_seq',(select packinglistdtid from packinglistdt order by packinglistdtid desc limit 1) + 1,false);" &
                         " delete from packinglisthd ph where ph.createdon >= " & DateFormatyyyyMMdd(startdate) & " and ph.createdon <= " & DateFormatyyyyMMdd(enddate) & ";"


                Dim mymessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If



                ProgressReport(2, "Get Reference Data")

                'Get po with supplierinvoicenum
                Dim ds1 As New DataSet
                Dim sqlstr1 = "select pohd,polineno,supplierinvoicenum from miro m" &
                         " left join pomiro pm on pm.miroid = m.miroid" &
                         " left join podtl pd on pd.podtlid = pm.podtlid" &
                         " where miropostingdate >= " & DateFormatyyyyMMdd(DateTimePicker1.Value) &
                         " order by pohd,polineno,miropostingdate;" &
                         "select pohd,polineno,supplierinvoicenum from miro m" &
                         " left join pomiro pm on pm.miroid = m.miroid" &
                         " left join podtl pd on pd.podtlid = pm.podtlid" &
                         " where miropostingdate = '2000-01-01'" &
                         " order by pohd,polineno,miropostingdate;" &
                         " select delivery,housebill from packinglisthousebill;"

                If DbAdapter1.TbgetDataSet(sqlstr1, ds1) Then
                    Dim idx1(1) As DataColumn               '
                    idx1(0) = ds1.Tables(1).Columns(0)       'po    
                    idx1(1) = ds1.Tables(1).Columns(1)       'polineno
                    ds1.Tables(1).PrimaryKey = idx1

                    Dim idx2(1) As DataColumn
                    idx2(0) = ds1.Tables(2).Columns(0)       'delivery   
                    idx2(1) = ds1.Tables(2).Columns(1)       'housebil
                    ds1.Tables(2).PrimaryKey = idx2

                    'fill ds.tables(1) with Po + polineno
                    For i = 0 To ds1.Tables(0).Rows.Count - 1
                        If Not IsDBNull(ds1.Tables(0).Rows(i).Item(2)) Then
                            Dim pkey1(1) As Object
                            pkey1(0) = ds1.Tables(0).Rows(i).Item(0)
                            pkey1(1) = ds1.Tables(0).Rows(i).Item(1)
                            Dim result1 = ds1.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result1) Then
                                Dim dr As DataRow = ds1.Tables(1).NewRow
                                dr.Item(0) = pkey1(0)
                                dr.Item(1) = pkey1(1)
                                dr.Item(2) = ds1.Tables(0).Rows(i).Item(2)
                                If Not IsDBNull(dr.Item(0)) Then
                                    ds1.Tables(1).Rows.Add(dr)
                                End If

                            End If
                        Else
                            Dim dr As DataRow = ds1.Tables(1).NewRow
                            dr.Item(0) = ds1.Tables(0).Rows(i).Item(0)
                            dr.Item(1) = ds1.Tables(0).Rows(i).Item(1)
                            Try
                                ds1.Tables(1).Rows.Add(dr)
                            Catch ex As Exception
                                'Debug.Print("dup")
                            End Try

                            'Debug.Print("item2(blank) {0} {1}", ds.Tables(0).Rows(i).Item(0), ds.Tables(0).Rows(i).Item(1))
                        End If
                        
                    Next
                End If

                'Fill Header



                ProgressReport(2, "Initialize Table..")

                sqlstr = "select delivery from packinglisthd ph where ph.delivery= 0;"

                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "PackingListHD"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0


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
                    'If i = 68052 Then
                    'Debug.Print("hello")
                    'End If
                    myrecord = mylist(i)
                    If i >= 0 Then
                        createdon = DbAdapter1.dateformatdotdate(myrecord(15))
                        'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
                        If createdon >= startdate.Date AndAlso createdon <= enddate.Date Then
                            Dim ValidBillOfLading = GetBillofLading(myrecord(31), myrecord(32))
                            Dim pkey0(0) As Object
                            pkey0(0) = myrecord(5)
                            result = DS.Tables(0).Rows.Find(pkey0)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item(0) = myrecord(5)
                                DS.Tables(0).Rows.Add(dr)
                                'delivery bigint,createdby text, createdon date,  reference text, shippingpoint text,
                                'deliverydate date,  incoterm text,  incoterm2 text,  biloflading text,
                                'meansoftranstype text,  meansoftransid text,  documentdate date,vendorcode bigint,
                                'PackingListHDSB.Append(DbAdapter1.validlong(myrecord(5)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(14)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(15)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(3)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(26)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(2)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(27)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(28)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(11)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(12)) & vbTab &
                                '                       DbAdapter1.validchar(myrecord(13)) & vbTab &
                                '                       DbAdapter1.dateformatdot(myrecord(30)) & vbTab &
                                '                       DbAdapter1.validlong(myrecord(9)) & vbCrLf)

                                If myrecord(3) = "" Then
                                    Dim pkey1(1) As Object
                                    pkey1(0) = myrecord(7)
                                    pkey1(1) = CInt(myrecord(8))
                                    Dim dr1 = ds1.Tables(1).Rows.Find(pkey1)
                                    If Not IsNothing(dr1) Then
                                        Dim mystring As String = String.Empty

                                        If Not IsDBNull(dr1.Item(2)) Then
                                            myrecord(3) = dr1.Item(2)
                                        End If

                                    End If
                                End If

                                PackingListHDSB.Append(validlong(myrecord(5)) & vbTab &
                                                       validstr(myrecord(14)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(15)) & vbTab &
                                                       validstr(myrecord(3)) & vbTab &
                                                       validstr(myrecord(26)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(2)) & vbTab &
                                                       validstr(myrecord(27)) & vbTab &
                                                       validstr(myrecord(28)) & vbTab &
                                                       validstr(myrecord(11)) & vbTab &
                                                       validstr(myrecord(12)) & vbTab &
                                                       validstr(myrecord(13)) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(30)) & vbTab &
                                                       validlong(myrecord(9)) & vbTab &
                                                       validstr(ValidBillOfLading) & vbCrLf)
                            End If
                            'plant integer,  delivery bigint,  deliveryitem integer,  pohd bigint,  poitem integer,  
                            'cmmf bigint,  description text,  deliveredqty numeric,  unit text,  netweight numeric,
                            'nunit text,  grossweight numeric,  gunit text,  volume numeric,  vunit text,

                            'PackingListDtSB.Append(DbAdapter1.validint(myrecord(4)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(5)) & vbTab &
                            '                       DbAdapter1.validint(myrecord(6)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(7)) & vbTab &
                            '                       DbAdapter1.validint(myrecord(8)) & vbTab &
                            '                       DbAdapter1.validlong(myrecord(16)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(17)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(18)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(19)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(20)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(21)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(22)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(23)) & vbTab &
                            '                       DbAdapter1.validdec(myrecord(24)) & vbTab &
                            '                       DbAdapter1.validchar(myrecord(25)) & vbCrLf)
                            PackingListDtSB.Append(validint(myrecord(4)) & vbTab &
                                                   validlong(myrecord(5)) & vbTab &
                                                   validint(myrecord(6)) & vbTab &
                                                   validlong(myrecord(7)) & vbTab &
                                                   validint(myrecord(8)) & vbTab &
                                                   validlong(myrecord(16)) & vbTab &
                                                   validstr(myrecord(17)) & vbTab &
                                                   validreal(myrecord(18)) & vbTab &
                                                   validstr(myrecord(19)) & vbTab &
                                                   validreal(myrecord(20)) & vbTab &
                                                   validstr(myrecord(21)) & vbTab &
                                                   validreal(myrecord(22)) & vbTab &
                                                   validstr(myrecord(23)) & vbTab &
                                                   validreal(myrecord(24)) & vbTab &
                                                   validstr(myrecord(25)) & vbCrLf)

                            'fill packinglisthousebill with new data
                            If myrecord(5) <> "" And ValidBillOfLading <> "" Then
                                Dim pkey2(1) As Object
                                pkey2(0) = myrecord(5)
                                pkey2(1) = ValidBillOfLading
                                Dim result2 = ds1.Tables(2).Rows.Find(pkey2)
                                If IsNothing(result2) Then
                                    Dim dr As DataRow = ds1.Tables(2).NewRow
                                    dr.Item(0) = pkey2(0)
                                    dr.Item(1) = pkey2(1)
                                    If Not IsDBNull(dr.Item(0)) Then
                                        ds1.Tables(2).Rows.Add(dr)
                                    End If
                                    PackingListHousebillSB.Append(validlong(myrecord(5)) & vbTab &
                                                       validstr(ValidBillOfLading) & vbCrLf)
                                End If
                            End If
                        End If
                    End If
                Next


            End With
        End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If PackingListHDSB.Length > 0 Then
                ProgressReport(2, "Copy PackingListHD")
                'delivery bigint,createdby text, createdon date,  reference text, shippingpoint text,
                'deliverydate date,  incoterm text,  incoterm2 text,  biloflading text,
                'meansoftranstype text,  meansoftransid text,  documentdate date,vendorcode bigint,
                sqlstr = "copy packinglisthd(delivery,createdby,createdon,reference,shippingpoint,deliverydate,incoterm,incoterm2,biloflading,meansoftranstype,meansoftransid,  documentdate ,vendorcode,housebill) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PackingListHDSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PackingListHD" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If PackingListDtSB.Length > 0 Then
                ProgressReport(2, "Copy PackingListDt")
                'plant integer,  delivery bigint,  deliveryitem integer,  pohd bigint,  poitem integer,  
                'cmmf bigint,  description text,  deliveredqty numeric,  unit text,  netweight numeric,
                'nunit text,  grossweight numeric,  gunit text,  volume numeric,  vunit text,
                sqlstr = "copy packinglistdt(plant,delivery,deliveryitem,pohd,poitem,cmmf,description,deliveredqty,unit,netweight,nunit,grossweight,gunit,volume,vunit) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PackingListDtSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PackingListdt" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If PackingListHousebillSB.Length > 0 Then
                ProgressReport(2, "Copy PackingListHousebill")
                'plant integer,  delivery bigint,  deliveryitem integer,  pohd bigint,  poitem integer,  
                'cmmf bigint,  description text,  deliveredqty numeric,  unit text,  netweight numeric,
                'nunit text,  grossweight numeric,  gunit text,  volume numeric,  vunit text,
                sqlstr = "copy packinglisthousebill(delivery,housebill) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PackingListHousebillSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PackingListHousebill" & "::" & errmessage)
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
    Private Function GetBillofLading(ByVal Billoflading As String, ByVal BOL As String) As String
        Dim myret As String = Billoflading
        If BOL.Length > 0 Then
            myret = BOL
        End If
        Return myret
    End Function



    Private Sub FormImportPackingList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class