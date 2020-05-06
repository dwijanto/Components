Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass


Public Class ImportWOR
    Public Enum FileTypeEnum
        FG = 0
        CP = 1
    End Enum

    Public myFileType As FileTypeEnum

    Dim myCount As Integer = 0
    Dim listcount As Integer = 0
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)

    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)


    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim ReadFileStatus As Boolean = False
    Dim Dataset1 As DataSet
    Dim sb As StringBuilder
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim Source As String
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim mySelectedPath As String
    Dim startdate As Date
    Dim enddate As Date
    Dim deletedata As Boolean
    Dim ButtonDeleteClick As Boolean = False


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not myThread.IsAlive Then
            ButtonDeleteClick = False
            startdate = DateTimePicker1.Value
            enddate = DateTimePicker2.Value
            deletedata = CheckBox1.Checked
            With FolderBrowserDialog1
                .RootFolder = Environment.SpecialFolder.Desktop
                .SelectedPath = "c:\"
                .Description = "Select the source directory"
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .SelectedPath

                    Try
                        myThread = New System.Threading.Thread(myThreadDelegate)
                        myThread.SetApartmentState(ApartmentState.MTA)
                        myThread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End With
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()

        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Read Folder..")

        If DbAdapter1.getproglock("FImportweeklyreportFGComp", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
        Else
            ProgressReport(2, "Read Folder..")

            If ButtonDeleteClick Then
                ProgressReport(2, String.Format("Replace rows ..........."))
                'delete record bigger than receptiondate
                Dim mymessage As String = String.Empty
                If Not DbAdapter1.deleteWOR(startdate, enddate, mymessage) Then
                    ProgressReport(2, mymessage)
                Else
                    DbAdapter1.getproglock("FImportweeklyreportFGComp", HelperClass1.UserInfo.DisplayName, 0)
                    ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                End If
                ButtonDeleteClick = False
            Else
                ReadFileStatus = ImportTextFile(mySelectedPath, errSB)
                If ReadFileStatus Then
                    sw.Stop()
                    ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                    DbAdapter1.getproglock("FImportweeklyreportFGComp", HelperClass1.UserInfo.DisplayName, 0)
                Else
                    If errSB.Length > 100 Then
                        'Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                        Using mystream As New StreamWriter(mySelectedPath & "\error.txt")
                            mystream.WriteLine(errSB.ToString)
                        End Using
                        'Process.Start(Application.StartupPath & "\error.txt")
                        Process.Start(mySelectedPath & "\error.txt")
                        ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                    Else
                        ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errSB.ToString))
                    End If
                End If
            End If
            ProgressReport(5, "Set to continuous mode again")
            sw.Stop()
        End If
    End Sub


    Private Shared Function WaitForAll(ByVal events As ManualResetEvent()) As Boolean
        Dim result As Boolean = False
        Try
            If events IsNot Nothing Then
                For i As Integer = 0 To events.Length - 1
                    events(i).WaitOne()
                Next
                result = True
            End If
        Catch
            result = False
        End Try
        Return result
    End Function
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3
                    ToolStripStatusLabel2.Text = message
                Case 4
                    ToolStripStatusLabel1.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                    'ToolStripStatusLabel1.Text = message
                    'ToolStripStatusLabel1.Text = "Preparing Data .." & myvalue(0) & "/" & myvalue(1)
            End Select

        End If

    End Sub

    Private Sub FormImportData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Load the query in background

        'myQueryThread.Start()
    End Sub

    Private Function ImportTextFile(ByVal FileName As String, ByRef errSB As StringBuilder) As Boolean
        Dim errmsg As String = String.Empty
        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of myData)
        Dim mycsvfile As String = String.Empty
        Dim i As Long
        Dim count As Long
        Try
            'Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)

            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.csv")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            Dim myindex As Integer

            For Each fi As IO.FileInfo In arrFI
                mycsvfile = fi.FullName
                ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))

                Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(";")
                        .HasFieldsEnclosedInQuotes = False
                        count = 0

                        Do Until .EndOfData
                            'If count > 0 Then
                            'If tcount = 121853 Then
                            '    Debug.Print("debug")
                            'End If
                            'If count = 20944 Then
                            '    Debug.Print("debug")
                            'End If
                            myrecord = .ReadFields
                            If myrecord.Length = 1 Then
                                errSB.Append(String.Format("Wrong Delimiter: Filename {0} {2}", fi.FullName, vbCrLf))
                                Exit For
                            End If
                            If count > 0 Then
                                Dim mydate As Date
                                If Not "Header,Confirmation,Shipment".Contains(myrecord(0)) Then
                                    errSB.Append(String.Format("Bad Records: Rows {0}, Filename {1} {2}", count + 1, fi.FullName, vbCrLf))
                                End If
                                If myrecord.Length > myindex Then
                                    mydate = CDate(myrecord(myindex))
                                    If mydate >= startdate.Date And mydate <= enddate.Date Then
                                        Dim mydata As New myData(fi.FullName, count + 1, myrecord)
                                        myList.Add(mydata)
                                    End If
                                Else
                                    'Captured Bad Records
                                    errSB.Append(String.Format("Bad Records: Rows {0}, Filename {1} {2}", count + 1, fi.FullName, vbCrLf))
                                End If
                            Else
                                If myrecord(15) = "Reception Date" Or myrecord(15) = "SPO Creation Date" Then
                                    myindex = 15 'Components
                                    myFileType = FileTypeEnum.CP
                                Else
                                    myindex = 21 'FG
                                    myFileType = FileTypeEnum.FG
                                End If
                            End If

                            tcount += 1
                            'End If
                            count += 1
                        Loop
                    End With
                End Using
            Next
            If myList.Count = 0 Or errSB.Length > 0 Then
                'errSB.Append("Text File Wrong Format")
                ProgressReport(5, "Set To Marque")
                Return myret
            End If
            'get dataset
            Dim DS As New DataSet

            If deletedata Then
                'get initial keys from Database fro related table
                ProgressReport(2, String.Format("Replace rows ..........."))
                'delete record bigger than receptiondate
                If Not DbAdapter1.deleteWOR(startdate, enddate) Then
                    Return False
                End If
            End If

            ProgressReport(2, String.Format("Fill Data ..........."))

            If Not FillDataset(DS, errmsg) Then
                Return False
            End If

            'Create object for handleing row creation
            Dim WOR As New WOR(DS)

            ProgressReport(2, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")

            For i = 0 To myList.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & myList.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                'If myList(i).data(45) <> "" Then
                '    Debug.Print(myList(i).data(45))
                'End If
                If i = 1867 Then
                    Debug.Print(i)
                End If
                If Not WOR.buildSB(errmsg, myList(i)) Then
                    errSB.Append(errmsg & vbCrLf)
                    'Return False
                End If

                'End If
            Next
            If errSB.Length > 0 Then
                Return False
            End If

            'Added Dispose to clear memory.
            DS.Dispose()

            ProgressReport(6, "Set To Marque")
            ProgressReport(2, String.Format("Copy To Db"))
            If Not WOR.copyToDb(errmsg, Me) Then
                errSB.Append(errmsg)
                Return False
            End If
            myret = True

        Catch ex As Exception
            errSB.Append(String.Format("Filename {0}, Record Number {1} ", mycsvfile, count + 1) & ex.Message & vbCrLf)
        End Try
        'copy

        'myret = True
        'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(SW.Elapsed.Minutes, "00"), Format(SW.Elapsed.Seconds, "00"), SW.Elapsed.Milliseconds.ToString))
        Return myret
    End Function

    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As String) As Boolean
        Dim myret As Boolean = False

        'Dim Sqlstr As String = " select sebasiasalesorder from cxsalesorder;" &
        '                       " select cxsalesorderdtlid,sebasiasalesorder,solineno from cxsalesorderdtl;" &
        '                       " select sebasiapono from cxsebpo;" &
        '                       " select cxsebpodtlid,sebasiapono,polineno from cxsebpodtl;" &
        '                       " select customercode from customer;" &
        '                       " select count(1) from cxsalesorderdtl;" &
        '                       " select count(1) from cxsebpodtl;" &
        '                       " select cxrelsalesdocpoid,cxsalesorderdtlid,cxsebpodtlid from cxrelsalesdocpo;" &
        '                       " select count(1) from cxrelsalesdocpo;" &
        '                       " select relsalesdocpoid,ordertype,latestupdate,updatesince,shipfrom from cxsebodtp;" &
        '                       " select count(1) from cxsebodtp;" &
        '                       " select count(1) from cxconf;" 
        Dim Sqlstr As String = " select sebasiasalesorder,finalcustomerorder from cxsalesorder;" &
                       " select cxsalesorderdtlid,sebasiasalesorder,solineno from cxsalesorderdtl;" &
                       " select sebasiapono from cxsebpo;" &
                       " select cxsebpodtlid,sebasiapono,polineno from cxsebpodtl;" &
                       " select customercode from customer;" &
                       " select cxsalesorderdtlid from cxsalesorderdtl order by cxsalesorderdtlid desc limit 1;" &
                       " select cxsebpodtlid from cxsebpodtl order by cxsebpodtlid desc limit 1;" &
                       " select cxrelsalesdocpoid,cxsalesorderdtlid,cxsebpodtlid from cxrelsalesdocpo;" &
                       " select cxrelsalesdocpoid from cxrelsalesdocpo order by cxrelsalesdocpoid desc limit 1;" &
                       " select relsalesdocpoid,ordertype,latestupdate,updatesince,shipfrom from cxsebodtp;" &
                       " select cxsebodtpid from cxsebodtp order by cxsebodtpid desc limit 1;" &
                       " select cxconfid from cxconf order by cxconfid desc limit 1;" &
                       " select cmmf,itemid from cmmf;"

        '" select cxconfstatusid from cxconfstatus order by cxconfstatusid desc limit 1" &
        '" select cxconfotherid from cxconfother order by cxconfotherid desc limit 1" &
        '" select cxshipmentid from cxshipment order by cxshipmentid desc limit 1" &
        '" select cxshipmentotherid from cxshipmentother order by cxshipmentotherid desc limit 1"

        If DbAdapter1.TbgetDataSet(Sqlstr, DS, errmessage) Then
            DS.Tables(0).TableName = "cxsalesorder"
            DS.Tables(1).TableName = "cxsalesorderdtl"
            DS.Tables(2).TableName = "cxsebpo"
            DS.Tables(3).TableName = "cxsebpodtl"
            DS.Tables(4).TableName = "customer"
            DS.Tables(5).TableName = "seqsalesorderdtl"
            DS.Tables(6).TableName = "seqsebpodtl"
            DS.Tables(7).TableName = "cxrelsalesdocpo"
            DS.Tables(8).TableName = "seqrelsalesdocpo"
            DS.Tables(9).TableName = "cxsebodtp"
            DS.Tables(10).TableName = "seqsebodtp"
            'DS.Tables(12).TableName = "seqconfstatus" '11 - 12, the last seqconf
            DS.Tables(11).TableName = "seqconf"
            'DS.Tables(13).TableName = "seqconfother"
            'DS.Tables(13).TableName = "seqshipment"
            'DS.Tables(13).TableName = "seqshipmentother"
            DS.Tables(12).TableName = "cmmf"

            Dim idx0(0) As DataColumn               'cxsalesorder
            idx0(0) = DS.Tables(0).Columns(0)       'sebasiasalesorder
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(1) As DataColumn               'cxsalesorderdtl
            idx1(0) = DS.Tables(1).Columns(1)       'sebasiasalesorder    
            idx1(1) = DS.Tables(1).Columns(2)       'solineno
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn               'cxsebpo
            idx2(0) = DS.Tables(2).Columns(0)       'sebasiapono
            DS.Tables(2).PrimaryKey = idx2

            Dim idx3(1) As DataColumn               'cxsebpodtl
            idx3(0) = DS.Tables(3).Columns(1)       'sebasiapono
            idx3(1) = DS.Tables(3).Columns(2)       'polineno
            DS.Tables(3).PrimaryKey = idx3


            Dim idx4(0) As DataColumn               'Customer Master
            idx4(0) = DS.Tables(4).Columns(0)       'customercode
            DS.Tables(4).PrimaryKey = idx4

            Dim idx7(1) As DataColumn               'cxrelsalesdocpo
            idx7(0) = DS.Tables(7).Columns(1)       'cxsalesorderdtlid
            idx7(1) = DS.Tables(7).Columns(2)       'cxsebpodtlid
            DS.Tables(7).PrimaryKey = idx7


            Dim idx12(0) As DataColumn               'Customer Master
            idx12(0) = DS.Tables(12).Columns(0)       'customercode
            DS.Tables(12).PrimaryKey = idx12
        Else
            Return False
        End If
        myret = True
        Return myret
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ButtonDeleteClick = True
        If MessageBox.Show("Delete this period?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
            If Not myThread.IsAlive Then
                startdate = DateTimePicker1.Value
                enddate = DateTimePicker2.Value
                deletedata = CheckBox1.Checked
                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                MsgBox("Please wait until the current process is finished")
            End If
        End If

    End Sub
End Class

Public Class WOR



    Public Property ds As DataSet
    Public Property cxSalesOrder As New StringBuilder
    Public Property cxSalesOrderdtl As New StringBuilder
    Public Property cxsebpo As New StringBuilder
    Public Property cxsebpodtl As New StringBuilder
    Public Property cxrelsalesdocpo As New StringBuilder

    Public Property cxSebodtp As New StringBuilder
    Public Property cxshipment As New StringBuilder
    Public Property cxshipmentother As New StringBuilder
    Public Property cxconf As New StringBuilder
    Public Property cxconfother As New StringBuilder
    Public Property cxconfstatus As New StringBuilder

    'Public Property brandsb As New StringBuilder
    'Public Property loadingsb As New StringBuilder
    Public Property cmmfsb As New StringBuilder
    Public Property cmmfUpdateSB As New StringBuilder
    Public Property FinalCustomerOrderSB As New StringBuilder

    Public vendorcode As String

    Public updatedict As New Dictionary(Of Long, String)
    Public updateFinalCustomerOrder As New Dictionary(Of Long, String)


    Public Property customer As New StringBuilder
    Dim seqsalesorderdtl As Long
    Dim seqsebpodtl As Long
    Dim seqrelSalesDocPo As Long
    Dim seqsebodtp As Long
    Dim seqconf As Long

    'Dim seqconfstatus As Long
    'Dim seqconfother As Long
    'Dim seqshipment As Long
    'Dim seqshipmentother As Long

    Dim podetailid As Long
    Dim salesorderdtlid As Long
    Dim relsalesdocpoid As Long
    Dim confid As Long

    Public Sub New(ByVal ds As DataSet)
        Me.ds = ds
        seqsalesorderdtl = 0
        If ds.Tables(5).Rows.Count > 0 Then
            seqsalesorderdtl = ds.Tables(5).Rows(0).Item(0)
        End If
        seqsebpodtl = 0
        If ds.Tables(6).Rows.Count > 0 Then
            seqsebpodtl = ds.Tables(6).Rows(0).Item(0)
        End If
        seqrelSalesDocPo = 0
        If ds.Tables(8).Rows.Count > 0 Then
            seqrelSalesDocPo = ds.Tables(8).Rows(0).Item(0)
        End If
        seqsebodtp = 0
        If ds.Tables(10).Rows.Count > 0 Then
            seqsebodtp = ds.Tables(10).Rows(0).Item(0)
        End If
        seqconf = 0
        If ds.Tables(11).Rows.Count > 0 Then
            seqconf = ds.Tables(11).Rows(0).Item(0)
        End If
        'seqconfstatus = ds.Tables(11).Rows(0).Item(0)

        'seqconfother = ds.Tables(13).Rows(0).Item(0)
        'seqshipment = ds.Tables(14).Rows(0).Item(0)
        'seqshipmentother = ds.Tables(15).Rows(0).Item(0)

    End Sub

    Public Function buildSB(ByRef message As String, ByVal mydata As myData) As Boolean
        Dim myret As Boolean = False
        'Dim myprogress As String = String.Empty


        Try

            If mydata.data.length >= 47 Then
                myret = fg(mydata, message)
            Else
                myret = Comp(mydata, message)
            End If
        Catch ex As Exception
            message = ex.Message & "::" & message
        End Try
        Return myret
    End Function

    Public Function copyToDb(ByRef errMsg As String, ByVal myform As ImportWOR) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String
        Try
            If cxSalesOrder.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxSalesOrder"))
                sqlstr = "copy cxsalesorder(sebasiasalesorder,soldtoparty,curinq,unittp,finalcustomerorder)  from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxSalesOrder.ToString, myret)
                If Not myret Then
                    Return myret
                End If

            End If
            If cxSalesOrderdtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxSalesOrderDtl"))
                sqlstr = "copy cxsalesorderdtl(sebasiasalesorder,solineno,customerorderno,orderstatus,inquiryeta,inquiryetd,inquiryqty,currentinquiryeta,currentinquiryetd,currentinquiryqty) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxSalesOrderdtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If customer.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Customer"))
                sqlstr = "copy customer(customercode,customername) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, customer.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            ''brand
            'If brandsb.ToString <> "" Then
            '    myform.ProgressReport(2, String.Format("Copy Brand"))
            '    sqlstr = "copy brand(brandid,brandname) from stdin with null as 'Null';"
            '    errMsg = DbAdapter1.copy(sqlstr, brandsb.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If
            'End If

            ''loadingcode
            'If loadingsb.ToString <> "" Then
            '    myform.ProgressReport(2, String.Format("Copy loading"))
            '    sqlstr = "copy loading(loadingcode,loadingname) from stdin with null as 'Null';"
            '    errMsg = DbAdapter1.copy(sqlstr, loadingsb.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If
            'End If

            ''cmmf
            'If cmmfsb.ToString <> "" Then
            '    myform.ProgressReport(2, String.Format("Copy cmmf"))
            '    sqlstr = "copy cmmf(cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid) from stdin with null as 'Null';"
            '    errMsg = DbAdapter1.copy(sqlstr, cmmfsb.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If
            'End If
            ''update cmmf
            'If cmmfUpdateSB.Length > 0 Then
            '    myform.ProgressReport(2, "Update CMMF")
            '    'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
            '    sqlstr = "update cmmf set rir= foo.rir,itemid = foo.itemid,vendorcode = foo.vendorcode::integer,comfam = foo.comfam::integer,loadingcode = foo.loadingcode,brandid = foo.brandid::integer from (select * from array_to_set8(Array[" & cmmfUpdateSB.ToString &
            '             "]) as tb (id character varying,rir character varying,itemid character varying,materialdesc character varying,vendorcode character varying,comfam character varying,loadingcode character varying,brandid character varying))foo where cmmf = foo.id::bigint;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errMsg) Then
            '        Return False
            '    End If
            'End If

            If cxsebpo.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy PO"))
                sqlstr = "copy cxsebpo(sebasiapono,receptiondate) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxsebpo.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxsebpodtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy PO Dtl"))
                'sqlstr = "copy cxsebpodtl(sebasiapono,polineno,shiptoparty,cmmf,fob,unittp,osqty,comments) from stdin with null as 'Null';"
                sqlstr = "copy cxsebpodtl(sebasiapono,polineno,shiptoparty,cmmf,fob,unittp,osqty,comments,receptiondate) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxsebpodtl.ToString, myret)
                'errMsg = cxsebpodtl.ToString
                If Not myret Then
                    Return myret
                End If
            End If

            If cxrelsalesdocpo.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Rel Sales PO"))
                sqlstr = "copy cxrelsalesdocpo(cxsalesorderdtlid,cxsebpodtlid) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxrelsalesdocpo.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxSebodtp.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Odtp"))
                sqlstr = "copy cxsebodtp(relsalesdocpoid,ordertype,latestupdate,updatesince,shipfrom,currentconfirmedeta,currentconfirmedetd,currentconfirmedqty,deliveredqty,shipdate,shipdateeta,dicustomerorder) from stdin with null as 'Null';"
                'sqlstr = "copy cxsebodtp(relsalesdocpoid,ordertype,latestupdate,updatesince,shipfrom) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxSebodtp.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If


            If cxconf.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Confirmation"))
                sqlstr = "copy cxconf(sebodtpid,currentconfirmedeta,currentconfirmedetd,currentconfirmedqty) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxconf.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxconfother.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy ConfOther"))
                sqlstr = "copy cxconfother(sebodtpid,stconfirmedetd,stconfirmedqty) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxconfother.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxconfstatus.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Confirmation Status"))
                sqlstr = "copy cxconfstatus(cxconfid,confirmationstatus) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxconfstatus.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            '***********************
            If cxshipment.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Shipment"))
                sqlstr = "copy cxshipment(sebodtpid,deliveredqty,shipdate,shipdateeta) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxshipment.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxshipmentother.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Shipment Other"))
                sqlstr = "copy cxshipmentother(sebodtpid,ctrno,boatid,packinglist) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxshipmentother.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If


            ''brand
            'If brandsb.ToString <> "" Then
            '    myform.ProgressReport(2, String.Format("Copy Brand"))
            '    sqlstr = "copy brand(brandid,brandname) from stdin with null as 'Null';"
            '    errMsg = DbAdapter1.copy(sqlstr, brandsb.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If
            'End If

            ''loadingcode
            'If loadingsb.ToString <> "" Then
            '    myform.ProgressReport(2, String.Format("Copy loading"))
            '    sqlstr = "copy loading(loadingcode,loadingname) from stdin with null as 'Null';"
            '    errMsg = DbAdapter1.copy(sqlstr, loadingsb.ToString, myret)
            '    If Not myret Then
            '        Return False
            '    End If
            'End If

            'cmmf
            If cmmfsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cmmf"))
                sqlstr = "copy cmmf(cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cmmfsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            'update cmmf
            If cmmfUpdateSB.Length > 0 Then
                myform.ProgressReport(2, "Update CMMF")
                'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                'sqlstr = "update cmmf set rir= foo.rir,itemid = foo.itemid,vendorcode = foo.vendorcode::integer,comfam = foo.comfam::integer,loadingcode = foo.loadingcode,brandid = foo.brandid::integer from (select * from array_to_set8(Array[" & cmmfUpdateSB.ToString &
                '         "]) as tb (id character varying,rir character varying,itemid character varying,materialdesc character varying,vendorcode character varying,comfam character varying,loadingcode character varying,brandid character varying))foo where cmmf = foo.id::bigint;"
                sqlstr = "update cmmf set itemid = foo.itemid from (select * from array_to_set2(Array[" & cmmfUpdateSB.ToString &
                        "]) as tb (id character varying,itemid character varying))foo where cmmf = foo.id::bigint;"

                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errMsg) Then
                    Return False
                End If
            End If

            If FinalCustomerOrderSB.Length > 0 Then
                myform.ProgressReport(2, "Update Final Customer Order")
                'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                'sqlstr = "update cmmf set rir= foo.rir,itemid = foo.itemid,vendorcode = foo.vendorcode::integer,comfam = foo.comfam::integer,loadingcode = foo.loadingcode,brandid = foo.brandid::integer from (select * from array_to_set8(Array[" & cmmfUpdateSB.ToString &
                '         "]) as tb (id character varying,rir character varying,itemid character varying,materialdesc character varying,vendorcode character varying,comfam character varying,loadingcode character varying,brandid character varying))foo where cmmf = foo.id::bigint;"
                sqlstr = "update cxsalesorder set finalcustomerorder = foo.finalcustomerorder from (select * from array_to_set2(Array[" & FinalCustomerOrderSB.ToString &
                        "]) as tb (id character varying,finalcustomerorder character varying))foo where sebasiasalesorder = foo.id::bigint;"

                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errMsg) Then
                    Return False
                End If
            End If

            myret = True
        Catch ex As Exception
            errMsg = ex.Message

        End Try
        Return myret

    End Function

    Private Function validint(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CInt(p1)
        End If
    End Function
    Private Function validReal(ByVal data As Object) As Object
        If data = "" Then
            Return "Null"
        Else
            Return data
        End If
    End Function
    Private Function validstr(ByVal data As Object) As Object
        If data = "" Then
            Return "Null"
        End If
        data = Replace(data, "\", "/").Replace(vbTab, "")
        Return data
    End Function

    Private Function fg(ByVal mydata As Object, ByRef message As String) As Boolean
        Dim myret = False
        Dim myprogress As String = String.Empty
        Dim data = mydata.data
        Dim comments As String = String.Empty
        Dim salesorder() As String
        Try


            If data(0) = "Header" Then
                myprogress = "SalesOrder"
                salesorder = data(7).ToString.Split("/")
                'find sales order in dataset.table(cxsalesorder) 
                'if avail then no need to create record in table cxsalesorder
                Dim result As DataRow

                'Salesorder Hd
                Dim pkey0(0) As Object
                pkey0(0) = salesorder(0)
                result = ds.Tables(0).Rows.Find(pkey0)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(0).NewRow
                    dr.Item(0) = salesorder(0)
                    ds.Tables(0).Rows.Add(dr)
                    cxSalesOrder.Append(salesorder(0) & vbTab &
                                        data(3) & vbTab &
                                        data(20) & vbTab &
                                        validint(data(23)) & vbTab &
                                        validstr(data(45)) & vbCrLf)
                Else
                    'Update Final Customer Order
                    Dim myitem As String = "" & result.Item("finalcustomerorder")
                    If myitem.Trim <> data(45) Then
                        If Not updateFinalCustomerOrder.ContainsKey(salesorder(0)) Then
                            updateFinalCustomerOrder.Add(salesorder(0), data(45))
                            If FinalCustomerOrderSB.Length > 0 Then
                                FinalCustomerOrderSB.Append(",")
                            End If
                            FinalCustomerOrderSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", salesorder(0), validstr(data(45)).ToString.Replace("'", "''")))
                        End If
                    End If
                End If


                myprogress = "SalesOrder Dtl"
                'SalesOrder Dtl
                Dim pkey1(1) As Object
                pkey1(0) = salesorder(0)
                pkey1(1) = salesorder(1)
                result = ds.Tables(1).Rows.Find(pkey1)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(1).NewRow
                    seqsalesorderdtl += 1
                    dr.Item(0) = seqsalesorderdtl
                    dr.Item(1) = salesorder(0)
                    dr.Item(2) = salesorder(1)
                    ds.Tables(1).Rows.Add(dr)

                    salesorderdtlid = seqsalesorderdtl
                    cxSalesOrderdtl.Append(salesorder(0) & vbTab &
                                           salesorder(1) & vbTab &
                                           data(6) & vbTab &
                                           data(17) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(24))) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(25))) & vbTab &
                                           data(26) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(27))) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(28))) & vbTab &
                                           data(29) & vbCrLf)
                Else
                    salesorderdtlid = result.Item(0)
                End If

                myprogress = "Check Customer Ship to Party"
                Dim pkey2(0) As Object
                pkey2(0) = data(1)
                result = ds.Tables(4).Rows.Find(pkey2)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(4).NewRow
                    dr.Item(0) = data(1)
                    ds.Tables(4).Rows.Add(dr)
                    customer.Append(data(1) & vbTab &
                                    data(2) & vbCrLf)
                End If

                myprogress = "Check Customer SoldToParty"
                Dim pkey3(0) As Object
                pkey3(0) = data(3)
                result = ds.Tables(4).Rows.Find(pkey3)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(4).NewRow
                    dr.Item(0) = data(3)

                    ds.Tables(4).Rows.Add(dr)
                    customer.Append(data(3) & vbTab &
                                    data(4) & vbCrLf)
                End If

                myprogress = "Check PO Header"
                Dim pkey4(0) As Object
                Dim po = data(40).ToString.Split("/")
                pkey4(0) = po(0)
                result = ds.Tables(2).Rows.Find(pkey4)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(2).NewRow
                    dr.Item(0) = po(0)

                    ds.Tables(2).Rows.Add(dr)
                    cxsebpo.Append(po(0) & vbTab &
                                    DateFormatyyyyMMddString(data(21)) & vbCrLf)
                End If

                myprogress = "Check PO Detail"
                Dim pkey5(1) As Object
                pkey5(0) = po(0)
                pkey5(1) = po(1)
                result = ds.Tables(3).Rows.Find(pkey5)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(3).NewRow
                    seqsebpodtl += 1
                    dr.Item(0) = seqsebpodtl
                    dr.Item(1) = po(0)
                    dr.Item(2) = po(1)
                    ds.Tables(3).Rows.Add(dr)

                    podetailid = seqsebpodtl
                    'For j = 46 To data.length - 1
                    '    comments += data(j)
                    'Next
                    'comments = "" & data(46)
                    'cxsebpodtl.Append(po(0) & vbTab &
                    '                po(1) & vbTab &
                    '                data(1) & vbTab &
                    '                data(11) & vbTab &
                    '                validReal(data(22)) & vbTab &
                    '                validReal(data(23)) & vbTab &
                    '                validint(data(39)) & vbTab &
                    '                validstr(comments) & vbCrLf)
                    'cxsebpodtl.Append(po(0) & vbTab &
                    '               po(1) & vbTab &
                    '               data(1) & vbTab &
                    '               data(11) & vbTab &
                    '               validReal(data(22)) & vbTab &
                    '               validReal(data(23)) & vbTab &
                    '               validint(data(39)) & vbTab &
                    '               validstr(data(46)) & vbCrLf)
                    cxsebpodtl.Append(po(0) & vbTab &
                                  po(1) & vbTab &
                                  data(1) & vbTab &
                                  data(11) & vbTab &
                                  validReal(data(22)) & vbTab &
                                  validReal(data(23)) & vbTab &
                                  validint(data(39)) & vbTab &
                                  validstr(data(46)) & vbTab &
                                    DateFormatyyyyMMddString(data(21)) & vbCrLf)
                Else
                    podetailid = result.Item(0)
                End If

                myprogress = "relSalesDocPo"
                Dim pkey7(1) As Object
                pkey7(0) = salesorderdtlid
                pkey7(1) = podetailid
                result = ds.Tables(7).Rows.Find(pkey7)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(7).NewRow
                    seqrelSalesDocPo += 1
                    dr.Item(0) = seqrelSalesDocPo
                    dr.Item(1) = salesorderdtlid
                    dr.Item(2) = podetailid
                    ds.Tables(7).Rows.Add(dr)

                    relsalesdocpoid = seqrelSalesDocPo

                    cxrelsalesdocpo.Append(salesorderdtlid & vbTab & podetailid & vbCrLf)

                Else
                    relsalesdocpoid = result.Item(0)
                End If

                'new additional

                'myprogress = "Check Loading Code"
                'Dim pkey12(0) As Object
                'pkey12(0) = data(44)
                'result = ds.Tables(12).Rows.Find(pkey12)
                'If IsNothing(result) Then
                '    Dim dr As DataRow = ds.Tables(12).NewRow
                '    dr.Item(0) = data(44)

                '    ds.Tables(12).Rows.Add(dr)
                '    'loadingcode,loadingname
                '    loadingsb.Append(data(44) & vbTab &
                '                    data(45) & vbCrLf)
                'End If

                'myprogress = "Check Brandid"
                'If data(13) <> "" Then
                '    Dim pkey13(0) As Object
                '    pkey13(0) = data(13)
                '    result = ds.Tables(13).Rows.Find(pkey13)
                '    If IsNothing(result) Then
                '        Dim dr As DataRow = ds.Tables(13).NewRow
                '        dr.Item(0) = data(13)
                '        ds.Tables(13).Rows.Add(dr)
                '        'brandid,brandname
                '        brandsb.Append(data(13) & vbTab &
                '                        data(14) & vbCrLf)
                '    End If
                'End If

                myprogress = "Check CMMF"
                Dim pkey12(0) As Object
                pkey12(0) = data(11)
                result = ds.Tables(12).Rows.Find(pkey12)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(12).NewRow
                    dr.Item(0) = data(11)
                    dr.Item(1) = data(15)
                    ds.Tables(12).Rows.Add(dr)
                    'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                    'cmmfsb.Append(data(11) & vbTab &
                    '               data(10) & vbTab &
                    '               data(15) & vbTab &
                    '               data(16) & vbTab &
                    '               validstr(vendorcode) & vbTab &
                    '               validstr(data(12)) & vbTab &
                    '               validstr(data(44)) & vbTab &
                    '               validint(data(13)) & vbCrLf)
                Else
                    'update cmmf
                    'check cmmf
                    'Dim pkey16(0) As Object
                    'pkey16(0) = data(11)
                    'result = ds.Tables(16).Rows.Find(pkey16)
                    'If IsNothing(result) Then
                    '    Dim dr As DataRow = ds.Tables(16).NewRow
                    '    dr.Item(0) = data(11)
                    '    ds.Tables(16).Rows.Add(dr)
                    '    If cmmfUpdateSB.Length > 0 Then
                    '        cmmfUpdateSB.Append(",")
                    '    End If
                    '    cmmfUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying,'{6}'::character varying,'{7}'::character varying]",
                    '                                      data(11), data(10), data(15), data(16), validstr(vendorcode), validstr(data(12)), validstr(data(44)), validstr(data(13))))
                    'Else
                    '    'MessageBox.Show("updatecmmf")
                    'End If

                    'update
                    Dim myitem As String = "" & result.Item("itemid")
                    If myitem.Trim <> data(15) Then
                        If Not updatedict.ContainsKey(data(11)) Then
                            updatedict.Add(data(11), data(15))
                            If cmmfUpdateSB.Length > 0 Then
                                cmmfUpdateSB.Append(",")
                            End If
                            cmmfUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", data(11), data(15)))
                        End If
                    End If
                End If

            End If

            'All Data(0) Type 
            seqsebodtp += 1
            cxSebodtp.Append(relsalesdocpoid & vbTab &
                             data(0) & vbTab &
                             DateFormatyyyyMMddString(data(18)) & vbTab &
                             validstr(data(19)) & vbTab &
                             validstr(data(44)) & vbTab &
                             DateFormatyyyyMMddString(data(33)) & vbTab &
                             DateFormatyyyyMMddString(data(34)) & vbTab &
                             validint(data(35)) & vbTab &
                             validint(data(36)) & vbTab &
                             DateFormatyyyyMMddString(data(37)) & vbTab &
                             DateFormatyyyyMMddString(data(38)) & vbTab &
                             validstr(data(47)) & vbCrLf)
            'cxSebodtp.Append(relsalesdocpoid & vbTab &
            '                 data(0) & vbTab &
            '                 DateFormatyyyyMMddString(data(18)) & vbTab &
            '                 validstr(data(19)) & vbTab &
            '                 validstr(data(44)) & vbCrLf)

            If data(0) = "Header" Then
                cxconfother.Append(seqsebodtp & vbTab &
                                   DateFormatyyyyMMddString(data(31)) & vbTab &
                                   validint(data(32)) & vbCrLf)

                'add Confirmation for Header
                cxconf.Append(seqsebodtp & vbTab &
                             DateFormatyyyyMMddString(data(33)) & vbTab &
                             DateFormatyyyyMMddString(data(34)) & vbTab &
                             validint(data(35)) & vbCrLf)

                'add shipment for Header
                If IsNumeric(data(36)) Then
                    cxshipment.Append(seqsebodtp & vbTab &
                                  validint(data(36)) & vbTab &
                                  DateFormatyyyyMMddString(data(37)) & vbTab &
                                  DateFormatyyyyMMddString(data(38)) & vbCrLf)
                End If


            End If

            If data(0) = "Confirmation" Then
                'create cxconf
                seqconf += 1
                confid = seqconf
                cxconf.Append(seqsebodtp & vbTab &
                              DateFormatyyyyMMddString(data(33)) & vbTab &
                              DateFormatyyyyMMddString(data(34)) & vbTab &
                              validint(data(35)) & vbCrLf)
                'create cxconfother confstatus
                If data(30) <> "" Then
                    'cxconfstatus.Append(confid & vbTab & data(30) & vbCrLf)
                    cxconfstatus.Append(seqsebodtp & vbTab & data(30) & vbCrLf)
                End If
            End If
            If data(0) = "Shipment" Then
                'create shipment
                cxshipment.Append(seqsebodtp & vbTab &
                                  validint(data(36)) & vbTab &
                                  DateFormatyyyyMMddString(data(37)) & vbTab &
                                  DateFormatyyyyMMddString(data(38)) & vbCrLf)
                'create shipmentother
                cxshipmentother.Append(seqsebodtp & vbTab &
                                       validstr(data(41)) & vbTab &
                                       validstr(data(42)) & vbTab &
                                       validstr(data(43)) & vbCrLf
                                       )

            End If
            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2},Row Num {3} {4}", myprogress, ex.Message, mydata.filename, mydata.rownumber, vbCrLf)
        End Try
        Return myret
    End Function



    Private Function Comp(ByVal mydata As Object, ByRef message As String) As Boolean
        Dim myret As Boolean = False
        Dim myprogress As String = String.Empty
        Dim comments As String = String.Empty
        Dim salesorder() As String
        Dim data = mydata.data

        Try
            If mydata.rownumber = 1867 Then
                Debug.Print(mydata.rownumber)
            End If
            If data(0) = "Header" Then
                myprogress = "SalesOrder"
                salesorder = data(6).ToString.Split("/")
                'find sales order in dataset.table(cxsalesorder) 
                'if avail then no need to create record in table cxsalesorder
                Dim result As DataRow

                'Salesorder Hd
                Dim pkey0(0) As Object
                pkey0(0) = salesorder(0)
                result = ds.Tables(0).Rows.Find(pkey0)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(0).NewRow
                    dr.Item(0) = salesorder(0)
                    ds.Tables(0).Rows.Add(dr)

                    cxSalesOrder.Append(salesorder(0) & vbTab &
                                        data(3) & vbTab &
                                        data(14) & vbTab &
                                        validReal(data(17)) & vbTab &
                                        "Null" & vbCrLf)
                End If

                myprogress = "SalesOrder Dtl"
                'SalesOrder Dtl
                Dim pkey1(1) As Object
                pkey1(0) = salesorder(0)
                pkey1(1) = salesorder(1)
                result = ds.Tables(1).Rows.Find(pkey1)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(1).NewRow
                    seqsalesorderdtl += 1
                    dr.Item(0) = seqsalesorderdtl
                    dr.Item(1) = salesorder(0)
                    dr.Item(2) = salesorder(1)
                    ds.Tables(1).Rows.Add(dr)

                    salesorderdtlid = seqsalesorderdtl
                    cxSalesOrderdtl.Append(salesorder(0) & vbTab &
                                           salesorder(1) & vbTab &
                                           data(5) & vbTab &
                                           data(11) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(18))) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(19))) & vbTab &
                                           data(20) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(21))) & vbTab &
                                           DateFormatyyyyMMdd(CDate(data(22))) & vbTab &
                                           data(23) & vbCrLf)
                Else
                    salesorderdtlid = result.Item(0)
                End If

                myprogress = "Check Customer Ship to Party"
                Dim pkey2(0) As Object
                pkey2(0) = data(1)
                result = ds.Tables(4).Rows.Find(pkey2)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(4).NewRow
                    dr.Item(0) = data(1)
                    ds.Tables(4).Rows.Add(dr)
                    customer.Append(data(1) & vbTab &
                                    data(2) & vbCrLf)
                End If

                myprogress = "Check Customer SoldToParty"
                Dim pkey3(0) As Object
                pkey3(0) = data(3)
                result = ds.Tables(4).Rows.Find(pkey3)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(4).NewRow
                    dr.Item(0) = data(3)

                    ds.Tables(4).Rows.Add(dr)
                    customer.Append(data(3) & vbTab &
                                    data(4) & vbCrLf)
                End If

                myprogress = "Check PO Header"
                Dim pkey4(0) As Object
                Dim po = data(34).ToString.Split("/")
                pkey4(0) = po(0)
                result = ds.Tables(2).Rows.Find(pkey4)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(2).NewRow
                    dr.Item(0) = po(0)

                    ds.Tables(2).Rows.Add(dr)
                    cxsebpo.Append(po(0) & vbTab &
                                    DateFormatyyyyMMddString(data(15)) & vbCrLf)
                End If

                myprogress = "Check PO Detail"
                Dim pkey5(1) As Object
                pkey5(0) = po(0)
                pkey5(1) = po(1)
                result = ds.Tables(3).Rows.Find(pkey5)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(3).NewRow
                    seqsebpodtl += 1
                    dr.Item(0) = seqsebpodtl
                    dr.Item(1) = po(0)
                    dr.Item(2) = po(1)
                    ds.Tables(3).Rows.Add(dr)

                    podetailid = seqsebpodtl
                    'For j = 39 To data.length - 1
                    '    comments += data(j)
                    'Next
                    'cxsebpodtl.Append(po(0) & vbTab &
                    '                po(1) & vbTab &
                    '                data(1) & vbTab &
                    '                data(8) & vbTab &
                    '                validReal(data(16)) & vbTab &
                    '                validReal(data(17)) & vbTab &
                    '                validint(data(33)) & vbTab &
                    '                validstr(comments) & vbCrLf)
                    'cxsebpodtl.Append(po(0) & vbTab &
                    '               po(1) & vbTab &
                    '               data(1) & vbTab &
                    '               data(8) & vbTab &
                    '               validReal(data(16)) & vbTab &
                    '               validReal(data(17)) & vbTab &
                    '               validint(data(33)) & vbTab &
                    '               validstr(data(39)) & vbCrLf)
                    cxsebpodtl.Append(po(0) & vbTab &
                                  po(1) & vbTab &
                                  data(1) & vbTab &
                                  data(8) & vbTab &
                                  validReal(data(16)) & vbTab &
                                  validReal(data(17)) & vbTab &
                                  validint(data(33)) & vbTab &
                                  validstr(data(39)) & vbTab &
                                    DateFormatyyyyMMddString(data(15)) & vbCrLf)
                Else
                    podetailid = result.Item(0)
                End If

                myprogress = "relSalesDocPo"
                Dim pkey7(1) As Object
                pkey7(0) = salesorderdtlid
                pkey7(1) = podetailid
                result = ds.Tables(7).Rows.Find(pkey7)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(7).NewRow
                    seqrelSalesDocPo += 1
                    dr.Item(0) = seqrelSalesDocPo
                    dr.Item(1) = salesorderdtlid
                    dr.Item(2) = podetailid
                    ds.Tables(7).Rows.Add(dr)

                    relsalesdocpoid = seqrelSalesDocPo
                    'For j = 39 To data.length - 1
                    '    comments += data(j)
                    'Next
                    cxrelsalesdocpo.Append(salesorderdtlid & vbTab & podetailid & vbCrLf)

                Else
                    relsalesdocpoid = result.Item(0)
                End If

                'New Additional

                myprogress = "Check CMMF"
                Dim pkey12(0) As Object
                pkey12(0) = data(8)
                result = ds.Tables(12).Rows.Find(pkey12)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(12).NewRow
                    dr.Item(0) = data(8)
                    dr.Item(1) = data(9)
                    ds.Tables(12).Rows.Add(dr)
                    'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                    'cmmfsb.Append(data(11) & vbTab &
                    '               data(10) & vbTab &
                    '               data(15) & vbTab &
                    '               data(16) & vbTab &
                    '               validstr(vendorcode) & vbTab &
                    '               validstr(data(12)) & vbTab &
                    '               validstr(data(44)) & vbTab &
                    '               validint(data(13)) & vbCrLf)
                Else
                    'update cmmf
                    'check cmmf
                    'Dim pkey16(0) As Object
                    'pkey16(0) = data(11)
                    'result = ds.Tables(16).Rows.Find(pkey16)
                    'If IsNothing(result) Then
                    '    Dim dr As DataRow = ds.Tables(16).NewRow
                    '    dr.Item(0) = data(11)
                    '    ds.Tables(16).Rows.Add(dr)
                    '    If cmmfUpdateSB.Length > 0 Then
                    '        cmmfUpdateSB.Append(",")
                    '    End If
                    '    cmmfUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying,'{6}'::character varying,'{7}'::character varying]",
                    '                                      data(11), data(10), data(15), data(16), validstr(vendorcode), validstr(data(12)), validstr(data(44)), validstr(data(13))))
                    'Else
                    '    'MessageBox.Show("updatecmmf")
                    'End If

                    'update
                    Dim myitem As String = "" & result.Item("itemid")
                    If myitem.Trim <> data(9) Then
                        If Not updatedict.ContainsKey(data(8)) Then
                            updatedict.Add(data(8), data(9))
                            If cmmfUpdateSB.Length > 0 Then
                                cmmfUpdateSB.Append(",")
                            End If
                            cmmfUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", data(8), data(9)))
                        End If
                    End If
                End If

            End If

            'All Data(0) Type 
            seqsebodtp += 1
            cxSebodtp.Append(relsalesdocpoid & vbTab &
                             data(0) & vbTab &
                             DateFormatyyyyMMddString(data(12)) & vbTab &
                             validstr(data(13)) & vbTab &
                             validstr(data(38)) & vbTab &
                             DateFormatyyyyMMddString(data(27)) & vbTab &
                             DateFormatyyyyMMddString(data(28)) & vbTab &
                             validint(data(29)) & vbTab &
                             validint(data(30)) & vbTab &
                             DateFormatyyyyMMddString(data(31)) & vbTab &
                             DateFormatyyyyMMddString(data(32)) & vbTab &
                             validstr(data(40)) & vbCrLf)

            'cxSebodtp.Append(relsalesdocpoid & vbTab &
            '                 data(0) & vbTab &
            '                 DateFormatyyyyMMddString(data(12)) & vbTab &
            '                 validstr(data(13)) & vbTab &
            '                 validstr(data(38)) & vbCrLf)

            If data(0) = "Header" Then
                cxconfother.Append(seqsebodtp & vbTab &
                                   DateFormatyyyyMMddString(data(25)) & vbTab &
                                   validint(data(26)) & vbCrLf)


                'Add Confirmation
                cxconf.Append(seqsebodtp & vbTab &
                              DateFormatyyyyMMddString(data(27)) & vbTab &
                              DateFormatyyyyMMddString(data(28)) & vbTab &
                              validint(data(29)) & vbCrLf)
                If IsNumeric(data(30)) Then
                    'Add Shipment
                    cxshipment.Append(seqsebodtp & vbTab &
                                      validint(data(30)) & vbTab &
                                      DateFormatyyyyMMddString(data(31)) & vbTab &
                                      DateFormatyyyyMMddString(data(32)) & vbCrLf)
                End If

            End If

            If data(0) = "Confirmation" Then
                'create cxconf
                seqconf += 1
                confid = seqconf
                cxconf.Append(seqsebodtp & vbTab &
                              DateFormatyyyyMMddString(data(27)) & vbTab &
                              DateFormatyyyyMMddString(data(28)) & vbTab &
                              validint(data(29)) & vbCrLf)
                'create cxconfother


                'create cxconfstatus
                If data(24) <> "" Then
                    'cxconfstatus.Append(confid & vbTab & data(24) & vbCrLf)
                    cxconfstatus.Append(seqsebodtp & vbTab & data(24) & vbCrLf)
                End If
            End If
            If data(0) = "Shipment" Then
                'create shipment
                cxshipment.Append(seqsebodtp & vbTab &
                                  validint(data(30)) & vbTab &
                                  DateFormatyyyyMMddString(data(31)) & vbTab &
                                  DateFormatyyyyMMddString(data(32)) & vbCrLf)
                'create shipmentother
                cxshipmentother.Append(seqsebodtp & vbTab &
                                       validstr(data(35)) & vbTab &
                                       validstr(data(36)) & vbTab &
                                       validstr(data(37)) & vbCrLf
                                       )

            End If
            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2},Row Num {3}", myprogress, ex.Message, mydata.filename, mydata.rownumber)
        End Try
        Return myret
    End Function
End Class

Public Class myData
    Public Property filename As String
    Public Property rownumber As Long
    Public Property data As Object

    Public Sub New(ByVal filename As String, ByVal rownumber As Long, ByVal data As Object)
        Me.filename = filename
        Me.rownumber = rownumber
        Me.data = data
    End Sub
End Class