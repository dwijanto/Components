Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass

Public Class ImportWORFG
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)

    Dim addShipment As Boolean = False
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


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
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

        'add program lock
        If DbAdapter1.getproglock("FImportweeklyreportNewFormat", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
        Else
            ProgressReport(2, "Read Folder..")

            ReadFileStatus = ImportTextFile(mySelectedPath, errMsg)
            If ReadFileStatus Then
                sw.Stop()
                ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                'release program lock
                DbAdapter1.getproglock("FImportweeklyreportNewFormat", HelperClass1.UserInfo.DisplayName, 0)
            Else
                If errMsg.Length > 100 Then
                    Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                        mystream.WriteLine(errMsg.ToString)
                    End Using
                    Process.Start(Application.StartupPath & "\error.txt")
                    ProgressReport(3, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                Else
                    ProgressReport(3, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errMsg))
                End If

            End If
        End If
        ProgressReport(5, "Set to continuous mode again")
        sw.Stop()
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

    Private Function ImportTextFile(ByVal FileName As String, ByRef errMsg As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of myData)


        Dim i As Long
        Try
            'Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)

            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.csv")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            For Each fi As IO.FileInfo In arrFI
                ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))
                Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(";")
                        .HasFieldsEnclosedInQuotes = True
                        Dim count As Long = 0

                        Do Until .EndOfData
                            'If count > 0 Then
                            myrecord = .ReadFields
                            If myrecord.Length = 1 Then
                                Exit For
                            End If
                            If count > 0 Then
                                Dim mydata As New myData(fi.FullName, count + 1, myrecord)
                                myList.Add(mydata)
                            End If

                            tcount += 1
                            'End If
                            count += 1

                        Loop
                    End With
                End Using
            Next
            If myList.Count = 0 Then
                errMsg = "Text File Wrong Format"
                ProgressReport(5, "Set To Marque")
                Return myret
            End If
            'get dataset
            Dim DS As New DataSet
            DS.CaseSensitive = True
            'get initial keys from Database fro related table
            ProgressReport(2, String.Format("Delete rows ..........."))

            DbAdapter1.ExecuteStoreProcedure("deleteworfg")
            ProgressReport(2, String.Format("Get Initial Data ..........."))
            If Not FillDataset(DS, errMsg) Then
                Return False
            End If

            'Create object for handleing row creation
            Dim WOR As New WORFG(DS)

            ProgressReport(2, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")
            For i = 0 To myList.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & myList.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                If Not WOR.buildSB(errMsg, myList(i)) Then
                    Return False
                End If

                'End If
            Next
            ProgressReport(6, "Set To Marque")
            ProgressReport(2, String.Format("Copy To Db"))
            If WOR.ErrCheck.Length > 0 Then
                errMsg = WOR.ErrCheck.ToString
                'Using sw As New StreamWriter(Application.StartupPath & "\error.txt")
                '    sw.WriteLine(WOR.ErrCheck.ToString)
                'End Using
                'Process.Start(Application.StartupPath & "\error.txt")
            Else
                If Not WOR.copyToDb(errMsg, Me) Then
                    Return False
                Else
                    'check append shipment
                    If addShipment Then
                        If Not WOR.ImportShipment(errMsg, Me) Then
                            Return False
                        End If
                    End If
                End If
                myret = True
            End If
            


        Catch ex As Exception
            errMsg = String.Format("Row : {0} ", i) & ex.Message
        End Try
        'copy

        'myret = True
        'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(SW.Elapsed.Minutes, "00"), Format(SW.Elapsed.Seconds, "00"), SW.Elapsed.Milliseconds.ToString))
        Return myret
    End Function

    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As String) As Boolean
        Dim myret As Boolean = False
        Dim Sqlstr As String = " select sebasiasalesorder from ohd;" &
                               " select sebasiasalesorder,solineno,orderdtlid from odtl;" &
                               " select customercode,customername from customer;" &
                               " select loadingcode,loadingname from loading;" &
                               " select brandid,brandname from brand;" &
                               " select cmmf from cmmf;" &
                               " select max(officerid)as officerid,officername::character varying from officer group by officername having not officername = '';" &
                               " select vendorcode,vendorname::character varying from vendor where not vendorcode in (select cvalue::bigint from paramdt where paramname = 'vendorduplicate') order by vendorname;" &
                               " select cvalue,ivalue from paramdt where paramname = 'vendordict';" &
                               " select cvalue,ivalue from paramdt where paramname = 'vendormap';" &
                               " select sebasiapono from po;" &
                               " select ordertype,sebasiasalesorder,solineno from firstconfirmation;" &
                               " select confirmationid from confs;" &
                               " select dvalue from paramhd where paramname = 'LatestUpdate';" &
                               " select cmmf from cmmf where cmmf = 0"

        If DbAdapter1.TbgetDataSet(Sqlstr, DS, errmessage) Then
            DS.Tables(0).TableName = "ohd"
            DS.Tables(1).TableName = "odtl"
            DS.Tables(2).TableName = "customer"
            DS.Tables(3).TableName = "loading"
            DS.Tables(4).TableName = "brand"
            DS.Tables(5).TableName = "cmmf"
            DS.Tables(6).TableName = "sao"
            DS.Tables(7).TableName = "vendor"
            DS.Tables(8).TableName = "vendordict"
            DS.Tables(9).TableName = "vendormap"
            DS.Tables(10).TableName = "po"
            DS.Tables(11).TableName = "firstconfirmation"
            DS.Tables(12).TableName = "confs"
            DS.Tables(13).TableName = "cmmfupdate"

            Dim idx0(0) As DataColumn               'ohd
            idx0(0) = DS.Tables(0).Columns(0)       'sebasiasalesorder
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(1) As DataColumn               'odtl
            idx1(0) = DS.Tables(1).Columns(1)       'sebasiasalesorder    
            idx1(1) = DS.Tables(1).Columns(2)       'solineno
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn               'customer
            idx2(0) = DS.Tables(2).Columns(0)       'customercode
            DS.Tables(2).PrimaryKey = idx2

            Dim idx3(0) As DataColumn               'loading
            idx3(0) = DS.Tables(3).Columns(0)       'loadingcode
            DS.Tables(3).PrimaryKey = idx3

            Dim idx4(0) As DataColumn               'brand
            idx4(0) = DS.Tables(4).Columns(0)       'brandid
            DS.Tables(4).PrimaryKey = idx4

            Dim idx5(0) As DataColumn               'cmmf
            idx5(0) = DS.Tables(5).Columns(0)
            DS.Tables(5).PrimaryKey = idx5

            Dim idx6(0) As DataColumn               'sao
            idx6(0) = DS.Tables(6).Columns(1)
            DS.Tables(6).PrimaryKey = idx6

            Dim idx7(0) As DataColumn               'vendor
            idx7(0) = DS.Tables(7).Columns(1)
            DS.Tables(7).PrimaryKey = idx7

            Dim idx10(0) As DataColumn               'po (Final Customer Order)
            idx10(0) = DS.Tables(10).Columns(0)
            DS.Tables(10).PrimaryKey = idx10

            Dim idx11(2) As DataColumn               'first confirmation
            idx11(0) = DS.Tables(11).Columns(0)
            idx11(1) = DS.Tables(11).Columns(1)
            idx11(2) = DS.Tables(11).Columns(2)
            DS.Tables(11).PrimaryKey = idx11

            Dim idx12(0) As DataColumn               'confs
            idx12(0) = DS.Tables(12).Columns(0)
            DS.Tables(12).PrimaryKey = idx12
            myret = True

            Dim idx14(0) As DataColumn               'confs
            idx14(0) = DS.Tables(14).Columns(0)
            DS.Tables(14).PrimaryKey = idx14
            myret = True
        Else
            Return False
        End If

        Return myret
    End Function

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        addshipment = CheckBox1.Checked
    End Sub


End Class
Public Class WORFG

    Dim headerconfid As Long
    Dim seqship As Long
    Dim shipid As Long
    Dim myLatestUpdate As String
    Dim lsodtlid As Object
    Dim lsodtlidseq As Object
    Dim lsodtypeidseq As Object
    Dim lsshipidseq As Object
    Dim lspodtlidseq As Object
    Dim lspodtlid As Object

    Public Property ds As DataSet
    Public Property ohdSB As New StringBuilder
    Public Property odtlsb As New StringBuilder
    Public Property posb As New StringBuilder
    Public Property odtpsb As New StringBuilder
    Public Property confsb As New StringBuilder
    Public Property firstconfirmationsb As New StringBuilder
    Public Property confssb As New StringBuilder
    Public Property shipsb As New StringBuilder
    Public Property shipdtlsb As New StringBuilder
    Public Property customersb As New StringBuilder
    Public Property loadingsb As New StringBuilder
    Public Property brandsb As New StringBuilder
    Public Property cmmfsb As New StringBuilder
    Public Property cmmfUpdateSB As New StringBuilder


    Dim lsohdsb As New StringBuilder
    Dim lsodtlsb As New StringBuilder
    Dim lspohdsb As New StringBuilder
    Dim lspodtlsb As New StringBuilder
    Dim lsodtypesb As New StringBuilder
    Dim lsshipsb As New StringBuilder
    Dim lsshipdtlsb As New StringBuilder

    Public Property ErrCheck As New StringBuilder
    Public Property myVendorDict As New Dictionary(Of String, Long)
    'Public Property myVendorMapDict As New Dictionary(Of String, Long)

    Dim myVendorList As String
    Dim myVendorMapList As String
    Dim seqsalesorderdtl As Long
    Dim seqsebpodtl As Long
    Dim seqrelSalesDocPo As Long

    Dim seqodtl As Long = 0
    Dim seqodtp As Long = 0

    Dim seqconf As Long

    Dim podetailid As Long
    Dim salesorderdtlid As Long
    Dim relsalesdocpoid As Long
    Dim confid As Long
    Dim odtlid As Long

    Public Sub New(ByVal ds As DataSet)
        Me.ds = ds
        Me.myVendorDict = myVendorDict
        For Each dr As DataRow In ds.Tables(8).Rows
            myVendorDict.Add(dr.Item(0).ToString.Trim, dr.Item(1).ToString)
        Next
        'For Each dr As DataRow In ds.Tables(9).Rows
        '    myvendormapDict.Add(dr.Item(0).ToString.Trim, dr.Item(1).ToString)
        'Next

        For Each dr As DataRow In ds.Tables(8).Rows
            myVendorList = myVendorList + IIf(myVendorList = "", "", ",") + dr.Item(0)
        Next
        'For Each dr As DataRow In ds.Tables(9).Rows
        '    myVendorMapList = myVendorMapList + IIf(myVendorMapList = "", "", ",") + dr.Item(0)
        'Next
        myLatestUpdate = DateFormatyyyyMMdd(CDate(ds.Tables(13).Rows(0).Item(0)))
    End Sub

    Public Function buildSB(ByRef message As String, ByRef mydata As myData) As Boolean
        Dim myret As Boolean = False
        'Dim myprogress As String = String.Empty


        Try
            If mydata.data.length >= 41 Then
                myret = fg(mydata, message)
            Else
                message = String.Format("Data Length to short. File Name: {0} ,Row Number:{1}", mydata.filename, mydata.rownumber)
                myret = False
            End If
        Catch ex As Exception
            message = ex.Message & "::" & vbCrLf & String.Format("Data Length to short. File Name: {0} ,Row Number:{1}", mydata.filename, mydata.rownumber)
        End Try
        Return myret
    End Function

    Public Function copyToDb(ByRef errMsg As String, ByVal myform As ImportWORFG) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String
        Try
            If customersb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Customer"))
                sqlstr = "copy customer(customercode,customername) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, customersb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            If ohdSB.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy ohd"))
                sqlstr = "copy ohd(sebasiasalesorder,customercode,orderstatus,soldto,officerid)  from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, ohdSB.ToString, myret)
                If Not myret Then
                    Return False
                End If

            End If

            'brand
            If brandsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Brand"))
                sqlstr = "copy brand(brandid,brandname) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, brandsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'loadingcode
            If loadingsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy loading"))
                sqlstr = "copy loading(loadingcode,loadingname) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, loadingsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

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
                sqlstr = "update cmmf set rir= foo.rir,itemid = foo.itemid,vendorcode = foo.vendorcode::integer,comfam = foo.comfam::integer,loadingcode = foo.loadingcode,brandid = foo.brandid::integer from (select * from array_to_set8(Array[" & cmmfUpdateSB.ToString &
                         "]) as tb (id character varying,rir character varying,itemid character varying,materialdesc character varying,vendorcode character varying,comfam character varying,loadingcode character varying,brandid character varying))foo where cmmf = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errMsg) Then
                    Return False
                End If
            End If

            'odtl
            If odtlsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy odtl"))
                sqlstr = "copy odtl(cmmf,osqty,sebasiapono,comments,sebasiasalesorder,solineno,polineno,receptiondate,customerorderno,vendorcode) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, odtlsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If


            'PO
            If posb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy po"))
                sqlstr = "copy po(sebasiapono,finalcustomerorder) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, posb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'odtp
            'Using sw As New StreamWriter(Application.StartupPath & "\error.txt")
            '    sw.WriteLine(odtpsb.ToString)
            'End Using
            'Process.Start(Application.StartupPath & "\error.txt")
            If odtpsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy odtp"))
                sqlstr = "copy odtp(updatesince7,curinq,latestupdate,inquiryetd,currentinquiryetd,currentinquiryqty,fob,unittp,inquiryeta,inquiryqty,currentinquiryeta,orderdtlid,ordertype,shipfrom) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, odtpsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'conf
            If confsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy conf"))
                sqlstr = "copy conf(currentconfirmedeta,currentconfirmedetd,orderdtltypeid,currentconfirmedqty) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, confsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'firstconfirmation
            If firstconfirmationsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy firstconfirmation"))
                sqlstr = "copy firstconfirmation(ordertype,sebasiasalesorder,solineno,""1stconfirmedetd"",""1stconfirmedqty"") from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, firstconfirmationsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'confs
            If confssb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy confs"))
                sqlstr = "copy confs(confirmationid,confirmationstatus) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, confssb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'ship
            If shipsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy ship"))
                sqlstr = "copy ship(deliveredqty,shipdate,shipdateeta,orderdtltypeid) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, shipsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            'shipmentdtl
            If shipdtlsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy ship"))
                sqlstr = "copy shipmentdtl(shipmentid,ctrno,boatid,packinglist,shipfrom) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, shipdtlsb.ToString, myret)
                If Not myret Then
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
        'If IsDBNull(p1) Then
        '    Return "Null"
        'ElseIf p1 = "" Then
        '    Return "Null"
        'Else
        '    Return CInt(p1)
        'End If
        If IsNumeric(p1) Then
            Return p1
        Else
            Return "Null"
        End If
    End Function
    Private Function validReal(ByVal data As Object) As Object
        'If IsDBNull(data) Then
        '    Return "Null"
        'ElseIf data = "" Then
        '    Return "Null"
        'Else
        '    Return data
        'End If
        If IsNumeric(data) Then
            Return data
        Else
            Return "Null"
        End If
    End Function
    Private Function validstr(ByVal data As Object) As Object
        If IsDBNull(data) Then
            Return "Null"
        ElseIf data = "" Then
            Return "Null"
        End If
        Return data
    End Function

    Private Function fg(ByRef mydata As myData, ByRef message As String) As Boolean
        Dim myret = False
        Dim myprogress As String = String.Empty
        Dim data = mydata.data
        Dim comments As String = String.Empty
        Dim salesorder() As String
        Dim sebpono() As String

        Try

            'If mydata.data(11) = 8000033438 Then
            '    Debug.Print("Debug Mode")
            'End If
            'check valid data

            If Not ValidateData(mydata, ErrCheck) Then
                Return True
            End If

            salesorder = data(7).ToString.Split("/")
            sebpono = data(40).ToString.Split("/")

            If data(0) = "Header" Then
                Dim saoid As String = String.Empty
                Dim result As DataRow
                Dim vendorcode As String = String.Empty


                'officerseb
                myprogress = "SAO"
                Dim pkey6(0) As Object
                pkey6(0) = data(5)
                result = ds.Tables(6).Rows.Find(pkey6)
                If Not IsNothing(result) Then
                    saoid = result.Item(0).ToString
                End If

                myprogress = "Check Customer Ship to Party"
                Dim pkey2(0) As Object
                pkey2(0) = data(1)
                result = ds.Tables(2).Rows.Find(pkey2)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(2).NewRow
                    dr.Item(0) = data(1)
                    ds.Tables(2).Rows.Add(dr)
                    'customercode,customername
                    customersb.Append(data(1) & vbTab &
                                    data(2) & vbCrLf)
                End If

                myprogress = "Check Customer SoldToParty"
                pkey2(0) = data(3)
                result = ds.Tables(2).Rows.Find(pkey2)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(2).NewRow
                    dr.Item(0) = data(3)

                    ds.Tables(2).Rows.Add(dr)
                    'customercode,customername
                    customersb.Append(data(3) & vbTab &
                                    data(4) & vbCrLf)
                End If


                'vendorcode
                myprogress = "Check Vendorcode"
                Dim pkey7(0) As Object
                pkey7(0) = data(8)
                result = ds.Tables(7).Rows.Find(pkey7)
                If IsNothing(result) Then
                    If myVendorList.Contains(data(8)) Then
                        vendorcode = myVendorDict(data(8))
                    Else
                        ErrCheck.Append(String.Format("No Vendorcode for '{0}'", data(8)) & vbCrLf)
                    End If
                Else
                    vendorcode = result.Item(0)
                End If

                'myprogress = "Check VendorMap"
                ''Check Mapping Vendorcode
                'If myVendorMapList.Contains(data(8)) Then
                '    vendorcode = myVendorMapDict(data(8))
                'End If

                myprogress = "Check Loading Code"
                Dim pkey3(0) As Object
                pkey3(0) = data(44)
                result = ds.Tables(3).Rows.Find(pkey3)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(3).NewRow
                    dr.Item(0) = data(44)

                    ds.Tables(3).Rows.Add(dr)
                    'loadingcode,loadingname
                    loadingsb.Append(data(44) & vbTab &
                                    data(45) & vbCrLf)
                End If
                myprogress = "Check Brandid"
                If data(13) <> "" Then
                    Dim pkey4(0) As Object
                    pkey4(0) = data(13)
                    result = ds.Tables(4).Rows.Find(pkey4)
                    If IsNothing(result) Then
                        Dim dr As DataRow = ds.Tables(4).NewRow
                        dr.Item(0) = data(13)
                        ds.Tables(4).Rows.Add(dr)
                        'brandid,brandname
                        brandsb.Append(data(13) & vbTab &
                                        data(14) & vbCrLf)
                    End If
                End If

                myprogress = "Check CMMF"
                Dim pkey5(0) As Object
                pkey5(0) = data(11)
                result = ds.Tables(5).Rows.Find(pkey5)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(5).NewRow
                    dr.Item(0) = data(11)
                    ds.Tables(5).Rows.Add(dr)
                    'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                    cmmfsb.Append(data(11) & vbTab &
                                   data(10) & vbTab &
                                   data(15) & vbTab &
                                   data(16) & vbTab &
                                   validstr(vendorcode) & vbTab &
                                   validstr(data(12)) & vbTab &
                                   validstr(data(44)) & vbTab &
                                   validint(data(13)) & vbCrLf)
                Else
                    'update cmmf
                    'check cmmf
                    Dim pkey14(0) As Object
                    pkey14(0) = data(11)
                    result = ds.Tables(14).Rows.Find(pkey14)
                    If IsNothing(result) Then
                        Dim dr As DataRow = ds.Tables(14).NewRow
                        dr.Item(0) = data(11)
                        ds.Tables(14).Rows.Add(dr)
                        If cmmfUpdateSB.Length > 0 Then
                            cmmfUpdateSB.Append(",")
                        End If
                        cmmfUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying,'{6}'::character varying,'{7}'::character varying]",
                                                          data(11), data(10), data(15), data(16), validstr(vendorcode), validstr(data(12)), validstr(data(44)), validstr(data(13))))
                    Else
                        'MessageBox.Show("updatecmmf")
                    End If
                End If

                myprogress = "ohd"

                'find sales order in dataset.table(salesorder)
                'if avail, no need to create record
                'ohd

                Dim pkey0(0) As Object
                pkey0(0) = salesorder(0)
                result = ds.Tables(0).Rows.Find(pkey0)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(0).NewRow
                    dr.Item(0) = salesorder(0)
                    ds.Tables(0).Rows.Add(dr)
                    'sebasiasalesorder,customercode,orderstatus,soldto,officerid
                    ohdSB.Append(salesorder(0) & vbTab &
                                 data(1) & vbTab &
                                 data(17) & vbTab &
                                 data(3) & vbTab &
                                 saoid & vbCrLf)
                End If

                myprogress = "odtl"

                Dim pkey1(1) As Object
                pkey1(0) = salesorder(0)
                pkey1(1) = salesorder(1)
                result = ds.Tables(1).Rows.Find(pkey1)
                If IsNothing(result) Then
                    myprogress = "comments"
                    Dim myComment As String = String.Empty
                    For number = 46 To UBound(data)
                        myComment = myComment & IIf(myComment = "", "", " ") & Trim(data(number))
                    Next number

                    myprogress = "odtlsb"
                    seqodtl += 1
                    odtlid = seqodtl

                    Dim dr As DataRow = ds.Tables(1).NewRow
                    dr.Item(0) = salesorder(0)
                    dr.Item(1) = salesorder(1)
                    dr.Item(2) = odtlid
                    ds.Tables(1).Rows.Add(dr)
                    'cmmf,osqty,sebasiapono,comments,sebasiaslesorder,solineno,polineno,receptiondate,customerorderno,vendorcode
                    odtlsb.Append(data(11) & vbTab &
                                   validint(data(39)) & vbTab &
                                   sebpono(0) & vbTab &
                                   myComment & vbTab &
                                   salesorder(0) & vbTab &
                                   salesorder(1) & vbTab &
                                   sebpono(1) & vbTab &
                                    DateFormatyyyyMMddString(data(21)) & vbTab &
                                 data(6) & vbTab &
                                 validstr(vendorcode) & vbCrLf)

                    myprogress = "Check Final Customer Order"
                    'Check Final Customer Order
                    If UBound(data) > 40 Then
                        If data(45) <> "" Then
                            Dim pkey10(0) As Object
                            pkey10(0) = sebpono(0)
                            result = ds.Tables(10).Rows.Find(pkey10)
                            If IsNothing(result) Then
                                dr = ds.Tables(10).NewRow
                                dr.Item(0) = sebpono(0)
                                ds.Tables(10).Rows.Add(dr)
                                'sebasiapono,finalcustomerorder
                                posb.Append(sebpono(0) & vbTab &
                                                data(45) & vbCrLf)
                            End If
                        End If
                    End If
                Else
                    odtlid = result.Item(2)
                End If
            End If
            myprogress = "odtp"
            'Insert Odtp
            seqodtp += 1
            'updatesince7,curinq,latestupdate,inquiryetd,currentinquiryetd,currentinquiryqty,fob,unittp,inquiryeta,inquiryqty,currentinquiryeta,orderdtlid,ordertype,shipfrom
            If UBound(data) >= 41 Then
                odtpsb.Append(data(19) & vbTab &
                          data(20) & vbTab &
                          DateFormatyyyyMMddString(data(18)) & vbTab &
                          DateFormatyyyyMMddString(data(25)) & vbTab &
                          DateFormatyyyyMMddString(data(28)) & vbTab &
                          validint(data(29)) & vbTab &
                          validint(data(22)) & vbTab &
                          validint(data(23)) & vbTab &
                          DateFormatyyyyMMddString(data(24)) & vbTab &
                          validint(data(26)) & vbTab &
                          DateFormatyyyyMMddString(data(28)) & vbTab &
                          odtlid & vbTab &
                          data(0) & vbTab &
                          data(44) & vbCrLf)
            Else
                odtpsb.Append(data(19) & vbTab &
                          data(20) & vbTab &
                          DateFormatyyyyMMddString(data(18)) & vbTab &
                          DateFormatyyyyMMddString(data(25)) & vbTab &
                          DateFormatyyyyMMddString(data(28)) & vbTab &
                          validint(data(29)) & vbTab &
                          validint(data(22)) & vbTab &
                          validint(data(23)) & vbTab &
                          DateFormatyyyyMMddString(data(24)) & vbTab &
                          validint(data(26)) & vbTab &
                          DateFormatyyyyMMddString(data(28)) & vbTab &
                          odtlid & vbTab &
                          data(0) & vbTab &
                          "Null" & vbCrLf)
            End If


            myprogress = "Header and confirmation"
            If data(0) <> "Shipment" Then 'Header and Confirmation only
                seqconf += 1
                confid = seqconf
                'currentconfirmedeta,currentconfirmedetd,orderdtltypeid,currentconfirmedqty
                confsb.Append(DateFormatyyyyMMddString(data(33)) & vbTab &
                               DateFormatyyyyMMddString(data(34)) & vbTab &
                               seqodtp & vbTab &
                               validint(data(35)) & vbCrLf)
                If data(0) = "Header" Then
                    headerconfid = confid
                    If data(31) <> "" Then
                        myprogress = "FirstConfirmation"
                        'Find Firstconfirmation
                        Dim pkey11(2) As Object
                        pkey11(0) = "Header"
                        pkey11(1) = salesorder(0)
                        pkey11(2) = salesorder(1)
                        Dim result = ds.Tables(11).Rows.Find(pkey11)
                        Dim dr As DataRow
                        If IsNothing(result) Then
                            dr = ds.Tables(11).NewRow
                            dr.Item(0) = "Header"
                            dr.Item(1) = salesorder(0)
                            dr.Item(2) = salesorder(1)
                            ds.Tables(11).Rows.Add(dr)
                            'ordertype,sebasiasalesorder,solineno,""1stconfirmedetd"",""1stconfirmedqty""
                            firstconfirmationsb.Append("Header" & vbTab &
                                                       salesorder(0) & vbTab &
                                                       salesorder(1) & vbTab &
                                                        DateFormatyyyyMMddString(data(31)) & vbTab &
                                                       validint(data(32)) & vbCrLf)
                        End If
                    End If

                End If
                If data(0) = "Confirmation" Then
                    myprogress = "ConfirmationStatus"
                    If data(30) <> "" Then
                        'confirmationid,confirmationstatus
                        confssb.Append(confid & vbTab &
                                      data(30) & vbCrLf)


                        'Create also for header.
                        'Please Find a way to do this :)
                        'i think you can do it by find the last Header data
                        'Ok here we go, Find first from table(12) confs if not avail then you can create it.
                        Dim pkey12(0) As Object
                        pkey12(0) = headerconfid
                        Dim result = ds.Tables(12).Rows.Find(pkey12)
                        Dim dr As DataRow
                        If IsNothing(result) Then
                            dr = ds.Tables(12).NewRow
                            dr.Item(0) = headerconfid
                            ds.Tables(12).Rows.Add(dr)
                            confssb.Append(headerconfid & vbTab &
                                      data(30) & vbCrLf)
                        End If
                    End If
                End If
            End If
            myprogress = "Header and shipment"
            If data(0) <> "Confirmation" Then 'Header and Shipment only
                seqship += 1
                shipid = seqship
                myprogress = "Shipment"
                'deliveredqty,shipdate,shipdateeta,orderdtltypeid
                shipsb.Append(validint(data(36)) & vbTab &
                              DateFormatyyyyMMddString(data(37)) & vbTab &
                              DateFormatyyyyMMddString(data(38)) & vbTab &
                              seqodtp & vbCrLf)
                myprogress = "Shipmentdtl"
                'create shipment detail
                'shipmentid,ctrno,boatid,packinglist,shipfrom
                If UBound(data) >= 41 Then
                    shipdtlsb.Append(shipid & vbTab &
                                 validstr(data(41)) & vbTab &
                                 validstr(data(42)) & vbTab &
                                 validstr(data(43)) & vbTab &
                                 validstr(data(44)) & vbCrLf)
                End If
            End If

            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2}, Row Num {3} {4}", myprogress, ex.Message, mydata.filename, mydata.rownumber, vbCrLf)
        End Try
        Return myret
    End Function


    Private Function ValidateData(ByRef mydata As myData, ByRef errcheck As StringBuilder) As Boolean
        Dim myret As Boolean = False
        Try
            If mydata.data(17).ToString.Length > 10 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "OrderStatus more than 10 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(6).ToString.Length > 20 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "CustomerOrderNo more than 20 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(2).ToString.Length > 50 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "CustomerName more than 50 chars", mydata.filename, mydata.rownumber, vbCrLf))
                'Else
                '    mydata.data(2).ToString.Replace("'", "''")
            End If
            If mydata.data(4).ToString.Length > 50 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "ShipToPartyName more than 50 chars", mydata.filename, mydata.rownumber, vbCrLf))
                'Else
                '    mydata.data(4).ToString.Replace("'", "''")
            End If
            If mydata.data(8).ToString.Length > 50 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "Vendorname more than 50 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(16).ToString.Length > 50 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "MaterialDesc more than 50 chars", mydata.filename, mydata.rownumber, vbCrLf))
            Else
                mydata.data(16) = mydata.data(16).ToString.Replace("'", "''")
            End If
            If mydata.data(0).ToString.Length > 12 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "Header/item more than 12 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(19).ToString.Length > 10 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "Updatesince7 more than 10 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(20).ToString.Length > 10 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "currinq more than 10 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(44).ToString.Length > 30 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "shipfrom more than 30 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(30).ToString.Length > 15 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "ConfirmationStatus more than 15 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(14).ToString.Length > 30 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "BrandName more than 30 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            If mydata.data(45).ToString.Length > 30 Then
                errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2},", "FinalCustomerNumber more than 30 chars", mydata.filename, mydata.rownumber, vbCrLf))
            End If
            myret = True
        Catch ex As Exception
            errcheck.Append(String.Format("ErrorMessage: {0}, Filename : {1}, Rownumber :{2} {3}", ex.Message, mydata.filename, mydata.rownumber, vbCrLf))
        End Try
        Return myret

    End Function

    Public Function ImportShipment(ByRef errmsg As String, ByVal myform As ImportWORFG) As Boolean
        Dim myret As Boolean = False
        Dim ra As Long
        'Delete shipment bigger than latestupdate
        Dim sqlstr As String = "delete from lsodtype where lsodtypeid in  (select odtp.lsodtypeid from lsship ship" &
                 " left join lsodtype odtp on odtp.lsodtypeid= ship.lsodtypeid" &
                 " where latestupdate >= " & myLatestUpdate & ")"   '2012-02-09')"
        'update cmmf

        If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
            Return False
        End If

        'get Dataset from latestupdate on ward
        Dim ds As New DataSet
        sqlstr = "select odtp.ordertype::character varying, ohd.customercode, ohd.officerid::character varying, ohd.soldto, odtl.customerorderno::character varying, ohd.sebasiasalesorder,odtl.solineno,   odtl.vendorcode,odtp.curinq::character varying, odtl.cmmf,  ohd.orderstatus::character varying, odtp.latestupdate, odtl.receptiondate,odtp.fob,odtp.unittp,odtp.inquiryeta, odtp.inquiryetd,odtp.inquiryqty, odtp.currentinquiryeta,odtp.currentinquiryetd,odtp.currentinquiryqty, ship.deliveredqty, ship.shipdate, shipmentdtl.shipfrom::character varying,odtl.osqty, odtl.sebasiapono,odtl.polineno, odtl.comments::character varying, odtl.commentid,po.finalcustomerorder::character varying,odtp.updatesince7::character varying,ship.shipdateeta,shipmentdtl.boatid,shipmentdtl.ctrno,shipmentdtl.packinglist" & _
             " From ohd" & _
             " LEFT JOIN odtl ON ohd.sebasiasalesorder = odtl.sebasiasalesorder" & _
             " LEFT JOIN odtp ON odtp.orderdtlid = odtl.orderdtlid" & _
             " LEFT JOIN ship ON ship.orderdtltypeid = odtp.orderdtltypeid" & _
             " left join shipmentdtl on shipmentdtl.shipmentid = ship.shipmentid" & _
             " LEFT JOIN po ON po.sebasiapono = odtl.sebasiapono" & _
             " WHERE ordertype = 'Shipment' and latestupdate >= " & myLatestUpdate & _
             " ORDER BY odtp.orderdtltypeid;"

        If Not DbAdapter1.TbgetDataSet(sqlstr, ds, errmsg) Then
            Return False
        End If

        If Not insertshipment(ds, errmsg, myform) Then
            Return False
        End If

        'update LatestUpdatedate
        sqlstr = "update paramhd set dvalue = (select latestupdate from lsodtype order by latestupdate desc limit 1) where paramname = 'LatestUpdate';"
        If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
            Return False
        End If

        myret = True
        Return myret
    End Function

    Private Function insertshipment(ByVal ds As DataSet, ByRef errmsg As String, ByVal myform As ImportWORFG) As Boolean
        Dim myret As Boolean = False
        'get initial data
        Dim mydata As New DataSet
        If Not getinitialdata(mydata, errmsg) Then
            Return myret
        End If

        myform.ProgressReport(2, String.Format("Build Data row..........."))
        For Each dr As DataRow In ds.Tables(0).Rows
            'build stringbuilder
            If Not buildshipmentsb(dr, mydata, errmsg) Then
                Return myret
            End If
        Next

        'copy
        If Not copyshipment(errmsg, myform) Then
            Return myret
        End If

        myret = True
        Return myret
    End Function

    Public Function getinitialdata(ByRef mydata As DataSet, ByRef errmsg As String) As Boolean
        Dim myret As Boolean = False
        Dim process As String = String.Empty
        Dim Sqlstr As String = " select sebasiasalesorder from lsohd;" &
                               " select sebasiasalesorder,solineno,lsodtlid from lsodtl;" &
                               " select lsodtlid from lsodtl order by lsodtlid desc limit 1;" &
                               " select sebasiapono from lspohd;" &
                               " select sebasiapono,polineno,lspodtlid from lspodtl;" &
                               " select lspodtlid from lspodtl order by lspodtlid desc limit 1;" &
                               " select lsodtypeid from lsodtype order by lsodtypeid desc limit 1;" &
                               " select lsshipid from lsship order by lsshipid desc limit 1;" &
                               " select setval('lsodtl_lsodtlid_seq'::regclass,(select lsodtlid from lsodtl order by lsodtlid desc limit 1) + 1,false); " &
                               " select setval('lspodtl_lspodtlid_seq'::regclass,(select lspodtlid from lspodtl order by lspodtlid desc limit 1) + 1,false);" &
                               " select setval('lsodtype_lsodtypeid_seq'::regclass,(select lsodtypeid from lsodtype order by lsodtypeid desc limit 1) + 1,false);" &
                               " select setval('lsship_lsshipid_seq'::regclass,(select lsshipid from lsship order by lsshipid desc limit 1) + 1,false);" &
                               " select setval('lsshipdtl_lsshipdtlid_seq'::regclass,(select lsshipdtlid from lsshipdtl order by lsshipdtlid desc limit 1) + 1,false);"

        If DbAdapter1.TbgetDataSet(Sqlstr, mydata, errmsg) Then
            mydata.Tables(0).TableName = "lsohd"
            mydata.Tables(1).TableName = "lsodtl"
            mydata.Tables(3).TableName = "lspohd"
            mydata.Tables(4).TableName = "lspodtl"

            lsodtlidseq = mydata.Tables(2).Rows(0).Item(0)
            lspodtlidseq = mydata.Tables(5).Rows(0).Item(0)
            lsodtypeidseq = mydata.Tables(6).Rows(0).Item(0)
            lsshipidseq = mydata.Tables(7).Rows(0).Item(0)
            Try

                process = "ohd index"
                Dim idx0(0) As DataColumn               'ohd
                idx0(0) = mydata.Tables(0).Columns(0)       'sebasiasalesorder
                mydata.Tables(0).PrimaryKey = idx0

                process = "odtl index"
                Dim idx1(1) As DataColumn               'odtl
                idx1(0) = mydata.Tables(1).Columns(0)       'sebasiasalesorder    
                idx1(1) = mydata.Tables(1).Columns(1)       'solineno
                mydata.Tables(1).PrimaryKey = idx1

                process = "lspohd index"
                Dim idx3(0) As DataColumn               'lspohd
                idx3(0) = mydata.Tables(3).Columns(0)
                mydata.Tables(3).PrimaryKey = idx3

                process = "lspodtl index"
                Dim idx4(1) As DataColumn               'brand
                idx4(0) = mydata.Tables(4).Columns(0)       'brandid
                idx4(1) = mydata.Tables(4).Columns(1)       'brandid
                mydata.Tables(4).PrimaryKey = idx4

                myret = True
            Catch ex As Exception
                errmsg = ex.Message
            End Try

        Else
            Return False
        End If

        Return myret
    End Function

    Private Function buildshipmentsb(ByVal mydr As DataRow, ByRef mydata As DataSet, ByVal errmsg As String) As Boolean
        Dim myret As Boolean = False
        Dim result As DataRow
        Dim myprogress As String = String.Empty
        Try

        
            'progress lsohd
            myprogress = "lsohd"
            Dim pkey0(0) As Object
            pkey0(0) = mydr.Item("sebasiasalesorder")
            result = mydata.Tables(0).Rows.Find(pkey0)
            If IsNothing(result) Then
                Dim dr As DataRow = mydata.Tables(0).NewRow
                dr.Item(0) = mydr.Item("sebasiasalesorder")
                mydata.Tables(0).Rows.Add(dr)
                'sebasiasalesorder,soldto,officerid
                lsohdsb.Append(mydr.Item("sebasiasalesorder") & vbTab &
                             mydr.Item("soldto") & vbTab &
                             validstr(mydr.Item("officerid")) & vbCrLf)
            End If

        'progress lsodtl
            myprogress = "lsodtl"
            Dim pkey1(1) As Object
            pkey1(0) = mydr.Item("sebasiasalesorder")
            pkey1(1) = mydr.Item("solineno")
            result = mydata.Tables(1).Rows.Find(pkey1)
            If IsNothing(result) Then
                lsodtlidseq += 1
                lsodtlid = lsodtlidseq
                Dim dr As DataRow = mydata.Tables(1).NewRow
                dr.Item(0) = mydr.Item("sebasiasalesorder")
                dr.Item(1) = mydr.Item("solineno")
                dr.Item(2) = lsodtlid

                mydata.Tables(1).Rows.Add(dr)
                'sebasiasalesorder,solineno,customerorderno
                lsodtlsb.Append(mydr.Item("sebasiasalesorder") & vbTab &
                         mydr.Item("solineno") & vbTab &
                         validstr(mydr.Item("customerorderno")) & vbCrLf)
            Else
                lsodtlid = result.Item("lsodtlid")
            End If

            myprogress = "lspohd"
            'progress lspohd
            Dim pkey3(0) As Object
            pkey3(0) = mydr.Item("sebasiapono")
            result = mydata.Tables(3).Rows.Find(pkey3)
            If IsNothing(result) Then
                Dim dr As DataRow = mydata.Tables(3).NewRow
                dr.Item(0) = mydr.Item("sebasiapono")
                mydata.Tables(3).Rows.Add(dr)
                'sebasiapono,customercode,receptiondate,finalcustomerorder
                lspohdsb.Append(mydr.Item("sebasiapono") & vbTab &
                             mydr.Item("customercode") & vbTab &
                             DateFormatyyyyMMdd(mydr.Item("receptiondate")) & vbTab &
                             validstr(mydr.Item("finalcustomerorder")) & vbCrLf)
            End If

            'progress lspodtl
            myprogress = "lspodtl"
            Dim pkey4(1) As Object
            pkey4(0) = mydr.Item("sebasiapono")
            pkey4(1) = mydr.Item("polineno")
            result = mydata.Tables(4).Rows.Find(pkey4)
            If IsNothing(result) Then
                lspodtlidseq += 1
                lspodtlid = lspodtlidseq
                Dim dr As DataRow = mydata.Tables(4).NewRow
                dr.Item(0) = mydr.Item("sebasiapono")
                dr.Item(1) = mydr.Item("polineno")
                dr.Item(2) = lspodtlid
                mydata.Tables(4).Rows.Add(dr)
                'sebasiapono,polineno,cmmf,vendorcode,osqty,orderstatus,comments,commentid,fob,unittp,inquiryeta,inquiryetd,inquiryqty,
                'currentinquiryeta, currentinquiryetd, currentinquiryqty, curinq, shipfrom
                lspodtlsb.Append(mydr.Item("sebasiapono") & vbTab &
                             mydr.Item("polineno") & vbTab &
                             mydr.Item("cmmf") & vbTab &
                             mydr.Item("vendorcode") & vbTab &
                             validint(mydr.Item("osqty")) & vbTab &
                             validstr(mydr.Item("orderstatus")) & vbTab &
                             validstr(mydr.Item("comments")) & vbTab &
                             validint(mydr.Item("commentid")) & vbTab &
                             validReal(mydr.Item("fob")) & vbTab &
                             validReal(mydr.Item("unittp")) & vbTab &
                             DateFormatyyyyMMdd(mydr.Item("inquiryeta")) & vbTab &
                             DateFormatyyyyMMdd(mydr.Item("inquiryetd")) & vbTab &
                             validint(mydr.Item("inquiryqty")) & vbTab &
                             DateFormatyyyyMMdd(mydr.Item("currentinquiryeta")) & vbTab &
                             DateFormatyyyyMMdd(mydr.Item("currentinquiryetd")) & vbTab &
                             validReal(mydr.Item("currentinquiryqty")) & vbTab &
                             validstr(mydr.Item("curinq")) & vbTab &
                             validstr(mydr.Item("shipfrom")) & vbCrLf)
            Else
                lspodtlid = result.Item("lspodtlid")
            End If
            myprogress = "lsodtype"
            'progress lsodtype
            lsodtypeidseq += 1
            'lsodtlid,updatesince7,ordertype,latestupdate,lspodtlid
            lsodtypesb.Append(lsodtlid & vbTab &
                              validstr(mydr.Item("updatesince7")) & vbTab &
                              validstr(mydr.Item("ordertype")) & vbTab &
                              DateFormatyyyyMMdd(mydr.Item("latestupdate")) & vbTab &
                               lspodtlid & vbCrLf)

            'create lsship
            'deliveredqty,shipdate,lsodtypeid,shipdateeta,cslstatus
            myprogress = "lsship"
            lsshipidseq += 1
            lsshipsb.Append(validint(mydr.Item("deliveredqty")) & vbTab &
                            DateFormatyyyyMMdd(mydr.Item("shipdate")) & vbTab &
                            lsodtypeidseq & vbTab &
                            DateFormatyyyyMMdd(mydr.Item("shipdateeta")) & vbTab &
                            IIf(Math.Abs(DateDiff(DateInterval.Day, CDate(mydr.Item("shipdate")), CDate(mydr.Item("currentinquiryetd")))) > 7, False, True) & vbCrLf)

            'create lsshipdtl
            'lsshipid,ctrno,boatid,packinglist
            myprogress = "lsshipdtl"
            lsshipdtlsb.Append(lsshipidseq & vbTab &
                               validstr(mydr.Item("ctrno")) & vbTab &
                               validstr(mydr.Item("boatid")) & vbTab &
                               validstr(mydr.Item("packinglist")) & vbCrLf)
            myret = True
        Catch ex As Exception
            errmsg = ex.Message
        End Try
        Return myret
    End Function

    Private Function copyshipment(ByVal errmsg As String, ByVal myform As ImportWORFG) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String = String.Empty
        Try
            If lsohdsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lsohd"))
                sqlstr = "copy lsohd(sebasiasalesorder,soldto,officerid) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lsohdsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            If lsodtlsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lsodtl"))
                sqlstr = "copy lsodtl(sebasiasalesorder,solineno,customerorderno) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lsodtlsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            If lspohdsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lspohd"))
                sqlstr = "copy lspohd(sebasiapono,customercode,receptiondate,finalcustomerorder) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lspohdsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            If lspodtlsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lspodtlsb"))
                sqlstr = "copy lspodtl(sebasiapono,polineno,cmmf,vendorcode,osqty,orderstatus,comments,commentid,fob,unittp,inquiryeta,inquiryetd,inquiryqty,currentinquiryeta, currentinquiryetd, currentinquiryqty, curinq, shipfrom) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lspodtlsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            If lsodtypesb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lsodtype"))
                sqlstr = "copy lsodtype(lsodtlid,updatesince7,ordertype,latestupdate,lspodtlid) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lsodtypesb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If

            If lsshipsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lsship"))
                sqlstr = "copy lsship(deliveredqty,shipdate,lsodtypeid,shipdateeta,cslstatus) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lsshipsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            If lsshipdtlsb.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy lsodtl"))
                sqlstr = "copy lsshipdtl(lsshipid,ctrno,boatid,packinglist) from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, lsshipdtlsb.ToString, myret)
                If Not myret Then
                    Return False
                End If
            End If
            myret = True
        Catch ex As Exception
            errmsg = ex.Message
        End Try



        Return myret
    End Function

End Class