Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass

Public Class ImportIPLT

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
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        If Not myThread.IsAlive Then

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
        ProgressReport(2, TextBox2.Text & "Read Folder..")

        ReadFileStatus = ImportTextFile(mySelectedPath, errMsg)
        If ReadFileStatus Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(2, TextBox2.Text & "Done.")
            ProgressReport(5, "Set to continuous mode again")
        Else
            errSB.Append(errMsg & vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If

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
        If Me.TextBox1.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    TextBox2.Text = message
                Case 3
                    TextBox3.Text = message
                Case 4
                    TextBox1.Text = message
                Case 5
                    ProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ProgressBar1.Minimum = 1
                    ProgressBar1.Value = myvalue(0)
                    ProgressBar1.Maximum = myvalue(1)
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

        ProgressReport(2, "Scanning Text File...")
        ProgressReport(3, "Open Text File...")
        Dim i As Long
        Try
            'Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)

            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.XLS")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            For Each fi As IO.FileInfo In arrFI
                ProgressReport(3, String.Format("Read Text File...{0}", fi.FullName))
                Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        .HasFieldsEnclosedInQuotes = False
                        Dim count As Long = 0

                        Do Until .EndOfData
                            'If count > 0 Then
                            myrecord = .ReadFields
                            If count > 1 Then
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
                errMsg = "Nothing to process."
                Return myret
            End If
            'get dataset
            Dim DS As New DataSet

            'get initial keys from Database fro related table
            ProgressReport(3, String.Format("Delete rows ..........."))
            DbAdapter1.deleteIPLT()

            If Not FillDataset(DS, errMsg) Then
                Return False
            End If

            'Create object for handleing row creation
            Dim IPLT As New IPLT(DS)

            ProgressReport(3, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")
            For i = 0 To myList.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & myList.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                If Not IPLT.buildSB(errMsg, myList(i)) Then
                    Return False
                End If

                'End If
            Next
            ProgressReport(6, "Set To Marque")
            ProgressReport(3, String.Format("Copy To Db"))
            If Not IPLT.copyToDb(errMsg, Me) Then
                Return False
            End If
            myret = True

        Catch ex As Exception
            errMsg = String.Format("Row : {0} ", i) & ex.Message
        End Try
        'copy

        myret = True
        'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(SW.Elapsed.Minutes, "00"), Format(SW.Elapsed.Seconds, "00"), SW.Elapsed.Milliseconds.ToString))
        Return myret
    End Function

    Private Function validchar(ByVal strvalue As String) As Object
        If strvalue = "" Then
            Return ""
        Else
            'Return "'" & Trim(strvalue.Replace("'", "''").Replace("""", "")) & "'"
            Return Trim(strvalue.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Private Function validint(ByVal intvalue As String) As Object
        If intvalue = "" Then
            Return "NULL"
        Else
            Return CInt(intvalue.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Private Function validdec(ByVal decvalue As String) As Object
        If decvalue = "" Then
            Return "NULL"
        Else
            Return CDec(decvalue.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Private Function validdate(ByVal datevalue As String) As Object
        Dim mydata() As String
        If datevalue.Contains(".") Then
            mydata = datevalue.Split(".")
        Else
            mydata = datevalue.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return "NULL"
    End Function

    Private Function validboolean(ByVal booleanvalue As String) As String
        If booleanvalue = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function

    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As String) As Boolean
        Dim myret As Boolean = False

        Dim Sqlstr As String = " select delivery from cxdeliveryhd;" &
                               " select cxdeliverydtlid,delivery,deliveryitem  from cxdeliverydtl;" &
                               " select billingdoc from cxbillingdochd;" &
                               " select cxbillingdocdtlid,billingdoc,item from cxbillingdocdtl;" &
                               " select count(1) from cxdeliverydtl;" &
                               " select count(1) from cxbillingdocdtl;" &
                               " select carrier from forwarder where not carrier isnull;" &
                               " select sebpono from cxipltpohd;" &
                               " select cxipltpodtlid,sebpono,solineno from cxipltpodtl;" &
                               " select salesdoc from cxipltsalesdochd;" &
                               " select cxipltsalesdocdtlid,salesdoc,solineno from cxipltsalesdocdtl;" &
                               " select count(1) from cxipltpodtl;" &
                               " select count(1) from cxipltsalesdocdtl;" &
                               " select po from povendor;" &
                               " select cmmf from cmmf;" &
                               " select vendorcode from vendor;" &
                               " select officerid from officer;" &
                               " select shpt from shpt;" &
                               " select vendorcode from vendorssm;"


        If DbAdapter1.TbgetDataSet(Sqlstr, DS, errmessage) Then
            DS.Tables(0).TableName = "cxdeliveryhd"
            DS.Tables(1).TableName = "cxdeliverydtl"
            DS.Tables(2).TableName = "cxbillingdochd"
            DS.Tables(3).TableName = "cxbillingdocdtl"
            DS.Tables(4).TableName = "seqdeliverydtl"
            DS.Tables(5).TableName = "seqbilingdocdtl"
            DS.Tables(6).TableName = "forwarder"
            DS.Tables(7).TableName = "cxipltpohd"
            DS.Tables(8).TableName = "cxipltpodtl"
            DS.Tables(9).TableName = "cxsalesdochd"
            DS.Tables(10).TableName = "cxsalesdocdtl"
            DS.Tables(11).TableName = "seqipltpodtl"
            DS.Tables(12).TableName = "seqipltsalesdocdtl"
            DS.Tables(13).TableName = "povendor"
            DS.Tables(14).TableName = "cmmf"
            DS.Tables(15).TableName = "vendor"
            DS.Tables(16).TableName = "officer"
            DS.Tables(17).TableName = "shpt"
            DS.Tables(18).TableName = "vendorssm"

            Dim idx0(0) As DataColumn               'cxdeliveryhd
            idx0(0) = DS.Tables(0).Columns(0)       'delivery
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(1) As DataColumn               'cxdeliverydtl
            idx1(0) = DS.Tables(1).Columns(1)       'delivery    
            idx1(1) = DS.Tables(1).Columns(2)       'deliveryitem
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn               'cxbilingdochd
            idx2(0) = DS.Tables(2).Columns(0)       'billingdoc
            DS.Tables(2).PrimaryKey = idx2

            Dim idx3(1) As DataColumn               'cxbillingdocdtl
            idx3(0) = DS.Tables(3).Columns(1)       'billingdoc
            idx3(1) = DS.Tables(3).Columns(2)       'item
            DS.Tables(3).PrimaryKey = idx3

            Dim idx6(0) As DataColumn               'forwarder
            idx6(0) = DS.Tables(6).Columns(0)       'carrier
            DS.Tables(6).PrimaryKey = idx6


            Dim idx7(0) As DataColumn               'pohd
            idx7(0) = DS.Tables(7).Columns(0)
            DS.Tables(7).PrimaryKey = idx7


            Dim idx8(1) As DataColumn               'podtl
            idx8(0) = DS.Tables(8).Columns(0)
            idx8(1) = DS.Tables(8).Columns(1)
            DS.Tables(8).PrimaryKey = idx8


            Dim idx9(0) As DataColumn               'salesdoc
            idx9(0) = DS.Tables(9).Columns(0)
            DS.Tables(9).PrimaryKey = idx9


            Dim idx10(1) As DataColumn               'salesdocdtl
            idx10(0) = DS.Tables(10).Columns(0)
            idx10(1) = DS.Tables(10).Columns(1)
            DS.Tables(10).PrimaryKey = idx10

            Dim idx13(0) As DataColumn               'povendor
            idx13(0) = DS.Tables(13).Columns(0)
            DS.Tables(13).PrimaryKey = idx13

            Dim idx14(0) As DataColumn               'cmmf
            idx14(0) = DS.Tables(14).Columns(0)
            DS.Tables(14).PrimaryKey = idx14

            Dim idx15(0) As DataColumn               'vendor
            idx15(0) = DS.Tables(15).Columns(0)
            DS.Tables(15).PrimaryKey = idx15

            Dim idx16(0) As DataColumn               'officer
            idx16(0) = DS.Tables(16).Columns(0)
            DS.Tables(16).PrimaryKey = idx16

            Dim idx17(0) As DataColumn               'shpt
            idx17(0) = DS.Tables(17).Columns(0)
            DS.Tables(17).PrimaryKey = idx17

            Dim idx18(0) As DataColumn               'shpt
            idx18(0) = DS.Tables(18).Columns(0)
            DS.Tables(18).PrimaryKey = idx18
        Else
            Return False
        End If
        myret = True
        Return myret
    End Function
End Class

Public Class IPLT
    Public Property ds As DataSet
    Public Property cxDelivery As New StringBuilder
    Public Property cxDeliverydtl As New StringBuilder
    Public Property cxBillingdoc As New StringBuilder
    Public Property cxBillingdocdtl As New StringBuilder
    Public Property cxForwarder As New StringBuilder
    Public Property cxPOHd As New StringBuilder
    Public Property cxPodtl As New StringBuilder
    Public Property cxSalesHD As New StringBuilder
    Public Property cxSalesDtl As New StringBuilder
    Public Property cxpovendor As New StringBuilder
    Public Property cxiplt As New StringBuilder
    Public Property cxvendor As New StringBuilder
    Public Property cxofficer As New StringBuilder
    Public Property cxcmmf As New StringBuilder
    Public Property cxshpt As New StringBuilder
    Public Property vendorssm As New StringBuilder

    Dim seqdeliverydtl As Long
    Dim seqbillingdocdtl As Long
    Dim seqforwarder As Long
    Dim seqipltpodtl As Long
    Dim seqipltsalesdtl As Long

    Dim deliverydtlid As Long
    Dim billingdocdtlid As Long
    Dim forwarderid As Long
    Dim podtlid As Long
    Dim salesdtlid As Long

    Public Sub New(ByVal ds As DataSet)
        Me.ds = ds
        seqdeliverydtl = ds.Tables(4).Rows(0).Item(0)
        seqbillingdocdtl = ds.Tables(5).Rows(0).Item(0)
        seqipltpodtl = ds.Tables(11).Rows(0).Item(0)
        seqipltsalesdtl = ds.Tables(12).Rows(0).Item(0)
    End Sub

    Public Function buildSB(ByRef message As String, ByVal mydata As myData) As Boolean
        Dim myret As Boolean = False

        Dim myprogress As String = String.Empty
        Dim data = mydata.data
        Dim comments As String = String.Empty
        Dim result As DataRow
        Try

            myprogress = "Forwarder"

            'Forwarder
            Dim pkey6(0) As Object
            pkey6(0) = data(8)
            result = ds.Tables(6).Rows.Find(pkey6)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(6).NewRow
                dr.Item(0) = data(8)
                ds.Tables(6).Rows.Add(dr)
                cxForwarder.Append(data(8) & vbTab &
                                    data(9) & vbCrLf)
            End If

            myprogress = "Delivery HD"
            'DeliveyHD
            If data(10) <> "" Then
                Dim pkey0(0) As Object
                pkey0(0) = data(10)
                result = ds.Tables(0).Rows.Find(pkey0)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(0).NewRow
                    dr.Item(0) = data(10)
                    ds.Tables(0).Rows.Add(dr)

                    cxDelivery.Append(data(10) & vbTab &
                                        dateformatdotyyyymmdd(data(12)) & vbTab &
                                        validstr(data(13)) & vbTab &
                                        validstr(data(14)) & vbTab &
                                        validstr(data(15)) & vbTab &
                                        validreal(data(16)) & vbTab &
                                        validstr(data(17)) & vbTab &
                                        validstr(data(19)) & vbCrLf)
                End If
            End If
            myprogress = "Delivery Detail"
            'DeliveyHD
            If data(10) <> "" Then
                Dim pkey1(1) As Object
                pkey1(0) = data(10)
                pkey1(1) = data(11)
                result = ds.Tables(1).Rows.Find(pkey1)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(1).NewRow
                    seqdeliverydtl += 1
                    deliverydtlid = seqdeliverydtl
                    dr.Item(0) = deliverydtlid
                    dr.Item(1) = data(10)
                    dr.Item(2) = data(11)
                    ds.Tables(1).Rows.Add(dr)

                    cxDeliverydtl.Append(data(10) & vbTab &
                                        data(11) & vbTab &
                                        validreal(data(18)) & vbCrLf)
                Else
                    deliverydtlid = result.Item(0)
                End If
            End If


            myprogress = "BillingDoc HD"
            If data(1) <> "" Then
                Dim pkey2(0) As Object
                pkey2(0) = data(1)
                result = ds.Tables(2).Rows.Find(pkey2)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(2).NewRow
                    dr.Item(0) = data(1)
                    ds.Tables(2).Rows.Add(dr)

                    cxBillingdoc.Append(data(1) & vbTab &
                                        dateformatdotyyyymmdd(data(3)) & vbTab &
                                        validstr(data(7)) & vbCrLf)
                End If
            End If
            myprogress = "Delivery Detail"
            'DeliveyHD
            If data(1) <> "" Then
                Dim pkey3(1) As Object
                pkey3(0) = data(1)
                pkey3(1) = data(2)
                result = ds.Tables(1).Rows.Find(pkey3)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(3).NewRow
                    seqbillingdocdtl += 1
                    billingdocdtlid = seqbillingdocdtl
                    dr.Item(0) = billingdocdtlid
                    dr.Item(1) = data(1)
                    dr.Item(2) = data(2)
                    ds.Tables(3).Rows.Add(dr)

                    cxBillingdocdtl.Append(data(1) & vbTab &
                                           data(2) & vbTab &
                                        validint(data(4)) & vbTab &
                                        validstr(data(5)) & vbTab &
                                        validreal(data(6)) & vbCrLf)
                Else
                    deliverydtlid = result.Item(0)
                End If
            End If

            'Check VendorCode
            myprogress = "Check Vendor Code"
            Dim pkey15(0) As Object
            pkey15(0) = data(23)
            result = ds.Tables(15).Rows.Find(pkey15)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(15).NewRow
                dr.Item(0) = data(23)
                ds.Tables(7).Rows.Add(dr)

                cxvendor.Append(data(23) & vbTab &
                                    data(24) & vbCrLf)
            End If

            'Check Officer
            myprogress = "Check Officer"
            Dim pkey16(0) As Object
            pkey16(0) = data(39)
            result = ds.Tables(16).Rows.Find(pkey16)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(16).NewRow
                dr.Item(0) = data(39)
                ds.Tables(16).Rows.Add(dr)
                cxofficer.Append(data(39) & vbTab &
                                data(40) & vbCrLf)
            End If
            myprogress = "check ssm"
            pkey16(0) = data(57)
            result = ds.Tables(16).Rows.Find(pkey16)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(16).NewRow
                dr.Item(0) = data(57)
                ds.Tables(16).Rows.Add(dr)
                cxofficer.Append(data(57) & vbTab &
                                data(58) & vbCrLf)
            End If

            myprogress = "PO HD"

            Dim pkey7(0) As Object
            pkey7(0) = data(59)
            result = ds.Tables(7).Rows.Find(pkey7)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(7).NewRow
                dr.Item(0) = data(59)
                ds.Tables(7).Rows.Add(dr)
                cxPOHd.Append(data(59) & vbTab &
                                    data(23) & vbTab &
                                    data(25) & vbTab &                                    
                                    dateformatdotyyyymmdd(data(65)) & vbCrLf)
            End If

            myprogress = "vendorssm"

            Dim pkey18(0) As Object
            pkey18(0) = data(23)
            result = ds.Tables(18).Rows.Find(pkey18)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(18).NewRow
                dr.Item(0) = data(23)
                ds.Tables(18).Rows.Add(dr)
                vendorssm.Append(data(23) & vbTab &
                                    validlong(data(57)) & vbCrLf)
            End If

            'check cmmf (matl no,material,materialdesc,rri,family)
            If data(21) <> "" Then
                myprogress = "Check CMMF"
                Dim pkey14(0) As Object
                pkey14(0) = data(21)
                result = ds.Tables(14).Rows.Find(pkey14)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(14).NewRow
                    dr.Item(0) = data(21)
                    ds.Tables(14).Rows.Add(dr)
                    cxcmmf.Append(data(20) & vbTab &
                                    data(21) & vbTab &
                                    data(22) & vbTab &
                                    data(38) & vbTab &
                                    data(56) & vbCrLf)
                End If
            End If
            If cxcmmf.ToString <> "" Then
                MessageBox.Show("hello")
            End If

            myprogress = "povendor"


            Dim pkey13(0) As Object
            pkey13(0) = data(59)
            result = ds.Tables(13).Rows.Find(pkey13)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(13).NewRow
                dr.Item(0) = data(59)
                ds.Tables(13).Rows.Add(dr)
                cxpovendor.Append(data(59) & vbTab &
                                    data(23) & vbCrLf)
            End If


            myprogress = "POdtl"

            Dim pkey8(1) As Object
            pkey8(0) = data(59)
            pkey8(1) = data(60)
            result = ds.Tables(8).Rows.Find(pkey8)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(8).NewRow
                seqipltpodtl += 1
                podtlid = seqipltpodtl
                dr.Item(0) = podtlid
                dr.Item(1) = data(59)
                dr.Item(2) = data(60)
                ds.Tables(8).Rows.Add(dr)

                cxPodtl.Append(data(59) & vbTab &
                               data(60) & vbTab &
                                   data(21) & vbTab &
                                    validstr(data(47)) & vbTab &
                                    validstr(data(48)) & vbTab &
                                    validstr(data(20)) & vbTab &
                                    validstr(data(38)) & vbTab &
                                    validstr(data(56)) & vbCrLf)
            Else
                podtlid = result.Item(0)
            End If

            
            'cxipltsalesHD

            myprogress = "Sales HD"

            Dim pkey9(0) As Object
            pkey9(0) = data(27)
            result = ds.Tables(9).Rows.Find(pkey9)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(9).NewRow
                dr.Item(0) = data(27)
                ds.Tables(9).Rows.Add(dr)
                cxSalesHD.Append(data(27) & vbTab &
                                    data(33) & vbTab &
                                    data(37) & vbTab &
                                    data(34) & vbTab &
                                    validstr(data(39)) & vbTab &
                                    data(55) & vbTab &
                                    data(62) & vbTab &
                                    data(64) & vbTab &
                                     data(53) & vbTab &
                                     validstr(data(52)) & vbCrLf)
            End If

            'Check shpt
            If Not data(50) <> "" Then
                myprogress = "Check shpt"
                Dim pkey17(0) As Object
                pkey17(0) = data(50)
                result = ds.Tables(17).Rows.Find(pkey17)
                If IsNothing(result) Then
                    Dim dr As DataRow = ds.Tables(17).NewRow
                    dr.Item(0) = data(50)
                    ds.Tables(17).Rows.Add(dr)
                    cxshpt.Append(data(50) & vbTab &
                                    data(51) & vbCrLf)
                End If
            End If


            myprogress = "SalesDtl"

            Dim pkey10(1) As Object
            pkey10(0) = data(27)
            pkey10(1) = data(28)
            result = ds.Tables(10).Rows.Find(pkey10)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(10).NewRow
                seqipltsalesdtl += 1
                salesdtlid = seqipltsalesdtl
                dr.Item(0) = podtlid
                dr.Item(1) = data(27)
                dr.Item(2) = data(28)
                ds.Tables(10).Rows.Add(dr)

                cxSalesDtl.Append(data(27) & vbTab &
                               data(28) & vbTab &
                                   validint(data(29)) & vbTab &
                                    data(30) & vbTab &
                                    validreal(data(31)) & vbTab &
                                    validreal(data(32)) & vbTab &
                                    validreal(data(36)) & vbTab &
                                    data(49) & vbTab &
                                    validstr(data(50)) & vbTab &
                                    validint(data(61)) & vbCrLf)
            Else
                salesdtlid = result.Item(0)
            End If


            myprogress = "Iplt Miro"

 
            cxiplt.Append(podtlid & vbTab &
                          salesdtlid & vbTab &
                          validzerotonull(deliverydtlid) & vbTab &
                          validzerotonull(billingdocdtlid) & vbTab &
                          validlong(data(41)) & vbTab &
                          validint(data(42)) & vbTab &
                          dateformatdotyyyymmdd(data(43)) & vbTab &
                          validint(data(44)) & vbTab &
                          validreal(data(45)) & vbTab &
                          validreal(data(46)) & vbTab &
                                    validstr(data(8)) & vbTab &
                                    validstr(data(35)) & vbTab &
                                    validstr(data(63)) & vbCrLf)

            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2},Row Num {3}", myprogress, ex.Message, mydata.filename, mydata.rownumber)
        End Try
        Return myret
    End Function

    Public Function copyToDb(ByRef errMsg As String, ByVal myform As ImportIPLT) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String
        Try
            If cxForwarder.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Forwarder"))
                sqlstr = "copy forwarder(carrier,forwardername)  from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxForwarder.ToString, myret)
                If Not myret Then
                    Return myret
                End If

            End If
            If cxDelivery.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxDelivery"))
                sqlstr = "copy cxdeliveryhd(delivery,deliverydate,trpt,meansoftrid,container,volume,vun,un) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxDelivery.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxDeliverydtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxDeliverydtl"))
                sqlstr = "copy cxdeliverydtl(delivery,deliveryitem,net) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxDeliverydtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxBillingdoc.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxBillingDoc"))
                sqlstr = "copy cxbillingdochd(billingdoc,createdon,currency) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxBillingdoc.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxBillingdocdtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy CxBillingDocDtl"))
                sqlstr = "copy cxbillingdocdtl(billingdoc,item,billqty,su,netvalue) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxBillingdocdtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxPOHd.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxPOHd"))
                sqlstr = "copy cxipltpohd(sebpono,vendorcode,shiptopartycode,docdate) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxPOHd.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If vendorssm.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy vendorssm"))
                sqlstr = "copy vendorssm(vendorcode,ssm) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, vendorssm.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxPodtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxPOdtl"))
                sqlstr = "copy cxipltpodtl(sebpono,solineno,cmmf,activity,technofamilies,itemid,rir,comfam) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxPodtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxSalesHD.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxSalesHD"))
                sqlstr = "copy cxipltsalesdochd(salesdoc,curr2,crcy,customerpono,officer,curr3,oun,sc,customerid,incoterms2) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxSalesHD.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxSalesDtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxSalesDtl"))
                sqlstr = "copy cxipltsalesdocdtl(salesdoc,solineno,orderqty,su2,netprice,netvalue2,netvalue,prdt,shpt,quantity) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxSalesDtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxpovendor.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxpovendor"))
                sqlstr = "copy povendor(po,vendorcode) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxpovendor.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxiplt.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxiplt"))
                sqlstr = "copy cxiplt(cxipltpodtlid,cxipltsalesdocdtlid,cxdeliverydtlid,cxbillingdtlid,mironumber,myyear,invoicepostingdate,qtyshipped,supplieramount,supplierprice,carrier,supplierinvoice,created) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxiplt.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxvendor.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxvendor"))
                sqlstr = "copy vendor(vendorcode,vendorname) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxvendor.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxofficer.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxofficer"))
                sqlstr = "copy officer(officerid,officername) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxofficer.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxcmmf.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxcmmf"))
                sqlstr = "copy cmmf(itemid,cmmf,materialdesc,rir,comfam) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxcmmf.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxshpt.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy cxshpt"))
                sqlstr = "copy shpt(shpt,shippointdesc) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxshpt.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            'Public Property cxshpt As New StringBuilder

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
            Return CInt((Replace(p1, ",", "")))
        End If
    End Function
    Private Function validlong(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CLng((Replace(p1, ",", "")))
        End If
    End Function

    Private Function validstr(ByVal data As Object) As Object
        If data = "" Then
            Return "Null"
        End If
        Return data
    End Function

    Private Function dateformatdotyyyymmdd(ByVal data As Object) As Object
        Dim myret As String = "Null"
        If data = "" Then
            Return myret
        End If
        Dim mydate = data.ToString.Split(".")
        myret = "'" & mydate(2) & "-" & mydate(1) & "-" & mydate(0) & "'"
        Return myret
    End Function

    Private Function validreal(ByVal data As Object) As Object
        Dim myret As String = "Null"
        If data = "" Then
            Return myret
        End If
        Return CDec(Replace(data, ",", ""))
    End Function

    Private Function validzerotonull(ByVal podtlid As Long) As String
        Dim myret = podtlid
        If podtlid = 0 Then
            Return "Null"
        End If
        Return myret
    End Function

End Class
