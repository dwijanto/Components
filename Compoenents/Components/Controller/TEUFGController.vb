Imports System.Threading
Imports Microsoft.Office.Interop

Public Class TEUFGController
    Private myform As FormTEUFG
    Dim myadapter As DbAdapter = DbAdapter.getInstance
    Dim sqlstr As String = String.Empty
    Dim _ErrorMessage As String
    Dim _LastUpdate As Date

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Public ReadOnly Property LastUpdate As Date
        Get
            Return _LastUpdate
        End Get
    End Property

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property


    Public Sub New(ByVal obj As FormTEUFG)
        Me.myform = obj
    End Sub

    Public Sub run()
        If Not myThread.IsAlive Then
            myThread = New System.Threading.Thread(myThreadDelegate)
            myThread.SetApartmentState(ApartmentState.MTA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")           
        End If             
    End Sub

    Private Sub AssignForwarderId()
        Dim i As Long        
        Dim bs As New BindingSource

        'assign null to forwarderid
        ProgressReport(1, "Resetting Forwarder...")
        setNullForwarderId()
        
        'read all weeklyorder report from
        Dim Sqlstr = "select * from wkteunofix"
        Dim ds As New DataSet
        ProgressReport(1, "Read WOR Data..")
        If myadapter.TbgetDataSet(Sqlstr, ds) Then
            bs.DataSource = ds.Tables(0)
        End If
        ProgressReport(1, String.Format("Assign Forwarder ..."))
        For Each drv As DataRowView In bs.List            
            i = i + 1
            If i = 848 Then
                Debug.Print("Debug")
            End If
            'ProgressReport(7, String.Format("{0}/{1}", i, bs.Count))
            ProgressReport(1, String.Format("Assign Forwarder ...Processing {0}/{1}", i, bs.Count))
            Dim vendorcode As Long = 0
            If Not IsDBNull(drv.Row.Item("vendorcode")) Then
                vendorcode = drv.Row.Item("vendorcode")
            End If
            Dim loadingcode As String = String.Empty
            If Not IsDBNull(drv.Row.Item("loadingcode")) Then
                loadingcode = drv.Row.Item("loadingcode")
            Else
                MessageBox.Show(String.Format("Loading code is null for this cmmf : {0}", drv.Row.Item("cmmf")))
            End If
            Dim mycmmf As Long = 0
            If Not IsDBNull(drv.Row.Item("cmmf")) Then
                mycmmf = drv.Row.Item("cmmf")
            End If
            Dim zoneid As Integer = 0
            If Not IsDBNull(drv.Row.Item("zoneid")) Then
                zoneid = drv.Row.Item("zoneid")
            End If

            Dim myFWid = getForwarderId(drv.Row.Item("soldtoparty"), drv.Row.Item("customercode"), vendorcode, loadingcode, zoneid, mycmmf)

            'Update odtl with forwarderid
            If IsNothing(myFWid) Then
                myFWid = "NULL"
            End If
            Sqlstr = "Update cxsebpodtl set forwarderid = " & myFWid & _
                       " where cxsebpodtlid = " & drv.Row.Item("cxsebpodtlid")
            myadapter.ExecuteScalar(Sqlstr)

        Next
        
        'While Not rsTmp.EOF
        '    i = i + 1
        '    StatusBar1.Panels(1).Text = "Assign Forwarder , Processing line number: " & i & "/" & rsTmp.RecordCount

        '    Dim vendorcode As Long
        '    vendorcode = 0
        '    If Not IsNull(rsTmp!vendorcode) Then
        '        vendorcode = rsTmp!vendorcode
        '    End If

        '    Dim loadingcode As String
        '    loadingcode = ""
        '    If Not IsNull(rsTmp!loadingcode) Then
        '        loadingcode = rsTmp!loadingcode

        '    Else
        '        MsgBox("Loading code is null for this cmmf : " & rsTmp!cmmf)
        '    End If
        '    mycmmf = 0
        '    If Not IsNull(rsTmp!cmmf) Then
        '        mycmmf = rsTmp!cmmf
        '    End If

        '    Dim zoneid As Integer
        '    zoneid = 0
        '    If Not IsNull(rsTmp!zoneid) Then
        '        zoneid = rsTmp!zoneid
        '    End If
        '    'Use Function to get forwarderid
        '    'myFWid = getForwarderId(rsTmp!soldtoparty, rsTmp!Customercode, vendorcode, loadingcode, zoneid)
        '    myFWid = getForwarderId(rsTmp!soldtoparty, rsTmp!Customercode, vendorcode, loadingcode, zoneid, mycmmf)

        '    'Update odtl with forwarderid
        '    Sqlstr = "Update cxsebpodtl set forwarderid = " & myFWid & _
        '               " where cxsebpodtlid = " & rsTmp!cxsebpodtlid
        '    ocon.Execute(Sqlstr)

        '    rsTmp.MoveNext()
        '    DoEvents()
        'End While
    End Sub

    Sub DoWork()

        'Dim mydate2 As Date = myform.MonthSelection
        'Dim FEDate As Date = myform.ForecastStart
        If myform.CheckBox3.Checked Then
            Call AssignForwarderId()
            Call AssignForwarderIdForecast()
        Else
            'Check whether need AssignForwarderid or not
            'Sqlstr = "select forwarderid from odtl where  not(forwarderid is null) ;"
            sqlstr = "select forwarderid from cxsebpodtl where  not(forwarderid is null) limit 1;"
            Dim check As Long = getForwarderId(Sqlstr)
            If IsNothing(check) Then
                Call AssignForwarderId()
            End If
            Sqlstr = "Select forwarderid from forecastestimation where not(forwarderid is null) limit 1;"
            check = getForwarderId(Sqlstr)
            If IsNothing(check) Then
                Call AssignForwarderIdForecast()
            End If
        End If




        'ALL Data
        If Not (myform.CheckBox1.Checked) Then

            'sqlstr = String.Format("select *,'cietd' as status from vteui where currentinquiryetd >= '{0:yyyy-MM-dd}' and currentinquiryetd <= '{1:yyyy-MM-dd}'" &
            '     " Union All " &
            '     " select *,'cietd_past' as status from vteui where currentinquiryetd < '{0:yyyy-MM-dd}' " &
            '     " Union All " &
            '     " select *,'ccetd' as status from vteu where currentconfirmedetd >= '{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}'" &
            '     " Union All" &
            '     " select *,'ccetd_past' as status from vteu where currentconfirmedetd < '{0:yyyy-MM-dd}'" &
            '     " union all " &
            '     " select myyear,substring(weeketa::text,5,2)::double precision,mymonth,null,null,sapcustomercode,sapcustomername,null,null,null,forwardername,loadingcode,zone,pod,null,null,""SEB Asia SAP Vendor Code"",""SEB ASIA SAP Vendor Name"",cmmf,materialdesc,rir,period,qty,length,width,height,unit,gross,weightunit,pcperbox,volume,totalvolume,totalgrossweight,null,teu,brandname,'fe'  from forecastestimationteu where period >= '{2:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}'", myform.StartDate, myform.EndDate, myform.ForecastStart)
            sqlstr = String.Format("select *,'cietd' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)"" from vteui where currentinquiryetd >= '{0:yyyy-MM-dd}' and currentinquiryetd <= '{1:yyyy-MM-dd}'" &
                 " Union All " &
                 " select *,'cietd_past' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)"" from vteui where currentinquiryetd < '{0:yyyy-MM-dd}' " &
                 " Union All " &
                 " select *,'ccetd' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)"" from vteu where currentconfirmedetd >= '{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}'" &
                 " Union All" &
                 " select *,'ccetd_past' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)"" from vteu where currentconfirmedetd < '{0:yyyy-MM-dd}'" &
                 " union all " &
                 " select myyear,substring(weeketa::text,5,2)::double precision,mymonth,null,null,sapcustomercode,sapcustomername,null,null,null,forwardername,loadingcode,zone,pod,null,null,""SEB Asia SAP Vendor Code"",""SEB ASIA SAP Vendor Name"",cmmf,materialdesc,rir,period,qty,length,width,height,unit,gross,weightunit,pcperbox,volume,totalvolume,totalgrossweight,null,teu,brandname,'fe',getloadinggroup(loadingcode) as ""loadingcode(grouped)""  from forecastestimationteu where period >= '{2:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}'", myform.StartDate, myform.EndDate, myform.ForecastStart)

        Else
            'sqlstr = String.Format("(select *,'cietd' as status from vteuiex where currentinquiryetd >= '{0:yyyy-MM-dd}' and currentinquiryetd <= '{1:yyyy-MM-dd}'" & _
            '         " except select *,'cietd' as status from vteuexi where currentconfirmedetd >= '{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}')" & _
            '     " Union All " & _
            '     " select *,'ccetd' as status from vteuex where currentconfirmedetd >='{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}' " & _
            '     " Union All" & _
            '     " select myyear,substring(weeketa::text,5,2)::double precision,mymonth,null,null,sapcustomercode,sapcustomername,null,null,null,forwardername,loadingcode,zone,pod,null,null,""SEB Asia SAP Vendor Code"",""SEB ASIA SAP Vendor Name"",cmmf,materialdesc,rir,period,qty,length,width,height,unit,gross,weightunit,pcperbox,volume,totalvolume,totalgrossweight,null,teu,brandname,'fe'  from forecastestimationteuex where period >= '{2:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}'", myform.StartDate, myform.EndDate, myform.ForecastStart)
            'sqlstr = String.Format("select *,'cietd' as status from vteuiex where currentinquiryetd >= '{0:yyyy-MM-dd}' and currentinquiryetd <= '{1:yyyy-MM-dd}'" &
            '     " Union All " & _
            '     " select *,'cietd_past' as status from vteuiex where currentinquiryetd < '{0:yyyy-MM-dd}' " &
            '     " Union All " &
            '     " select *,'ccetd' as status from vteuex where currentconfirmedetd >='{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}' " &
            '     " Union All" & _
            '     " select *,'ccetd_past' as status from vteuex where currentconfirmedetd < '{0:yyyy-MM-dd}' " &
            '     " Union All " &
            '     " select myyear,substring(weeketa::text,5,2)::double precision,mymonth,null,null,sapcustomercode,sapcustomername,sapcustomercode,sapcustomername,null,forwardername,loadingcode,zone,pod,null,null,""SEB Asia SAP Vendor Code"",""SEB ASIA SAP Vendor Name"",cmmf,materialdesc,rir,period,qty,length,width,height,unit,gross,weightunit,pcperbox,volume,totalvolume,totalgrossweight,null,teu,brandname,'fe'  from forecastestimationteuex where period >= '{2:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}'", myform.StartDate, myform.EndDate, myform.ForecastStart)
            sqlstr = String.Format("select *,'cietd' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)""  from vteuiex where currentinquiryetd >= '{0:yyyy-MM-dd}' and currentinquiryetd <= '{1:yyyy-MM-dd}'" &
                 " Union All " & _
                 " select *,'cietd_past' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)""  from vteuiex where currentinquiryetd < '{0:yyyy-MM-dd}' " &
                 " Union All " &
                 " select *,'ccetd' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)""  from vteuex where currentconfirmedetd >='{0:yyyy-MM-dd}' and currentconfirmedetd <= '{1:yyyy-MM-dd}' " &
                 " Union All" & _
                 " select *,'ccetd_past' as status,getloadinggroup(loadingcode) as ""loadingcode(grouped)""  from vteuex where currentconfirmedetd < '{0:yyyy-MM-dd}' " &
                 " Union All " &
                 " select myyear,substring(weeketa::text,5,2)::double precision,mymonth,null,null,sapcustomercode,sapcustomername,sapcustomercode,sapcustomername,null,forwardername,loadingcode,zone,pod,null,null,""SEB Asia SAP Vendor Code"",""SEB ASIA SAP Vendor Name"",cmmf,materialdesc,rir,period,qty,length,width,height,unit,gross,weightunit,pcperbox,volume,totalvolume,totalgrossweight,null,teu,brandname,'fe',getloadinggroup(loadingcode) as ""loadingcode(grouped)""   from forecastestimationteuex where period >= '{2:yyyy-MM-dd}' and period <= '{1:yyyy-MM-dd}'", myform.StartDate, myform.EndDate, myform.ForecastStart)

        End If

        ProgressReport(4, "Execute Report")
        

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If myform.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            myform.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    myform.ToolStripStatusLabel1.Text = message
                Case 2
                    myform.ToolStripStatusLabel2.Text = message
                Case 4
                    Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
                    DirectoryBrowser.Description = "Which directory do you want to use?"
                    If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                        Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
                        Dim reportname = "TEUFG" '& GetCompanyName()
                        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
                        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

                        Dim myreport As New ExportToExcelFile(myform, sqlstr, filename, reportname, mycallback, PivotCallback, 2, "\templates\TEU.xltx")
                        myreport.Run(myform, New EventArgs)
                    End If
                Case 5
                    myform.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    myform.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    myform.ToolStripProgressBar1.Minimum = 1
                    myform.ToolStripProgressBar1.Value = myvalue(0)
                    myform.ToolStripProgressBar1.Maximum = myvalue(1)

            End Select

        End If

    End Sub

    
    Public Function getForwarderId(ByVal myCustomer As Long, ByVal myShipto As Long, ByVal myVendorcode As Long, ByVal myLoadingCode As String, ByVal myZoneid As Object, ByVal cmmf As Double) As String
        'Dim Sqlstr As String
        'Dim rstmp1 As ADODB.Recordset
        Dim myRule As Object
        Dim myFWid As String
        Dim Brandid As Object = Nothing
        Dim sqlstr As String = String.Empty
        'Find Rule in table customer
        myRule = getRule(myCustomer)
        If IsDBNull(myRule) Then
            myRule = Nothing
        End If
        If IsDBNull(myZoneid) Then
            myZoneid = Nothing
        End If

        If Not IsNothing(myRule) Then
            If myRule = "Y0" Then 'Check for special rule
                'Check Zone
                If myZoneid = 1 Then  'rsTmp!zoneid = 1 Then 'Hardcode to Southchina = 1
                    myRule = "Y1"
                Else
                    myRule = "Y2"
                End If
            ElseIf myRule = "Y5" Then
                Brandid = getBrandId(cmmf)
                If IsDBNull(Brandid) Then
                    MessageBox.Show(String.Format("Brand is null for this cmmf : {0}", cmmf))
                    Brandid = 0
                End If
            End If
            'Find custfw base on custzf Rule

            Select Case myRule
                Case "Y1"
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and sparty = {1} and vendorcode  = {2};", myCustomer, myShipto, myVendorcode)
                Case "Y2"
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and sparty = {1}  and zoneid = {2};", myCustomer, myShipto, IIf(IsNothing(myZoneid), "Null", myZoneid))
                Case "Y3"
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and loadingcode = '{1}' and zoneid = {2}", myCustomer, myLoadingCode, IIf(IsNothing(myZoneid), "Null", myZoneid))
                Case "Y4" 'search based on vendorcode
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and vendorcode  = {1} and zoneid = {2}", myCustomer, myVendorcode, IIf(IsNothing(myZoneid), "Null", myZoneid))
                Case "Y5" 'search based on brand
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and vendorcode  = {1} and brandid  = {2} and zoneid = {3}", myCustomer, myVendorcode, Brandid, IIf(IsNothing(myZoneid), "Null", myZoneid))
                Case "Y6"
                    'Search based on 1. ShiptoParty and Vendorcode
                    sqlstr = String.Format("select forwarderid from custzf where customerid = {0} and sparty = {1} and vendorcode  = {2}", myCustomer, myShipto, myVendorcode)
                Case Else
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
            End Select
        End If

        myFWid = getForwarderId(sqlstr)
        If IsNothing(myFWid) Then
            Select Case myRule
                Case "Y1"
                    'Find without zone
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid isnull and sparty isnull and vendorcode isnull", myCustomer)
                    myFWid = getForwarderId(sqlstr)
                Case "Y2"
                    'Find with zone
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and sparty isnull ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
                    myFWid = getForwarderId(sqlstr)
                    If IsNothing(myFWid) Then
                        sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid isnull and sparty isnull ", myCustomer)
                        myFWid = getForwarderId(sqlstr)
                    End If
                Case "Y3"
                    'Find with Zone
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and loadingcode isnull ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
                    myFWid = getForwarderId(sqlstr)
                    If IsNothing(myFWid) Then
                        'find without zone
                        sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid isnull  and loadingcode isnull ", myCustomer)
                        myFWid = getForwarderId(sqlstr)
                    End If
                Case "Y4"
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and vendorcode isnull ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
                    myFWid = getForwarderId(sqlstr)
                    If IsNothing(myFWid) Then
                        'find without zone
                        sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid isnull  and loadingcode isnull ", myCustomer)
                        myFWid = getForwarderId(sqlstr)
                    End If
                Case "Y5"
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and brandid isnull ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
                    myFWid = getForwarderId(sqlstr)
                    If IsNothing(myFWid) Then
                        'find without zone
                        sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid isnull  and loadingcode isnull ", myCustomer)
                        myFWid = getForwarderId(sqlstr)
                    End If

                Case "Y6"
                    'Find NonGroup
                    'Sqlstr = "Select forwarderid from forwarder where forwardername = 'NON GROUP'"
                    'Set rstmp1 = ocon.Execute(Sqlstr)
                    'myFWid = rstmp1!forwarderid
                    'New rule for Y6
                    '                2. ShiptoParty and Vendorcode isnull and Zone
                    '                3. ShiptoParty isnull and Vendorcode isnull and Zone
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and vendorcode isnull and sparty = {2}", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid), myShipto)

                    myFWid = getForwarderId(sqlstr)
                    If IsNothing(myFWid) Then
                        'find without zone
                        sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid = {1} and sparty isnull ", myCustomer, IIf(IsNothing(myZoneid), "Null", myZoneid))
                        myFWid = getForwarderId(sqlstr)
                    End If

                Case Else
                    'Find Without zone
                    sqlstr = String.Format("Select forwarderid from custzf where customerid = {0} and zoneid is null ", myCustomer)
                    myFWid = getForwarderId(sqlstr)
            End Select

        End If

        getForwarderId = myFWid
    End Function

    Public Function getRule(ByVal mycustomer As String) As Object
        Dim Sqlstr = String.Format("Select Rule from customer c where c.customercode = {0}", mycustomer) 'rsTmp!soldto
        Dim myresult As Object = String.Empty
        myadapter.ExecuteScalar(Sqlstr, myresult)
        Return myresult
    End Function

    Private Sub setNullForwarderId()
        Dim Sqlstr = "update cxsebpodtl set forwarderid = Null where " & _
                 " cxsebpodtlid in (select cxsebpodtlid from cxsebpodtl" & _
                 " left join ekko e on e.po = sebasiapono" & _
                 " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" & _
                 " where groupsbuid = 1)"
        myadapter.ExNonQuery(Sqlstr)
    End Sub

    Private Function getBrandId(ByVal cmmf As Double) As Object
        Dim Sqlstr = String.Format("select brandid from cmmf where cmmf = {0}", cmmf)
        Dim myresult As Object = String.Empty
        myadapter.ExecuteScalar(Sqlstr, myresult)
        Return myresult
    End Function

    Private Function getForwarderId(ByVal sqlstr As String) As String
        Dim myresult As Object = String.Empty
        myadapter.ExecuteScalar(sqlstr, myresult)
        Return myresult
    End Function


    Private Sub AssignForwarderIdForecast()
        Dim i As Long
        Dim myFWid As String = String.Empty
        Dim BS As New BindingSource
        'assign null to forwarderid
        Dim Sqlstr = "Update forecastestimation set forwarderid = Null"
        myadapter.ExecuteScalar(Sqlstr)

        'read all forecast report from     
        Sqlstr = "select feid,sapcustomercode,""SEB Asia SAP Vendor Code"",loadingcode,zoneid,cmmf from forecastestimationteu"
        Dim DS As New DataSet

        If myadapter.TbgetDataSet(Sqlstr, DS) Then            
            BS.DataSource = DS.Tables(0)
        End If
        'ProgressReport(1, String.Format("Assign Forwarder Forecast , Processing line number: {0}/{1}", i, BS.Count))
        ProgressReport(1, String.Format("Assign Forwarder Forecast ..."))
        For Each drv In BS.List
            i = i + 1
            'ProgressReport(7, String.Format("{0},{1}", i, BS.Count))
            ProgressReport(1, String.Format("Assign Forwarder Forecast ...Processing {0}/{1}", i, BS.Count))
            'Use Function to get forwarderid
            If Not (IsNothing(drv.row.item("sapcustomercode"))) Then
                'myFWid = getForwarderId(rsTmp!sapcustomercode, rsTmp!sapcustomercode, IIf(IsNull(rsTmp![SEB Asia SAP Vendor Code]), 0, rsTmp![SEB Asia SAP Vendor Code]), IIf(IsNull(rsTmp!loadingcode), 0, rsTmp!loadingcode), IIf(IsNull(rsTmp!zoneid), 0, rsTmp!zoneid))
                myFWid = getForwarderId(drv.row.item("sapcustomercode"), drv.row.item("sapcustomercode"), IIf(IsNothing(drv.row.item("SEB Asia SAP Vendor Code")), 0, drv.row.item("SEB Asia SAP Vendor Code")), IIf(IsNothing(drv.row.item("loadingcode")), 0, drv.row.item("loadingcode")), IIf(IsNothing(drv.row.item("zoneid")), 0, drv.row.item("zoneid")), drv.row.item("cmmf"))
            End If
            'Update odtl with forwarderid
            If IsNothing(myFWid) Then
                myFWid = "NULL"
            End If
            Sqlstr = "Update forecastestimation set forwarderid = " & myFWid & _
                   " where feid= " & drv.row.item("feid")
            myadapter.ExecuteScalar(Sqlstr)
        Next
        
    End Sub

    Private Sub FormattingReport()

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        owb.Names.Add("dbRange", RefersToR1C1:="=OFFSET('ALL DATA'!R1C1,0,0,COUNTA('ALL DATA'!C1),COUNTA('ALL DATA'!R1))")
        owb.Worksheets(1).select()
        Dim osheet = owb.Worksheets(1)
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlDatabase, SourceData:="dbrange", Version:=6))
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        
    End Sub


    Public Function getLastUpdate() As Boolean
        Dim sqlstr As String = "select startdate,status from programlocking where progname = 'FImFE'"

        Return myadapter.ExecuteScalar(sqlstr, _LastUpdate, message:=_ErrorMessage)
    End Function



End Class
