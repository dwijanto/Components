Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class FormGenerateParetoChart
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date
    Dim Dataset1 As DataSet
    Dim Filename As String = String.Empty
    Dim exclude As Boolean = True
    Public Property mySASL As String
    Dim groupsbu As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            ProgressReport(5, "")
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"
            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                Filename = DirectoryBrowser.SelectedPath

                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
            End If

        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()

        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        Dim status As Boolean = False
        Dim message As String = String.Empty
        sw.Start()


        ProgressReport(5, "Export To Excel..")

        status = GenerateReport(Message)

        If Status Then
            sw.Stop()
            ProgressReport(5, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'ProgressReport(2, TextBox2.Text & "Done.")
            ProgressReport(5, "")
            If MsgBox("File name: " & Filename & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(Filename)
            End If
            ProgressReport(5, "")
        Else
            ProgressReport(5, message)
        End If
        sw.Stop()
    End Sub


    Private Function GenerateReport(ByRef errmsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim hwnd As System.IntPtr
        Dim StopWatch As New Stopwatch
        Dim myfilter As New System.Text.StringBuilder
        myfilter.Append(" and sp.shipdate >= " & DateFormatyyyyMMddString(DateTimePicker1.Value.ToString) & " and sp.shipdate <= " & DateFormatyyyyMMddString(DateTimePicker2.Value.ToString))
        'myfilter.Append(" and abs(sp.shipdate - sd.currentinquiryetd) <= 7")
        myfilter.Append(mySASL)
        myfilter.Append(groupsbu)
        If exclude Then
            myfilter.Append(" and vendorname !~* 'MEYER'")
        End If
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()
        'Cursor.Current = Cursors.WaitCursor
        Dim ds As New DataSet
        'mySASL = " and abs(sp.shipdate - sd.currentinquiryetd) <= 7"
        'Dim sqlstr = "SELECT c1.customername" &
        '" FROM cxsebodtp od" &
        '" LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
        '" LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        '" LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
        '" LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
        '" LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
        '" left join povendor pv on pv.po = ph.sebasiapono" &
        '" left join vendor v on v.vendorcode = pv.vendorcode" &
        '" LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
        '" LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" &
        '" left join stdcmntdtl st on  st.stdcmntdtlid = pd.commentid" &
        '" WHERE od.ordertype::text = 'Shipment'::text and not commentid isnull " & myfilter.ToString &
        'mySASL &
        '" group by c1.customername  order by c1.customername"
        Dim sqlstr = "SELECT count(0)" &
        " FROM cxsebodtp od" &
        " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
        " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        " LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
        " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
        " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
        " left join ekko e on e.po = ph.sebasiapono" &
        " left join vendor v on v.vendorcode = e.vendorcode" &
        " LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
        " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" &
        " left join stdcmntdtl st on  st.stdcmntdtlid = pd.commentid" &
        " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" &
        " WHERE od.ordertype::text = 'Shipment'::text and not commentid isnull " & myfilter.ToString
        If Not DbAdapter1.TbgetDataSet(sqlstr, ds, errmsg) Then
            Return result
        End If

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty

        Try
            'Create Object Excel 
            ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(5, "Opening Template...")
            ProgressReport(5, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ParetoTemplate.xltx")

            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")
            'For i = 3 To (ds.Tables(0).Rows.Count + 3)
            '    oWb.Worksheets.Add(After:=oWb.Worksheets(i))
            'Next
            'backOrder

            Dim obj As New ThreadPoolObj

            'Get Filter

            'obj.osheet = oWb.Worksheets(2 + ds.Tables(0).Rows.Count)
            obj.osheet = oWb.Worksheets(2)

            'obj.strsql = "SELECT pd.commentid,c1.customername as factory,pd.comments,st.stdcmntdtlcode as frequency ,1 as count,1 as runningtotal,1 as pct," &
            '             " CASE WHEN abs(sp.shipdate - sd.currentinquiryetd) <= 7 THEN 1 " &
            '             " ELSE 0 END AS ""sasl<=7""" &
            '             " FROM cxsebodtp od" &
            '             " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
            '             " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
            '             " LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
            '             " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
            '             " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
            '             " left join povendor pv on pv.po = ph.sebasiapono" &
            '             " left join vendor v on v.vendorcode = pv.vendorcode" &
            '             " LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
            '             " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" &
            '             " left join stdcmntdtl st on  st.stdcmntdtlid = pd.commentid" &
            '             " WHERE od.ordertype::text = 'Shipment'::text and not commentid isnull " & myfilter.ToString &
            '             " ORDER BY od.cxsebodtpid;"
            obj.strsql = "SELECT pd.commentid,c1.customername as factory,pd.comments,st.stdcmntdtlcode as frequency ,1 as count,1 as runningtotal,1 as pct," &
             " CASE WHEN abs(sp.shipdate - sd.currentinquiryetd) <= 7 THEN 1 " &
             " ELSE 0 END AS ""sasl<=7""" &
             " FROM cxsebodtp od" &
             " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
             " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
             " LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
             " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
             " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
             " left join ekko e on e.po = ph.sebasiapono" &
             " left join vendor v on v.vendorcode = e.vendorcode" &
             " LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
             " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" &
             " left join stdcmntdtl st on  st.stdcmntdtlid = pd.commentid" &
             " left join purchasinggroup pg on pg.purchasinggroup = e.purchasinggroup" &
             " WHERE od.ordertype::text = 'Shipment'::text and not commentid isnull " & myfilter.ToString &
             " ORDER BY od.cxsebodtpid;"

            obj.osheet.Name = "DATA"
            'get dataset

            FillWorksheet(obj.osheet, obj.strsql, DbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                'For i = 0 To ds.Tables(0).Rows.Count
                '    ProgressReport(5, "Generating Pivot Tables.." & i)
                '    If i = 0 Then
                '        'Generate All
                '        CreatePivotTable(oWb, i, "All")
                '        'createchart(oWb, i, "All")
                '    Else
                '        'Generate each factory
                '        'CreatePivotTable(oWb, i, ds.Tables(0).Rows(i - 1).Item(0))
                '        'createchart(oWb, i, ds.Tables(0).Rows(i - 1).Item(0))
                '    End If
                'Next
                CreatePivotTable(oWb, 0, "All")
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            Filename = ValidateFileName(Filename, Filename & "\" & String.Format("PARETOCHART-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))
            ProgressReport(5, "Done ")
            ProgressReport(2, "Saving File ...")
            oWb.SaveAs(Filename)
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            errmsg = ex.Message
        Finally
            'ProgressReport(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result
    End Function
    Private Sub createchart(ByVal oWb As Excel.Workbook, ByVal sheetnum As Integer, ByVal message As String)
        'Create Name Range
        Try
            oWb.Names.Add(Name:="myyearWeek", RefersToR1C1:="=OFFSET(Data!R2C8,0,0,COUNTA(Data!C1)-1,1)")
            oWb.Names.Add(Name:="myAverage", RefersToR1C1:="=OFFSET(Data!R2C2,0,0,COUNTA(Data!C1)-1,1)")
            oWb.Names.Add(Name:="PCTValue", RefersToR1C1:="=OFFSET(Data!R2C3,0,0,COUNTA(Data!C1)-1,1)")
            oWb.Names.Add(Name:="myTarget", RefersToR1C1:="=OFFSET(Data!R2C4,0,0,COUNTA(Data!C1)-1,1)")
            oWb.Names.Add(Name:="myOrder", RefersToR1C1:="=OFFSET(Data!R2C5,0,0,COUNTA(Data!C1)-1,1)")
            oWb.Names.Add(Name:="myssl", RefersToR1C1:="=OFFSET(Data!R2C9,0,0,COUNTA(Data!C1)-1,1)")
            'Assign To Chart

            Dim osheet = oWb.Worksheets(1)
            Dim myChart = osheet.ChartObjects(1).Chart
            myChart.SeriesCollection(1).XValues = "=Data!myyearWeek"
            myChart.SeriesCollection(1).Values = "=Data!myorder"
            myChart.SeriesCollection(1).Name = "Count of OrderType"
            myChart.SeriesCollection(2).Values = "=Data!PCTValue"
            myChart.SeriesCollection(2).Name = "%-7<=SASL<=7 days"
            myChart.SeriesCollection(3).Values = "=Data!myTarget"
            myChart.SeriesCollection(3).Name = "Target 85%"
            myChart.SeriesCollection(4).Name = "SSL"
            myChart.SeriesCollection(4).Values = "=Data!myssl"
        Catch ex As Exception
            message = ex.Message
        End Try

    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    Me.ToolStripStatusLabel1.Text = message
                Case 3
                    Me.ToolStripStatusLabel1.Text = message
                Case 4
                    Me.ToolStripStatusLabel1.Text = message
                Case 5
                    Me.ToolStripStatusLabel1.Text = message

            End Select

        End If

    End Sub
    '    Private Sub CreatePareto()
    '        On Error GoTo ErrHdl
    '        Dim oXL As Excel.Application
    '        Dim oWB As Excel.Workbook
    '        Dim oSheet As Excel.Worksheet
    '        Dim oRange As Excel.Range
    '        Dim oChart As Chart

    '        Dim iRow As Integer
    '        Dim time1 As Date
    '        Dim myDate1 As String
    '        Dim mydate2 As String
    '        Dim mySelectedPath As String
    '        Dim i As Integer
    '        Dim strdata As String
    '        Dim maX1 As Long
    '        Dim maxR As Long
    '        Dim maxC As Integer
    '        Dim maX2 As Long
    '        Dim maX3 As Long
    '        Dim maxPvt As Integer
    '        Dim mySheetName As String
    '        Dim myWhere As String
    '        Dim sqlstrP As String
    '        Dim mySelectDate As String

    '        'Prepare data
    '        myDate1 = "'" & Year(DTPicker1.Value) & "-" & Month(DTPicker1.Value) & "-" & Day(DTPicker1.Value) & "'"
    '        mydate2 = "'" & Year(DTPicker2.Value) & "-" & Month(DTPicker2.Value) & "-" & Day(DTPicker2.Value) & "'"



    '        mySelectDate = " ship.shipdate >= " & myDate1 & " and ship.shipdate <= " & mydate2
    '        myWhere = " Where " & mySelectDate & myLTsign '& " and odtp.ordertype = 'Header'"

    '        sqlstrP = "select chd.cmnttxhdid, sbu.sbuname as sbu ,cdtl.cmnttxdtlname as problem, count(cmnttxdtlname) as frequency" & _
    '                 " FROM cmnttxhd chd" & _
    '                 " LEFT JOIN cmnttxdtl cdtl ON cdtl.cmnttxhdid = chd.cmnttxhdid" & _
    '                 " LEFT JOIN convcomnt cv ON cv.cmnttxdtlid = cdtl.cmnttxdtlid" & _
    '                 " LEFT JOIN stdcmntdtl s ON s.stdcmntdtlid = cv.stdcmntdtlid" & _
    '                 " LEFT JOIN lspodtl podtl ON podtl.commentid = s.stdcmntdtlid" & _
    '                 " LEFT JOIN cmmf on cmmf.cmmf = podtl.cmmf" & _
    '                 " LEFT JOIN activity ac ON cmmf.rir = ac.activitycode" & _
    '                 " LEFT JOIN sbu ON ac.sbuid = sbu.sbuid" & _
    '                 " LEFT JOIN lsodtype odtp ON odtp.lspodtlid = podtl.lspodtlid" & _
    '                 " LEFT JOIN lsship ship ON ship.lsodtypeid = odtp.lsodtypeid" & _
    '                 " left join lsodtl odtl on odtl.lsodtlid = odtp.lsodtlid" & _
    '                 " Left join lsohd ohd on ohd.sebasiasalesorder = odtl.sebasiasalesorder" & _
    '                 myWhere & _
    '                 " GROUP BY sbu.sbuname, chd.cmnttxhdid, chd.cmnttxhdname, cdtl.cmnttxdtlname" & _
    '                 " ORDER BY sbu.sbuname, count(cdtl.cmnttxdtlname) DESC"

    '        rsTmp = ocon.Execute(sqlstrP)

    '        If rsTmp.AbsolutePosition < 0 Then
    '            MsgBox("Data not available")
    '            Err.Raise(1)
    '        End If

    '        mySelectedPath = BrowseForFolder(Me.hWnd)
    '        If mySelectedPath = "" Then
    '            Exit Sub
    '        End If

    '        StatusBar1.Panels(1).Text = "Preparing Temporary Table...."
    '        'Clean table
    '        StatusBar1.Panels(1).Text = "Creating Chart...."
    '        mySelectedPath = IIf(Right(mySelectedPath, 1) = "\", Mid(mySelectedPath, 1, Len(mySelectedPath) - 1), mySelectedPath)
    '        time1 = Now

    '        oXL = CreateObject("excel.application")
    '        oWB = oXL.Workbooks.Open(App.Path & "\Template\WOR.xls")
    '        DoEvents()
    '        oXL.Visible = False 'True

    '        'prepare how many excel need based on SBU :: New criteria :: find only for sbu in transaction
    '        sqlstrP = "select distinct sbu.sbuid,sbu.sbuname" & _
    '                  " FROM cmnttxhd chd" & _
    '                 " LEFT JOIN cmnttxdtl cdtl ON cdtl.cmnttxhdid = chd.cmnttxhdid" & _
    '                 " LEFT JOIN convcomnt cv ON cv.cmnttxdtlid = cdtl.cmnttxdtlid" & _
    '                 " LEFT JOIN stdcmntdtl s ON s.stdcmntdtlid = cv.stdcmntdtlid" & _
    '                 " LEFT JOIN lspodtl podtl ON podtl.commentid = s.stdcmntdtlid" & _
    '                 " LEFT JOIN cmmf on cmmf.cmmf = podtl.cmmf" & _
    '                 " LEFT JOIN activity ac ON cmmf.rir = ac.activitycode" & _
    '                 " LEFT JOIN sbu ON ac.sbuid = sbu.sbuid" & _
    '                 " LEFT JOIN lsodtype odtp ON odtp.lspodtlid = podtl.lspodtlid" & _
    '                 " LEFT JOIN lsship ship ON ship.lsodtypeid = odtp.lsodtypeid" & _
    '                 " left join lsodtl odtl on odtl.lsodtlid = odtp.lsodtlid" & _
    '                 " Left join lsohd ohd on ohd.sebasiasalesorder = odtl.sebasiasalesorder" & _
    '                 " where " & mySelectDate & myLTsign & _
    '                 " ORDER BY sbu.sbuid,sbu.sbuname"

    '        rsSBU = ocon.Execute(sqlstrP)

    '        oXL.DisplayAlerts = False
    '        For i = 2 To oWB.Sheets.Count
    '            oWB.Sheets(2).Delete()
    '        Next i

    '        For i = 1 To rsSBU.RecordCount + 1
    '            oWB.Sheets.Add(After:=oWB.Sheets(i))
    '        Next i

    '        oSheet = oWB.Worksheets(rsSBU.RecordCount + 1)

    '        'Query First Data All
    '        myDate1 = "'" & Year(DTPicker1.Value) & "-" & Month(DTPicker1.Value) & "-" & Day(DTPicker1.Value) & "'"
    '        mydate2 = "'" & Year(DTPicker2.Value) & "-" & Month(DTPicker2.Value) & "-" & Day(DTPicker2.Value) & "'"

    '        oSheet.Name = "CHART DATA"
    '        For i = 0 To rsSBU.RecordCount - 1
    '            If Not IsNull(rsSBU!SBUid) Then


    '                Select Case i
    '                    Case 0
    '                        sqlstrP = "delete from pareto; Select setval('pareto_myid_seq'::regclass,1);"
    '                        ocon.Execute(sqlstrP)

    '                        sqlstrP = "select chd.cmnttxhdid, cdtl.cmnttxdtlname as problem, count(cmnttxdtlname) as frequency" & _
    '                                 " FROM cmnttxhd chd" & _
    '                                 " LEFT JOIN cmnttxdtl cdtl ON cdtl.cmnttxhdid = chd.cmnttxhdid" & _
    '                                 " LEFT JOIN convcomnt cv ON cv.cmnttxdtlid = cdtl.cmnttxdtlid" & _
    '                                 " LEFT JOIN stdcmntdtl s ON s.stdcmntdtlid = cv.stdcmntdtlid" & _
    '                                 " LEFT JOIN lspodtl podtl ON podtl.commentid = s.stdcmntdtlid" & _
    '                                 " LEFT JOIN lsodtype odtp ON odtp.lspodtlid = podtl.lspodtlid" & _
    '                                 " LEFT JOIN lsship ship ON ship.lsodtypeid = odtp.lsodtypeid" & _
    '                                 " left join lsodtl odtl on odtl.lsodtlid = odtp.lsodtlid" & _
    '                                 " Left join lsohd ohd on ohd.sebasiasalesorder = odtl.sebasiasalesorder" & _
    '                                 myWhere & _
    '                                 " GROUP BY  chd.cmnttxhdid, chd.cmnttxhdname, cdtl.cmnttxdtlname" & _
    '                                 " ORDER BY count(cdtl.cmnttxdtlname) DESC"

    '                        rsTmp = ocon.Execute(sqlstrP)
    '                        While Not rsTmp.EOF
    '                            If rsTmp!Frequency >= CInt(Text2.Text) Then
    '                                sqlstrP = "insert into pareto(pid,problem,frequency) values(" & rsTmp!cmnttxhdid & "," & escapeString(Trim(rsTmp!problem)) & "," & rsTmp!Frequency & ");"
    '                                ocon.Execute(sqlstrP)
    '                            End If
    '                            rsTmp.MoveNext()

    '                            DoEvents()
    '                        End While

    '                        sqlstrP = "SELECT a1.problem, a1.frequency, SUM(a2.frequency) as Running_Total,a1.pid, a1.frequency/(SELECT SUM(frequency) FROM pareto) as Pct_To_Total,  SUM(a2.frequency)/(SELECT SUM(frequency) FROM pareto) as gt " & _
    '                               " FROM pareto a1, pareto a2" & _
    '                               " Where a1.myid >= a2.myid" & _
    '                               " GROUP BY  a1.myid,a1.problem, a1.frequency,a1.pid" & _
    '                               " ORDER BY  a1.myid"
    '                    Case Else
    '                        oSheet = oWB.Worksheets("CHART DATA")
    '                        oWB.Worksheets("CHART DATA").Select()
    '                        sqlstrP = "delete from pareto;Select setval('pareto_myid_seq'::regclass,1);"

    '                        ocon.Execute(sqlstrP)

    '                        sqlstrP = "select chd.cmnttxhdid,sbu.sbuname as sbu, cdtl.cmnttxdtlname as problem, count(cmnttxdtlname) as frequency" & _
    '                                 " FROM cmnttxhd chd" & _
    '                                 " LEFT JOIN cmnttxdtl cdtl ON cdtl.cmnttxhdid = chd.cmnttxhdid" & _
    '                                 " LEFT JOIN convcomnt cv ON cv.cmnttxdtlid = cdtl.cmnttxdtlid" & _
    '                                 " LEFT JOIN stdcmntdtl s ON s.stdcmntdtlid = cv.stdcmntdtlid" & _
    '                                 " LEFT JOIN lspodtl podtl ON podtl.commentid = s.stdcmntdtlid" & _
    '                                 " LEFT JOIN cmmf on cmmf.cmmf = podtl.cmmf" & _
    '                                 " LEFT JOIN activity ac ON cmmf.rir = ac.activitycode" & _
    '                                 " LEFT JOIN sbu ON ac.sbuid = sbu.sbuid" & _
    '                                 " LEFT JOIN lsodtype odtp ON odtp.lspodtlid = podtl.lspodtlid" & _
    '                                 " LEFT JOIN lsship ship ON ship.lsodtypeid = odtp.lsodtypeid" & _
    '                                 " left join lsodtl odtl on odtl.lsodtlid = odtp.lsodtlid" & _
    '                                 " Left join lsohd ohd on ohd.sebasiasalesorder = odtl.sebasiasalesorder" & _
    '                                 myWhere & "  and sbu.sbuid = " & rsSBU!SBUid & _
    '                                 " GROUP BY  chd.cmnttxhdid, sbu.sbuname, chd.cmnttxhdname, cdtl.cmnttxdtlname" & _
    '                                 " ORDER BY count(cdtl.cmnttxdtlname) DESC"
    '                        rsTmp = ocon.Execute(sqlstrP)
    '                        While Not rsTmp.EOF
    '                            If rsTmp!Frequency >= CInt(Text2.Text) Then
    '                                sqlstrP = "insert into pareto(pid,problem,frequency) values(" & rsTmp!cmnttxhdid & "," & escapeString(Trim(rsTmp!problem)) & "," & rsTmp!Frequency & ");"
    '                                ocon.Execute(sqlstrP)
    '                            End If
    '                            rsTmp.MoveNext()
    '                            DoEvents()
    '                        End While

    '                        sqlstrP = "SELECT a1.problem, a1.frequency, SUM(a2.frequency) as Running_Total,a1.pid, a1.frequency/(SELECT SUM(frequency) FROM pareto) as Pct_To_Total,  SUM(a2.frequency)/(SELECT SUM(frequency) FROM pareto ) as gt " & _
    '                               " FROM pareto a1, pareto a2" & _
    '                               " Where a1.myid >= a2.myid" & _
    '                               " GROUP BY a1.myid, a1.problem, a1.frequency,a1.pid" & _
    '                               " ORDER BY a1.myid"
    '                End Select
    '                rsTmp = ocon.Execute(sqlstrP)
    '                If rsTmp.AbsolutePosition > 0 Then


    '                    iRow = 1
    '                    If i = 0 Then
    '                        StatusBar1.Panels(1).Text = "Creating Chart....(1/" & rsSBU.RecordCount & ") All SBU"
    '                        oSheet.Cells(1, 1) = "All SBU"

    '                    Else

    '                        oSheet.Cells(iRow, i * 5 + 1) = Trim(rsSBU!sbuname)
    '                        StatusBar1.Panels(1).Text = "Creating Chart...." & "(" & i + 1 & "/" & rsSBU.RecordCount & ")" & Trim(rsSBU!sbuname)
    '                    End If
    '                    oSheet.Cells(2, 1 + i * 5) = "Problem"
    '                    oSheet.Cells(2, 2 + i * 5) = "Frequency"
    '                    oSheet.Cells(2, 3 + i * 5) = "Percentage"
    '                    oSheet.Cells(2, 4 + i * 5) = "Running Total"
    '                    oSheet.Cells(2, 5 + i * 5) = "Percentage to Total"
    '                    iRow = 3
    '                    While Not rsTmp.EOF
    '                        oSheet.Cells(iRow, 1 + i * 5) = Trim(rsTmp!problem)
    '                        oSheet.Cells(iRow, 2 + i * 5) = rsTmp!Frequency
    '                        oSheet.Cells(iRow, 3 + i * 5) = rsTmp!pct_to_total
    '                        oSheet.Cells(iRow, 4 + i * 5) = rsTmp!Running_Total
    '                        oSheet.Cells(iRow, 5 + i * 5) = rsTmp!gt
    '                        iRow = iRow + 1
    '                        rsTmp.MoveNext()

    '                    End While
    '                    oSheet.Range(oSheet.Cells(2, 5 + i * 5), oSheet.Cells(rsTmp.RecordCount + 2, 5 + i * 5)).NumberFormat = "0%"
    '                    rsTmp.MoveLast()
    '                    myGT = rsTmp!Running_Total

    '                    maX3 = rsTmp.RecordCount + 2
    '                    mySheetName = oSheet.Name

    '                    'Create Chart

    '                    oSheet = oWB.Worksheets(i + 1)
    '                    oWB.Worksheets(i + 1).Select()
    '                    oSheet.Name = IIf(i = 0, "All SBU", Trim(rsSBU!sbuname))
    '                    'using embedded chart if excel version 2007
    '                    If oXL.Version < 12 Then
    '                        oChart = oXL.Charts.Add
    '                    Else
    '                        oChart = oSheet.Shapes.AddChart.Chart
    '                    End If
    '                    oChart.SeriesCollection.NewSeries()
    '                    oChart.SeriesCollection(1).Values = "='CHART DATA'!R3C" & 2 + (i * 5) & ":R" & maX3 & "C" & 2 + (i * 5)
    '                    oChart.SeriesCollection(1).Name = "='CHART DATA'!R2C" & 2 + i * 5
    '                    oChart.SeriesCollection(1).XValues = "='CHART DATA'!R3C" & 1 + i * 5 & ":R" & maX3 & "C" & 1 + i * 5
    '                    oChart.SeriesCollection(1).ChartType = xlColumnClustered

    '                    oChart.SeriesCollection.NewSeries()
    '                    oChart.SeriesCollection(2).Values = "='CHART DATA'!R3C" & 5 + i * 5 & ":R" & maX3 & "C" & 5 + i * 5
    '                    oChart.SeriesCollection(2).Name = "='CHART DATA'!R2C" & 5 + i * 5
    '                    oChart.SeriesCollection(2).ChartType = xlLineMarkers
    '                    oChart.SeriesCollection(2).AxisGroup = 2

    '                    oChart.Axes(xlValue).MinimumScale = 0
    '                    oChart.Axes(xlValue).MaximumScale = myGT
    '                    oChart.Axes(xlValue, xlSecondary).MinimumScale = 0
    '                    oChart.Axes(xlValue, xlSecondary).MaximumScale = 1

    '                    oChart.HasTitle = True
    '                    oChart.ChartTitle.Characters.Text = "Pareto Chart " & Chr(10) & Option3(myLT3Index).Caption
    '                    oChart.HasLegend = True
    '                    oChart.SeriesCollection(1).Interior.colorindex = 10
    '                    oChart.SeriesCollection(2).MarkerSize = 7
    '                    oChart.Location(xlLocationAsObject, oSheet.Name)
    '                    oSheet.Shapes(1).Top = 10
    '                    oSheet.Shapes(1).Left = 10
    '                    oSheet.Shapes(1).ScaleWidth(1.7, msoFalse, msoScaleFromTopLeft)
    '                    oSheet.Shapes(1).ScaleHeight(1.7, msoFalse, msoScaleFromTopLeft)
    '                Else
    '                    oSheet = oWB.Worksheets(i + 1)
    '                    oWB.Worksheets(i + 1).Select()
    '                    oSheet.Name = Trim(rsSBU!sbuname) & " (No Data)"
    '                End If
    '            End If
    '            rsSBU.MoveNext()
    '            DoEvents()

    '        Next i

    '        'Create Database
    '        oSheet = oWB.Worksheets(i + 2)
    '        oWB.Worksheets(i + 2).Select()


    '        StatusBar1.Panels(1).Text = "Retrieving data from server..."
    '        oSheet.Name = "Database"
    '        mycon = "ODBC;DRIVER={PostgreSQL ANSI};DATABASE=" & myDatabase & ";SERVER=" & myServer & ";PORT=" & myPort & ";UID=" & pUserid & ";PWD=" & pPassword & ";SSLmode=disable;ReadOnly=0;" & _
    '                "Protocol=7.4;FakeOidIndex=0;ShowOidColumn=0;RowVersioning=0;ShowSystemTables=0;ConnSettings=;Fetch=100;Socket=4096;UnknownSizes=0;" & _
    '                "MaxVarcharSize=255;MaxLongVarcharSize=8190;Debug=0;CommLog=0;Optimizer=1;Ksqo=1;UseDeclareFetch=0;TextAsLongVarchar=1;" & _
    '                "UnknownsAsLongVarchar=0;BoolsAsChar=1;Parse=0;CancelAsFreeStmt=0;ExtraSysTablePrefixes=dd_;LFConversion=1;UpdatableCursors=1;" & _
    '                "DisallowPremature=0;TrueIsMinus1=0;BI=0;ByteaAsLongVarBinary=0;UseServerSidePrepare=0;LowerCaseIdentifier=0;XaOpt=1"
    '        DoEvents()

    '        With oSheet.QueryTables.Add(mycon, oSheet.Range("A1"))
    '            .CommandText = Sqlstr & " and " & mySelectDate
    '            .FieldNames = True
    '            .RowNumbers = False
    '            .FillAdjacentFormulas = False
    '            .PreserveFormatting = True
    '            .RefreshOnFileOpen = False
    '            .BackgroundQuery = True
    '            .RefreshStyle = xlInsertDeleteCells
    '            .SavePassword = True
    '            .SaveData = True
    '            .AdjustColumnWidth = True
    '            .RefreshPeriod = 0
    '            .PreserveColumnInfo = True
    '            .Refresh(BackgroundQuery:=False)
    '            DoEvents()
    '        End With

    '        'Create Filter
    '        oRange = oSheet.Range("1:1")
    '        oRange.AutoFilter()
    '        oSheet.Columns("A:N").EntireColumn.AutoFit()
    '        oSheet.Columns("J:L").NumberFormat = "dd-MMM-yyyy"


    '        StatusBar1.Panels(1).Text = "Processing Time: " & Format(DateAdd("s", DateDiff("s", time1, Now), "00:00:00"), "HH:mm:ss") & " Done!"
    '        myDate1 = Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & Format(Day(DTPicker1.Value), "00")
    '        mydate2 = Year(DTPicker2.Value) & Format(Month(DTPicker2.Value), "00") & Format(Day(DTPicker2.Value), "00")
    '        oWB.Worksheets(1).Select()
    '        oWB.SaveAs(mySelectedPath & "\" & "Pareto-" & myDate1 & "-" & mydate2 & ".xls")
    '        oXL.Visible = True

    'errExit:
    '        rsTmp = Nothing
    '        rsSBU = Nothing

    '        If Not (oXL Is Nothing) Then
    '            oXL.DisplayAlerts = False
    '            oXL.DisplayAlerts = True
    '            oXL = Nothing
    '            oWB = Nothing
    '            oSheet = Nothing

    '        End If

    '        Exit Sub
    'ErrHdl:

    '        If Err.Number = 1 Then
    '            rsTmp = Nothing
    '            Exit Sub
    '        End If
    '        Call ErrHandler()
    '        GoTo errExit

    '    End Sub

    Private Sub FormGenerateParetoChart_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today
    End Sub

    Private Sub CreatePivotTable(ByVal oWb As Excel.Workbook, ByVal i As Integer, ByVal p3 As String)
        Dim osheet = oWb.Worksheets(i + 1)

        If i = 0 Then
            oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C15", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        Else
            oWb.Worksheets("Sheet1").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R6C15", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        End If


        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With

        osheet.PivotTables("PivotTable1").pivotfields("factory").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("frequency").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("count"), " Count", Excel.XlConsolidationFunction.xlSum)
        osheet.pivottables("PivotTable1").pivotfields("frequency").autosort(Excel.XlSortOrder.xlDescending, " Count", osheet.pivottables("PivotTable1").pivotcolumnaxis.pivotlines(1), 1)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("runningtotal"), " Running Total", Excel.XlConsolidationFunction.xlSum)
        With osheet.pivottables("PivotTable1").pivotfields(" Running Total")
            .calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
            .basefield = "frequency"            
        End With

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("pct"), " Percentage", Excel.XlConsolidationFunction.xlSum)
        With osheet.pivottables("PivotTable1").pivotfields(" Percentage")
            .calculation = Excel.XlPivotFieldCalculation.xlPercentOfColumn
            .numberformat = "0.00%"
        End With

        oWb.Names.Add("SUMFREQUENCY", RefersToR1C1:="=OFFSET(" & osheet.name & "!R8C15,0,0,COUNTA(" & osheet.name & "!C15)-3,1)")
        oWb.Names.Add("SUMCOUNT", RefersToR1C1:="=OFFSET(" & osheet.name & "!R8C16,0,0,COUNTA(" & osheet.name & "!C15)-3,1)")
        oWb.Names.Add("SUMRT", RefersToR1C1:="=OFFSET(" & osheet.name & "!R8C17,0,0,COUNTA(" & osheet.name & "!C15)-3,1)")

        'Create Chart

        Dim myChart = osheet.ChartObjects(1).Chart
        myChart.SeriesCollection(1).XValues = "=" & osheet.name & "!SUMFREQUENCY"
        myChart.SeriesCollection(1).Values = "=" & osheet.name & "!SUMCOUNT"
        myChart.SeriesCollection(1).Name = "COUNT"
        myChart.SeriesCollection(2).XValues = "=" & osheet.name & "!SUMFREQUENCY"
        myChart.SeriesCollection(2).Values = "=" & osheet.name & "!SUMRT"
        myChart.SeriesCollection(2).Name = "Running Total"
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
         exclude = CheckBox1.Checked
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged


        If RadioButton1.Checked Then
            groupsbu = ""
        ElseIf RadioButton2.Checked Then
            groupsbu = " and groupsbuid = 1"
        ElseIf RadioButton3.Checked Then
            groupsbu = " and groupsbuid = 2"
        End If
    End Sub
End Class