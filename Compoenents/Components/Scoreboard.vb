Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Xml
Imports System.Xml.Xsl

Public Class Scoreboard
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myQueryThread As New System.Threading.Thread(QueryDelegate)

    Dim exclude As Boolean = True
    Dim excludeComments As Boolean = False
    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim ReadFileStatus As Boolean = False
    Dim Dataset1 As DataSet
    Dim sb As StringBuilder
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim Source As String
    Dim excelcaption As String
    Dim hwnd As System.IntPtr
    Dim Marketbs As New BindingSource
    Dim VendorBs As New BindingSource
    Dim SAOBs As New BindingSource
    Dim comboid As String = String.Empty
    Dim selectedCheckedListbox As CheckedListBox
    Dim mycriteria As String = String.Empty
    Dim myexception As String = String.Empty
    Dim startdate As Date
    Dim enddate As Date
    Dim commentstartdate As Date
    Dim commentenddate As Date
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If myQueryThread.IsAlive Then
            ProgressReport(5, "Checkedlistbox still populating. Please wait.")
            Exit Sub
        End If

        If Not myThread.IsAlive Then
            'get Criteria
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            ToolStripStatusLabel3.Text = ""

            startdate = DateTimePicker1.Value.Date
            enddate = DateTimePicker2.Value.Date
            commentstartdate = DateTimePicker3.Value.Date
            commentenddate = DateTimePicker4.Value.Date

            If RadioButton1.Checked Then
                selectedCheckedListbox = CheckedListBox1
            ElseIf RadioButton2.Checked Then
                selectedCheckedListbox = CheckedListBox2
            ElseIf RadioButton3.Checked Then
                selectedCheckedListbox = CheckedListBox3


            End If
            ProgressReport(5, "")
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                FileName = DirectoryBrowser.SelectedPath

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
        sw.Start()


        Dim chkstate As CheckState
        chkstate = selectedCheckedListbox.GetItemCheckState(0)

        For Each item As Object In selectedCheckedListbox.CheckedItems
            ProgressReport(5, "Export To Excel..")
            Dim dr As DataRowView = DirectCast(item, DataRowView)
            Dim myvalue = dr.Item(0)
            If myvalue <> "All" Then
                If myvalue = "All Market" Or myvalue = "All Supplier" Or myvalue = "All SAO" Then
                    mycriteria = ""
                Else
                    If RadioButton1.Checked Then
                        mycriteria = "shiptoparty = " & dr.Item(1)
                        myexception = ""
                    ElseIf RadioButton2.Checked Then
                        mycriteria = "vendorcode = " & dr.Item(1)
                        myexception = ""
                    ElseIf RadioButton3.Checked Then
                        mycriteria = "sao = " & escapestr(dr.Item(0))
                        myexception = ""
                    End If
                End If
                Dim sr As New ScoreboardReport
                sr.filename = FileName
                sr.errormsg = errMsg
                sr.ds = Dataset1
                sr.criteria = mycriteria
                sr.exception = myexception
                sr.startdate = startdate
                sr.enddate = enddate
                sr.dr = dr
                sr.commentstartdate = commentstartdate
                sr.commentenddate = commentenddate
                Status = GenerateReport(sr)

                If Status Then
                    sw.Stop()
                    ProgressReport(5, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                    'ProgressReport(2, TextBox2.Text & "Done.")
                    ProgressReport(5, "")
                    If MsgBox("File name: " & FileName & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                        Process.Start(FileName)
                    End If
                    ProgressReport(5, "")
                Else
                    errSB.Append(sr.errormsg & vbCrLf)
                    ProgressReport(5, errSB.ToString)
                End If
                sw.Stop()
            End If
        Next
    End Sub

    'Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
    Private Function GenerateReport(ByVal sr As ScoreboardReport) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()
        'Cursor.Current = Cursors.WaitCursor


        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty


        'Need these variable to kill excel
        Dim aprocesses() As Process = Nothing '= Process.GetProcesses
        Dim aprocess As Process = Nothing
        Try
            'Create Object Excel 
            ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(5, "Opening Template...")
            ProgressReport(5, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ScoreboardComponentsTemplate.xltx")
            'For i = 0 To 6
            '    oWb.Worksheets.Add()
            'Next

            Dim events As New List(Of ManualResetEvent)()
            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")
            'backOrder
            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(3)
            Dim myfilter As New System.Text.StringBuilder
            If sr.criteria <> "" Then
                myfilter.Append(" and " & sr.criteria)
            End If
            If exclude Then
                myfilter.Append(" and vendorname !~* 'MEYER'")
            End If
            Dim commentlist As String = String.Empty
            For Each dr As DataRow In Dataset1.Tables(3).Rows
                commentlist = commentlist + IIf(commentlist = "", "", ",") & dr.Item(0).ToString.Trim
            Next
            If excludeComments Then
                myfilter.Append(" and validcomments(cmnttxdtlname,'" & commentlist & "',shipdate," & DateFormatyyyyMMdd(sr.commentstartdate) & "::date," & DateFormatyyyyMMdd(sr.commentenddate) & "::date)")
            End If

            obj.strsql = "select * from cxscoreboard where shipdate >= " & DateFormatyyyyMMdd(sr.startdate) & " and shipdate <= " & DateFormatyyyyMMdd(sr.enddate) & myfilter.ToString
            obj.Name = "WOR"
            If CreateWorksheet(obj) > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable1(oWb, 2, sr.startdate)
                ProgressReport(5, "Creating Charts..")
                CreateChart1(oWb, 1, sr)
            End If

            obj.osheet = oWb.Worksheets(4)
            myfilter.Clear()
            If sr.criteria <> "" Then
                myfilter.Append(" and " & sr.criteria)
            End If
            If exclude Then
                myfilter.Append(" and vendorname !~* 'MEYER'")
            End If


            obj.strsql = "select * from cxipltview where postingdate >= " & DateFormatyyyyMMdd(sr.startdate) & " and postingdate <= " & DateFormatyyyyMMdd(sr.enddate) & myfilter.ToString
            obj.Name = "IPLT"
            If CreateWorksheet(obj) > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable2(oWb, 2, sr.startdate)
                ProgressReport(5, "Creating Charts..")
                CreateChart2(oWb, 1, sr)
            End If


            obj.osheet = oWb.Worksheets(5)
            myfilter.Clear()
            If sr.criteria <> "" Then
                myfilter.Append(" and " & sr.criteria)
            End If
            If exclude Then
                myfilter.Append(" and vendorname !~* 'MEYER'")
            End If
            obj.strsql = "select * from cxipltview where invoicepostingdate >= " & DateFormatyyyyMMdd(sr.startdate) & " and invoicepostingdate <= " & DateFormatyyyyMMdd(sr.enddate) & myfilter.ToString
            obj.Name = "IPLT Miro"
            If CreateWorksheet(obj) > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable3(oWb, 2, sr.startdate)
                ProgressReport(5, "Creating Charts..")
                CreateChart3(oWb, 1, sr)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            FileName = ValidateFileName(FileName, FileName & "\" & String.Format("Scoreboard {3}-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day), sr.dr.Item(0).ToString))
            ProgressReport(5, "Done ")
            ProgressReport(2, "Saving File ...")
            oWb.SaveAs(FileName)
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            sr.errormsg = ex.Message
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
    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.CheckedListBox1.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Me.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try

        Else
            Select Case id
                Case 2
                    'TextBox2.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 3
                    'TextBox3.Text = message
                    Me.ToolStripStatusLabel2.Text = message
                Case 4
                    'TextBox1.Text = message
                    Me.ToolStripStatusLabel3.Text = message
                Case 5
                    'ToolStripStatusLabel1.Text = message
                    'ComboBox1.DataSource = bs
                    'ComboBox1.DisplayMember = "typeofitem"
                    'ComboBox1.ValueMember = "typeofitemid"
                    Me.ToolStripStatusLabel3.Text = message

                Case 6
                    CheckedListBox1.DataSource = Marketbs
                    CheckedListBox1.DisplayMember = "customername"
                    CheckedListBox1.ValueMember = "shiptoparty"
                    CheckedListBox2.DataSource = VendorBs
                    CheckedListBox2.DisplayMember = "vendorname"
                    CheckedListBox2.ValueMember = "vendorcode"
                    CheckedListBox3.DataSource = SAOBs
                    CheckedListBox3.DisplayMember = "officersebname"
                    CheckedListBox3.ValueMember = "ofsebid"
                Case 7

            End Select

        End If

    End Sub

    Private Sub FormOrderStatusReport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        'Load the query in background
        myQueryThread.Start()
    End Sub

    Sub DoQuery()
        Dataset1 = New DataSet
        Dim sqlstr As String = "select 'All' as customername,0 as shiptoparty union all select 'All Market',1 as shiptoparty union all (select distinct c.customername,s.shiptoparty from cxsebpodtl s" &
                               " left join customer c on c.customercode = s.shiptoparty order by c.customername);" &
                               " select 'All' as vendorname,0 as vendorcode union all select 'All Supplier',1  union all(select v.vendorname,p.vendorcode from (select distinct vendorcode from povendor ) as p" &
                               " left join vendor v on v.vendorcode = p.vendorcode order by v.vendorname);" &
                               " select 'All' as officersebname,0 as ofsebid union all select 'All SAO',1  union all(select distinct o.officersebname ,c.ofsebid from (select distinct s.shiptoparty from cxsebpodtl s ) as s" &
                               " left join customer c on c.customercode = s.shiptoparty " &
                               " left join officerseb o on o.ofsebid = c.ofsebid where not o.ofsebid isnull order by officersebname );" &
                               " select customername from orderlinemembers om" &
                               " left join orderline o on o.orderlineid = om.orderlineid " &
                               " where o.orderlinename = 'ExcludeComments'"

        If DbAdapter1.TbgetDataSet(sqlstr, Dataset1) Then
            Dataset1.Tables(0).TableName = "Market"
            Marketbs.DataSource = Dataset1.Tables(0)
            VendorBs.DataSource = Dataset1.Tables(1)
            SAOBs.DataSource = Dataset1.Tables(2)
            Dataset1.Tables(3).TableName = "ExcludeComments"
            ProgressReport(6, "")

        Else
            ProgressReport(5, "Error while loading Dataset.")
        End If

    End Sub

    Private Function CreateWorksheet(ByVal obj As Object) As Long

        Dim osheet = DirectCast(obj, ThreadPoolObj).osheet
        osheet.Name = DirectCast(obj, ThreadPoolObj).Name
        ProgressReport(5, "Waiting for the query to be executed..." & DirectCast(obj, ThreadPoolObj).osheet.Name)
        Dim sqlstr = DirectCast(obj, ThreadPoolObj).strsql
        FillWorksheet(osheet, sqlstr, DbAdapter1)
        Dim lastrow = osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row
        Return lastrow

        'DirectCast(obj.signal, ManualResetEvent).Set()
    End Function

    Private Sub CreatePivotTable1(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "WOR!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)

        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With


        osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next
        osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").numberformat = "MMM-yy"
        osheet.Columns("A:A").numberformat = "mmm-yy"
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetsasl", "=0.85", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetfcr", "=0.85", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetcltslt", "=28", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetfsl", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetoplt", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetotd", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetssl", "=0.70", True)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ordertype"), " count", Excel.XlConsolidationFunction.xlCount)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl"), " sasl", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl<=7"), " sasl<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetsasl"), " targetsasl", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr"), " fcr", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr<=7"), " fcr<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfcr"), " targetfcr", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("clt"), " clt", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("slt"), " slt", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetcltslt"), " targetcltslt", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl"), " fsl", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl<=7"), " fsl<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfsl"), " targetfsl", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("pireply"), " pireply", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetoplt"), " targetoplt", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("0-5days"), " 0-5days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("not conf"), " not conf", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("OTD<=7"), " OTD<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetotd"), " targetotd", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ssl"), " ssl", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetssl"), " targetssl", Excel.XlConsolidationFunction.xlAverage)


        osheet.PivotTables("PivotTable1").PivotFields(" sasl").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" sasl<=7").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetsasl").numberformat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" fcr").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" fcr<=7").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetfcr").numberformat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" clt").NumberFormat = "0.0"
        osheet.PivotTables("PivotTable1").PivotFields(" slt").numberformat = "0.0"
        osheet.PivotTables("PivotTable1").PivotFields("targetcltslt").NumberFormat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" fsl").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" fsl<=7").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetfsl").numberformat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" pireply").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" 0-5days").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetoplt").numberformat = "0%"
        osheet.PivotTables("PivotTable1").PivotFields("not conf").numberformat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" OTD<=7").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetotd").numberformat = "0%"

        osheet.PivotTables("PivotTable1").PivotFields(" ssl").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" targetssl").numberformat = "0%"

        oWb.Names.Add("MonthRange", RefersToR1C1:="=OFFSET('PivotTables'!R8C1,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("ordercount", RefersToR1C1:="=OFFSET('PivotTables'!R8C2,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("sasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C3,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("saslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C4,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetsasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C5,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("fcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C6,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("fcrle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C7,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetfcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C8,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("clt", RefersToR1C1:="=OFFSET('PivotTables'!R8C9,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("slt", RefersToR1C1:="=OFFSET('PivotTables'!R8C10,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetcltslt", RefersToR1C1:="=OFFSET('PivotTables'!R8C11,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("fsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C12,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("fslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C13,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetfsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C14,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("pireply", RefersToR1C1:="=OFFSET('PivotTables'!R8C15,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetoplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C16,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("lefive", RefersToR1C1:="=OFFSET('PivotTables'!R8C17,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("notconf", RefersToR1C1:="=OFFSET('PivotTables'!R8C18,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("otdle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C19,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetotd", RefersToR1C1:="=OFFSET('PivotTables'!R8C20,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("ssl", RefersToR1C1:="=OFFSET('PivotTables'!R8C21,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("targetssl", RefersToR1C1:="=OFFSET('PivotTables'!R8C22,0,0,COUNTA('PivotTables'!C1)-3,1)")

        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "PivotTables"


    End Sub

    Private Sub CreatePivotTable2(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C27", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable2").calculatedfields.add("targetiplt", "=0.95", True)
        With osheet.PivotTables("PivotTable2")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With


        osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("AA9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable2").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable2").Pivotfields("postingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Columns("AA:AA").numberformat = "mmm-yy"

        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("postingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("iplt<=5"), " iplt<=5", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable2").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable2").PivotFields(" iplt<=5").numberformat = "0%"

        oWb.Names.Add("myrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C27,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("iplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C28,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("targetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C29,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("ipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C30,0,0,COUNTA('PivotTables'!C28)-4,1)")

        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "PivotTables"

        osheet.Cells.EntireColumn.AutoFit()
    End Sub
    Private Sub CreatePivotTable3(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT Miro!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C35", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable3")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With

        osheet.PivotTables("PivotTable3").calculatedfields.add("targetiplt", "=0.95", True)

        osheet.PivotTables("PivotTable3").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("AI9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable3").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable3").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable3").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable3").Pivotfields("miropostingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Columns("AI:AI").numberformat = "mmm-yy"

        'osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("miropostingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("LeadTime"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("miro iplt<=5"), " miro iplt<=5", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable3").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable3").PivotFields(" miro iplt<=5").numberformat = "0%"

        oWb.Names.Add("miromyrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C35,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("miroiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C36,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("mirotargetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C37,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("miroipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C38,0,0,COUNTA('PivotTables'!C36)-4,1)")

        osheet.Cells.EntireColumn.AutoFit()
    End Sub
    'Private Sub CreatePivotTable(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
    '    Dim osheet As Excel.Worksheet

    '    oWb.Worksheets(isheet).select()
    '    osheet = oWb.Worksheets(isheet)
    '    oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "WOR!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)

    '    With osheet.PivotTables("PivotTable1")
    '        .ingriddropzones = True
    '        .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
    '    End With


    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Range("A9").Group(True, True, Periods:={False, False, False, False, True, False, True})
    '    osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

    '    For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
    '        Dim obj = DirectCast(item, Excel.PivotItem)
    '        If obj.Value.ToString <> startdate.Year.ToString Then
    '            obj.Visible = False
    '        End If
    '    Next
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").numberformat = "MMM-yy"

    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetsasl", "=0.85", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetfcr", "=0.85", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetcltslt", "=28", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetfsl", "=0.90", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetoplt", "=0.90", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetotd", "=0.90", True)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ordertype"), " count", Excel.XlConsolidationFunction.xlCount)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl"), " sasl", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl<=7"), " sasl<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetsasl"), " targetsasl", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr"), " fcr", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr<=7"), " fcr<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfcr"), " targetfcr", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("clt"), " clt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("slt"), " slt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetcltslt"), " targetcltslt", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl"), " fsl", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl<=7"), " fsl<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfsl"), " targetfsl", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("pireply"), " pireply", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetoplt"), " targetoplt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("0-5days"), " 0-5days", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("not conf"), " not conf", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("OTD<=7"), " OTD<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetotd"), " targetotd", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").PivotFields(" sasl").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" sasl<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetsasl").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" fcr").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" fcr<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetfcr").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" clt").NumberFormat = "0.0"
    '    osheet.PivotTables("PivotTable1").PivotFields(" slt").numberformat = "0.0"
    '    osheet.PivotTables("PivotTable1").PivotFields("targetcltslt").NumberFormat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" fsl").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" fsl<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetfsl").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" pireply").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" 0-5days").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetoplt").numberformat = "0%"
    '    osheet.PivotTables("PivotTable1").PivotFields("not conf").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" OTD<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetotd").numberformat = "0%"

    '    oWb.Names.Add("MonthRange", RefersToR1C1:="=OFFSET('PivotTables'!R8C1,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("sasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C3,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("saslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C4,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetsasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C5,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C6,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fcrle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C7,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetfcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C8,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("clt", RefersToR1C1:="=OFFSET('PivotTables'!R8C9,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("slt", RefersToR1C1:="=OFFSET('PivotTables'!R8C10,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetcltslt", RefersToR1C1:="=OFFSET('PivotTables'!R8C11,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C12,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C13,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetfsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C14,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("pireply", RefersToR1C1:="=OFFSET('PivotTables'!R8C15,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetoplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C16,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("lefive", RefersToR1C1:="=OFFSET('PivotTables'!R8C17,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("notconf", RefersToR1C1:="=OFFSET('PivotTables'!R8C18,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("otdle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C19,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetotd", RefersToR1C1:="=OFFSET('PivotTables'!R8C20,0,0,COUNTA('PivotTables'!C1)-3,1)")

    '    oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C27", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
    '    osheet.PivotTables("PivotTable2").calculatedfields.add("targetiplt", "=0.95", True)
    '    With osheet.PivotTables("PivotTable2")
    '        .ingriddropzones = True
    '        .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
    '    End With


    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Range("AA9").Group(True, True, Periods:={False, False, False, False, True, False, True})
    '    osheet.PivotTables("PivotTable2").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

    '    For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
    '        Dim obj = DirectCast(item, Excel.PivotItem)
    '        If obj.Value.ToString <> startdate.Year.ToString Then
    '            obj.Visible = False
    '        End If
    '    Next

    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Columns("AA:AA").numberformat = "mmm-yy"

    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("postingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("iplt<=5"), " iplt<=5", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable2").PivotFields(" Nbr of days").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable2").PivotFields(" targetiplt").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable2").PivotFields(" iplt<=5").numberformat = "0%"

    '    oWb.Names.Add("myrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C27,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("iplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C28,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("targetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C29,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("ipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C30,0,0,COUNTA('PivotTables'!C28)-4,1)")

    '    osheet.Cells.EntireColumn.AutoFit()
    '    osheet.Name = "PivotTables"


    'End Sub



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged
        Dim obj = DirectCast(sender, RadioButton)
        Select Case obj.Text
            Case "Market"
                CheckedListBox1.Enabled = True
                CheckedListBox2.Enabled = False
                CheckedListBox3.Enabled = False
            Case "Supplier"
                CheckedListBox1.Enabled = False
                CheckedListBox2.Enabled = True
                CheckedListBox3.Enabled = False
            Case "SAO"
                CheckedListBox1.Enabled = False
                CheckedListBox2.Enabled = False
                CheckedListBox3.Enabled = True
        End Select
    End Sub


    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged, CheckedListBox2.SelectedIndexChanged, CheckedListBox3.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)
    End Sub


    Private Function getselected(ByVal sender As Object) As String
        Dim myobj = DirectCast(sender, CheckedListBox)


        Dim sb As New StringBuilder
        Return sb.ToString
    End Function

    Private Sub CreateChart1(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As Components.ScoreboardReport)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
        osheet.Cells(12, 1) = sr.dr.Item(0)

        Dim ochart As New Excel.Chart

        ochart = osheet.ChartObjects("SASL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!sasl"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetsasl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!saslle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("FCR").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!clt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetfcr"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!fcrle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("CLT").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!clt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!slt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!targetcltslt"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("FSL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!fsl"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetfsl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!fslle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("OPLT").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!pireply"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetoplt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!lefive"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(4).Values = "='PivotTables'!notconf"
        ochart.SeriesCollection(4).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("OTD").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!otdle7"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetotd"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("SSL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!ordercount"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetssl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!ssl"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

    End Sub
    Private Sub CreateChart2(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As Components.ScoreboardReport)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
        osheet.Cells(12, 1) = sr.dr.Item(0)

        Dim ochart As New Excel.Chart

        ochart = osheet.ChartObjects("IPLT").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MyRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!iplt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetiplt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!myRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!ipltle5"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!myRange"

    End Sub
    Private Sub CreateChart3(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As Components.ScoreboardReport)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
        osheet.Cells(12, 1) = sr.dr.Item(0)

        Dim ochart As New Excel.Chart

        ochart = osheet.ChartObjects("IPLT Miro").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!miroMyRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!miroiplt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!mirotargetiplt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!miromyRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!miroipltle5"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!miromyRange"

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        exclude = CheckBox1.Checked
        excludeComments = CheckBox2.Checked
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub



    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Dim obj = CType(sender, CheckBox)
        DateTimePicker3.Enabled = obj.Checked
        DateTimePicker4.Enabled = obj.Checked
    End Sub
End Class

Public Class ScoreboardReport
    Public Property filename As String
    Public Property ds As DataSet
    Public Property errormsg As String
    Public Property criteria As String
    Public Property startdate As Date
    Public Property enddate As Date
    Public Property exception As String
    Public Property dr As DataRowView
    Public Property commentstartdate As Date
    Public Property commentenddate As Date


End Class