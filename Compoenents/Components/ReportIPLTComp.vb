Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class ReportIPLTComp
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date
    Dim saodate As String

    Dim Dataset1 As DataSet
    Dim Filename As String = String.Empty
    Dim exclude As Boolean = True
    Dim myYearWeek As String = String.Empty

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            'get Criteria
            'myYearWeek = String.Format("{0} {1:00}", TextBox1.Text, CInt(TextBox2.Text))
            startdate = DateTimePicker1.Value
            enddate = DateTimePicker2.Value
            saodate = DateTimePicker3.Value.Month & "/1/" & DateTimePicker3.Value.Year

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

        status = GenerateReport(message)

        If status Then
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

    'Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String, ByVal dataset1 As DataSet) As Boolean
    Private Function GenerateReport(ByRef errmsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False
        Dim hwnd As System.IntPtr
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

        Try
            'Create Object Excel 
            ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(5, "Opening Template...")
            ProgressReport(5, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\IPLTComponentTemplate.xltx")

            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")
            'backOrder
            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(2)
            Dim myfilter As New System.Text.StringBuilder


            If exclude Then
                myfilter.Append(" and vendorname !~* 'MEYER'")
            End If

            obj.strsql = "select * from cxipltview where invoicepostingdate >= " & DateFormatyyyyMMdd(startdate) & " and invoicepostingdate <= " & DateFormatyyyyMMdd(enddate) & myfilter.ToString
            obj.osheet.Name = "DATA"

            FillWorksheet(obj.osheet, obj.strsql, DbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable(oWb, 3, startdate)
                createchart(oWb, 1, errmsg)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            Filename = ValidateFileName(Filename, Filename & "\" & String.Format("IPLTComp-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))
            ProgressReport(5, "Done ")
            ProgressReport(2, "Saving File ...")
            oWb.SaveAs(Filename)
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            errmsg = ex.Message
        Finally
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

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
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

            End Select

        End If

    End Sub

    Private Sub FormOrderStatusReport_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Application.DoEvents()

    End Sub

    Private Function CreateWorksheet(ByVal obj As Object) As Long
        Dim osheet = DirectCast(obj, ThreadPoolObj).osheet
        osheet.Name = DirectCast(obj, ThreadPoolObj).Name
        ProgressReport(5, "Waiting for the query to be executed..." & DirectCast(obj, ThreadPoolObj).osheet.Name)
        Dim sqlstr = DirectCast(obj, ThreadPoolObj).strsql
        FillWorksheet(osheet, sqlstr, DbAdapter1)
        Dim lastrow = osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row
        Return lastrow
    End Function



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Function getselected(ByVal sender As Object) As String
        Dim myobj = DirectCast(sender, CheckedListBox)


        Dim sb As New StringBuilder
        Return sb.ToString
    End Function

    Private Sub createchart(ByVal oWb As Excel.Workbook, ByVal sheetnum As Integer, ByVal message As String)
        'Create Name Range
        Try
            Dim osheet = oWb.Worksheets(1)
            Dim myChart = osheet.ChartObjects(1).Chart
            myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(1).Values = "=PivotTables!miroiplt"
            myChart.SeriesCollection(1).Name = "Average of Lead Time"
            myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetiplt"
            myChart.SeriesCollection(2).Name = "Target 95%"
            myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrange"
            myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7"
            myChart.SeriesCollection(3).Name = "%<=7 Days"


            myChart = osheet.ChartObjects(2).Chart
            myChart.SeriesCollection(1).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(1).Values = "=PivotTables!miroipltsao"
            myChart.SeriesCollection(1).Name = "Average of Lead Time"
            myChart.SeriesCollection(2).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(2).Values = "=PivotTables!mirotargetipltsao"
            myChart.SeriesCollection(2).Name = "Target 95%"
            myChart.SeriesCollection(3).XValues = "=PivotTables!miromyrangesao"
            myChart.SeriesCollection(3).Values = "=PivotTables!miroipltle7sao"
            myChart.SeriesCollection(3).Name = "%<=7 Days"
        Catch ex As Exception
            message = ex.Message
        End Try

    End Sub

    Private Sub CreatePivotTable(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add("targetiplt", "=0.95", True)

        osheet.PivotTables("PivotTable1").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A8").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable1").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").Pivotfields("miropostingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Columns("A:A").numberformat = "mmm-yy"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("LeadTime"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("miro iplt<=7"), " %iplt<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("miro iplt>=8"), " %iplt>=8", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable1").PivotFields(" %iplt<=7").numberformat = "0%"
        osheet.PivotTables("PivotTable1").PivotFields(" %iplt>=8").numberformat = "0%"
        osheet.Name = "PivotTables"
        oWb.Names.Add("miromyrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C1,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("miroiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C2,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("mirotargetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C3,0,0,COUNTA('PivotTables'!C1)-3,1)")
        oWb.Names.Add("miroipltle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C4,0,0,COUNTA('PivotTables'!C1)-3,1)")

        osheet.Cells.EntireColumn.AutoFit()

        'Second PivotTable
        oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable("PivotTables!R7C10", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable2")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.PivotTables("PivotTable2").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable2").Pivotfields("miropostingdaterange").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable2").PivotFields("miropostingdaterange").CurrentPage = saodate

        osheet.PivotTables("PivotTable2").Pivotfields("sao").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("LeadTime"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("miro iplt<=7"), " %iplt<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("miro iplt>=8"), " %iplt>=8", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable2").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable2").PivotFields(" %iplt<=7").numberformat = "0%"
        osheet.PivotTables("PivotTable2").PivotFields(" %iplt>=8").numberformat = "0%"
        osheet.Name = "PivotTables"
        oWb.Names.Add("miromyrangesao", RefersToR1C1:="=OFFSET('PivotTables'!R9C10,0,0,COUNTA('PivotTables'!C10)-4,1)")
        oWb.Names.Add("miroipltsao", RefersToR1C1:="=OFFSET('PivotTables'!R9C11,0,0,COUNTA('PivotTables'!C10)-4,1)")
        oWb.Names.Add("mirotargetipltsao", RefersToR1C1:="=OFFSET('PivotTables'!R9C12,0,0,COUNTA('PivotTables'!C10)-4,1)")
        oWb.Names.Add("miroipltle7sao", RefersToR1C1:="=OFFSET('PivotTables'!R9C13,0,0,COUNTA('PivotTables'!C10)-4,1)")

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        exclude = CheckBox1.Checked
    End Sub
End Class