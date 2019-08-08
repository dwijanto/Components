Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class FormSASLSSL
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date
    Dim Dataset1 As DataSet
    Dim Filename As String = String.Empty
    Dim exclude As Boolean = True

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If Not myThread.IsAlive Then
            'get Criteria
            startdate = DateTimePicker1.Value.Date
            enddate = DateTimePicker2.Value.Date
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
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")
            'backOrder
            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(3)
            Dim myfilter As New System.Text.StringBuilder

            If exclude Then
                myfilter.Append(" and vendorname !~* 'MEYER'")
            End If
            obj.strsql = "select * from cxscoreboard where shipdate >= " & DateFormatyyyyMMdd(startdate) & " and shipdate <= " & DateFormatyyyyMMdd(enddate) & myfilter.ToString
            obj.osheet.Name = "WOR"

            FillWorksheet(obj.osheet, obj.strsql, DbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable1(oWb, 1, startdate)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            Filename = ValidateFileName(Filename, Filename & "\" & String.Format("SASLSSLReport-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))
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
        osheet.Range("A9").Group(36528, True, 7, Periods:={False, False, False, True, False, False, False})
        osheet.PivotTables("PivotTable1").Pivotfields("week").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("shipdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetsasl", "=0.85", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetfcr", "=0.85", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetcltslt", "=28", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetfsl", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetoplt", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetotd", "=0.90", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("targetssl", "=0.70", True)

                osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl"), " sasl", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl<=7"), " sasl<=7", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetsasl"), " targetsasl", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ssl"), " ssl", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetssl"), " targetssl", Excel.XlConsolidationFunction.xlAverage)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ordertype"), " count", Excel.XlConsolidationFunction.xlCount)

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

    

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        exclude = CheckBox1.Checked
    End Sub

End Class