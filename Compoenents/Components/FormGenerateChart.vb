Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class FormGenerateChart
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date
    Dim Dataset1 As DataSet
    Dim Filename As String = String.Empty
    Dim exclude As Boolean = True
    Dim myYearWeek As String = String.Empty

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            'get Criteria
            myYearWeek = String.Format("{0} {1:00}", TextBox1.Text, CInt(TextBox2.Text))

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



    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating, TextBox2.Validating
        Dim obj = DirectCast(sender, TextBox)
        If Not IsNumeric(obj.Text) Then
            ErrorProvider1.SetError(obj, "Numeric value required.")
            e.Cancel = True
        Else
            ErrorProvider1.SetError(obj, "")
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
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\SASLSSLWeekly.xltx")

            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")
            'backOrder
            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(2)
            Dim myfilter As New System.Text.StringBuilder


            myfilter.Append(" Where yearweek >= '" & myYearWeek & "'")


            obj.strsql = "SELECT myweek,sasl,pctsasl,targetsasl,countordertype,cxweeklyevolutionid,myyear,('WK' || to_char(myweek,'99')) as yearweek,pctssl,myyear || to_char(myweek,'99') as myorder" & _
                         " FROM cxweeklyevolution" & _
                         myfilter.ToString & " order by myorder asc"
            obj.osheet.Name = "DATA"

            FillWorksheet(obj.osheet, obj.strsql, DbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                'CreatePivotTable1(oWb, 1, startdate)
                createchart(oWb, 1, errmsg)
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





End Class