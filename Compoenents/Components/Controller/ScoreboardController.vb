Imports System.Text
Imports Microsoft.Office.Interop

Imports Components.SharedClass
Imports Components.ExportToExcelFile
Imports System.IO

Public Enum DepartmentEnum
    FinishedGoods = 1
    Components = 2
End Enum

Public Class ScoreboardController
    Dim myModel As ScoreboardModel = New ScoreboardModel
    Private WithEvents doBackground1 As DoBackground
    Private Parent As Object

    Public Department As DepartmentEnum = DepartmentEnum.FinishedGoods
    Public WMF As Boolean = False
    Public GROUPSUPPLIER As Boolean = False

    Public ExcludeSiS As Boolean = True
    Public OnlySIS As Boolean = False
    Public OnlyWMF As Boolean = False
    Public startdate As Date
    Public enddate As Date
    Public currentmonth As Date
    Public fslstartdate As Date

    Private SISList As String
    Private WMFList As String
    Public WMFCriteria As String = String.Empty
    Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
    Dim Pivot15 As Long
    Dim Pivot2 As Long
    Dim Pivot3 As Long
    Dim Pivot6 As Integer
    Dim Pivot4 As Long
    Dim whiteSupplier As Boolean
    Dim wSLow As Double
    Dim myTarget As Object
    Dim myTargetSSL As Object
    Dim myPareto As Boolean
    Dim myParetoGrp As Boolean
    Dim ParetoBar As Integer
    Dim ParetoBarGrp As Integer
    Dim myChart As Excel.Chart
    Dim smileySASL As Object
    Dim smileySSL As Object
    Dim smileyFSSLN As Object
    Dim smileyFSSL As Object
    Dim smileyOPLT As Object
    Dim smileyIPLT As Object
    Dim smileySCR As Object

    Public Sub New(ByVal parent As Object)
        Me.Parent = parent
        doBackground1 = New DoBackground(parent)
    End Sub

    Public Sub GetInitialData()
        WMFCriteria = ""
        If WMF Then
            WMFCriteria = String.Format(" left join ekko e on e.vendorcode = vp.vendorcode" &
                          " where e.purchasinggroup in ({0})", WMFList)
        ElseIf GROUPSUPPLIER Then
            WMFCriteria = String.Format(" where v.vendorcode in (select supplierid from groupsupplier)")
        End If
        If IsNothing(doBackground1.myThread) Then
            Select Case Department
                Case DepartmentEnum.FinishedGoods
                    doBackground1.doThread(AddressOf doQueryFG)
                Case DepartmentEnum.Components
                    doBackground1.doThread(AddressOf doQueryCP)
            End Select
        Else
            If Not doBackground1.myThread.IsAlive Then
                Select Case Department
                    Case DepartmentEnum.FinishedGoods
                        doBackground1.doThread(AddressOf doQueryFG)
                    Case DepartmentEnum.Components
                        doBackground1.doThread(AddressOf doQueryCP)
                End Select
            End If
        End If
        

    End Sub

    Public Sub LoadData()
        doBackground1.doThread(AddressOf doWork)
    End Sub

    Sub GenerateReport()
        doBackground1.doThread(AddressOf doGenerateReport)
    End Sub

    Private Sub doGenerateReport()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()

        Dim chkstate As CheckState
        chkstate = Parent.CheckedListBox1.GetItemCheckState(0)
        Try
            For Each item As Object In Parent.CheckedListBox1.CheckedItems
                doBackground1.ProgressReport(5, "Marque")
                sw.Start()
                Dim dr As DataRowView = DirectCast(item, DataRowView)
                Dim myvalue = dr.Item(0)
                Dim mycriteria As String = String.Empty
                Dim myexception As String = String.Empty
                Dim status As Boolean

                doBackground1.ProgressReport(1, String.Format("Working on {0}", myvalue))

                If myvalue <> "All" Then
                    'If myvalue = "All Supplier" Then
                    '    mycriteria = ""
                    'Else
                    '    'Dim rd1 As Boolean = Parent.RadioButton1.checked()


                    '    'If RadioButton1.Checked Then
                    '    '    mycriteria = "shiptoparty = " & dr.Item(1)
                    '    '    myexception = ""
                    '    'ElseIf RadioButton2.Checked Then
                    '    '    mycriteria = "vendorcode = " & dr.Item(1)
                    '    '    myexception = ""
                    '    'ElseIf RadioButton3.Checked Then
                    '    '    mycriteria = "sao = " & escapestr(dr.Item(0))
                    '    '    myexception = ""
                    '    'End If
                    'End If
                    Dim sr As New ScoreboardReport
                    sr.filename = String.Format("{0}\{1}-{2}-", Parent.FullNameDirectory, "Scoreboard-MT", dr.Item(0))
                    sr.errormsg = errMsg
                    sr.ds = New DataSet
                    sr.criteria = mycriteria
                    sr.exception = myexception
                    sr.startdate = startdate
                    sr.enddate = enddate
                    sr.dr = dr

                    status = GenerateReport(sr)

                    If Not status Then
                        errSB.Append(String.Format("{0} {1} {2}", myvalue, sr.errormsg, vbCrLf))
                    End If
                End If
            Next
        Catch ex As Exception
            doBackground1.ProgressReport(1, ex.Message)
        Finally
            sw.Stop()
            If errSB.Length > 0 Then
                'Error Found
                doBackground1.ProgressReport(1, errSB.ToString)
            Else
                doBackground1.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            End If

            doBackground1.ProgressReport(6, "Continuous")
        End Try
    End Sub

    Sub doWork()

    End Sub

    Sub doQueryFG()
        doBackground1.ProgressReport(5, "Marque")
        doBackground1.ProgressReportCallback = AddressOf doQueryCallBack
        Try

            If myModel.LoadInitialDataFG(WMFCriteria) Then
                doBackground1.ProgressReport(4, "Calling Callback")
                doBackground1.ProgressReport(1, "LoadInitialData Done.")
            Else
                doBackground1.ProgressReport(1, myModel.ErrorMessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            doBackground1.ProgressReport(6, "Continuous")
        End Try
    End Sub

    Private Sub doQueryCP()
        doBackground1.ProgressReport(5, "Marque")
        doBackground1.ProgressReportCallback = AddressOf doQueryCallBack
        Try
            If myModel.LoadInitialDataCP(WMFCriteria) Then
                doBackground1.ProgressReport(4, "Calling Callback")
                doBackground1.ProgressReport(1, "LoadInitialData Done.")
            Else
                doBackground1.ProgressReport(1, myModel.ErrorMessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            doBackground1.ProgressReport(6, "Continuous")
        End Try
    End Sub

    Private Sub doBackground_Callback(ByVal sender As Object, ByVal e As System.EventArgs) Handles doBackground1.Callback
        Select Case sender
            Case 4
            Case 8
        End Select
    End Sub

    Private Sub doQueryCallBack()
        Parent.CheckedListBox1.DataSource = myModel.VendorBS
        Parent.CheckedListBox1.DisplayMember = "vendorname"
        Parent.CheckedListBox1.ValueMember = "vendorcode"
        WMFList = myModel.WMFList
        SISList = myModel.SISList
    End Sub

    Private Function GenerateReport(ByVal sr As ScoreboardReport) As Boolean
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Dim sqlstrOTDCSL As String = String.Empty
        Dim sqlstrFSL As String = String.Empty
        Dim sqlstrOPLT As String = String.Empty
        Dim sqlstrIPLT As String = String.Empty

        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()

        Application.DoEvents()

        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As Integer
        Try
            'Create Object Excel 
            doBackground1.ProgressReport(1, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            doBackground1.ProgressReport(1, "Opening Template...")
            doBackground1.ProgressReport(1, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\Templates\ScoreboardLGrp.xltx")

            For i = 3 To 7
                oWb.Sheets.Add(After:=oWb.Sheets(i))
            Next i

            'Preparing Query For
            'OTDCSL,FSL,OPLT,IPLT
            Dim myVendorcode As String = sr.dr.Item(1) 'Vendorcode
            Dim myVendorname As String = sr.dr.Item(0)

            Dim myCriteria As String = String.Empty
            Dim mycriteria1 As String = String.Empty
            Dim pg As String = String.Empty
            Dim pg1 As String = String.Empty
            Dim soldtoparty As String = String.Empty
            Dim soldtoparty1 As String = String.Empty

            If myVendorcode = 1 Then 'ALL Suppliers
                If OnlyWMF Then
                    myCriteria = " and substring(vendorcode::text,1,1) = '1'"
                    mycriteria1 = " and substring(e.vendorcode::text,1,1) = ''1''"
                Else
                    myCriteria = ""
                    mycriteria1 = ""
                End If
                myVendorcode = 0
                myVendorname = ""
            Else
                myCriteria = String.Format(" and vendorcode = {0}", sr.dr.Item(1))
                mycriteria1 = String.Format(" and e.vendorcode = {0}", sr.dr.Item(1))
            End If

            If ExcludeSiS Then
                soldtoparty = String.Format(" and not(sh.soldtoparty in ({0})) ", SISList)
                soldtoparty1 = String.Format(" and not(soldtoparty in ({0}))", SISList)
            ElseIf OnlySIS Then
                soldtoparty = String.Format(" and (sh.soldtoparty in ({0}))", SISList)
                soldtoparty1 = String.Format(" and (soldtoparty in ({0}))", SISList)
            ElseIf OnlyWMF Then
                pg = String.Format(" and (e.purchasinggroup in ({0}))", WMFList.Replace("'", "''"))
                pg1 = String.Format(" and purchasinggroup in ({0})", WMFList)
            End If

            Dim mydate1 As String = String.Format("'{0:yyyy-MM-dd}'", sr.startdate)
            Dim mydate2 As String = String.Format("'{0:yyyy-MM-dd}'", sr.enddate)
            Dim newdate As String = String.Format("'{0:yyyy-MM-dd}'", fslstartdate)
            Dim LastYearDate As String = String.Format("'{0:yyyy-MM-dd}'", DateAdd(DateInterval.Year, -1, sr.startdate))

            If Department = DepartmentEnum.FinishedGoods Then
                sqlstrOTDCSL = "select * from getotdsslnetfggrp3(" & LastYearDate & "," & mydate2 & ",'" & mycriteria1 & pg & soldtoparty & "');"
                sqlstrFSL = "select * from getfsslnetfg2(" & newdate & ",'" & mycriteria1 & pg & soldtoparty & "')"

                sqlstrOPLT = "Select * from cxopltfgspnew" & _
                         " where podocdate >= " & mydate1 & " and podocdate <= " & String.Format("'{0:yyyy-MM-dd}'", sr.enddate) & pg1 & soldtoparty1 & myCriteria
                sqlstrIPLT = "Select *,monthyear(miropostingdate) from cxipltfg" & _
                         " where miropostingdate >= " & String.Format("'{0:yyyy-MM-dd}'", sr.startdate) & " and miropostingdate <= " & String.Format("'{0:yyyy-MM-dd}'", sr.enddate) & myCriteria & pg1 & soldtoparty1
            Else
                sqlstrOTDCSL = "select * from getotdsslnetcompgrp(" & String.Format("'{0:yyyy-MM-dd}'", DateAdd(DateInterval.Year, -1, sr.startdate)) & "," & mydate2 & ",'" & mycriteria1 & pg & soldtoparty & "');"
                sqlstrFSL = "select * from getfsslnetcomp(" & newdate & ",'" & mycriteria1 & pg & soldtoparty & "')"
                sqlstrOPLT = "Select * from cxopltcompspnew" & _
                         " where podocdate >= " & mydate1 & " and podocdate <= " & mydate2 & pg1 & soldtoparty1 & myCriteria
                sqlstrIPLT = "Select *,monthyear(miropostingdate) from cxipltcomp" & _
                         " where miropostingdate >= " & mydate1 & " and miropostingdate <= " & mydate2 & myCriteria & pg1 & soldtoparty1
            End If

            Dim q1 As New QueryWorksheet With {.DataSheet = 3,
                                              .SheetName = "DATAOTDSASL",
                                              .Sqlstr = sqlstrOTDCSL}
            Dim q2 As New QueryWorksheet With {.DataSheet = 4,
                                               .SheetName = "DATAFSL",
                                               .Sqlstr = sqlstrFSL}
            Dim q3 As New QueryWorksheet With {.DataSheet = 5,
                                               .SheetName = "DATAOPLT",
                                               .Sqlstr = sqlstrOPLT}
            Dim q4 As New QueryWorksheet With {.DataSheet = 6,
                                               .SheetName = "DATAIPLT",
                                               .Sqlstr = sqlstrIPLT}
            myQueryWorksheetList.Add(q1)
            myQueryWorksheetList.Add(q2)
            myQueryWorksheetList.Add(q3)
            myQueryWorksheetList.Add(q4)

            doBackground1.ProgressReport(1, "Creating Worksheet...")

            'Retrieve Data

            For Each myquery In myQueryWorksheetList
                oWb.Worksheets(myquery.DataSheet).select()
                oSheet = oWb.Worksheets(myquery.DataSheet)
                oSheet.Name = myquery.SheetName
                doBackground1.ProgressReport(1, String.Format("Retrieving data from server...{0}", myquery.SheetName))

                Components.ExportToExcelFile.FillWorksheet(oSheet, myquery.Sqlstr)
                Dim orange = oSheet.Range("A1")
                Dim lastrow = GetLastRow(oXl, oSheet, orange)

                Select Case myquery.SheetName
                    Case "DATAOTDSASL"
                        Pivot15 = lastrow
                        oWb.Names.Add(Name:="DBRangeOTDSASL", RefersToR1C1:="=OFFSET(DATAOTDSASL!R1C1,0,0,COUNTA(DATAOTDSASL!C1),COUNTA(DATAOTDSASL!R1))")
                    Case "DATAFSL"
                        Pivot2 = lastrow
                        oWb.Names.Add(Name:="DBRangeFSL", RefersToR1C1:="=OFFSET(DATAFSL!R1C1,0,0,COUNTA(DATAFSL!C1),COUNTA(DATAFSL!R1))")
                    Case "DATAOPLT"
                        Pivot3 = lastrow
                        Pivot6 = 0
                        If InStr(1, myModel.VendorSASL, myVendorcode, vbTextCompare) > 0 Then
                            Pivot6 = 1
                        End If
                        oWb.Names.Add(Name:="DBRangeOPLT", RefersToR1C1:="=OFFSET(DATAOPLT!R1C1,0,0,COUNTA(DATAOPLT!C1),COUNTA(DATAOPLT!R1))")
                    Case "DATAIPLT"
                        Pivot4 = lastrow
                        oWb.Names.Add(Name:="DBRangeIPLT", RefersToR1C1:="=OFFSET(DATAIPLT!R1C1,0,0,COUNTA(DATAIPLT!C1),COUNTA(DATAIPLT!R1))")
                End Select

                If lastrow > 1 Then
                    doBackground1.ProgressReport(1, "Formatting Report..")
                    'Delegate for modification
                    'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                    'FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Next

            doBackground1.ProgressReport(1, "Exporting to excel...Generating Pivot Table")

            oWb.Worksheets(2).Select()
            oSheet = oWb.Worksheets(2)
            oSheet.Name = "PivotTables"

            Dim CurrYear As Boolean
            Dim LastYear As Boolean

            Dim SCRChart As Boolean

            whiteSupplier = False
            wSLow = 0.7

            myTarget = "=" & myModel.getTarget(myVendorcode, Department, "sasl")
            myTargetSSL = "=" & myModel.getTargetVendor(myVendorcode)
            If myTargetSSL = "=0.9" Then
                whiteSupplier = True
                wSLow = 0.8
            End If


            'Pivot Table Part
            If Pivot15 > 1 Then
                'Current Year
                CurrYear = CreatePivotCurrentYear(oWb, oSheet, myTarget, myTargetSSL)
                LastYear = CreatePivotLastYear(oWb, oSheet, myTarget)
                Call CreatePivotFull(oWb, oSheet, myTarget)
            End If

            If Pivot2 > 1 Then
                oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeFSL").CreatePivotTable(oSheet.Name & "!R6C53", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                oSheet.PivotTables("PivotTable2").PivotFields("currentconfirmedetd").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable2").PivotFields("currentconfirmedetd").Caption = "Current Confirmed ETD"
                oSheet.PivotTables("PivotTable2").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable2").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable2").PivotFields("vendorname").Position = 1
                'oSheet.PivotTables("PivotTable2").PivotFields("week").Orientation = xlRowField
                oSheet.PivotTables("PivotTable2").CalculatedFields.Add("%FSL", "='fsl'/ordercount", True)
                oSheet.PivotTables("PivotTable2").CalculatedFields.Add("%FSSL", "='fssl'/ordercount", True)
                oSheet.PivotTables("PivotTable2").CalculatedFields.Add("%FSSLN", "='fsslnet'/ordercount", True)
                oSheet.PivotTables("PivotTable2").CalculatedFields.Add("TargetFSL", "= " & myModel.getTarget(myVendorcode, Department, "fsl"), True)
                oSheet.PivotTables("PivotTable2").CalculatedFields.Add("TargetFSSL", "= " & myModel.getTarget(myVendorcode, Department, "fssl"), True)

                oSheet.PivotTables("PivotTable2").AddDataField(oSheet.PivotTables("PivotTable2").PivotFields("ordercount"), " OC-RT", Excel.XlConsolidationFunction.xlSum)
                oSheet.Range("BA7").Group(Start:=36836, End:=True, By:=7, Periods:={False, False, False, True, False, False, False})
                With oSheet.PivotTables("PivotTable2").PivotFields(" OC-RT")
                    .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                    .BaseField = "Current Confirmed ETD"
                End With

                oSheet.PivotTables("PivotTable2").AddDataField(oSheet.PivotTables("PivotTable2").PivotFields("fsl"), " FSL-RT", Excel.XlConsolidationFunction.xlSum)
                With oSheet.PivotTables("PivotTable2").PivotFields(" FSL-RT")
                    .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                    .BaseField = "Current Confirmed ETD"
                End With
                oSheet.PivotTables("PivotTable2").AddDataField(oSheet.PivotTables("PivotTable2").PivotFields("fssl"), " FSSL-RT", Excel.XlConsolidationFunction.xlSum)
                With oSheet.PivotTables("PivotTable2").PivotFields(" FSSL-RT")
                    .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                    .BaseField = "Current Confirmed ETD"
                End With

                oSheet.PivotTables("PivotTable2").AddDataField(oSheet.PivotTables("PivotTable2").PivotFields("fsslnet"), " FSSLN-RT", Excel.XlConsolidationFunction.xlSum)
                With oSheet.PivotTables("PivotTable2").PivotFields(" FSSLN-RT")
                    .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                    .BaseField = "Current Confirmed ETD"
                End With

                oSheet.PivotTables("PivotTable2").PivotFields("Current Confirmed ETD").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable2").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField


                oWb.Worksheets("PivotTables").PivotTables("PivotTable2").PivotCache.CreatePivotTable(oSheet.Name & "!R6C61", "PivotTable2B", Excel.XlPivotTableVersionList.xlPivotTableVersion10)

                oSheet.PivotTables("PivotTable2B").PivotFields("currentconfirmedetd").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable2B").PivotFields("currentconfirmedetd").Caption = "Current Confirmed ETD"
                oSheet.PivotTables("PivotTable2B").PivotFields("week").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable2B").PivotFields("Current Confirmed ETD").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

                oSheet.PivotTables("PivotTable2B").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable2B").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable2B").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable2B").PivotFields("vendorname").Position = 1
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("ordercount"), " Sum of Order Count", Excel.XlConsolidationFunction.xlCount)
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("%FSL"), " %FSL", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("%FSSL"), " %FSSL", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("%FSSLN"), " %FSSLN", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("TargetFSL"), " TargetFSL", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable2B").AddDataField(oSheet.PivotTables("PivotTable2B").PivotFields("TargetFSSL"), " TargetFSSL", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable2B").PivotFields(" %FSL").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable2B").PivotFields(" %FSSL").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable2B").PivotFields(" %FSSLN").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable2B").PivotFields(" TargetFSL").NumberFormat = "0%"
                oSheet.PivotTables("PivotTable2B").PivotFields(" TargetFSSL").NumberFormat = "0%"
                oSheet.PivotTables("PivotTable2B").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                oSheet.PivotTables("PivotTable2B").ColumnGrand = False
                oSheet.PivotTables("PivotTable2B").RowGrand = False
                'add other information
                Dim obj As Object


                obj = oSheet.Cells(1, 60)
                Dim myRow As Integer
                obj.FormulaR1C1 = "=COUNTA(C61:C61)"
                If obj.Value >= 3 Then

                    '**** check avail data on Week + 8 if not then find Week - 1

                    'Dim newdate As Date

                    myRow = obj.Value - 2 + 8


                    'Set smileyFSL = oSheet.Cells(myRow + 1, 68)
                    smileyFSSLN = oSheet.Cells(myRow + 1, 68)
                    smileyFSSL = oSheet.Cells(myRow + 1, 69)

                    For i = 0 To 7
                        newdate = fslstartdate.AddDays(8 * (7 - i))
                        obj = oSheet.Cells(myRow, 68)
                        obj.Value = "=GETPIVOTDATA("" FSSLN-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))"
                        smileyFSSLN.Value = "=IF(GETPIVOTDATA("" FSSLN-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")) > 0.95,0,IF(GETPIVOTDATA("" FSSLN-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")) >= 0.79,1,2))"
                        obj = oSheet.Cells(myRow, 69)
                        obj.Value = "=GETPIVOTDATA("" FSSL-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))"

                        smileyFSSL.Value = "=IF(GETPIVOTDATA("" FSSL-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")) > GETPIVOTDATA("" TargetFSSL"",$BI$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")),0,IF(GETPIVOTDATA("" FSSL-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")) >= " & wSLow & " ,1,2))"
                        If Not obj.Text = "#REF!" Then
                            Exit For
                        End If

                    Next

                End If
                obj = Nothing
            End If

            If Pivot3 > 1 Then
                oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeOPLT").CreatePivotTable(oSheet.Name & "!R6C79", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersion10)

                '******************
                oSheet.PivotTables("PivotTable3").PivotFields("podocdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable3").PivotFields("podocdate").Caption = "Doc Date"
                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("LeadTime"), "Average of Lead Time", Excel.XlConsolidationFunction.xlAverage)
                oSheet.Range("CA7").Group(Start:=True, End:=True, Periods:={False, False, False, False, True, False, True})
                oSheet.PivotTables("PivotTable3").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable3").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable3").PivotFields("vendorname").Position = 1
                oSheet.PivotTables("PivotTable3").PivotFields("monthyear").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable3").PivotFields("Doc Date").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable3").CalculatedFields.Add("% 0-5 Days", "='0-5 days'/('Not Conf' + 'Nb Confirmed')", True)
                oSheet.PivotTables("PivotTable3").CalculatedFields.Add("% Not Conf", "='Not Conf'/('Not Conf' + 'Nb Confirmed')", True)
                oSheet.PivotTables("PivotTable3").CalculatedFields.Add("Target 0-5", "=" & myModel.getTarget(myVendorcode, Department, "oplt"), True)

                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("0-5 days"), " 0-5 days", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("Not Conf"), "Nb of Not Conf", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("% 0-5 Days"), "%0-5 Days", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("% Not Conf"), "%Not Conf", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable3").AddDataField(oSheet.PivotTables("PivotTable3").PivotFields("Target 0-5"), "Target 0-5 Days", Excel.XlConsolidationFunction.xlSum)

                oSheet.PivotTables("PivotTable3").PivotFields("Average of Lead Time").NumberFormat = "0.0"
                oSheet.PivotTables("PivotTable3").PivotFields("%0-5 Days").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable3").PivotFields("%Not Conf").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable3").PivotFields("Target 0-5 Days").NumberFormat = "0%"
                oSheet.PivotTables("PivotTable3").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField

                'Dim obj As Object

                smileyOPLT = oSheet.Cells(8, 88)
                smileyOPLT.Value = "=if(GETPIVOTDATA(""%0-5 Days"",$CA$6) > GETPIVOTDATA(""Target 0-5 Days"",$CA$6),0,2 )"

            End If

            If Pivot4 > 1 Then
                oWb.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeIPLT").CreatePivotTable(oSheet.Name & "!R6C105", "PivotTable4", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                oSheet.PivotTables("PivotTable4").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable4").PivotFields("miropostingdate").Caption = "Miropostingdate"

                oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                'oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate1").Orientation = xlPageField
                'Dim p As Excel.PivotItem

                'hide blank shipmentdate
                Dim mycount As Integer = oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems.Count
                'If oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems(oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems.Count) = "(blank)" Then
                If oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems(mycount).value = "(blank)" Then
                    If oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems.Count <> 1 Then
                        'oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems(oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems.Count).Visible = False
                        oSheet.PivotTables("PivotTable4").PivotFields("shipmentdate").PivotItems(mycount).Visible = False
                    End If

                End If


                oSheet.PivotTables("PivotTable4").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable4").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField

                oSheet.PivotTables("PivotTable4").PivotFields("vendorname").Position = 1

                oSheet.PivotTables("PivotTable4").CalculatedFields.Add("% <=7 Days", "='<=7'/'ordercount'", True)
                'oSheet.PivotTables("PivotTable4").CalculatedFields.Add "Target", "=0.98", True
                oSheet.PivotTables("PivotTable4").CalculatedFields.Add("Target", "=" & myModel.getTarget(myVendorcode, Department, "iplt"), True)
                oSheet.PivotTables("PivotTable4").PivotFields("monthyear").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable4").PivotFields("Miropostingdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable4").AddDataField(oSheet.PivotTables("PivotTable4").PivotFields("leadtime"), "Average of Lead Time", Excel.XlConsolidationFunction.xlAverage)
                oSheet.PivotTables("PivotTable4").AddDataField(oSheet.PivotTables("PivotTable4").PivotFields("<=7"), " <=7", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable4").AddDataField(oSheet.PivotTables("PivotTable4").PivotFields("ordercount"), " Number Of Lines", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable4").AddDataField(oSheet.PivotTables("PivotTable4").PivotFields("% <=7 Days"), " % <=7 Days", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable4").AddDataField(oSheet.PivotTables("PivotTable4").PivotFields("Target"), " Target", Excel.XlConsolidationFunction.xlSum)

                oSheet.PivotTables("PivotTable4").PivotFields("Average of Lead Time").NumberFormat = "0.0"
                oSheet.PivotTables("PivotTable4").PivotFields(" % <=7 Days").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable4").PivotFields(" Target").NumberFormat = "0%"

                oSheet.Range("DA7").Group(Start:=True, End:=True, Periods:={False, False, False, False, True, False, True})

                oSheet.PivotTables("PivotTable4").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField


                smileyIPLT = oSheet.Cells(8, 113)
                smileyIPLT.Value = "=if(GETPIVOTDATA("" % <=7 Days"",$DA$6) > GETPIVOTDATA("" Target"",$DA$6),0,2 )"
            End If


            If Pivot15 > 1 And CurrYear Then

                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C131", "PivotTable5", DefaultVersion:=Excel.XlPivotTableVersionList.xlPivotTableVersion10)

                oSheet.PivotTables("PivotTable5").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable5").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField

                oSheet.PivotTables("PivotTable5").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable5").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable5").PivotFields("vendorname").Position = 1

                oSheet.PivotTables("PivotTable5").CalculatedFields.Add("%SCR", "='scr'/'firstconfirmation'", True)
                'oSheet.PivotTables("PivotTable5").CalculatedFields.Add "%NotOnTime", "='NB NOT OTD'/'OrderCount'", True
                If Department = DepartmentEnum.FinishedGoods Then
                    oSheet.PivotTables("PivotTable5").CalculatedFields.Add("TargetSCRHigh", "=" & myModel.getTarget(myVendorcode, Department, "scrhigh"), True)
                    oSheet.PivotTables("PivotTable5").CalculatedFields.Add("TargetSCRLow", "=" & myModel.getTarget(myVendorcode, Department, "scrlow"), True)
                Else
                    oSheet.PivotTables("PivotTable5").CalculatedFields.Add("TargetSCRHigh", "=" & myModel.getTarget(myVendorcode, Department, "scrhighcomp"), True)
                    oSheet.PivotTables("PivotTable5").CalculatedFields.Add("TargetSCRLow", "=" & myModel.getTarget(myVendorcode, Department, "scrlowcomp"), True)
                End If

                'request by skong 2013-11-07
                oSheet.PivotTables("PivotTable5").CalculatedFields.Add("Customer LT", "='customerlt_weight'/'weight'", True)
                oSheet.PivotTables("PivotTable5").CalculatedFields.Add("Supplier LT", "='supplierlt_weight'/'weight'", True)
                'end request

                oSheet.PivotTables("PivotTable5").PivotFields("monthyear").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable5").PivotFields("Shipdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("%SCR"), "% SCR", Excel.XlConsolidationFunction.xlSum)
                'oSheet.PivotTables("PivotTable5").PivotFields("% On Time").Calculation = xlPercentOfColumn
                'oSheet.PivotTables("PivotTable5").AddDataField oSheet.PivotTables("PivotTable5").PivotFields("%NotOnTime"), "% Not On Time", xlSum
                'oSheet.PivotTables("PivotTable5").AddDataField oSheet.PivotTables("PivotTable5").PivotFields("OrderCount"), "% Total", xlAverage
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("TargetSCRHigh"), "% TargetHigh", Excel.XlConsolidationFunction.xlAverage)
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("TargetSCRLow"), "% TargetLow", Excel.XlConsolidationFunction.xlAverage)
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("firstconfirmation"), " Nbr of Line Confirmed successful", Excel.XlConsolidationFunction.xlSum)
                'oSheet.PivotTables("PivotTable5").AddDataField oSheet.PivotTables("PivotTable5").PivotFields("customerlt_weight"), "Customer LT", xlAverage
                'oSheet.PivotTables("PivotTable5").AddDataField oSheet.PivotTables("PivotTable5").PivotFields("supplierlt_weight"), "Supplier LT", xlAverage
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("Customer LT"), " Customer LT", Excel.XlConsolidationFunction.xlAverage)
                oSheet.PivotTables("PivotTable5").AddDataField(oSheet.PivotTables("PivotTable5").PivotFields("Supplier LT"), " Supplier LT", Excel.XlConsolidationFunction.xlAverage)


                oSheet.PivotTables("PivotTable5").PivotFields("% SCR").NumberFormat = "0.0%"
                'oSheet.PivotTables("PivotTable5").PivotFields("% Not On Time").NumberFormat = "0.0%"
                'oSheet.PivotTables("PivotTable5").PivotFields("% Total").NumberFormat = "0%"
                oSheet.PivotTables("PivotTable5").PivotFields("% TargetHigh").NumberFormat = "0%"
                oSheet.PivotTables("PivotTable5").PivotFields("% TargetLow").NumberFormat = "0%"


                oSheet.PivotTables("PivotTable5").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                oSheet.PivotTables("PivotTable5").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable5").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable5").PivotFields("Years").CurrentPage = Year(startdate)
                'oSheet.PivotTables("PivotTable5").ColumnGrand = False
                Dim obj As Object

                obj = oSheet.Cells(1, 130)
                'Dim myRow As Integer
                obj.FormulaR1C1 = "=COUNTA(C131:C131)"
                If obj.Value > 3 Then
                    SCRChart = True

                    'Dim obj As Object
                    obj = oSheet.Cells(8, 140)
                    obj.Value = "=iferror(GETPIVOTDATA(""% SCR"",$EA$6),0)"
                    Dim myValue As Double
                    myValue = obj.Value

                    smileySCR = oSheet.Cells(9, 140)
                    smileySCR.Value = "=if(" & myValue & " > GETPIVOTDATA(""% TargetHigh"",$EA$6),0,IF(" & myValue & " >GETPIVOTDATA(""% TargetLow"",$EA$6),1,2))"

                End If
                obj = Nothing

            End If

            'Create Pivot Pareto
            If Pivot15 > 1 And CurrYear Then
                'Create from cache

                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C157", "PivotTablePareto", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                'check currentmonth pareto
                Dim CreatePareto As Boolean
                CreatePareto = False
                For Each p In oSheet.PivotTables("PivotTablePareto").PivotFields("years").PivotItems
                    If p.Value = Format(currentmonth, "yyyy").ToString Then
                        CreatePareto = True
                        Exit For
                    End If
                Next
                'For Each p In oSheet.PivotTables("PivotTablePareto").PivotFields("monthyear").PivotItems
                '    If p.Value = Format(DTPicker3, "MMM-yy") Then
                '       CreatePareto = True
                '    End If
                'Next
                If Not CreatePareto Then
                    myPareto = False
                Else


                    'With oSheet.PivotTables("PivotTablePareto").PivotFields("monthyear")
                    '    .Orientation = xlPageField
                    '    .CurrentPage = Format(DTPicker3, "MMM-yy")
                    'End With
                    With oSheet.PivotTables("PivotTablePareto").PivotFields("years")
                        .Orientation = Excel.XlPivotFieldOrientation.xlPageField
                        .CurrentPage = Format(currentmonth, "yyyy")
                    End With


                    oSheet.PivotTables("PivotTablePareto").PivotFields("mgtmsg").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                    oSheet.PivotTables("PivotTablePareto").AddDataField(oSheet.PivotTables("PivotTablePareto").PivotFields("weight"), "Sum of weight", Excel.XlConsolidationFunction.xlSum)
                    oWb.Names.Add(Name:="createpareto", RefersToR1C1:="=COUNTA(PivotTables!C157)-4>0")
                    oWb.Names.Add(Name:="countparetobar", RefersToR1C1:="=COUNTA(PivotTables!R7C157:R7C182)-2")
                    Dim obj As Object
                    obj = oSheet.Range("iscreatepareto")
                    obj.FormulaR1C1 = "=createpareto"
                    myPareto = obj.Value

                    oWb.Names.Add(Name:="BarAxis" & 1, RefersToR1C1:="=OFFSET(PivotTables!R8C157,0,0,COUNTA(PivotTables!C157)-4,1)")
                    If myPareto Then
                        'oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems("CUSTOMER ISSUE (S)").Visible = False

                        'oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems("SSL GROSS").Visible = False

                        'Dim p As PivotItem
                        For Each p In oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems
                            If p.Value = "SSL GROSS" Then
                                If oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems.Count = 1 Then
                                    myPareto = False
                                Else
                                    p.Visible = False
                                End If

                            End If
                        Next

                        oSheet.PivotTables("PivotTablePareto").PivotFields("mgtmsg").AutoSort(Excel.XlSortOrder.xlDescending, "Sum of weight") 'ActiveSheet.PivotTables("PivotTablePareto").PivotColumnAxis.PivotLines(2), 1
                        oSheet.PivotTables("PivotTablePareto").PivotFields("sum of weight").NumberFormat = "0.0"
                    End If
                    obj = oSheet.Range("paretobar")
                    obj.FormulaR1C1 = "=countparetobar"
                    ParetoBar = obj.Value
                    For i = 1 To ParetoBar
                        oWb.Names.Add(Name:="PBar" & i, RefersToR1C1:="=OFFSET(PivotTables!R8C" & 157 + i & ",0,0,COUNTA(PivotTables!C157)-4,1)")
                        oWb.Names.Add(Name:="PBarName" & i, RefersToR1C1:="=PivotTables!R7C" & 157 + i)
                    Next i
                End If

                'Pivot Table Pareto Group

                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C183", "PivotTableParetoGrp", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                'check currentmonth pareto
                Dim CreateParetoGrp As Boolean
                CreateParetoGrp = False
                For Each p In oSheet.PivotTables("PivotTableParetoGrp").PivotFields("years").PivotItems
                    If p.Value = Format(currentmonth, "yyyy").ToString Then
                        CreateParetoGrp = True
                        Exit For
                    End If
                Next
                'For Each p In oSheet.PivotTables("PivotTablePareto").PivotFields("monthyear").PivotItems
                '    If p.Value = Format(DTPicker3, "MMM-yy") Then
                '       CreatePareto = True
                '    End If
                'Next
                If Not CreateParetoGrp Then
                    myParetoGrp = False
                Else


                    'With oSheet.PivotTables("PivotTablePareto").PivotFields("monthyear")
                    '    .Orientation = xlPageField
                    '    .CurrentPage = Format(DTPicker3, "MMM-yy")
                    'End With
                    With oSheet.PivotTables("PivotTableParetoGrp").PivotFields("years")
                        .Orientation = Excel.XlPivotFieldOrientation.xlPageField
                        .CurrentPage = Format(currentmonth, "yyyy")
                    End With


                    oSheet.PivotTables("PivotTableParetoGrp").PivotFields("grpname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    oSheet.PivotTables("PivotTableParetoGrp").PivotFields("catissues1").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                    oSheet.PivotTables("PivotTableParetoGrp").AddDataField(oSheet.PivotTables("PivotTableParetoGrp").PivotFields("weight"), "Sum of weight", Excel.XlConsolidationFunction.xlSum)
                    oWb.Names.Add(Name:="createparetogrp", RefersToR1C1:="=COUNTA(PivotTables!C183)-4>0")
                    oWb.Names.Add(Name:="countparetobargrp", RefersToR1C1:="=COUNTA(PivotTables!R7C183:R7C208)-2")
                    Dim obj As Object
                    obj = oSheet.Range("iscreateparetogrp")
                    obj.FormulaR1C1 = "=createparetogrp"
                    myParetoGrp = obj.Value

                    oWb.Names.Add(Name:="BarAxisGrp" & 1, RefersToR1C1:="=OFFSET(PivotTables!R8C183,0,0,COUNTA(PivotTables!C183)-4,1)")
                    If myParetoGrp Then
                        'oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems("CUSTOMER ISSUE (S)").Visible = False

                        'oSheet.PivotTables("PivotTablePareto").PivotFields("catissues1").PivotItems("SSL GROSS").Visible = False

                        'Dim p As PivotItem
                        For Each p In oSheet.PivotTables("PivotTableParetoGrp").PivotFields("catissues1").PivotItems
                            If p.Value = "SSL GROSS" Then
                                If oSheet.PivotTables("PivotTableParetoGrp").PivotFields("catissues1").PivotItems.Count = 1 Then
                                    myPareto = False
                                Else
                                    p.Visible = False
                                End If

                            End If
                        Next

                        oSheet.PivotTables("PivotTableParetoGrp").PivotFields("grpname").AutoSort(Excel.XlSortOrder.xlDescending, "Sum of weight") 'ActiveSheet.PivotTables("PivotTablePareto").PivotColumnAxis.PivotLines(2), 1
                        oSheet.PivotTables("PivotTableParetoGrp").PivotFields("sum of weight").NumberFormat = "0.0"
                    End If
                    obj = oSheet.Range("paretobargrp")
                    obj.FormulaR1C1 = "=countparetobargrp"
                    ParetoBarGrp = obj.Value
                    For i = 1 To ParetoBarGrp
                        oWb.Names.Add(Name:="PBarGrp" & i, RefersToR1C1:="=OFFSET(PivotTables!R8C" & 183 + i & ",0,0,COUNTA(PivotTables!C183)-4,1)")
                        oWb.Names.Add(Name:="PBarNameGrp" & i, RefersToR1C1:="=PivotTables!R7C" & 183 + i)
                    Next i
                End If


                'Create Pareto Category
                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C209", "PivotTableParetoCategory", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                With oSheet.PivotTables("PivotTableParetoCategory").PivotFields("years")
                    .Orientation = Excel.XlPivotFieldOrientation.xlPageField
                    .CurrentPage = Format(currentmonth, "yyyy")
                End With

                For Each p In oSheet.PivotTables("PivotTableParetoCategory").PivotFields("catissues1").PivotItems
                    If p.Value = "SSL GROSS" Then
                        If oSheet.PivotTables("PivotTableParetoCategory").PivotFields("catissues1").PivotItems.Count = 1 Then
                            myPareto = False
                        Else
                            p.Visible = False
                        End If

                    End If
                Next

                oSheet.PivotTables("PivotTableParetoCategory").PivotFields("catissues1").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTableParetoCategory").AddDataField(oSheet.PivotTables("PivotTableParetoCategory").PivotFields("weight"), "Sum of weight", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTableParetoCategory").AddDataField(oSheet.PivotTables("PivotTableParetoCategory").PivotFields("weight"), "Sum of weight2", Excel.XlConsolidationFunction.xlSum)
                With oSheet.PivotTables("PivotTableParetoCategory").PivotFields("Sum of weight2")
                    .Calculation = Excel.XlPivotFieldCalculation.xlPercentOfColumn
                    .NumberFormat = "0.00%"
                End With
                With oSheet.PivotTables("PivotTableParetoCategory").DataPivotField
                    .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                    .Position = 1
                End With

                'Create Pareto Category Full

                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C235", "PivotTableParetoCategoryFull", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                With oSheet.PivotTables("PivotTableParetoCategoryFull").PivotFields("years")
                    .Orientation = Excel.XlPivotFieldOrientation.xlPageField
                    .CurrentPage = Format(currentmonth, "yyyy")
                End With

                oSheet.PivotTables("PivotTableParetoCategoryFull").PivotFields("catissues1").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTableParetoCategoryFull").AddDataField(oSheet.PivotTables("PivotTableParetoCategoryFull").PivotFields("weight"), "Sum of weight", Excel.XlConsolidationFunction.xlSum)


            End If

            If Pivot15 > 1 Then
                'Create New Pivot Table Purchasing Group
                oWb.Worksheets(7).Select()
                oSheet = oWb.Worksheets(7)
                oSheet.Name = "SSL_By_Purchasing_Group"
                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R8C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                My.Application.DoEvents()

                '******************
                'oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Orientation = xlRowField
                'oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Caption = "Shipdate"
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup"), " weight", xlSum
                'oSheet.PivotTables("PivotTable1").PivotFields(" weight").NumberFormat = "0"

                'oSheet.Range("G7").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)

                oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Caption = "Shipdate"
                oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("sslnet"), " weight", xlSum
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("PCTSSLNET"), " %SSLNET", Excel.XlConsolidationFunction.xlSum)

                oSheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("Years").Position = 1
                oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Position = 1

                'oSheet.PivotTables("PivotTable1").PivotFields("Shipdate").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "PCT", "='sasl'/ordercount", True

                'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "Target", myTarget, True

                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSL"), " %SSL", xlSum

                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("deliveredqty"), " Sum of deliveryqty", xlSum

                'oSheet.PivotTables("PivotTable1c").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("TargetSASL"), " Target SASL", xlSum
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSL"), " Target SSL", xlSum
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSLNET"), " Target SSL NET", xlSum
                'oSheet.PivotTables("PivotTable1c").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSASL"), " %SASL", xlSum
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSL"), " %SSL", xlSum
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSLNET"), " %SSLNET", xlSum
                'oSheet.PivotTables("PivotTable1c").PivotFields(" Target SASL").NumberFormat = "0%"
                'oSheet.PivotTables("PivotTable1").PivotFields(" Target SSL").NumberFormat = "0%"
                'oSheet.PivotTables("PivotTable1").PivotFields(" Target SSL NET").NumberFormat = "0%"
                'oSheet.PivotTables("PivotTable1c").PivotFields(" %SASL").NumberFormat = "0.0%"
                'oSheet.PivotTables("PivotTable1").PivotFields(" %SSL").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable1").PivotFields(" %SSLNET").NumberFormat = "0.0%"
                'oSheet.PivotTables("PivotTable1").PivotFields(" Sum of deliveryqty").NumberFormat = "#,##0"

                'oSheet.PivotTables("PivotTable1").DataPivotField.Orientation = xlColumnField
                'oSheet.PivotTables("PivotTable1").DataPivotField.Position = 1
                For Each p In oSheet.PivotTables("PivotTable1").PivotFields("Years").PivotItems
                    If p.value = startdate.Year.ToString Then
                        oSheet.PivotTables("PivotTable1").PivotFields("Years").CurrentPage = startdate.Year
                        Exit For
                    End If

                Next

                oSheet.PivotTables("PivotTable1").DisplayErrorString = True
            End If
            If Pivot15 > 1 Then
                'SSL Purchasing Group
                oWb.Worksheets(8).Select()
                oSheet = oWb.Worksheets(8)
                oSheet.Name = "SSL_By_Purchasing_SBU"
                oWb.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R8C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
                My.Application.DoEvents()

                '******************
                'oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Orientation = xlRowField
                'oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Caption = "Shipdate"
                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup"), " weight", xlSum
                'oSheet.PivotTables("PivotTable1").PivotFields(" weight").NumberFormat = "0"

                'oSheet.Range("G7").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)

                oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
                oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Caption = "Shipdate"

                'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("sslnet"), " weight", xlSum
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("weight"), " Weight", Excel.XlConsolidationFunction.xlSum)
                oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("PCTSSLNET"), " %SSLNET", Excel.XlConsolidationFunction.xlSum)


                oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
                oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Orientation = Excel.XlPivotFieldOrientation.xlPageField

                oSheet.PivotTables("PivotTable1").PivotFields("sbunamesp").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                oSheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlRowField
                With oSheet.PivotTables("PivotTable1").DataPivotField
                    .Orientation = Excel.XlPivotFieldOrientation.xlRowField
                    .Position = 2
                End With

                oSheet.PivotTables("PivotTable1").PivotFields(" %SSLNET").NumberFormat = "0.0%"
                oSheet.PivotTables("PivotTable1").PivotFields(" Weight").NumberFormat = "#,##0.0"
                oSheet.PivotTables("PivotTable1").DisplayErrorString = True
            End If


            '=COUNTA(PivotTables!$FA:$FA)-4>1
            If CurrYear Then
                'oWB.Names.Add Name:="MonthRangeSASLY", RefersToR1C1:="=OFFSET(PivotTables!R8C16,0,0,COUNTA(PivotTables!C16)-5,1)"
                oWb.Names.Add(Name:="MonthRangeSSLNETY", RefersToR1C1:="=OFFSET(PivotTables!R8C16,0,0,COUNTA(PivotTables!C16)-6,1)")
                'oWB.Names.Add Name:="OrderCountSASLY", RefersToR1C1:="=OFFSET(PivotTables!R8C17,0,0,COUNTA(PivotTables!C16)-5,1)"
                oWb.Names.Add(Name:="OrderCountSSLNETY", RefersToR1C1:="=OFFSET(PivotTables!R8C17,0,0,COUNTA(PivotTables!C16)-6,1)")
                'oWB.Names.Add Name:="TargetSASLY", RefersToR1C1:="=OFFSET(PivotTables!R8C19,0,0,COUNTA(PivotTables!C16)-5,1)"
                oWb.Names.Add(Name:="TargetSSLNETY", RefersToR1C1:="=OFFSET(PivotTables!R8C20,0,0,COUNTA(PivotTables!C16)-6,1)")
                oWb.Names.Add(Name:="TargetSSLY", RefersToR1C1:="=OFFSET(PivotTables!R8C19,0,0,COUNTA(PivotTables!C16)-6,1)")
                'oWB.Names.Add Name:="PercentageSASLY", RefersToR1C1:="=OFFSET(PivotTables!R8C21,0,0,COUNTA(PivotTables!C16)-5,1)"
                oWb.Names.Add(Name:="PercentageSSLNETY", RefersToR1C1:="=OFFSET(PivotTables!R8C22,0,0,COUNTA(PivotTables!C16)-6,1)")
                oWb.Names.Add(Name:="PercentageSSLY", RefersToR1C1:="=OFFSET(PivotTables!R8C21,0,0,COUNTA(PivotTables!C16)-6,1)")
            End If

            Dim fullyear As Boolean
            If CurrYear And LastYear Then
                fullyear = True
                'oWB.Names.Add Name:="MonthRangeSASLF", RefersToR1C1:="=OFFSET(PivotTables!R9C29,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="MonthRangeSSLNETF", RefersToR1C1:="=OFFSET(PivotTables!R9C29,0,0,COUNTA(PivotTables!C29)-4,1)")
                'oWB.Names.Add Name:="OrderCountSASLYmin1F", RefersToR1C1:="=OFFSET(PivotTables!R9C30,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="OrderCountSSLNETYmin1F", RefersToR1C1:="=OFFSET(PivotTables!R9C30,0,0,COUNTA(PivotTables!C29)-4,1)")
                'oWB.Names.Add Name:="OrderCountSASLYF", RefersToR1C1:="=OFFSET(PivotTables!R9C31,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="OrderCountSSLNETYF", RefersToR1C1:="=OFFSET(PivotTables!R9C31,0,0,COUNTA(PivotTables!C29)-4,1)")
                'oWB.Names.Add Name:="TargetSASLYF", RefersToR1C1:="=OFFSET(PivotTables!R9C32,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="TargetSSLNETYF", RefersToR1C1:="=OFFSET(PivotTables!R9C35,0,0,COUNTA(PivotTables!C29)-4,1)")
                oWb.Names.Add(Name:="TargetSSLYF", RefersToR1C1:="=OFFSET(PivotTables!R9C33,0,0,COUNTA(PivotTables!C29)-4,1)")
                'oWB.Names.Add Name:="PercentageSASLYmin1F", RefersToR1C1:="=OFFSET(PivotTables!R9C36,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="PercentageSSLNETYmin1F", RefersToR1C1:="=OFFSET(PivotTables!R9C38,0,0,COUNTA(PivotTables!C29)-4,1)")
                'oWB.Names.Add Name:="PercentageSASLYF", RefersToR1C1:="=OFFSET(PivotTables!R9C37,0,0,COUNTA(PivotTables!C29)-3,1)"
                oWb.Names.Add(Name:="PercentageSSLNETYF", RefersToR1C1:="=OFFSET(PivotTables!R9C39,0,0,COUNTA(PivotTables!C29)-4,1)")
                oWb.Names.Add(Name:="PercentageSSLYmin1F", RefersToR1C1:="=OFFSET(PivotTables!R9C36,0,0,COUNTA(PivotTables!C29)-4,1)")
                oWb.Names.Add(Name:="PercentageSSLYF", RefersToR1C1:="=OFFSET(PivotTables!R9C37,0,0,COUNTA(PivotTables!C29)-4,1)")
            End If


            oWb.Names.Add(Name:="MonthRangeFSL", RefersToR1C1:="=OFFSET(PivotTables!R8C62,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="OrderCount", RefersToR1C1:="=OFFSET(PivotTables!R8C63,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="FSL", RefersToR1C1:="=OFFSET(PivotTables!R8C64,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="FSSL", RefersToR1C1:="=OFFSET(PivotTables!R8C65,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="FSSLN", RefersToR1C1:="=OFFSET(PivotTables!R8C66,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="TargetFSL", RefersToR1C1:="=OFFSET(PivotTables!R8C67,0,0,COUNTA(PivotTables!C61)-4,1)")
            oWb.Names.Add(Name:="TargetFSSL", RefersToR1C1:="=OFFSET(PivotTables!R8C68,0,0,COUNTA(PivotTables!C61)-4,1)")

            oWb.Names.Add(Name:="MonthRangeOPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C81,0,0,COUNTA(PivotTables!C80)-3,1)")
            oWb.Names.Add(Name:="AverageLTOPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C82,0,0,COUNTA(PivotTables!C80)-3,1)")
            oWb.Names.Add(Name:="Percentage04OPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C85,0,0,COUNTA(PivotTables!C80)-3,1)")
            oWb.Names.Add(Name:="PercentageNotConfOPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C86,0,0,COUNTA(PivotTables!C80)-3,1)")
            oWb.Names.Add(Name:="TargetOPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C87,0,0,COUNTA(PivotTables!C80)-3,1)")


            oWb.Names.Add(Name:="MonthRangeIPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C107,0,0,COUNTA(PivotTables!C106)-4,1)")
            oWb.Names.Add(Name:="AverageLTIPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C108,0,0,COUNTA(PivotTables!C106)-4,1)")
            oWb.Names.Add(Name:="PercentageIPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C111,0,0,COUNTA(PivotTables!C106)-4,1)")
            oWb.Names.Add(Name:="TargetIPLT", RefersToR1C1:="=OFFSET(PivotTables!R8C112,0,0,COUNTA(PivotTables!C106)-4,1)")


            oWb.Names.Add(Name:="MonthRangeSCR", RefersToR1C1:="=OFFSET(PivotTables!R8C132,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="PercentageSCR", RefersToR1C1:="=OFFSET(PivotTables!R8C133,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="SCRHigh", RefersToR1C1:="=OFFSET(PivotTables!R8C134,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="SCRLow", RefersToR1C1:="=OFFSET(PivotTables!R8C135,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="CountNbOfOrder", RefersToR1C1:="=OFFSET(PivotTables!R8C136,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="CustomerLT", RefersToR1C1:="=OFFSET(PivotTables!R8C137,0,0,COUNTA(PivotTables!C132)-4,1)")
            oWb.Names.Add(Name:="SupplierLT", RefersToR1C1:="=OFFSET(PivotTables!R8C138,0,0,COUNTA(PivotTables!C132)-4,1)")



            'Create Chart
            oWb.Worksheets(1).Select()
            oSheet = oWb.Worksheets(1)

            If Department = DepartmentEnum.FinishedGoods Then
                oSheet.Cells(4, 1) = "Logistics Indicators for Finished Goods"
            Else
                oSheet.Cells(4, 1) = "Logistics Indicators for Components"
            End If
            '    If addMacro Then
            'oSheet.Cells(5, 1) = "As " & Format(Date, "dd MMM yyyy")
            '    Else
            oSheet.Cells(5, 1) = myVendorname & " as " & Format(Date.Today, "dd MMM yyyy")
            '    End If


            'oSheet.Cells(10, 8) = "1/1/" & DTPicker1.year - 1 & "-" & DTPicker2.Day & "/" & DTPicker2.month & "/" & DTPicker2.year - 1
            oSheet.Range("sslperiod").Value = "1/1/" & startdate.Year - 1 & "-" & enddate.Day & "/" & enddate.Month & "/" & enddate.Year - 1

            If whiteSupplier Then
                oSheet.Range("sgreenlabel").Value = ">=90%"
                oSheet.Range("syellowlabel").Value = "80-89%"
                oSheet.Range("sredlabel").Value = "<=79%"

                oSheet.Range("fgreenlabel").Value = ">=90%"
                oSheet.Range("fyellowlabel").Value = "80-89%"
                oSheet.Range("fredlabel").Value = "<=79%"
            End If

            myChart = oSheet.ChartObjects("SASL").Chart
            If Pivot15 > 1 Then
                If fullyear Then
                    myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(1).Values = "='PivotTables'!OrderCountSSLNETYF"
                    myChart.SeriesCollection(1).AxisGroup = 1

                    myChart.SeriesCollection(2).Name = "=""Target SSL gross"""
                    myChart.SeriesCollection(2).Values = "='PivotTables'!TargetSSLYF"
                    myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(2).AxisGroup = 2

                    myChart.SeriesCollection(3).Name = "=""Target SSL net"""
                    myChart.SeriesCollection(3).Values = "='PivotTables'!TargetSSLNETYF"
                    myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(3).AxisGroup = 2


                    myChart.SeriesCollection(4).Values = "='PivotTables'!PercentageSSLYF"
                    myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(4).AxisGroup = 2

                    myChart.SeriesCollection(5).Values = "='PivotTables'!PercentageSSLNETYF"
                    myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(5).AxisGroup = 2

                    myChart.SeriesCollection(6).Values = "='PivotTables'!OrderCountSSLNETYmin1F"
                    myChart.SeriesCollection(6).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(6).AxisGroup = 1

                    myChart.SeriesCollection(7).Values = "='PivotTables'!PercentageSSLYmin1F"
                    myChart.SeriesCollection(7).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(7).AxisGroup = 2

                    myChart.SeriesCollection(8).Values = "='PivotTables'!PercentageSSLNETYmin1F"
                    myChart.SeriesCollection(8).XValues = "='PivotTables'!MonthRangeSSLNETF"
                    myChart.SeriesCollection(8).AxisGroup = 2

                    'Delete from the highest
                    myChart.SeriesCollection(7).Delete()
                    myChart.SeriesCollection(4).Delete()
                    myChart.SeriesCollection(2).Delete()



                    oSheet.Range("ymin1").Value = startdate.Year - 1
                    oSheet.Range("sslnetymin1").Value = "=IFERROR(GETPIVOTDATA("" %SSLNET"",PivotTables!$A$6,""Years""," & startdate.Year - 1 & "),""N/A"")"
                    oSheet.Range("ymin1ytd").Value = startdate.Year - 1 & " YTD"
                    oSheet.Range("sslnetymin1ytd").Value = "=IFERROR(GETPIVOTDATA("" SSLNET-RT"",PivotTables!$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & ")/GETPIVOTDATA("" W-RT"",PivotTables!$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & "),""N/A"")"
                    oSheet.Range("yytd").Value = startdate.Year & " YTD"
                    oSheet.Range("sslnetyytd").Value = "=IFERROR(GETPIVOTDATA("" sum of sslnet"",PivotTables!$A$6,""Years""," & startdate.Year & ")/GETPIVOTDATA("" weight"",PivotTables!$A$6,""Years""," & startdate.Year & "),""N/A"")"
                    Dim obj As Object
                    obj = oSheet.Range("sslnetsmiley")
                    obj.Value = "=char(74 + PivotTables!" & smileySASL.Address & ")"

                ElseIf CurrYear Then
                    'myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeSASLY"
                    myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeSSLNETY"
                    'myChart.SeriesCollection(1).Values = "='PivotTables'!OrderCountSASLY"
                    myChart.SeriesCollection(1).Values = "='PivotTables'!OrderCountSSLNETY"

                    'myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeSASLY"
                    myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeSSLNETY"
                    'myChart.SeriesCollection(2).Values = "='PivotTables'!TargetSASLY"
                    myChart.SeriesCollection(2).Values = "='PivotTables'!TargetSSLNETY"

                    'myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeSASLY"
                    myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeSSLNETY"
                    myChart.SeriesCollection(3).Values = "='PivotTables'!TargetSSLY"

                    'myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeSASLY"
                    myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeSSLNETY"
                    'myChart.SeriesCollection(4).Values = "='PivotTables'!PercentageSASLY"
                    myChart.SeriesCollection(4).Values = "='PivotTables'!PercentageSSLNETY"

                    'myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeSASLY"
                    myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeSSLNETY"
                    myChart.SeriesCollection(5).Values = "='PivotTables'!PercentageSSLY"

                    '            oSheet.Cells(9, 10).Value = DTPicker1.year & " YTD"
                    '            oSheet.Cells(11, 10).Value = "=IFERROR(GETPIVOTDATA("" sum of sasl"",PivotTables!$A$6,""Years""," & DTPicker1.year & ")/GETPIVOTDATA("" weight"",PivotTables!$A$6,""Years""," & DTPicker1.year & "),""N/A"")"
                    '            oSheet.Cells(12, 10).Value = "=IFERROR(GETPIVOTDATA("" Sum of ssl"",PivotTables!$A$6,""Years""," & DTPicker1.year & ")/GETPIVOTDATA("" weight"",PivotTables!$A$6,""Years""," & DTPicker1.year & "),""N/A"")"
                    oSheet.Range("yytd").Value = startdate.Year & " YTD"
                    oSheet.Range("sslnetyytd").Value = "=IFERROR(GETPIVOTDATA("" sum of sslnet"",PivotTables!$A$6,""Years""," & startdate.Year & ")/GETPIVOTDATA("" weight"",PivotTables!$A$6,""Years""," & startdate.Year & "),""N/A"")"
                    'oSheet.Range("sslgrossyytd").Value = "=IFERROR(GETPIVOTDATA("" Sum of ssl"",PivotTables!$A$6,""Years""," & DTPicker1.year & ")/GETPIVOTDATA("" weight"",PivotTables!$A$6,""Years""," & DTPicker1.year & "),""N/A"")"

                    'oSheet.Cells(11, 10).Value = "=IFERROR(GETPIVOTDATA("" SASL-RT"",PivotTables!$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year & ")/GETPIVOTDATA("" W-RT"",PivotTables!$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year & "),""N/A"")"
                    'oSheet.Cells(12, 10).Value = "=IFERROR(GETPIVOTDATA("" SSL-RT"",PivotTables!$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year & ")/GETPIVOTDATA("" W-RT"",PivotTables!$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year & "),""N/A"")"

                    myChart.SeriesCollection(8).Delete()
                    myChart.SeriesCollection(7).Delete()
                    myChart.SeriesCollection(6).Delete()
                    Dim obj As Object
                    obj = oSheet.Cells(11, 2)
                    obj.Value = "=char(74 + PivotTables!" & smileySASL.Address & " + PivotTables!" & smileySSL.Address & ")"
                ElseIf LastYear Then
                    myChart.Parent.Delete()

                End If

            Else
                myChart.Parent.Delete()
            End If



            myChart = oSheet.ChartObjects("FSL").Chart
            If Pivot2 > 1 Then
                myChart = oSheet.ChartObjects("FSL").Chart
                myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeFSL"
                myChart.SeriesCollection(1).Values = "='PivotTables'!OrderCount"




                myChart.SeriesCollection(2).Values = "='PivotTables'!TargetFSSL"
                myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeFSL"

                myChart.SeriesCollection(3).Values = "='PivotTables'!TargetFSL"
                myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeFSL"

                'myChart.SeriesCollection(4).Values = "='PivotTables'!FSL"
                'myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeFSL"

                myChart.SeriesCollection(4).Values = "='PivotTables'!FSSL"
                myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeFSL"

                myChart.SeriesCollection(5).Values = "='PivotTables'!FSSLN"
                myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeFSL"


                myChart.SeriesCollection(4).Delete()
                myChart.SeriesCollection(2).Delete()

                '****Loop for avail week + i
                Dim fsllabel As String = String.Empty
                Dim fssllabel As String = String.Empty

                For i = 0 To 7

                    newdate = fslstartdate.AddDays(8 * (7 - i))
                    oSheet.Range("fsslnvalue").Value = "=IFERROR(GETPIVOTDATA("" FSSLN-RT"",PivotTables!$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",PivotTables!$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & ")),""N/A"")"
                    If Not oSheet.Range("fsslnvalue").Value.ToString = "N/A" Then
                        fsllabel = "Average FSSLN WK - WK +" & (8 - i) & "="
                        fssllabel = "Average FSSL W to W +" & (8 - i) & "="
                        Exit For
                    End If
                Next
                'oSheet.Cells(10, 18) = fsllabel
                'oSheet.Cells(11, 18) = fssllabel
                oSheet.Range("labelfssln").Value = fsllabel
                'oSheet.Range("labelfssl") = fssllabel
                '******

                'Value = "=GETPIVOTDATA("" FSL-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))"
                '   Set obj = oSheet.Cells(myRow, 68)
                '   obj.Value = "=GETPIVOTDATA("" FSSL-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))/GETPIVOTDATA("" OC-RT"",$BA$6,""Current Confirmed ETD"",DATE(" & dateformatcomma(newdate) & "))"




                '        myChart.SeriesCollection(5).Values = "='PivotTables'!FSLplus8"
                '        myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeFSL"
                '
                '        myChart.SeriesCollection(6).Values = "='PivotTables'!FSSLplus8"
                '        myChart.SeriesCollection(6).XValues = "='PivotTables'!MonthRangeFSL"
                Dim obj As Object
                obj = oSheet.Range("fsslnsmiley") 'oSheet.Cells(11, 14)
                'obj.Value = "=char(74 + PivotTables!" & smileyFSL.Address & " + PivotTables!" & smileyFSSL.Address & ")"
                obj.Value = "=char(74 + PivotTables!" & smileyFSSLN.Address & ")"
                'Set obj = oSheet.Range("fsslsmiley") 'oSheet.Cells(11, 15)
                'obj.Value = "=char(74 + PivotTables!" & smileyFSSL.Address & ")"
                'oSheet.Cells(11, 14) = "=char(74 + " & "PivotTables!" & smileyFSL.Address & " + PivotTables!" & smileyFSSL.Address & ")"
            Else
                'oSheet.Cells(11, 14) = "L"
                'oSheet.Range("smileyfssln") = "L"
                myChart.Parent.Delete()
            End If
            myChart = oSheet.ChartObjects("OPLT").Chart
            If Pivot3 > 1 Then

                myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeOPLT"
                myChart.SeriesCollection(1).Values = "='PivotTables'!AverageLTOPLT"
                myChart.SeriesCollection(2).Values = "='PivotTables'!Percentage04OPLT"
                myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeOPLT"
                myChart.SeriesCollection(3).Values = "='PivotTables'!TargetOPLT"
                myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeOPLT"
            Else
                'myChart.Parent.Delete
            End If

            'Set obj = oSheet.Range("paretocurrentdate")
            'obj.Value = DTPicker3

            If myPareto Then
                For i = 1 To ParetoBar

                    myChart = oSheet.ChartObjects("PARETO2").Chart
                    myChart.SeriesCollection.NewSeries()
                    myChart.SeriesCollection(i).XValues = "='PivotTables'!BarAxis1"
                    myChart.SeriesCollection(i).Values = "='PivotTables'!PBar" & i
                    myChart.SeriesCollection(i).Name = "='PivotTables'!PBarName" & i
                    'Assign Color
                    Select Case myChart.SeriesCollection(i).Name
                        Case "CUSTOMER ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 65535
                        Case "FORWARDER ISSUE"
                            myChart.SeriesCollection(i).Interior.Color = 49407
                        Case "INTERNAL ISSUES"
                            myChart.SeriesCollection(i).Interior.Color = 10498160
                        Case "SUPPLIER ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 255
                        Case "OTHERS"
                            myChart.SeriesCollection(i).Interior.Color = 5287936
                        Case "SIS ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 6250335
                        Case "CUSTOMS"
                            myChart.SeriesCollection(i).Interior.Color = 5263615
                    End Select
                Next

                myChart.Legend.Top = 3
                myChart.Legend.Left = 0
                myChart.Legend.Width = 600
                myChart.PlotArea.Top = 50
                myChart.PlotArea.Height = 600
                myChart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Orientation = 45



            Else
                'myChart.Parent.Delete
            End If

            If myParetoGrp Then
                For i = 1 To ParetoBarGrp
                    myChart = oSheet.ChartObjects("PARETO1").Chart
                    myChart.SeriesCollection.NewSeries()
                    myChart.SeriesCollection(i).XValues = "='PivotTables'!BarAxisGrp1"
                    myChart.SeriesCollection(i).Values = "='PivotTables'!PBarGrp" & i
                    myChart.SeriesCollection(i).Name = "='PivotTables'!PBarNameGrp" & i
                    'Assign Color
                    Select Case myChart.SeriesCollection(i).Name
                        Case "CUSTOMER ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 65535
                        Case "FORWARDER ISSUE"
                            myChart.SeriesCollection(i).Interior.Color = 49407
                        Case "INTERNAL ISSUES"
                            myChart.SeriesCollection(i).Interior.Color = 10498160
                        Case "SUPPLIER ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 255
                        Case "OTHERS"
                            myChart.SeriesCollection(i).Interior.Color = 5287936
                        Case "SIS ISSUE (S)"
                            myChart.SeriesCollection(i).Interior.Color = 6250335
                        Case "CUSTOMS"
                            myChart.SeriesCollection(i).Interior.Color = 5263615
                    End Select
                Next

                myChart.Legend.Top = 3
                myChart.Legend.Left = 0
                myChart.Legend.Width = 600
                myChart.PlotArea.Top = 50
                myChart.PlotArea.Height = 600
                myChart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Orientation = 45
                myChart.Axes(Excel.XlAxisType.xlCategory).TickLabels.Font.Bold = True

            Else
                'myChart.Parent.Delete
            End If

            'Total Group
            oSheet.Range("total1").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""CUSTOMER ISSUE (S)"")),"""",GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""CUSTOMER ISSUE (S)""))"
            oSheet.Range("total2").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""FORWARDER ISSUE"")),"""",GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""FORWARDER ISSUE""))"
            oSheet.Range("total3").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""SUPPLIER ISSUE (S)"")),"""",GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""SUPPLIER ISSUE (S)""))"
            oSheet.Range("total4").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""INTERNAL ISSUES"")),"""",GETPIVOTDATA(""Sum of weight"",PivotTables!$HA$6,""catissues1"",""INTERNAL ISSUES""))"

            oSheet.Range("pct_1").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""CUSTOMER ISSUE (S)"")),"""",GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""CUSTOMER ISSUE (S)""))"
            oSheet.Range("pct_2").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""FORWARDER ISSUE"")),"""",GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""FORWARDER ISSUE""))"
            oSheet.Range("pct_3").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""SUPPLIER ISSUE (S)"")),"""",GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""SUPPLIER ISSUE (S)""))"
            oSheet.Range("pct_4").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""INTERNAL ISSUES"")),"""",GETPIVOTDATA(""Sum of weight2"",PivotTables!$HA$6,""catissues1"",""INTERNAL ISSUES""))"

            oSheet.Range("totalweight").Value = "=IF(ISERROR(GETPIVOTDATA(""Sum of weight"",PivotTables!$IA$6)),"""",GETPIVOTDATA(""Sum of weight"",PivotTables!$IA$6))"

            'On Error GoTo errorchart4
            myChart = oSheet.ChartObjects("OPLT").Chart
            If Pivot4 > 1 Then

                myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeIPLT"
                myChart.SeriesCollection(4).Values = "='PivotTables'!AverageLTIPLT"
                myChart.SeriesCollection(5).Values = "='PivotTables'!PercentageIPLT"
                myChart.SeriesCollection(5).XValues = "='PivotTables'!MonthRangeIPLT"
                myChart.SeriesCollection(6).Values = "='PivotTables'!TargetIPLT"
                myChart.SeriesCollection(6).XValues = "='PivotTables'!MonthRangeIPLT"
                If Pivot3 = 1 Then
                    For i = 1 To 3
                        myChart.SeriesCollection(1).Delete()
                    Next i
                Else
                    'Set obj = oSheet.Cells(51, 14)
                    'obj.Value = "=char(74 + PivotTables!" & smileyOPLT.Address & " + PivotTables!" & smileyIPLT.Address & ")"
                    Dim obj As Object
                    obj = oSheet.Range("opltsmiley") 'oSheet.Cells(51, 14)
                    obj.Value = "=char(74 + PivotTables!" & smileyOPLT.Address & ")"
                    obj = oSheet.Range("ipltsmiley") 'oSheet.Cells(51, 16)
                    obj.Value = "=char(74 + PivotTables!" & smileyIPLT.Address & ")"
                End If
            Else
                If Pivot3 > 1 Then
                    For i = 4 To 6
                        myChart.SeriesCollection(4).Delete()
                    Next i
                End If

            End If
            If Pivot3 = 1 And Pivot4 = 1 Then
                myChart.Parent.Delete()
            End If

            '    Set myChart = oSheet.ChartObjects("IPLT").Chart
            '    myChart.Parent.Delete

            myChart = oSheet.ChartObjects("SCR").Chart
            If SCRChart Then

                myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeSCR"
                myChart.SeriesCollection(1).Values = "='PivotTables'!CountNbOfOrder"

                myChart.SeriesCollection(2).Values = "='PivotTables'!PercentageSCR"
                myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeSCR"

                myChart.SeriesCollection(3).Values = "='PivotTables'!SCRHigh"
                myChart.SeriesCollection(3).XValues = "='PivotTables'!MonthRangeSCR"

                'myChart.SeriesCollection(4).Values = "='PivotTables'!SCRLow"
                'myChart.SeriesCollection(4).XValues = "='PivotTables'!MonthRangeSCR"
                Dim obj As Object
                obj = oSheet.Range("scrsmiley") 'oSheet.Cells(51, 2)
                obj.Value = "=char(74 + PivotTables!" & smileySCR.Address & ")"

            Else
                myChart.Parent.Delete()
            End If

            myChart = oSheet.ChartObjects("LT").Chart
            If SCRChart Then

                myChart.SeriesCollection(1).XValues = "='PivotTables'!MonthRangeSCR"
                myChart.SeriesCollection(1).Values = "='PivotTables'!CustomerLT"

                myChart.SeriesCollection(2).XValues = "='PivotTables'!MonthRangeSCR"
                myChart.SeriesCollection(2).Values = "='PivotTables'!SupplierLT"

            Else
                myChart.Parent.Delete()
            End If




            'StatusBar1.Panels(1).Text = "Processing Time: " & Format(DateAdd("s", DateDiff("s", time1, Now), "00:00:00"), "HH:mm:ss") & " Done!"

            oWb.Worksheets(1).Select()
            Dim myfilename As String
            myfilename = sr.filename & IIf(Department = DepartmentEnum.FinishedGoods, "FG-", "COMP-") & String.Format("{0:yyyyMMdd}", startdate) & "-" & String.Format("{0:yyyyMMdd}", enddate) & ".xlsx"
            myfilename = Replace(myfilename, "&", "")
            myfilename = ValidateFileName(myfilename)

            'remove connection
            'Call RemoveConnection(oWB)
            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()

            'If addMacro Then
            '    oWb.SaveAs(myFilename, xlOpenXMLWorkbookMacroEnabled) ' mySelectedPath & "\" & "Scoreboard1-" & List1(MyIndex) & "-" & DateFormatShort(DTPicker1) & "-" & DateFormatShort(DTPicker2) & IIf(addMacro, ".xlsm", ".xlsx")
            'Else
            '    oWb.SaveAs(myFilename) ' mySelectedPath & "\" & "Scoreboard1-" & List1(MyIndex) & "-" & DateFormatShort(DTPicker1) & "-" & DateFormatShort(DTPicker2) & ".xlsx"
            'End If
            doBackground1.ProgressReport(1, String.Format("Saving File ...{0}", myfilename))
            oWb.SaveAs(myfilename)





            'FileName = ValidateFileName(FileName, FileName & "\" & String.Format("Scoreboard {3}-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day), sr.dr.Item(0).ToString))
            'ProgressReport(5, "Done ")
            'ProgressReport(2, "Saving File ...")
            'oWb.SaveAs(FileName)
            'ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
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

        Return True
    End Function

    Sub FormattingReport()
        Throw New NotImplementedException
    End Sub

    Private Function CreatePivotCurrentYear(ByVal oWB As Excel.Workbook, ByVal oSheet As Excel.Worksheet, ByVal myTarget As String, ByVal myTargetSSL As String) As Boolean
        Dim myret As Boolean
        Try
            oWB.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRangeOTDSASL").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
            My.Application.DoEvents()

            '******************
            oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1").PivotFields("shipdate").Caption = "Shipdate"
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("weight"), " weight", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").PivotFields(" weight").NumberFormat = "0"

            oSheet.Range("A7").Group(Start:=True, End:=True, Periods:={False, False, False, False, True, False, True})
            oSheet.PivotTables("PivotTable1").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField

            oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Position = 1
            oSheet.PivotTables("PivotTable1").PivotFields("Years").Subtotals = {True, False, False, False, False, False, False, False, False, False, False, False}
            oSheet.PivotTables("PivotTable1").PivotFields("Shipdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("PCTSSL", "=ssl /weight", True)
            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("PCTSSLNET", "=sslnet /weight", True)
            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("TargetSSL", myTarget, True)
            oSheet.PivotTables("PivotTable1").CalculatedFields.Add("TargetSSLNET", myTargetSSL, True)

            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("ssl"), " Sum of ssl", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("sslnet"), " Sum of sslnet", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").PivotFields(" Sum of ssl").NumberFormat = "0"
            oSheet.PivotTables("PivotTable1").PivotFields(" Sum of sslnet").NumberFormat = "0"
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("deliveredqty"), " Sum of deliveryqty", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("TargetSSL"), " Target SSL", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("TargetSSLNET"), " Target SSL NET", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("PCTSSL"), " %SSL", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("PCTSSLNET"), " %SSLNET", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("weight"), " W-RT", Excel.XlConsolidationFunction.xlSum)
            With oSheet.PivotTables("PivotTable1").PivotFields(" W-RT")
                .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                .BaseField = "Shipdate"
                .NumberFormat = "0"
            End With
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("ssl"), " SSL-RT", Excel.XlConsolidationFunction.xlSum)
            With oSheet.PivotTables("PivotTable1").PivotFields(" SSL-RT")
                .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                .BaseField = "Shipdate"
                .NumberFormat = "0"
            End With
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("sslnet"), " SSLNET-RT", Excel.XlConsolidationFunction.xlSum)
            With oSheet.PivotTables("PivotTable1").PivotFields(" SSLNET-RT")
                .Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
                .BaseField = "Shipdate"
                .NumberFormat = "0"
            End With

            oSheet.PivotTables("PivotTable1").PivotFields(" Target SSL").NumberFormat = "0%"
            oSheet.PivotTables("PivotTable1").PivotFields(" Target SSL NET").NumberFormat = "0%"
            oSheet.PivotTables("PivotTable1").PivotFields(" %SSL").NumberFormat = "0.0%"
            oSheet.PivotTables("PivotTable1").PivotFields(" %SSLNET").NumberFormat = "0.0%"
            oSheet.PivotTables("PivotTable1").PivotFields(" Sum of deliveryqty").NumberFormat = "#,##0"

            oSheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            oSheet.PivotTables("PivotTable1").DataPivotField.Position = 1
            oSheet.PivotTables("PivotTable1").RowGrand = False
            oSheet.PivotTables("PivotTable1").ColumnGrand = False

            '************* Create Current Year
            oWB.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C16", "PivotTable1c", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
            My.Application.DoEvents()

            '******************
            oSheet.PivotTables("PivotTable1c").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1c").PivotFields("shipdate").Caption = "Shipdate"
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("weight"), " weight", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1c").PivotFields(" weight").NumberFormat = "0"

            'oSheet.Range("G7").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)
            oSheet.PivotTables("PivotTable1c").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1c").PivotFields("Years").Position = 1
            oSheet.PivotTables("PivotTable1c").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1c").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1c").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1c").PivotFields("vendorname").Position = 1

            oSheet.PivotTables("PivotTable1c").PivotFields("Shipdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
            'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "PCT", "='sasl'/ordercount", True

            'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "Target", myTarget, True
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("deliveredqty"), " Sum of deliveryqty", Excel.XlConsolidationFunction.xlSum)

            'oSheet.PivotTables("PivotTable1c").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("TargetSASL"), " Target SASL", xlSum
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSL"), " Target SSL", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSLNET"), " Target SSL NET", Excel.XlConsolidationFunction.xlSum)
            'oSheet.PivotTables("PivotTable1c").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSASL"), " %SASL", xlSum
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSL"), " %SSL", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1c").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSLNET"), " %SSLNET", Excel.XlConsolidationFunction.xlSum)
            'oSheet.PivotTables("PivotTable1c").PivotFields(" Target SASL").NumberFormat = "0%"
            oSheet.PivotTables("PivotTable1c").PivotFields(" Target SSL").NumberFormat = "0%"
            oSheet.PivotTables("PivotTable1c").PivotFields(" Target SSL NET").NumberFormat = "0%"
            'oSheet.PivotTables("PivotTable1c").PivotFields(" %SASL").NumberFormat = "0.0%"
            oSheet.PivotTables("PivotTable1c").PivotFields(" %SSL").NumberFormat = "0.0%"
            oSheet.PivotTables("PivotTable1c").PivotFields(" %SSLNET").NumberFormat = "0.0%"
            oSheet.PivotTables("PivotTable1c").PivotFields(" Sum of deliveryqty").NumberFormat = "#,##0"

            oSheet.PivotTables("PivotTable1c").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            oSheet.PivotTables("PivotTable1c").DataPivotField.Position = 1
            oSheet.PivotTables("PivotTable1c").PivotFields("Years").CurrentPage = startdate.Year

            Dim obj As Object


            obj = oSheet.Cells(1, 15)
            Dim myRow As Integer
            obj.FormulaR1C1 = "=COUNTA(C16:C16)"
            If obj.Value > 4 Then
                myRow = obj.Value - 4 + 8
                obj = oSheet.Cells(myRow, 23)
                obj.Value = "=GETPIVOTDATA("" %SSLNET"",$P$6)"
                smileySASL = oSheet.Cells(myRow + 1, 23)
                smileySASL.Value = "=IF(GETPIVOTDATA("" %SSLNET"",$P$6) >= 0.95,0,if(GETPIVOTDATA("" %SSLNET"",$P$6) >= 0.79,1,2)) "


                obj = oSheet.Cells(myRow, 24)
                obj.Value = "=GETPIVOTDATA("" %SSL"",$P$6)"
                smileySSL = oSheet.Cells(myRow + 1, 24)
                smileySSL.Value = "=IF(GETPIVOTDATA("" %SSL"",$P$6) > GETPIVOTDATA("" Target SSL"",$P$6),0,IF(GETPIVOTDATA("" %SSL"",$P$6) >=" & wSLow & ",1,2)) "
                CreatePivotCurrentYear = True
            End If
            obj = Nothing
            myret = True
        Catch ex As Exception
            myret = False
        End Try
        Return myret
    End Function

    Private Function CreatePivotLastYear(ByVal oWB As Excel.Workbook, ByVal oSheet As Excel.Worksheet, ByVal myTarget As String) As Boolean
        Dim myret As Boolean = False
        Try
            oWB.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C26", "PivotTable1b", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
            My.Application.DoEvents()

            '******************
            oSheet.PivotTables("PivotTable1b").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
            oSheet.PivotTables("PivotTable1b").PivotFields("shipdate").Caption = "Shipdate"
            oSheet.PivotTables("PivotTable1b").AddDataField(oSheet.PivotTables("PivotTable1b").PivotFields("weight"), " weight", Excel.XlConsolidationFunction.xlSum)
            oSheet.PivotTables("PivotTable1b").PivotFields(" weight").NumberFormat = "0"

            'oSheet.Range("G7").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)
            oSheet.PivotTables("PivotTable1b").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1b").PivotFields("Years").Position = 1
            oSheet.PivotTables("PivotTable1b").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1b").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1b").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1b").PivotFields("vendorname").Position = 1
            oSheet.PivotTables("PivotTable1b").PivotFields("Years").CurrentPage = startdate.Year - 1

            Dim obj As Object
            obj = oSheet.Cells(1, 23)
            obj.FormulaR1C1 = "=COUNTA(C26:C26)"
            If obj.Value > 5 Then
                myret = True
            End If
            obj = Nothing
        Catch ex As Exception
            myret = False
        End Try
        Return myret
    End Function


    Private Sub CreatePivotFull(ByVal oWB As Excel.Workbook, ByVal oSheet As Excel.Worksheet, ByVal myTarget As String)

        'oWB.PivotCaches.Add(xlDatabase, "DBRangeOTDSASL").CreatePivotTable oSheet.Name & "!R6C13", "PivotTable1c", xlPivotTableVersion10
        oWB.Worksheets("PivotTables").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C29", "PivotTable1d", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        My.Application.DoEvents()

        '******************
        oSheet.PivotTables("PivotTable1d").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1d").PivotFields("shipdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1d").PivotFields("shipdate").Caption = "Shipdate"
        oSheet.PivotTables("PivotTable1d").AddDataField(oSheet.PivotTables("PivotTable1d").PivotFields("weight"), " weight", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1d").PivotFields(" weight").NumberFormat = "0"
        'oSheet.Range("M7").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)

        oSheet.PivotTables("PivotTable1d").PivotFields("sopdescription").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1d").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1d").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1d").PivotFields("vendorname").Position = 1

        oSheet.PivotTables("PivotTable1d").PivotFields("Shipdate").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "PCT", "='sasl'/OrderCount", True
        'oSheet.PivotTables("PivotTable1c").CalculatedFields.Add "Target", myTarget, True
        'oSheet.PivotTables("PivotTable1d").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("TargetSASL"), " Target SASL", xlSum
        oSheet.PivotTables("PivotTable1d").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSL"), " Target SSL", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1d").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("TargetSSLNET"), " Target SSL NET", Excel.XlConsolidationFunction.xlSum)

        'oSheet.PivotTables("PivotTable1d").AddDataField oSheet.PivotTables("PivotTable1c").PivotFields("PCTSASL"), " %SASL(", xlSum")
        oSheet.PivotTables("PivotTable1d").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSL"), " %SSL", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1d").AddDataField(oSheet.PivotTables("PivotTable1c").PivotFields("PCTSSLNET"), " %SSLNET", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1d").PivotFields(" Target SASL").NumberFormat = "0%"
        oSheet.PivotTables("PivotTable1d").PivotFields(" Target SSL").NumberFormat = "0%"
        oSheet.PivotTables("PivotTable1d").PivotFields(" Target SSL NET").NumberFormat = "0%"
        'oSheet.PivotTables("PivotTable1d").PivotFields(" %SASL").NumberFormat = "0.0%"
        oSheet.PivotTables("PivotTable1d").PivotFields(" %SSL").NumberFormat = "0%"
        oSheet.PivotTables("PivotTable1d").PivotFields(" %SSLNET").NumberFormat = "0.0%"

        oSheet.PivotTables("PivotTable1d").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1d").DataPivotField.Position = 1
        oSheet.PivotTables("PivotTable1d").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1d").DisplayErrorString = True
        oSheet.PivotTables("PivotTable1d").ColumnGrand = False
        oSheet.PivotTables("PivotTable1d").RowGrand = False
        Dim obj As Object
        obj = oSheet.Cells(1, 28)
        Dim myRow As Integer
        obj.FormulaR1C1 = "=COUNTA(C29:C29)"
        If obj.Value >= 3 Then
            myRow = obj.Value - 2 + 9
            'obj = oSheet.Cells(myRow, 40)
            'obj.Value = "=GETPIVOTDATA("" SSLNET-RT"",$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year - 1 & ")/GETPIVOTDATA("" W-RT"",$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year - 1 & ")"
            'obj = oSheet.Cells(myRow, 41)
            'obj.Value = "=GETPIVOTDATA("" Sum of sslnet"",$A$6,""Years""," & DTPicker3.year & ")/GETPIVOTDATA("" weight"",$A$6,""Years""," & DTPicker1.year & ")"
            'obj = oSheet.Cells(myRow, 42)
            'obj.Value = "=GETPIVOTDATA("" SSL-RT"",$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year - 1 & ")/GETPIVOTDATA("" W-RT"",$A$6,""Shipdate""," & month(DTPicker3.Value) & ",""Years""," & DTPicker1.year - 1 & ")"
            'obj = oSheet.Cells(myRow, 43)
            'obj.Value = "=GETPIVOTDATA("" Sum of ssl"",$A$6,""Years""," & DTPicker3.year & ")/GETPIVOTDATA("" weight"",$A$6,""Years""," & DTPicker1.year & ")"
            'obj = oSheet.Cells(myRow + 1, 44)
            'obj.Value = "=GETPIVOTDATA("" %SSLNET"",$A$6,""Years""," & DTPicker1.year - 1 & ")"
            'obj = oSheet.Cells(myRow + 1, 45)
            'obj.Value = "=GETPIVOTDATA("" %SSL"",$A$6,""Years""," & DTPicker1.year - 1 & ")"
            obj = oSheet.Cells(myRow, 40)
            obj.Value = "=GETPIVOTDATA("" SSLNET-RT"",$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & ")/GETPIVOTDATA("" W-RT"",$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & ")"
            obj = oSheet.Cells(myRow, 41)
            obj.Value = "=GETPIVOTDATA("" Sum of sslnet"",$A$6,""Years""," & currentmonth.Year & ")/GETPIVOTDATA("" weight"",$A$6,""Years""," & startdate.Year & ")"
            obj = oSheet.Cells(myRow, 42)
            obj.Value = "=GETPIVOTDATA("" SSL-RT"",$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & ")/GETPIVOTDATA("" W-RT"",$A$6,""Shipdate""," & Month(currentmonth) & ",""Years""," & startdate.Year - 1 & ")"
            obj = oSheet.Cells(myRow, 43)
            obj.Value = "=GETPIVOTDATA("" Sum of ssl"",$A$6,""Years""," & currentmonth.Year & ")/GETPIVOTDATA("" weight"",$A$6,""Years""," & startdate.Year & ")"
            obj = oSheet.Cells(myRow + 1, 44)
            obj.Value = "=GETPIVOTDATA("" %SSLNET"",$A$6,""Years""," & startdate.Year - 1 & ")"
            obj = oSheet.Cells(myRow + 1, 45)
            obj.Value = "=GETPIVOTDATA("" %SSL"",$A$6,""Years""," & startdate.Year - 1 & ")"

        End If
        obj = Nothing


errExit:
        Exit Sub
ErrHdl:

        GoTo errExit

    End Sub
    Private Function dateformatcomma(ByVal mydate As Date) As String
        dateformatcomma = Year(mydate) & "," & Month(mydate) & "," & mydate.Day

    End Function

    Public Function ValidateFileName(ByVal myFullPathFilename As String) As String

        Dim bFCheck As Boolean
        bFCheck = True
        Dim myCopy As String
        Dim i As Integer
        myCopy = ""
        i = 0

        While bFCheck
            If Not IO.File.Exists(Path.GetDirectoryName(myFullPathFilename) & "\" & Path.GetFileNameWithoutExtension(myFullPathFilename) & myCopy & Path.GetExtension(myFullPathFilename)) Then
                bFCheck = False
            Else
                i = i + 1
                myCopy = "_(" & i & ")"
            End If
            My.Application.DoEvents()
        End While
        bFCheck = True
        ValidateFileName = Path.GetDirectoryName(myFullPathFilename) & "\" & Path.GetFileNameWithoutExtension(myFullPathFilename) & myCopy & Path.GetExtension(myFullPathFilename)
    End Function
End Class
