Imports System.Threading
Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Class FormRTurnoverExtendYearCur
    Implements INotifyPropertyChanged
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myInitThread As New System.Threading.Thread(AddressOf DoQuery)
    Dim myDeleteThread As New System.Threading.Thread(AddressOf DoDelete)
    Private ComboboxDS As New DataSet
    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged
    Private myReport As RTurnoverExtendYearCur
    Private BaseItem As String

    Public Property TurnoverHistoryAction As Integer
        Get
            If RadioButton6.Checked = True Then
                Return TurnoverHistoryActionEnum.DoNotSavePeriod
            ElseIf RadioButton7.Checked = True Then
                Return TurnoverHistoryActionEnum.SavePeriod
            Else
                Return TurnoverHistoryActionEnum.DeletePeriod
            End If
        End Get
        Set(ByVal value As Integer)
            If value = TurnoverHistoryActionEnum.DoNotSavePeriod Then
                RadioButton6.Checked = True
            ElseIf value = TurnoverHistoryActionEnum.SavePeriod Then
                RadioButton7.Checked = True
            ElseIf value = TurnoverHistoryActionEnum.DeletePeriod Then
                RadioButton8.Checked = True
            End If
        End Set
    End Property

    Public Property ProductType As Integer
        Get
            If RadioButton1.Checked = True Then
                Return ProductTypeEnum.FinishedGood            
            Else
                Return ProductTypeEnum.Components
            End If
        End Get
        Set(ByVal value As Integer)
            If value = ProductTypeEnum.FinishedGood Then
                RadioButton1.Checked = True           
            ElseIf value = ProductTypeEnum.Components Then
                RadioButton2.Checked = True
            End If
        End Set
    End Property

    Public Property ReportType As Integer
        Get
            If RadioButton3.Checked = True Then
                Return ReportTypeEnum.ALLData
            ElseIf RadioButton4.Checked = True Then
                Return ReportTypeEnum.WMF
            Else
                Return ReportTypeEnum.SEBAsia
            End If
        End Get
        Set(ByVal value As Integer)
            If value = ReportTypeEnum.ALLData Then
                RadioButton3.Checked = True
            ElseIf value = ReportTypeEnum.WMF Then
                RadioButton4.Checked = True
            ElseIf value = ReportTypeEnum.SEBAsia Then
                RadioButton5.Checked = True
            End If
        End Set
    End Property

    Sub DoQuery()
        Using myProcess = New RTurnoverExtendYearCur
            ProgressReport(1, "Preparing Interface...Please wait")
            ProgressReport(6, "Marquee")
            If myProcess.loadCombobox(ComboboxDS) Then
                ProgressReport(8, "Initialize Combobox")
            Else
                ProgressReport(1, myProcess.errorMessage)
            End If
            ProgressReport(5, "Continuous")
            ProgressReport(1, "Preparing Interface...Done")
        End Using        
    End Sub

    Sub DoDelete()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Delete Record...Please wait")
        If Not myReport.DeletePeriod() Then
            ProgressReport(1, myReport.errorMessage)
        End If
        ProgressReport(6, "Continuous")
        ProgressReport(1, "Delete Record...Done. Refreshing Period...")

        Using myprocess = New RTurnoverExtendYearCur
            If myprocess.loadCombobox(ComboboxDS) Then
                ProgressReport(8, "Initialize Combobox")
            Else
                ProgressReport(1, myprocess.errorMessage)
            End If
            ProgressReport(5, "Continuous")
            ProgressReport(1, "Delete Record...Done. Refreshing Period....Done")
        End Using

    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myInitThread.Start()
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
                Case 4

                Case (5)
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                Case 8
                    Me.ComboBox1.DataSource = ComboboxDS.Tables(0)
                    Me.ComboBox1.DisplayMember = "period"
                    Me.ComboBox1.ValueMember = "period"
            End Select

        End If

    End Sub
    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged, RadioButton7.CheckedChanged, RadioButton8.CheckedChanged
        onPropertyChanged("TurnoverHistoryAction")
    End Sub
    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        onPropertyChanged("ReportType")
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        onPropertyChanged("ProductType")
        Panel1.Enabled = IIf(ProductType = ProductTypeEnum.FinishedGood, True, False)
    End Sub

    Sub onPropertyChanged(ByVal name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        BaseItem = DirectCast(ComboBox1.SelectedItem, DataRowView).Item("period")
        myReport = New RTurnoverExtendYearCur(ProductType, ReportType, TurnoverHistoryAction, DateTimePicker1.Value, BaseItem, DateTimePicker2.Value)
        'Get File Location
        'Generate Excel

        'Delete Period Function
        If TurnoverHistoryAction = TurnoverHistoryActionEnum.DeletePeriod Then
            If Not myDeleteThread.IsAlive Then
                myDeleteThread = New System.Threading.Thread(AddressOf DoDelete)
                myDeleteThread.Start()
            Else
                MessageBox.Show("Please wait, current process still running...")
            End If           
            Exit Sub
        End If
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        'SaveFileDialog1.FileName = String.Format("Turnover_{0:yyyyMMMdd}_{1}.xlsx", Today.Date, IIf(ProductType = ProductTypeEnum.FinishedGood, "FG", "CP"))
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'Dim mypath As String = System.IO.Path.GetDirectoryName(FolderBrowserDialog1.SelectedPath)
            Dim mypath As String = FolderBrowserDialog1.SelectedPath
            Dim reportname = String.Format("Turnover-{0}", IIf(ProductType = ProductTypeEnum.FinishedGood, "FG", "CP"))
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 12,
                                                            .SheetName = "DATA",
                                                            .Sqlstr = myReport.getQueryData
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 13,
                                                            .SheetName = "DATASUMMARY",
                                                            .Sqlstr = myReport.getQueryDataSummary
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            Dim myExcelReport As New ExportToExcelFile(Me, myQueryWorksheetList, mypath, reportname, mycallback, PivotCallback)
            myExcelReport.Run(Me, e)
        End If
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim withbilloflading As Boolean = False
    '    If MessageBox.Show("Include bill of lading?", "Bill of lading", MessageBoxButtons.YesNo) = DialogResult.Yes Then
    '        withbilloflading = True
    '    End If
    '    Dim dr As DataRow = CType(bs1.Current, DataRowView).Row
    '    Dim myQueryWorksheetList As New List(Of QueryWorksheet)
    '    Me.ToolStripStatusLabel1.Text = ""
    '    Me.ToolStripStatusLabel2.Text = ""

    '    Dim mymessage As String = String.Empty

    '    Dim myuser As String = String.Empty
    '    Dim dateupload As Date
    '    Dim sqlstr As String
    '    Dim sqlstr1 As String = String.Empty

    '    'If ComboBox1.Text <> "" Then
    '    myuser = "'" & dr.Item("username") & "'"
    '    dateupload = dr.Item("dateupload")
    '    If withbilloflading Then
    '        sqlstr = "select tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost amount,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"

    '        sqlstr1 = " with foo as (select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice, invoicehardcopydtid from (select  tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
    '                 " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
    '                 " order by invoicehardcopydtid) " &
    '                 " select foo.supplierinvoice,foo.amount,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
    '                 " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"

    '    Else
    '        sqlstr = "select tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
    '        'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopynobill(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
    '        '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
    '        '         " order by invoicehardcopydtid "

    '        sqlstr1 = "with foo as (select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,invoicehardcopydtid   from (select  tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
    '                " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
    '                " order by invoicehardcopydtid) " &
    '                " select foo.supplierinvoice,foo.amount,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
    '                " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"
    '    End If
    '    'sqlstr = "select tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
    '    'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
    '    '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
    '    '         " order by invoicehardcopydtid "

    '    'Else
    '    '    'sqlstr = "select tb.*from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"
    '    '    MessageBox.Show("Please select User Name.")
    '    '    Exit Sub
    '    'End If

    '    Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
    '    DirectoryBrowser.Description = "Which directory do you want to use?"
    '    If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
    '        Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
    '        Dim reportname = "DocumentHardCopy" '& GetCompanyName()
    '        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
    '        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

    '        Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
    '                                                        .SheetName = "DETAIL",
    '                                                        .Sqlstr = sqlstr
    '                                                        }
    '        myQueryWorksheetList.Add(myqueryworksheet)

    '        myqueryworksheet = New QueryWorksheet With {.DataSheet = 2,
    '                                                        .SheetName = "Total",
    '                                                        .Sqlstr = sqlstr1
    '                                                        }
    '        myQueryWorksheetList.Add(myqueryworksheet)

    '        'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
    '        Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

    '        myreport.Run(Me, e)

    '    End If

    'End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        If sender.name = "DATA" Then            
            'Check for saving Period
            If TurnoverHistoryAction = TurnoverHistoryActionEnum.SavePeriod Then
                myReport.AddNewPeriod()
            End If       
        End If

        

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim oWB = DirectCast(sender, Excel.Workbook)
        Dim oXL As Excel.Application = oWB.Parent
        Dim oSheet As Excel.Worksheet
        Dim pvt As Excel.PivotItem
        Dim Time1 As DateTime = Now
        Dim iws = 1

        'oWB.Worksheets(1).Select
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)

        oWB.Names.Add(Name:="DBSUMMARYRange", RefersToR1C1:="=OFFSET(DATASUMMARY!R1C1,0,0,COUNTA(DATASUMMARY!C4),COUNTA(DATASUMMARY!R1))")
        'oWB.Names.Add Name:="DBSUMMARYRange", RefersToR1C1:="=DATASUMMARY!EXTERNALDATA_1"
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))


        oWB.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBSUMMARYRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "Evo by Period N-GP"
        Application.DoEvents()


        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("SUM TTL", "=forecast+firmorder+shipment", True)
        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Sum TTL Variance", "=SUM TTL", True)

        oSheet.PivotTables("PivotTable1").ColumnGrand = False
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("sbuname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("sbuname").PivotItems
                If pvt.Value = "(blank)" Or pvt.Value = "NON AFFECTED" Then
                    pvt.Visible = False
                End If
            Next
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("factory").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("factory").PivotItems
                If pvt.Value = "(blank)" Then
                    pvt.Visible = False
                End If
            Next
        End If

        oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        'vendor show non group only
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False

            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            'End If
        Next

        oSheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("myyear").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("inqconf").Orientation = Excel.XlPivotFieldOrientation.xlColumnField


        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Sum TTL Variance"), "TTL Variance from 1st Period", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("period").AutoSort(Excel.XlSortOrder.xlDescending, "period")
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("SUM TTL"), "TTL", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), " FORECAST", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("firmorder"), " FIRMORDER", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("shipment"), " SHIPMENT", Excel.XlConsolidationFunction.xlSum)


        oSheet.PivotTables("PivotTable1").PivotFields(" FORECAST").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" FIRMORDER").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" SHIPMENT").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields("TTL").NumberFormat = "0.0"

        With oSheet.PivotTables("PivotTable1").PivotFields("TTL Variance from 1st Period")
            .Calculation = Excel.XlPivotFieldCalculation.xlDifferenceFrom
            .BaseField = "period"
            .BaseItem = myReport.BaseItem 'Combo1.Text
            .NumberFormat = "0.0_);[Red](0.0)"
        End With

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("inqconf").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf pvt.Value = 13 Then
                pvt.Visible = False
            End If
        Next
        With oSheet.Cells.Font
            .Name = "Calibri"
            .Size = 8
        End With
        oSheet.Cells.EntireColumn.AutoFit()




        'Forecast Pricing



        iws = iws + 1
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oWB.Worksheets(iws).Select()
            oSheet = oWB.Worksheets(iws)

            'oWB.Names.Add Name:="DBSUMMARYRange", RefersToR1C1:="=OFFSET(DATASUMMARY!R1C1,0,0,COUNTA(DATASUMMARY!C4),COUNTA(DATASUMMARY!R1))"
            'oWB.Names.Add Name:="DBSUMMARYRange", RefersToR1C1:="=DATASUMMARY!EXTERNALDATA_1"
            'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
            ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
            oWB.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
            oSheet.Name = "Forecast Pricing"

            'oWB.PivotCaches.Add(xlDatabase, "DBSUMMARYRange").CreatePivotTable oSheet.Name & "!R6C1", "PivotTable1", xlPivotTableVersion10
            'oSheet.Name = "Evo by Period N-GP"
            Application.DoEvents()


            'oSheet.PivotTables("PivotTable1").CalculatedFields.Add "SUM TTL", "=forecast+firmorder+shipment", True
            'oSheet.PivotTables("PivotTable1").CalculatedFields.Add "Sum TTL Variance", "=SUM TTL", True

            oSheet.PivotTables("PivotTable1").ColumnGrand = False
            '   If mypg = 1 Then
            oSheet.PivotTables("PivotTable1").PivotFields("sbuname").Orientation = Excel.XlPivotFieldOrientation.xlColumnField 'xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("sbuname").PivotItems
                If pvt.Value = "(blank)" Or pvt.Value = "NON AFFECTED" Then
                    pvt.Visible = False
                End If
            Next
            '   Else
            '        oSheet.PivotTables("PivotTable1").PivotFields("factory").Orientation = xlRowField 'xlPageField
            '        'hide nonaffected and blank
            '        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("factory").PivotItems
            '            If pvt.Value = "(blank)" Then
            '                pvt.Visible = False
            '            End If
            '        Next
            '   End If

            oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'vendor show non group only
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").PivotItems
                If pvt.Value = "(blank)" Then
                    pvt.Visible = False
                ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                    pvt.Visible = False
                End If
                'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
                '    pvt.Visible = False
                'End If
            Next

            oSheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField

            'oSheet.PivotTables("PivotTable1").PivotFields("myyear").Orientation = xlColumnField
            'oSheet.PivotTables("PivotTable1").PivotFields("inqconf").Orientation = xlColumnField


            'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("Sum TTL Variance"), "TTL Variance from 1st Period", Excel.XlConsolidationFunction.xlSum
            'oSheet.PivotTables("PivotTable1").PivotFields("period").AutoSort xlDescending, "period"
            'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("SUM TTL"), "TTL", xlSum
            oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), " FORECAST", Excel.XlConsolidationFunction.xlSum)
            'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("firmorder"), " FIRMORDER", xlSum
            'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("shipment"), " SHIPMENT", xlSum


            oSheet.PivotTables("PivotTable1").PivotFields(" FORECAST").NumberFormat = "0.0"
            'oSheet.PivotTables("PivotTable1").PivotFields(" FIRMORDER").NumberFormat = "0.0"
            'oSheet.PivotTables("PivotTable1").PivotFields(" SHIPMENT").NumberFormat = "0.0"
            'oSheet.PivotTables("PivotTable1").PivotFields("TTL").NumberFormat = "0.0"

            '   With oSheet.PivotTables("PivotTable1").PivotFields("TTL Variance from 1st Period")
            '        .Calculation = xlDifferenceFrom
            '        .BaseField = "period"
            '        .BaseItem = Combo1.Text
            '        .NumberFormat = "0.0_);[Red](0.0)"
            '   End With

            '   For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("inqconf").PivotItems
            '       If pvt.Value = "(blank)" Then
            '           pvt.Visible = False
            '       ElseIf pvt.Value = 13 Then
            '           pvt.Visible = False
            '       End If
            '   Next
            With oSheet.Cells.Font
                .Name = "Calibri"
                .Size = 8
            End With
            oSheet.Cells.EntireColumn.AutoFit()

        End If



        'Finance
        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)

        oWB.Names.Add(Name:="DBSUMMARYRange", RefersToR1C1:="=OFFSET(DATASUMMARY!R1C1,0,0,COUNTA(DATASUMMARY!C4),COUNTA(DATASUMMARY!R1))")
        'oWB.Names.Add Name:="DBSUMMARYRange", RefersToR1C1:="=DATASUMMARY!EXTERNALDATA_1"
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))


        oWB.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBSUMMARYRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "Evo by Period N-GP (Finance)"
        Application.DoEvents()


        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("SUM TTL", "=forecastoc+firmorderoc+shipmentoc", True)
        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Sum TTL Variance", "=SUM TTL", True)

        oSheet.PivotTables("PivotTable1").ColumnGrand = False
        oSheet.PivotTables("PivotTable1").RowGrand = False

        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("sbuname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("sbuname").PivotItems
                If pvt.Value = "(blank)" Or pvt.Value = "NON AFFECTED" Then
                    pvt.Visible = False
                End If
            Next
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("factory").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("factory").PivotItems
                If pvt.Value = "(blank)" Then
                    pvt.Visible = False
                End If
            Next
        End If

        oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        'vendor show non group only
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            '    pvt.Visible = False
            'End If
        Next

        oSheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("inqconf").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        With oSheet.PivotTables("PivotTable1").PivotFields("crcy")
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
            .Position = 1
        End With

        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Sum TTL Variance"), "TTL Variance from 1st Period", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("period").AutoSort(Excel.XlSortOrder.xlDescending, "period")
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("SUM TTL"), "TTL", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecastoc"), " FORECAST", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("firmorderoc"), " FIRMORDER", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("shipmentoc"), " SHIPMENT", Excel.XlConsolidationFunction.xlSum)


        oSheet.PivotTables("PivotTable1").PivotFields(" FORECAST").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" FIRMORDER").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" SHIPMENT").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields("TTL").NumberFormat = "0.0"

        With oSheet.PivotTables("PivotTable1").PivotFields("TTL Variance from 1st Period")
            .Calculation = Excel.XlPivotFieldCalculation.xlDifferenceFrom
            .BaseField = "period"
            .BaseItem = myReport.BaseItem 'Combo1.Text
            .NumberFormat = "0.0_);[Red](0.0)"
        End With

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("inqconf").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf pvt.Value = 13 Then
                pvt.Visible = False
            End If
        Next

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("crcy").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            End If
        Next
        With oSheet.Cells.Font
            .Name = "Calibri"
            .Size = 8
        End With
        oSheet.Cells.EntireColumn.AutoFit()


        '**************





        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets(1).PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "Evo by Period N-GP Supplier"
        Application.DoEvents()

        oSheet.PivotTables("PivotTable1").ColumnGrand = False

        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("sbuname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("sbuname").PivotItems
                If pvt.Value = "(blank)" Or pvt.Value = "NON AFFECTED" Then
                    pvt.Visible = False
                End If
            Next
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("factory").Orientation = Excel.XlPivotFieldOrientation.xlPageField
            'hide nonaffected and blank
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("factory").PivotItems
                If pvt.Value = "(blank)" Then
                    pvt.Visible = False
                End If
            Next
        End If

        oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        'vendor show non group only
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("vendorcode").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            '    pvt.Visible = False
            'End If
        Next

        oSheet.PivotTables("PivotTable1").PivotFields("period").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("inqconf").Orientation = Excel.XlPivotFieldOrientation.xlColumnField

        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Sum TTL Variance"), "TTL Variance from 1st Period", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("period").AutoSort(Excel.XlSortOrder.xlDescending, "period")
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("SUM TTL"), "TTL", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("forecast"), " FORECAST", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("firmorder"), " FIRMORDER", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("shipment"), " SHIPMENT", Excel.XlConsolidationFunction.xlSum)


        oSheet.PivotTables("PivotTable1").PivotFields(" FORECAST").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" FIRMORDER").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields(" SHIPMENT").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields("TTL").NumberFormat = "0.0"

        With oSheet.PivotTables("PivotTable1").PivotFields("TTL Variance from 1st Period")
            .Calculation = Excel.XlPivotFieldCalculation.xlDifferenceFrom
            .BaseField = "period"
            .BaseItem = myReport.BaseItem 'Combo1.Text
            .NumberFormat = "0.0_);[Red](0.0)"
        End With

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("inqconf").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            End If
        Next
        With oSheet.Cells.Font
            .Name = "Calibri"
            .Size = 8
        End With
        oSheet.Cells.EntireColumn.AutoFit()

        'oSheet.PivotTables("PivotTable1").PivotFields("vendorname").AutoSort xlDescending, "TTL", oSheet.PivotTables("PivotTable1").PivotColumnAxis.PivotLines(13), 1



        iws = iws + 1
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        oWB.Names.Add(Name:="DBRange", RefersToR1C1:="=OFFSET(DATA!R1C1,0,0,COUNTA(DATA!C1),COUNTA(DATA!R1))")
        'oWB.Names.Add Name:="DBRange", RefersToR1C1:="=DATA!EXTERNALDATA_1"
        '   oXL.ActiveWindow.Zoom = 80
        '   oSheet.Cells.EntireColumn.AutoFit

        oWB.PivotCaches.Add(Excel.XlPivotTableSourceType.xlDatabase, "DBRange").CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "non group Year Compare"
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"

        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If


        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        'Vendor SAP Show only SSEAC
        'SBU Hide blank and Non affected
        'familysbu hide blank and nonaffected

        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlRowField



        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Purchase Value in Millions", "='Purchase Value' /1000000", True)
        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Variance Y - Y1", "='Year" & Year(myReport.myLastDate) & " Purchase Value' - 'Year" & Year(myReport.myFirstDate) & " Purchase Value'", True)
        oSheet.PivotTables("PivotTable1").CalculatedFields.Add("Quantity Variance YvsY-1", "='Year" & Year(myReport.myLastDate) & " QTY' - 'Year" & Year(myReport.myFirstDate) & " QTY'", True)

        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Year" & Year(myReport.myFirstDate) & " Purchase Value"), "Purchase Value (MUSD) " & Year(myReport.myFirstDate), Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Year" & Year(myReport.myLastDate) & " Purchase Value"), "Purchase Value (MUSD) " & Year(myReport.myLastDate), Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Variance Y - Y1"), " Variance Y - Y1", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Variance Y - Y1"), " Accu Variance Y - Y1", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Variance Y - Y1").Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Variance Y - Y1").BaseField = "MONTH INQCONF"

        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Year" & Year(myReport.myFirstDate) & " QTY"), "Volume Quantity (M pcs) " & Year(myReport.myFirstDate), Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Year" & Year(myReport.myLastDate) & " QTY"), "Volume Quantity (M pcs) " & Year(myReport.myLastDate), Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Quantity Variance YvsY-1"), " Volume Quantity Variance YvsY-1", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Quantity Variance YvsY-1"), " Accu Volume Quantity Variance YvsY-1", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Volume Quantity Variance YvsY-1").Calculation = Excel.XlPivotFieldCalculation.xlRunningTotal
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Volume Quantity Variance YvsY-1").BaseField = "MONTH INQCONF"
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), " Purchase Value in Millions", xlSum



        oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value (MUSD) " & Year(myReport.myFirstDate)).NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value (MUSD) " & Year(myReport.myLastDate)).NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields("Volume Quantity (M pcs) " & Year(myReport.myFirstDate)).NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields("Volume Quantity (M pcs) " & Year(myReport.myLastDate)).NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"

        oSheet.PivotTables("PivotTable1").PivotFields(" Variance Y - Y1").NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Variance Y - Y1").NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields(" Volume Quantity Variance YvsY-1").NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields(" Accu Volume Quantity Variance YvsY-1").NumberFormat = "0.00;[Red]-0.00" '"#,##0.00"
        'oSheet.PivotTables("PivotTable1").PivotFields(" Purchase Value in Millions").NumberFormat = "#,##0.00"
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").AutoSort(Excel.XlSortOrder.xlDescending, " Purchase Value in Millions")

        oSheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}



        'vendor
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            '    pvt.Visible = False
            'End If
        Next

        'month remove blank
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            End If
        Next

        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

        End If

        oSheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").LayoutForm = Excel.XlLayoutFormType.xlTabular
        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()


        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "N-GP Supplier  " & Year(myReport.myLastDate)
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"

        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        'Vendor SAP Show only SSEAC
        'SBU Hide blank and Non affected
        'familysbu hide blank and nonaffected

        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField


        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), "Sum of Purchase Value (M USD)", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of Purchase Value (M USD)").NumberFormat = "#,##0.0"

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").AutoSort(Excel.XlSortOrder.xlDescending, "Sum of Purchase Value (M USD)")

        'select year

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value < Year(myReport.myLastDate) Then
                pvt.Visible = False

            End If
        Next
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
        End If


        'vendor
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            ' End If
        Next

        oSheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").LayoutForm = Excel.XlLayoutFormType.xlTabular

        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()





        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        'oWB.Worksheets(1).PivotCache.CreatePivotTable oSheet.Name & "!R6C1", "PivotTable1", xlPivotTableVersion10

        oSheet.Name = "N-GP BU " & Year(myReport.myLastDate)
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"

        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField




        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField



        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), "Sum of Purchase Value (M USD)", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of Purchase Value (M USD)").NumberFormat = "#,##0.0"

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}




        'select year

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value < Year(myReport.myLastDate) Then
                pvt.Visible = False

            End If
        Next
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
        End If
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            ' End If
        Next

        oSheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
        'oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").LayoutForm = Excel.XlLayoutFormType.xlTabular


        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()



        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "N-GP BU Qty"
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"

        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField


        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("industrialfamily").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("description").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("description").Caption = "Description of Model"
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField



        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("osqty"), "Sum of QuantityActual(K PCS)", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of QuantityActual(K PCS)").NumberFormat = "#,##0"

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").PivotFields("industrialfamily").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        oSheet.PivotTables("PivotTable1").PivotFields("Description of Model").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'select year

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value < Year(myReport.myLastDate) Then
                pvt.Visible = False

            End If
        Next
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
        End If
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) = 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) = 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            'End If
        Next

        '    oSheet.PivotTables("PivotTable1").RowAxisLayout xlCompactRow
        '    oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").LayoutForm = xlTabular

        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()




        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)
        oSheet.Name = "GP Supplier " & Year(myReport.myLastDate)
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"

        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField


        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField



        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), "Sum of Purchase Value(M USD)", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of Purchase Value(M USD)").NumberFormat = "#,##0.0"

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").AutoSort(Excel.XlSortOrder.xlDescending, "Sum of Purchase Value")

        'select year

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value < Year(myReport.myLastDate) Then
                pvt.Visible = False

            End If
        Next

        If myReport.ProductType = ProductTypeEnum.FinishedGood Then

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
        End If
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) <> 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) <> 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            'End If
        Next

        oSheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").LayoutForm = Excel.XlLayoutFormType.xlTabular

        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()





        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'Statusbar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)

        oSheet.Name = "GP BU " & Year(myReport.myLastDate)
        Application.DoEvents()
        'Pivot Table1

        oSheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleMedium8"


        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        'Vendor SAP Show only SSEAC
        'SBU Hide blank and Non affected
        'familysbu hide blank and nonaffected

        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("MONTH INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField



        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), "Sum of Purchase Value (M USD)", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of Purchase Value (M USD)").NumberFormat = "#,##0.0"

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}




        'select year

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value < Year(myReport.myLastDate) Then
                pvt.Visible = False

            End If
        Next
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next

            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("familysbu").PivotItems
                If InStr(1, "NON AFFECTED,(blank)", pvt.Value) <> 0 Then
                    pvt.Visible = False
                End If

            Next
        End If
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If pvt.Value = "(blank)" Then
                pvt.Visible = False
            ElseIf Mid(pvt.Value, 1, 2) <> 99 Then
                pvt.Visible = False
            End If
            'If Mid(pvt.Value, 1, 2) <> 99 Or pvt.Value = "(blank)" Then
            ' pvt.Visible = False
            ' End If
        Next

        oSheet.PivotTables("PivotTable1").RowAxisLayout(Excel.XlLayoutRowType.xlCompactRow)
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").LayoutForm = Excel.XlLayoutFormType.xlTabular


        oXL.ActiveWindow.Zoom = 80
        oSheet.Cells.EntireColumn.AutoFit()


        iws = iws + 1
        oWB.Worksheets(iws).Select()
        oSheet = oWB.Worksheets(iws)
        'StatusBar1.Panels(1).Text = "Generating PivotTable(" & iws & "/11)... "
        ProgressReport(1, String.Format("Generating PivotTable {0}/11", iws))
        oWB.Worksheets("non group Year Compare").PivotTables("PivotTable1").PivotCache.CreatePivotTable(oSheet.Name & "!R8C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersion10)

        oSheet.Name = "TH_B"
        Application.DoEvents()

        oSheet.PivotTables("PivotTable1").PivotFields("SPM").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        If myReport.ProductType = ProductTypeEnum.FinishedGood Then
            oSheet.PivotTables("PivotTable1").PivotFields("SBU").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        Else
            oSheet.PivotTables("PivotTable1").PivotFields("Sold To Party Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        End If
        oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        oSheet.PivotTables("PivotTable1").PivotFields("confirmstatus").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Orientation = Excel.XlPivotFieldOrientation.xlRowField

        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        oSheet.PivotTables("PivotTable1").PivotFields("adjustedmonth").Orientation = Excel.XlPivotFieldOrientation.xlColumnField


        oSheet.PivotTables("PivotTable1").AddDataField(oSheet.PivotTables("PivotTable1").PivotFields("Purchase Value in Millions"), "Sum of Purchase Value in Millions", Excel.XlConsolidationFunction.xlSum)
        oSheet.PivotTables("PivotTable1").PivotFields("Sum of Purchase Value in Millions").NumberFormat = "0.0"
        oSheet.PivotTables("PivotTable1").PivotFields("TYPE").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").PivotFields("familysbu").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("adjustedmonth").PivotItems
            If InStr(1, "(blank)", pvt.Value) = 1 Then
                pvt.Visible = False
            End If
        Next
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("Vendor Name").PivotItems
            If InStr(1, "(blank)", pvt.Value) = 1 Then
                pvt.Visible = False
            End If
        Next

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SAP Vendor Code").PivotItems
            If InStr(1, "(blank)", pvt.Value) = 1 Then
                pvt.Visible = False
            End If
        Next
        If myReport.ProductType = ProductTypeEnum.Components Then


            For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("SBU").PivotItems
                If InStr(1, "(blank)", pvt.Value) = 1 Then
                    pvt.Visible = False
                End If
            Next
        End If
        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("confirmstatus").PivotItems
            If InStr(1, "N,(blank)", pvt.Value) = 1 Then
                pvt.Visible = False
            End If
        Next

        For Each pvt In oSheet.PivotTables("PivotTable1").PivotFields("YEAR_INQCONF").PivotItems
            If pvt.Value <> Year(myReport.myLastDate) Then
                pvt.Visible = False
            End If
        Next
        With oSheet.Cells.Font
            .Name = "Calibri"
            .Size = 8
        End With
        oSheet.Cells.EntireColumn.AutoFit()


        Dim mymessage As String
        mymessage = "Processing Time: " & Format(DateAdd("s", DateDiff("s", time1, Now), "00:00:00"), "HH:mm:ss")
        'StatusBar1.Panels(1).Text = mymessage
        ProgressReport(1, String.Format("{0}", mymessage))

        oWB.Worksheets(1).Select()
    End Sub





End Class