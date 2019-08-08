Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass
Public Class FormShipmentAndFreightReport

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty


        'Dim sqlstr As String = "select * from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date,'SLEUNG') as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)"
        Dim sqlstr As String = "select *, case when trpttype = 'By Air' then 1 else 0 end as ""Orders By Air"",case when trpttype = 'By Air' then 0 else 1 end as ""Orders By Sea"" from getdeliverycountcalc1(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date) as tb(deliveryhd bigint,deliveryitem integer,deliverydate date,createdon date, billoflading character varying,meansoftransid character varying,createdby character varying, trpt character varying,trpttype character(20),volumecbm integer, shpt character varying,shippointdesc character(30),tradedistrict character(30),deliveredqty numeric,volume numeric, vun character varying,totalweight numeric,un character varying,vendorcode bigint,vendorname character(50),forwardername character(50),salesdoc bigint, salesdocitem integer,custpo character varying, custpono character varying,creationdate date,shiptoparty bigint,shiptopartyname character(50),soldtoparty bigint, soldtopartyname character(50),sao character varying,cmmf bigint,qty bigint,su character varying,rdeliverydate date,purchasinggroup character(3),comp integer, fg integer,pod character(50),zoneid integer,zone character(30),countdata integer,""Estimated 20 TO TEU"" numeric,""Estimated 40 TO TEU"" numeric,""Consol TO TEU"" numeric,""TEU COUNT"" numeric,consol numeric,""MIX"" integer, ""Total MIX"" numeric,conso20 integer,conso40 integer,conso40hq integer,fcl20 integer,fcl40 integer,fcl40hq integer,teutotal integer)"
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "FreightReport" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim datasheet As Integer = 5
            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.77\Users_I\Logistic Dept\KPI & Reporting\templates\FreightReportTemplate.xltx")
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        owb.Names.Add("DB_TWO", RefersToR1C1:="=OFFSET('DB_TWO'!R1C1,0,0,COUNTA('DB_TWO'!C1),COUNTA('DB_TWO'!R1))")
        owb.Worksheets(3).select()
        Dim osheet = owb.Worksheets(3)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        osheet.PivotTables("PivotTable3").PivotCache.Refresh()
        osheet.PivotTables("PivotTable4").PivotCache.Refresh()
        osheet.PivotTables("PivotTable5").PivotCache.Refresh()
        osheet.PivotTables("PivotTable6").PivotCache.Refresh()
        osheet.PivotTables("PivotTable7").PivotCache.Refresh()
        osheet.PivotTables("PivotTable8").PivotCache.Refresh()

        owb.Worksheets(4).select()
        osheet = owb.Worksheets(4)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        osheet.PivotTables("PivotTable4").PivotCache.Refresh()
        osheet.PivotTables("PivotTable7").PivotCache.Refresh()
        osheet.PivotTables("PivotTable8").PivotCache.Refresh()
        osheet.PivotTables("PivotTable9").PivotCache.Refresh()
    End Sub
End Class