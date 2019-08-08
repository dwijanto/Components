Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass
Public Class FormReportKPIChart

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty


        'Dim sqlstr As String = "select * from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date,'SLEUNG') as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)"
        Dim sqlstr As String = "select * from sp_kpidata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "," & DateFormatyyyyMMdd(DateTimePicker2.Value) & ") as tb(ordertype character varying,shiptoparty bigint, shiptopartyname character(50),soldtoparty bigint, soldtopartyname character(50), customerorderno character varying, sebasiasalesorder bigint," &
                               "solineno integer, vendorcode bigint,vendorname character(50),rir character(2), cmmf bigint,comfam integer,itemid character(15),materialdesc character(50), orderstatus character varying," &
                               "latestupdate date, updatesince character varying, curinq character varying,receptiondate date, fob real,unittp real, inquiryeta date,inquiryetd date, inquiryqty integer, " &
                               "currentinquiryetd date, currentinquiryqty integer, confirmationstatus character varying, stconfirmedetd date,stconfirmedqty integer,currentconfirmedeta date, currentconfirmedetd date," &
                               "currentconfirmedqty integer, deliveredqty integer, shipdate date,shipdateeta date,osqty integer, sebasiapono bigint, polineno integer,ctrno character varying, boatid character varying," &
                               "packinglist character varying,shipfrom character varying, comments character varying, cmnttxdtlname character(50),sao character varying, purchasinggroup character varying, " &
                               "sbu character(30),status text, shipmentline integer,week double precision,shipdate2 date, bu character(30),familysbu character(30), sp text,spm text,cmaxtext text,cmaxtext2 character varying,cmintext text,imaxrank integer,iminrank integer,igaptext text,igaptext2 character varying,icount integer,customerdemand integer,ishipvs1stietd integer,ishortline integer,shipvscietd integer,ifail1stconf integer,ic1 integer,ic3 integer,il2andsasladjust integer,il4andsasladjust integer,saslscoreboard integer,il2minsasladjust integer,il4minsasladjust integer,il1plusl2andsasladjust integer,il3plusl4adjust integer,ishipvs1stietd_weight numeric,shipvscietd_weight numeric,weight numeric)"
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "SASLSSLChart" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim datasheet As Integer = 3
            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.77\Users_I\Logistic Dept\KPI & Reporting\templates\SASL_SSL_Pie_Chart_FG_Template.xltx")
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        owb.Names.Add("kpidata", RefersToR1C1:="=OFFSET('Db'!R1C1,0,0,COUNTA('Db'!C1),COUNTA('DB'!R1))")
        owb.Worksheets(2).select()
        Dim osheet = owb.Worksheets(2)
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        osheet.PivotTables("PivotTable3").PivotCache.Refresh()
        osheet.PivotTables("PivotTable6").PivotCache.Refresh()
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        osheet.PivotTables("PivotTable5").PivotCache.Refresh()
    End Sub
End Class