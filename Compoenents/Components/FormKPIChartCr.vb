Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class FormKPIChartCr
    Dim PG As Integer = 1
    Dim mythread As New Thread(AddressOf doInitialObjects)
    Dim DoWorkThread As New Thread(AddressOf DoWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim DS As DataSet
    Dim fieldname As String
    Dim title As String
    Dim criteria As String = String.Empty
    Dim FilenameSelected As String = String.Empty

    Dim myobj As New CheckedListBox
    Dim selectedCheckedListBox As Integer = 1

    Dim date1 As Date
    Dim date2 As Date
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty



        'vendor
        'sp
        'bu/Factory



        Dim reportadditionalname As String = String.Empty
        If RadioButton3.Checked Then
            myobj = CheckedListBox1
            criteria = " and pv.vendorcode = "
        ElseIf RadioButton4.Checked Then
            myobj = CheckedListBox2
            If PG = 1 Then
                criteria = " and (vp.getspspm1(v.vendorcode,s.sbuid::bigint)).sp = "
            Else
                criteria = " and vsc.sp = "
            End If

            selectedCheckedListBox = 2
        ElseIf RadioButton5.Checked Then
            myobj = CheckedListBox3
            If PG = 1 Then
                criteria = " and sbu1.sbuname = "
                fieldname = "BU Name"
            Else
                criteria = " and c1.customercode = "
                fieldname = "Factory"
            End If
            selectedCheckedListBox = 3
        End If


        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog

        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            date1 = DateTimePicker1.Value
            date2 = DateTimePicker2.Value
            FilenameSelected = DirectoryBrowser.SelectedPath

            If Not DoWorkThread.IsAlive Then
                ToolStripStatusLabel1.Text = ""
                DoWorkThread = New Thread(AddressOf DoWork)
                DoWorkThread.Start()
            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
            


        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        ProgressReport(2, "Add KPI Data. =OFFSET(Db!R1C1,0,0,COUNTA(Db!C1),COUNTA(Db!R1))")
        owb.Names.Add("kpidata", RefersToR1C1:="=OFFSET(Db!R1C1,0,0,COUNTA(Db!C1),COUNTA(Db!R1))")
        'owb.Names.Add("kpidata", RefersToR1C1:="=OFFSET(INDIRECT(ADDRESS(1,1)),0,0,COUNTA(Db!C1),COUNTA(Db!R1))")
        ProgressReport(2, "After Add KPI Data.")
        owb.Worksheets(2).select()
        Dim osheet = owb.Worksheets(2)
        ProgressReport(2, "PivotTable1 Refresh Data.")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        ProgressReport(2, "PivotTable3 Refresh Data.")
        osheet.PivotTables("PivotTable3").PivotCache.Refresh()
        ProgressReport(2, "PivotTable6 Refresh Data.")
        osheet.PivotTables("PivotTable6").PivotCache.Refresh()
        ProgressReport(2, "PivotTable2 Refresh Data.")
        osheet.PivotTables("PivotTable2").PivotCache.Refresh()
        ProgressReport(2, "PivotTable5 Refresh Data.")
        osheet.PivotTables("PivotTable5").PivotCache.Refresh()
        osheet = owb.Worksheets(1)
        osheet.cells(3, 6) = fieldname
        osheet.cells(3, 7) = title
        ProgressReport(1, "String Format.")
        osheet.cells(4, 7) = String.Format("{0:dd-MMM-yyyy} to {1:dd-MMM-yyyy}", date1, date2)
    End Sub


    Private Sub InitialObjects()
        If Not mythread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            mythread = New Thread(AddressOf doInitialObjects)
            mythread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If


    End Sub
    Sub doInitialObjects()
        '
        'get supplier for selected pg
        'get sp for selected pg
        'get bu for selected pg
        DS = New DataSet
        Dim sqlstr As New StringBuilder
        ProgressReport(1, "Loading Data. Please wait!")

        If PG = 1 Then
            sqlstr.Append("select 'Select All' as vendorname,0 as vendorcode union all select 'All Supplier' as vendorname,1 as vendorcode union all (select distinct v.vendorname::character varying, v.vendorcode from vp.vendorbusp vp " &
            " left join vendor v on v.vendorcode = vp.vendorcode order by v.vendorname::character varying);")
            sqlstr.Append("select 'Select All' as officersebname union all select 'All SP' as officersebname union all (select officersebname from (select distinct spid from vp.vendorbusp v" &
                 " where not spid isnull)  as foo" &
                 " left join officerseb of on of.ofsebid = foo.spid" &
                 " order by officersebname);")
            sqlstr.Append("select 'Sellect All' as sbuname,0 as sbuid union all select 'All BU' as sbuname,1 as sbuid union all (select sbuname::character varying,sbuid from sbu where sp order by sbuname);")
        Else
            sqlstr.Append("select 'Select All' as vendorname,0 as vendorcode union all select 'All Supplier' as vendorname,1 as vendorcode union all  (select distinct v.vendorname, v.vendorcode from vendorspcomp vp" &
                " left join vendor v on v.vendorcode = vp.vendorcode order by v.vendorname);")
            sqlstr.Append("select 'Select All' as officersebname union all select 'All SP' as officersebname union all  (select distinct sp as officersebname from vendorspcomp order by sp);")
            sqlstr.Append("select 0::bigint as customercode,'Select All' as customername union all select 1::bigint as customercode,'All Factory' as customername union all (select distinct pd.shiptoparty as customercode,c.customername " &
                 " from cxsebpodtl pd" &
                 " left join ekko po on po.po = pd.sebasiapono" &
                 " left join purchasinggroup pg on pg.purchasinggroup = po.purchasinggroup" &
                 " left join customer c on c.customercode = pd.shiptoparty" &
                 " where pg.groupsbuid = 2 order by c.customername);") 'groupsbu allways 2

        End If


        Dim message As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr.ToString, DS, message) Then
            ProgressReport(4, "Fill CheckListBox")
            ProgressReport(1, "Loading Data. Done.")
        Else
            ProgressReport(1, message)
        End If


    End Sub


    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 4
                    With CheckedListBox1                        
                        .DataSource = DS.Tables(0)
                        .DisplayMember = "vendorname"
                        .ValueMember = "vendorcode"
                    End With
                    With CheckedListBox2
                        .DataSource = DS.Tables(1)
                        .DisplayMember = "officersebname"
                        .ValueMember = "officersebname"
                    End With

                    With CheckedListBox3
                        .DataSource = DS.Tables(2)
                        If PG = 1 Then
                            .DisplayMember = "sbuname"
                            .ValueMember = "sbuid"
                            RadioButton5.Text = "BU"
                        Else
                            .DisplayMember = "customername"
                            .ValueMember = "customercode"
                            RadioButton5.Text = "Factory"
                        End If

                        
                    End With
            End Select
        End If

    End Sub


    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged, RadioButton4.CheckedChanged, RadioButton5.CheckedChanged
        Dim obj = DirectCast(sender, RadioButton)
        If obj.Checked Then
            Select Case obj.Name
                Case "RadioButton3"
                    CheckedListBox1.Enabled = True
                    CheckedListBox2.Enabled = False
                    CheckedListBox3.Enabled = False
                Case "RadioButton4"
                    CheckedListBox1.Enabled = False
                    CheckedListBox3.Enabled = False
                    CheckedListBox2.Enabled = True
                Case "RadioButton5"
                    CheckedListBox1.Enabled = False
                    CheckedListBox2.Enabled = False
                    CheckedListBox3.Enabled = True

            End Select
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        Dim obj = DirectCast(sender, RadioButton)
        If obj.Checked Then
            Select Case obj.Text
                Case "Finished Goods"
                    PG = 1
                Case "Components"
                    PG = 2
            End Select
            CheckedListBox1.DataSource = Nothing
            CheckedListBox2.DataSource = Nothing
            CheckedListBox3.DataSource = Nothing
            InitialObjects()
        End If
    End Sub



    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged, CheckedListBox2.SelectedIndexChanged, CheckedListBox3.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)

    End Sub

    Sub DoWork()
        For i = 1 To myobj.Items.Count - 1
            If myobj.GetItemChecked(i) Then
                Dim drv As DataRowView = myobj.Items(i)
                Dim mycriteria As String = String.Empty
                Select Case selectedCheckedListBox
                    Case 1
                        mycriteria = criteria & drv.Row("vendorcode")
                        title = drv.Row("vendorname")
                        fieldname = "Supplier"

                    Case 2
                        mycriteria = criteria & "''" & drv.Row(0) & "''"
                        title = drv.Row(0)
                        fieldname = "Supply Planner"
                    Case 3
                        mycriteria = criteria & "''" & drv.Row(0) & "''"
                        title = drv.Row(1)
                End Select
                If i = 1 Then
                    mycriteria = ""
                    title = "ALL Data"
                End If

                Dim sqlstr As String
                If PG = 1 Then
                    sqlstr = "select * from sp_kpidata(" & DateFormatyyyyMMdd(date1) & "," & DateFormatyyyyMMdd(date2) & ",1,'" & mycriteria & "') as tb(ordertype character varying,shiptoparty bigint, shiptopartyname character(50),soldtoparty bigint, soldtopartyname character(50), customerorderno character varying, sebasiasalesorder bigint," &
                                       "solineno integer, vendorcode bigint,vendorname character(50),rir character(2), cmmf bigint,comfam integer,itemid character(15),materialdesc character(50), orderstatus character varying," &
                                       "latestupdate date, updatesince character varying, curinq character varying,receptiondate date, fob real,unittp real, inquiryeta date,inquiryetd date, inquiryqty integer, " &
                                       "currentinquiryetd date, currentinquiryqty integer, confirmationstatus character varying, stconfirmedetd date,stconfirmedqty integer,currentconfirmedeta date, currentconfirmedetd date," &
                                       "currentconfirmedqty integer, deliveredqty integer, shipdate date,shipdateeta date,osqty integer, sebasiapono bigint, polineno integer,ctrno character varying, boatid character varying," &
                                       "packinglist character varying,shipfrom character varying, comments character varying, cmnttxdtlname character(50),sao character varying, purchasinggroup character varying, " &
                                       "sbu character(30),status text, shipmentline integer,week double precision,shipdate2 date, bu character(30),familysbu character(30), sp text,spm text,cmaxtext text,cmaxtext2 text,cmintext text,imaxrank integer,iminrank integer,igaptext text,igaptext2 text,icount integer,customerdemand integer,ishipvs1stietd integer,ishortline integer,shipvscietd integer,ifail1stconf integer,ic1 integer,ic3 integer,il2andsasladjust integer,il4andsasladjust integer,saslscoreboard integer,il2minsasladjust integer,il4minsasladjust integer,il1plusl2andsasladjust integer,il3plusl4adjust integer,ishipvs1stietd_weight numeric,shipvscietd_weight numeric,weight numeric,saslscoreboard_weight numeric)"

                Else
                    sqlstr = "select * from sp_kpidatacomp(" & DateFormatyyyyMMdd(date1) & "," & DateFormatyyyyMMdd(date2) & ",'" & mycriteria & "') as tb(ordertype character varying,shiptoparty bigint, shiptopartyname character(50),soldtoparty bigint, soldtopartyname character(50), customerorderno character varying, sebasiasalesorder bigint," &
                                       "solineno integer, vendorcode bigint,vendorname character(50),rir character(2), cmmf bigint,comfam integer,itemid character(15),materialdesc character(50), orderstatus character varying," &
                                       "latestupdate date, updatesince character varying, curinq character varying,receptiondate date, fob real,unittp real, inquiryeta date,inquiryetd date, inquiryqty integer, " &
                                       "currentinquiryetd date, currentinquiryqty integer, confirmationstatus character varying, stconfirmedetd date,stconfirmedqty integer,currentconfirmedeta date, currentconfirmedetd date," &
                                       "currentconfirmedqty integer, deliveredqty integer, shipdate date,shipdateeta date,osqty integer, sebasiapono bigint, polineno integer,ctrno character varying, boatid character varying," &
                                       "packinglist character varying,shipfrom character varying, comments character varying, cmnttxdtlname character(50),sao character varying, purchasinggroup character varying, " &
                                       "sbu character(30),status text, shipmentline integer,week double precision,shipdate2 date, bu character(30),familysbu character(30), sp character varying,cmaxtext text,cmaxtext2 text,cmintext text,imaxrank integer,iminrank integer,igaptext text,igaptext2 text,icount integer,customerdemand integer,ishipvs1stietd integer,ishortline integer,shipvscietd integer,ifail1stconf integer,ic1 integer,ic3 integer,il2andsasladjust integer,il4andsasladjust integer,saslscoreboard integer,il2minsasladjust integer,il4minsasladjust integer,il1plusl2andsasladjust integer,il3plusl4adjust integer,ishipvs1stietd_weight numeric,shipvscietd_weight numeric,weight numeric,saslscoreboard_weight numeric)"

                End If

                Dim filename = FilenameSelected 'Application.StartupPath & "\PrintOut"
                Dim reportname = "SASLSSLChart" & "-" & title '& GetCompanyName()
                Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
                Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
                Dim datasheet As Integer = 3
                Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\\172.22.10.77\Users_I\Logistic Dept\KPI & Reporting\templates\SASL_SSL_Pie_Chart_FG_Template.xltx")

                'To avoid heavy workload in server, do not using Thread
                'myreport.Run(Me, New System.EventArgs)
                myreport.DoWork()


            End If
        Next
    End Sub


End Class