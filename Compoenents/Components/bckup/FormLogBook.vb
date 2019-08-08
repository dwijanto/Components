Imports System.Threading
Imports Components.SharedClass
Imports Components.PublicClass

Public Class FormLogBook
    Dim myThreadDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myWorkDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myWork As New System.Threading.Thread(myWorkDelegate)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim combobs As BindingSource
    Dim startdate As Date
    Dim enddate As Date
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim bs1 As BindingSource
    Dim bs2 As BindingSource
    Dim bs3 As BindingSource
    Dim bs4 As BindingSource
    Dim bs5 As BindingSource
    Dim bs6 As BindingSource

    Dim myuser As String = String.Empty
    Dim myOfficer As String = String.Empty
    Dim MyDS As DataSet
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        With startdateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        With enddateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        Dim host1 = New ToolStripControlHost(startdateDTP)
        Dim host2 = New ToolStripControlHost(enddateDTP)
        ToolStrip1.Items.Insert(3, host1)
        ToolStrip1.Items.Insert(5, host2)
        myThread.Start()

    End Sub

    Sub DoQuery()
        'Get All user from PackingListDtl
        Dim sqlstr = "select ''::text as username union all (select distinct username from accountinghd order by username);"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            combobs = New BindingSource
            combobs.DataSource = DS.Tables(0)
            If DS.Tables(0).Rows.Count > 0 Then

                ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
            End If

        End If
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
                    Me.ToolStripComboBox1.ComboBox.DataSource = combobs
                    Me.ToolStripComboBox1.ComboBox.DisplayMember = "username"
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
                    'Fill DataGridView

                    DataGridView1.AutoGenerateColumns = False
                    DataGridView2.AutoGenerateColumns = False
                    DataGridView3.AutoGenerateColumns = False
                    DataGridView4.AutoGenerateColumns = False
                    DataGridView5.AutoGenerateColumns = False
                    DataGridView6.AutoGenerateColumns = False

                    DataGridView1.DataSource = bs1
                    DataGridView2.DataSource = bs2
                    DataGridView3.DataSource = bs3
                    DataGridView4.DataSource = bs4
                    DataGridView5.DataSource = bs5
                    DataGridView6.DataSource = bs6
                    'Label4.Text = "Record" & IIf(CType(bs1.DataSource, DataTable).Rows.Count > 1, "s", "") & " Found :" & CType(bs1.DataSource, DataTable).Rows.Count.ToString
                    displayrecordcount()
                    
            End Select

        End If

    End Sub


    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If myWork.IsAlive Then
            MessageBox.Show("Process still running in background. Please wait...")
            Exit Sub
        End If
        myuser = ""
        myOfficer = ""
        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""

        If ToolStripComboBox1.Text <> "" Then
            myuser = " and username = '" & ToolStripComboBox1.Text & "'"
            myOfficer = " and officer =  '" & ToolStripComboBox1.Text & "'"
        End If
        startdate = startdateDTP.Value.Date
        enddate = enddateDTP.Value.Date
        myWork = New Thread(AddressOf DoWork)
        myWork.Start()
        'For i = 0 To DataGridView3.Columns.Count - 1
        '    MessageBox.Show(i & " " & DataGridView3.Columns(i).Width)
        'Next
    End Sub

    Sub DoWork()

        Dim sqlstrAccountingNonPartial = "select * from (select docno::character varying,myyear,postingdate,username,reference,mironumber::character varying,pohd::character varying,polineno,cmmf::character varying,qty,amount,ah.crcy,ah.exrate,e.vendorcode::character varying,vendorname::character varying from accountinghd ah" &
                                         " inner join miro m on m.mironumber = ah.miro" &
                                         " left join pomiro pm on pm.miroid = m.miroid" &
                                         " left join podtl pd on pd.podtlid = pm.podtlid" &
                                         " left join ekko e on e.po = pd.pohd" &
                                         " left join vendor v on v.vendorcode = e.vendorcode" &
                                         " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser &
                                         " ) foo inner Join " &
                                         " (select pohd,poitem,max(plh.delivery)as delivery,max(deliveryitem) as deliveryitem" &
                                         " from packinglisthd plh" &
                                         " left join packinglistdt pld on plh.delivery = pld.delivery" &
                                         " group by pohd,poitem" &
                                         " having(count(pohd) = 1)" &
                                         " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno" &
                                         " order by docno;"
        'Dim sqlstrAccountingPartial = "select * from (select docno::character varying,myyear,postingdate,username,reference,mironumber::character varying,pohd::character varying,polineno,cmmf::character varying,amount,ah.crcy,ah.exrate,e.vendorcode::character varying,vendorname::character varying from accountinghd ah" &
        '                                 " inner join miro m on m.mironumber = ah.miro" &
        '                                 " left join pomiro pm on pm.miroid = m.miroid" &
        '                                 " left join podtl pd on pd.podtlid = pm.podtlid" &
        '                                 " left join ekko e on e.po = pd.pohd" &
        '                                 " left join vendor v on v.vendorcode = e.vendorcode" &
        '                                 " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser &
        '                                 " ) foo inner Join " &
        '                                 " (select pohd,poitem,max(plh.delivery)as delivery,max(deliveryitem) as deliveryitem" &
        '                                 " from packinglisthd plh" &
        '                                 " left join packinglistdt pld on plh.delivery = pld.delivery" &
        '                                 " group by pohd,poitem" &
        '                                 " having(count(pohd) > 1)" &
        '                                 " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno" &
        '                                 " order by docno;"
        Dim sqlstrAccountingPartial = "select foo.docno::character varying,myyear,postingdate,username,reference,mironumber::character varying,foo.pohd::character varying,polineno,cmmf::character varying,qty,amount,crcy,exrate,vendorcode::character varying,vendorname::character varying,delivery,item as deliveryitem from " &
                                        " (select docno::character varying,myyear,postingdate,username,reference,mironumber::character varying,pohd::character varying,polineno,cmmf::character varying,qty,amount,ah.crcy,ah.exrate,e.vendorcode::character varying,vendorname::character varying from accountinghd ah " &
                                        " inner join miro m on m.mironumber = ah.miro " &
                                        " left join pomiro pm on pm.miroid = m.miroid " &
                                        " left join podtl pd on pd.podtlid = pm.podtlid " &
                                        " left join ekko e on e.po = pd.pohd " &
                                        " left join vendor v on v.vendorcode = e.vendorcode " &
                                        " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & ") foo " &
                                        " inner Join" &
                                        " (select pohd,poitem " &
                                        " from packinglisthd plh " &
                                        " left join packinglistdt pld on plh.delivery = pld.delivery " &
                                        " group by pohd,poitem having(count(pohd) > 1) " &
                                        " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno " &
                                        " left join packinglistdocument pc on pc.docno::character varying = foo.docno and pc.pohd::character varying = foo.pohd and pc.poitem = foo.polineno" &
                                        " order by foo.docno;"


        Dim sqlstrBillingNonPartial = "select foo2.billingdocument::character varying,foo2.billingtype,foo2.createdon,br.status, foo2.salesdoc::character varying,foo2.salesdocitem,foo2.sebasiapono::character varying,foo2.polineno,delivery::character varying,deliveryitem,deliveredqty,billedqty,requiredqty,officer" &
                                        " from " &
                                        " (select bh.billingdocument,billingtype,salesdoc,salesdocitem,sebasiapono,polineno,createdon,officer,billedqty,requiredqty" &
                                        " from billinghd bh" &
                                        " left join billingdtl bt on bt.billingdocument = bh.billingdocument" &
                                        " left join cxsalesorderdtl sd on sd.sebasiasalesorder = bt.salesdoc and sd.solineno = bt.salesdocitem" &
                                        " left join cxrelsalesdocpo cr on cr.cxsalesorderdtlid = sd.cxsalesorderdtlid" &
                                        " left join cxsebpodtl cpd on cpd.cxsebpodtlid = cr.cxsebpodtlid" &
                                        " where not sebasiapono isnull and bt.requiredqty - bt.billedqty = 0 and createdon >= " & DateFormatyyyyMMdd(startdate) & " and createdon <=  " & DateFormatyyyyMMdd(enddate) & myOfficer & ") as foo2 " &
                                        " left join (select plh.delivery,pld.deliveryitem,pohd,poitem,deliveredqty,housebill,createdby from packinglistdt pld" &
                                        " left join packinglisthd plh on plh.delivery = pld.delivery" &
                                        " left join housebill h on h.po = pld.pohd and h.containerno = plh.biloflading ) foo3 on foo3.pohd = foo2.sebasiapono and foo3.poitem = foo2.polineno" &
                                        " left join billingdocreversal br on br.billingdoc = foo2.billingdocument and br.salesdoc = foo2.salesdoc and br.item = foo2.salesdocitem" &
                                        " order by billingdocument;"


        
        Dim sqlstrBillingPartial = "select foo2.billingdocument::character varying,foo2.billingtype,foo2.createdon,br.status, foo2.salesdoc::character varying,foo2.salesdocitem,foo2.sebasiapono::character varying,foo2.polineno,foo3.delivery::character varying,deliveryitem,deliveredqty,billedqty,requiredqty,officer" &
                                        " from " &
                                        " (select bh.billingdocument,billingtype,salesdoc,salesdocitem,sebasiapono,polineno,createdon,officer,billedqty,requiredqty" &
                                        " from billinghd bh" &
                                        " left join billingdtl bt on bt.billingdocument = bh.billingdocument" &
                                        " left join cxsalesorderdtl sd on sd.sebasiasalesorder = bt.salesdoc and sd.solineno = bt.salesdocitem" &
                                        " left join cxrelsalesdocpo cr on cr.cxsalesorderdtlid = sd.cxsalesorderdtlid" &
                                        " left join cxsebpodtl cpd on cpd.cxsebpodtlid = cr.cxsebpodtlid" &
                                        " where not sebasiapono isnull and bt.requiredqty - bt.billedqty <> 0 and createdon >= " & DateFormatyyyyMMdd(startdate) & " and createdon <=  " & DateFormatyyyyMMdd(enddate) & myOfficer & ") as foo2 " &
                                        " left join billingdocreversal br on br.billingdoc = foo2.billingdocument and br.salesdoc = foo2.salesdoc and br.item = foo2.salesdocitem" &
                                        " left join packinglistdocument plc on plc.docno = foo2.billingdocument and plc.pohd = foo2.sebasiapono and plc.poitem = foo2.polineno" &
                                        " left join (select plh.delivery,pld.deliveryitem,pohd,poitem,deliveredqty,housebill,createdby from packinglistdt pld" &
                                        " left join packinglisthd plh on plh.delivery = pld.delivery" &
                                        " left join housebill h on h.po = pld.pohd and h.containerno = plh.biloflading ) foo3 on foo3.delivery = plc.delivery and foo3.deliveryitem = plc.item" &
                                        " order by billingdocument;"

        Dim sqlstrMiroNonPartial = "select foo.mironumber::character varying,foo.miropostingdate,foo.supplierinvoicenum::character varying,foo.pohd::character varying,foo.polineno,foo.vendorcode::character varying,foo.vendorname::character varying,foo.cmmf::character varying,foo.materialdesc,foo.amount,foo.qty,foo.unitprice,foo.crcy,foo.username,foo2.delivery::character varying, foo2.deliveryitem,foo2.shipdate from (select m.mironumber,m.miropostingdate,m.supplierinvoicenum,pd.pohd,pd.polineno,m.vendorcode,v.vendorname,pd.cmmf,c.materialdesc,pm.amount,pm.qty,pm.unitprice,pm.crcy,mu.username from miro m" &
                                   " left join mirouser mu on mu.miro = m.mironumber" &
                                   " left join pomiro pm on pm.miroid = m.miroid" &
                                   " left join podtl pd on pd.podtlid = pm.podtlid" &
                                   " left join pohd ph on ph.pohd = pd.pohd" &
                                   " left join vendor v on v.vendorcode = m.vendorcode" &
                                   " left join cmmf c on c.cmmf = pd.cmmf" &
                                   " where miropostingdate>= " & DateFormatyyyyMMdd(startdate) & " and miropostingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & " ) foo" &
                                   " inner Join" &
                                   " (select pohd,poitem,max(plh.delivery)as delivery,max(deliveryitem) as deliveryitem ,max(deliverydate) as shipdate" &
                                   " from packinglisthd plh " &
                                   " left join packinglistdt pld on plh.delivery = pld.delivery " &
                                   " group by pohd,poitem having(count(pohd) = 1) order by pohd,poitem ) foo2 " &
                                   " on foo2.pohd::character varying = foo.pohd::character varying and foo2.poitem = foo.polineno order by foo.mironumber;"

        Dim sqlstrMiroPartial = "select foo.mironumber::character varying,foo.miropostingdate,foo.supplierinvoicenum::character varying,foo.pohd::character varying,foo.polineno,foo.vendorcode::character varying,foo.vendorname::character varying,foo.cmmf::character varying,foo.materialdesc,foo.amount,foo.qty,foo.unitprice,foo.crcy,foo.username,pc.delivery::character varying, pc.item from (select m.mironumber,m.miropostingdate,m.supplierinvoicenum,pd.pohd,pd.polineno,m.vendorcode,v.vendorname,pd.cmmf,c.materialdesc,pm.amount,pm.qty,pm.unitprice,pm.crcy,mu.username from miro m" &
                                   " left join mirouser mu on mu.miro = m.mironumber" &
                                   " left join pomiro pm on pm.miroid = m.miroid" &
                                   " left join podtl pd on pd.podtlid = pm.podtlid" &
                                   " left join pohd ph on ph.pohd = pd.pohd" &
                                   " left join vendor v on v.vendorcode = m.vendorcode" &
                                   " left join cmmf c on c.cmmf = pd.cmmf" &
                                   " where miropostingdate>= " & DateFormatyyyyMMdd(startdate) & " and miropostingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & " ) foo" &
                                   " inner Join" &
                                   " (select pohd,poitem " &
                                   " from packinglisthd plh " &
                                   " left join packinglistdt pld on plh.delivery = pld.delivery " &
                                   " group by pohd,poitem having(count(pohd) > 1) " &
                                   " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd::character varying and foo2.poitem = foo.polineno " &
                                   " left join packinglistdocument pc on pc.docno::character varying = foo.mironumber::character varying and pc.pohd::character varying = foo.pohd::character varying and pc.poitem = foo.polineno" &
                                   " order by foo.mironumber;"

        Dim sqlstr = sqlstrAccountingNonPartial & sqlstrAccountingPartial & sqlstrBillingNonPartial & sqlstrBillingPartial & sqlstrMiroNonPartial & sqlstrMiroPartial

        'Dim MyDS As New DataSet
        myds = New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(6, "Marque")

        If Not DbAdapter1.TbgetDataSet(sqlstr, MyDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            bs1 = New BindingSource
            bs2 = New BindingSource
            bs3 = New BindingSource
            bs4 = New BindingSource
            bs5 = New BindingSource
            bs6 = New BindingSource
            bs1.DataSource = MyDS.Tables(0)
            bs2.DataSource = MyDS.Tables(1)
            bs3.DataSource = MyDS.Tables(2)
            bs4.DataSource = MyDS.Tables(3)
            bs5.DataSource = MyDS.Tables(4)
            bs6.DataSource = MyDS.Tables(5)

            ProgressReport(8, "Fill DataGridView")
        End If
        'ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", MyDS.Tables(0).Rows(0).Item(0)))
        ProgressReport(5, "Continues")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If IsNothing(bs4) Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        'Dim sqlstr = "select sebasiapono,polineno,billedqty,delivery,deliveryitem,deliveredqty from (select foo2.sebasiapono::character varying,foo2.polineno,billedqty" &
        '    " from" &
        '    " (select bh.billingdocument,salesdoc,salesdocitem,sebasiapono,polineno,billedqty" &
        '    " from billinghd bh " &
        '    " left join billingdtl bt on bt.billingdocument = bh.billingdocument " &
        '    " left join cxsalesorderdtl sd on sd.sebasiasalesorder = bt.salesdoc and sd.solineno = bt.salesdocitem " &
        '    " left join cxrelsalesdocpo cr on cr.cxsalesorderdtlid = sd.cxsalesorderdtlid " &
        '    " left join cxsebpodtl cpd on cpd.cxsebpodtlid = cr.cxsebpodtlid " &
        '    " where not sebasiapono isnull and bt.requiredqty - bt.billedqty <> 0 and createdon >= " & DateFormatyyyyMMdd(startdate) & " and createdon <=  " & DateFormatyyyyMMdd(enddate) & myOfficer & ") as foo2  " &
        '    " left join billingdocreversal br on br.billingdoc = foo2.billingdocument and br.salesdoc = foo2.salesdoc and br.item = foo2.salesdocitem " &
        '    " where status isnull" &
        '    " group by sebasiapono,polineno,billedqty" &
        '    " having count(sebasiapono) = 1)foo3" &
        '    " left join packinglistdt pld on pld.pohd = foo3.sebasiapono::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.billedqty"
        Dim sqlstr = "select sebasiapono,polineno,billedqty,max(delivery) as delivery,max(deliveryitem) as deliveryitem,max(deliveredqty) as deliveredqty from (select foo2.sebasiapono::character varying,foo2.polineno,billedqty" &
            " from" &
            " (select bh.billingdocument,salesdoc,salesdocitem,sebasiapono,polineno,billedqty" &
            " from billinghd bh " &
            " left join billingdtl bt on bt.billingdocument = bh.billingdocument " &
            " left join cxsalesorderdtl sd on sd.sebasiasalesorder = bt.salesdoc and sd.solineno = bt.salesdocitem " &
            " left join cxrelsalesdocpo cr on cr.cxsalesorderdtlid = sd.cxsalesorderdtlid " &
            " left join cxsebpodtl cpd on cpd.cxsebpodtlid = cr.cxsebpodtlid " &
            " where not sebasiapono isnull and bt.requiredqty - bt.billedqty <> 0 and createdon >= " & DateFormatyyyyMMdd(startdate) & " and createdon <=  " & DateFormatyyyyMMdd(enddate) & myOfficer & ") as foo2  " &
            " left join billingdocreversal br on br.billingdoc = foo2.billingdocument and br.salesdoc = foo2.salesdoc and br.item = foo2.salesdocitem " &
            " where status isnull" &
            " group by sebasiapono,polineno,billedqty" &
            " having count(sebasiapono) = 1)foo3" &
            " left join packinglistdt pld on pld.pohd = foo3.sebasiapono::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.billedqty" &
            " group by foo3.sebasiapono,polineno,billedqty having (count(foo3.sebasiapono) = 1) "

        Dim myds As New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, myds, mymessage) Then
            Dim pk(2) As DataColumn
            pk(0) = myds.Tables(0).Columns(0)
            pk(1) = myds.Tables(0).Columns(1)
            pk(2) = myds.Tables(0).Columns(2)
            myds.Tables(0).PrimaryKey = pk

            For i = 0 To bs4.Count - 1
                Dim dr As DataRow = CType(bs4.Item(i), DataRowView).Row
                Dim mykey(2) As Object
                mykey(0) = dr.Item("sebasiapono")
                mykey(1) = dr.Item("polineno")
                mykey(2) = dr.Item("billedqty")
                Dim result As DataRow = myds.Tables(0).Rows.Find(mykey)
                If Not IsNothing(result) Then
                    dr.Item("delivery") = result.Item("delivery")
                    dr.Item("deliveryitem") = result.Item("deliveryitem")
                    dr.Item("deliveredqty") = result.Item("deliveredqty")
                End If
            Next
        Else
            MessageBox.Show(mymessage)
        End If

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsNothing(bs2) Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        'Dim sqlstr As String = "select foo3.pohd,polineno,qty,delivery,deliveryitem from (select foo.pohd::character varying,polineno,qty " &
        '                       " from" &
        '                       " (select docno::character varying,pohd::character varying,polineno,qty" &
        '                       " from accountinghd ah  " &
        '                       " inner join miro m on m.mironumber = ah.miro  " &
        '                       " left join pomiro pm on pm.miroid = m.miroid  " &
        '                       " left join podtl pd on pd.podtlid = pm.podtlid  " &
        '                       " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & ") foo  " &
        '                       "    inner Join" &
        '                       " (select pohd,poitem  from packinglisthd plh " &
        '                       " left join packinglistdt pld on plh.delivery = pld.delivery  " &
        '                       " group by pohd,poitem having(count(pohd) > 1)  " &
        '                       " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno  " &
        '                       " group by foo.pohd,polineno,qty" &
        '                       " having count(foo.pohd) = 1)foo3" &
        '                       " left join packinglistdt pld on pld.pohd = foo3.pohd::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.qty"
        Dim sqlstr As String = "select foo3.pohd,polineno,qty,max(delivery) as delivery,max(deliveryitem) as deliveryitem from (select foo.pohd::character varying,polineno,qty " &
                       " from" &
                       " (select docno::character varying,pohd::character varying,polineno,qty" &
                       " from accountinghd ah  " &
                       " inner join miro m on m.mironumber = ah.miro  " &
                       " left join pomiro pm on pm.miroid = m.miroid  " &
                       " left join podtl pd on pd.podtlid = pm.podtlid  " &
                       " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & ") foo  " &
                       "    inner Join" &
                       " (select pohd,poitem  from packinglisthd plh " &
                       " left join packinglistdt pld on plh.delivery = pld.delivery  " &
                       " group by pohd,poitem having(count(pohd) > 1)  " &
                       " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno  " &
                       " group by foo.pohd,polineno,qty" &
                       " having count(foo.pohd) = 1)foo3" &
                       " left join packinglistdt pld on pld.pohd = foo3.pohd::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.qty" &
                       " group by foo3.pohd,polineno,qty having (count(foo3.pohd) = 1) "
        Dim myds As New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, myds, mymessage) Then
            Dim pk(2) As DataColumn
            pk(0) = myds.Tables(0).Columns(0)
            pk(1) = myds.Tables(0).Columns(1)
            pk(2) = myds.Tables(0).Columns(2)
            myds.Tables(0).PrimaryKey = pk

            For i = 0 To bs2.Count - 1
                Dim dr As DataRow = CType(bs2.Item(i), DataRowView).Row
                Dim mykey(2) As Object
                mykey(0) = dr.Item("pohd")
                mykey(1) = dr.Item("polineno")
                mykey(2) = dr.Item("qty")                
                Dim result As DataRow = myds.Tables(0).Rows.Find(mykey)
                If Not IsNothing(result) Then
                    dr.Item("delivery") = result.Item("delivery")
                    dr.Item("deliveryitem") = result.Item("deliveryitem")
                End If
            Next
        Else
            MessageBox.Show(mymessage)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If IsNothing(bs6) Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
       
        'Dim sqlstr As String = "select foo3.pohd,polineno,qty,max(delivery) as delivery,max(deliveryitem) as deliveryitem from (select foo.pohd::character varying,polineno,qty " &
        '               " from" &
        '               " (select docno::character varying,pohd::character varying,polineno,qty" &
        '               " from accountinghd ah  " &
        '               " inner join miro m on m.mironumber = ah.miro  " &
        '               " left join pomiro pm on pm.miroid = m.miroid  " &
        '               " left join podtl pd on pd.podtlid = pm.podtlid  " &
        '               " where not cmmf isnull  and postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & ") foo  " &
        '               "    inner Join" &
        '               " (select pohd,poitem  from packinglisthd plh " &
        '               " left join packinglistdt pld on plh.delivery = pld.delivery  " &
        '               " group by pohd,poitem having(count(pohd) > 1)  " &
        '               " order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno  " &
        '               " group by foo.pohd,polineno,qty" &
        '               " having count(foo.pohd) = 1)foo3" &
        '               " left join packinglistdt pld on pld.pohd = foo3.pohd::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.qty" &
        '               " group by foo3.pohd,polineno,qty having (count(foo3.pohd) = 1) "
        Dim sqlstr As String = "select foo3.pohd,polineno,qty,max(delivery) as delivery,max(deliveryitem) as deliveryitem " &
                               " from " &
                               " (select foo.pohd::character varying,polineno,qty" &
                               " from " &
                               " (select mironumber::character varying,pohd::character varying,polineno,qty " &
                               " from miro m " &
                               " left join mirouser mu on mu.miro = m.mironumber " &
                               " left join pomiro pm on pm.miroid = m.miroid " &
                               " left join podtl pd on pd.podtlid = pm.podtlid   " &
                               " where not cmmf isnull  and miropostingdate  >= " & DateFormatyyyyMMdd(startdate) & " and miropostingdate <= " & DateFormatyyyyMMdd(enddate) & myuser & ") foo " &
                               " inner Join " &
                               " (select pohd,poitem  " &
                               " from packinglisthd plh  " &
                               " left join packinglistdt pld on plh.delivery = pld.delivery   " &
                               " group by pohd,poitem having(count(pohd) > 1)   order by pohd,poitem ) foo2 on foo2.pohd::character varying = foo.pohd and foo2.poitem = foo.polineno   " &
                               " group by foo.pohd,polineno,qty having count(foo.pohd) = 1)foo3 " &
                               " left join packinglistdt pld on pld.pohd = foo3.pohd::bigint and pld.poitem = foo3.polineno and pld.deliveredqty = foo3.qty " &
                               " group by foo3.pohd,polineno,qty having (count(foo3.pohd) = 1) "
        Dim myds As New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, myds, mymessage) Then
            Dim pk(2) As DataColumn
            pk(0) = myds.Tables(0).Columns(0)
            pk(1) = myds.Tables(0).Columns(1)
            pk(2) = myds.Tables(0).Columns(2)
            myds.Tables(0).PrimaryKey = pk

            For i = 0 To bs6.Count - 1
                Dim dr As DataRow = CType(bs6.Item(i), DataRowView).Row
                Dim mykey(2) As Object
                mykey(0) = dr.Item("pohd")
                mykey(1) = dr.Item("polineno")
                mykey(2) = dr.Item("qty")
                Dim result As DataRow = myds.Tables(0).Rows.Find(mykey)
                If Not IsNothing(result) Then
                    dr.Item("delivery") = result.Item("delivery")
                    dr.Item("item") = result.Item("deliveryitem")
                End If
            Next
        Else
            MessageBox.Show(mymessage)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, TextBox6.TextChanged
        Dim myobj = CType(sender, TextBox)
        Select Case myobj.Name
            Case "TextBox1"
                Dim myfields() = {"docno", "mironumber", "reference", "pohd", "delivery", "cmmf", "vendorcode", "vendorname"}
                Try
                    bs1.Filter = ""
                    If TextBox1.Text <> "" Then
                        bs1.Filter = "[" & myfields(ComboBox1.SelectedIndex) & "] like '" & TextBox1.Text & "'"
                    End If
                Catch ex As Exception

                End Try
            Case "TextBox2"
                Dim myfields() = {"docno", "mironumber", "reference", "pohd", "delivery", "cmmf", "vendorcode", "vendorname"}
                Try
                    bs2.Filter = ""
                    If TextBox2.Text <> "" Then
                        bs2.Filter = "[" & myfields(ComboBox2.SelectedIndex) & "] like '" & TextBox2.Text & "'"
                    End If
                Catch ex As Exception

                End Try
            Case "TextBox3"
                Dim myfields() = {"billingdocument", "salesdoc", "sebasiapono", "delivery"}
                Try
                    bs3.Filter = ""
                    If TextBox3.Text <> "" Then
                        bs3.Filter = "[" & myfields(ComboBox3.SelectedIndex) & "] like '" & TextBox3.Text & "'"
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "TextBox4"
                Dim myfields() = {"billingdocument", "salesdoc", "sebasiapono", "delivery"}
                Try
                    bs4.Filter = ""
                    If TextBox4.Text <> "" Then
                        bs4.Filter = "[" & myfields(ComboBox4.SelectedIndex) & "] like '" & TextBox4.Text & "'"
                    End If
                Catch ex As Exception

                End Try
            Case "TextBox5"
                Dim myfields() = {"mironumber", "reference", "pohd", "delivery", "cmmf", "vendorcode", "vendorname"}
                Try
                    bs5.Filter = ""
                    If TextBox5.Text <> "" Then
                        bs5.Filter = "[" & myfields(ComboBox5.SelectedIndex) & "] like '" & TextBox5.Text & "'"
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Case "TextBox6"
                Dim myfields() = {"mironumber", "reference", "pohd", "delivery", "cmmf", "vendorcode", "vendorname"}
                Try
                    bs6.Filter = ""
                    If TextBox6.Text <> "" Then
                        bs6.Filter = "[" & myfields(ComboBox6.SelectedIndex) & "] like '" & TextBox6.Text & "'"
                    End If
                Catch ex As Exception

                End Try
        End Select
        displayrecordcount()

        'Try

        '    bs2.Filter = ""
        '    If TextBox2.Text <> "" Then
        '        bs2.Filter = "[" & myfields(ComboBox1.SelectedIndex) & "] like '" & TextBox1.Text & "'"
        '    End If
        '    bs3.Filter = ""
        '    If TextBox3.Text <> "" Then
        '        bs3.Filter = "[" & myfields(ComboBox1.SelectedIndex) & "] like '" & TextBox1.Text & "'"
        '    End If
        '    bs4.Filter = ""
        '    If TextBox4.Text <> "" Then
        '        bs4.Filter = "[" & myfields(ComboBox1.SelectedIndex) & "] like '" & TextBox1.Text & "'"
        '    End If

        '    displayrecordcount()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try

    End Sub

    Private Sub displayrecordcount()
        Label4.Text = "Record" & IIf(bs1.Count > 1, "s", "") & " Found :" & bs1.Count.ToString
        Label5.Text = "Record" & IIf(bs2.Count > 1, "s", "") & " Found :" & bs2.Count.ToString
        Label9.Text = "Record" & IIf(bs3.Count > 1, "s", "") & " Found :" & bs3.Count.ToString
        Label13.Text = "Record" & IIf(bs4.Count > 1, "s", "") & " Found :" & bs4.Count.ToString
        Label17.Text = "Record" & IIf(bs5.Count > 1, "s", "") & " Found :" & bs5.Count.ToString
        Label21.Text = "Record" & IIf(bs6.Count > 1, "s", "") & " Found :" & bs6.Count.ToString
    End Sub


    Private Sub DataGridView2_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick, DataGridView4.CellDoubleClick, DataGridView6.CellDoubleClick
        Dim bs As BindingSource = CType(sender, DataGridView).DataSource
        Dim dr = CType(bs.Current, DataRowView).Row
        Dim myform As FormDeliveryHelper
        If sender.name = "DataGridView2" Or sender.name = "DataGridView6" Then
            myform = New FormDeliveryHelper(dr.Item("pohd"), dr.Item("polineno"))
        Else
            myform = New FormDeliveryHelper(dr.Item("sebasiapono"), dr.Item("polineno"))
        End If

        'this process can be improve using Object Oriented.

        If myform.ShowDialog = DialogResult.OK Then
            Dim mydr As DataRow = CType(myform.bs.Current, DataRowView).Row
            dr.Item("delivery") = mydr.Item("delivery")
            If sender.name = "DataGridView6" Then
                dr.Item("item") = mydr.Item("deliveryitem")
            Else
                dr.Item("deliveryitem") = mydr.Item("deliveryitem")
            End If
            If sender.name = "DataGridView4" Then
                dr.Item("deliveredqty") = mydr.Item("deliveredqty")
            End If
        End If
    End Sub


    Private Sub FormLogBook_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not IsNothing(getChanges) Then
            Select Case MessageBox.Show("Save unsave records?", "Unsave Records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    ToolStripButton4.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Function getChanges() As Object
        Me.Validate()
        If IsNothing(MyDS) Then
            Return Nothing
        End If
        'bs1.EndEdit()
        bs2.EndEdit()
        bs3.EndEdit()
        bs4.EndEdit()
        bs5.EndEdit()
        bs6.EndEdit()
        Return MyDS.GetChanges
    End Function

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        
        Dim ds2 = getChanges()
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.logbook(Me, mye) Then
                'delete the modfied row for Merged
                Dim modifiedRows = From row In MyDS.Tables(1)
                                   Where row.RowState = DataRowState.Modified
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
                modifiedRows = From row In MyDS.Tables(2)
                                   Where row.RowState = DataRowState.Modified
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
                modifiedRows = From row In MyDS.Tables(3)
                                   Where row.RowState = DataRowState.Modified
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
                'modifiedRows = From row In MyDS.Tables(4)
                '   Where row.RowState = DataRowState.Modified
                'For Each row In modifiedRows.ToArray
                '    row.Delete()
                'Next
                modifiedRows = From row In MyDS.Tables(5)
                   Where row.RowState = DataRowState.Modified
                For Each row In modifiedRows.ToArray
                    row.Delete()
                Next
                MyDS.Merge(ds2)
                MyDS.AcceptChanges()
                MessageBox.Show("Saved.")
            Else
                MessageBox.Show(mye.message)
            End If
        End If
    End Sub

   


    Private Sub AccountingReportFGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountingReportFGToolStripMenuItem.Click
        Dim myform As New FormReportAccountingFG
        myform.Show()
    End Sub

    Private Sub AccountingReportCompToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccountingReportCompToolStripMenuItem.Click
        Dim myform As New FormReportAccountingComp
        myform.Show()
    End Sub
End Class