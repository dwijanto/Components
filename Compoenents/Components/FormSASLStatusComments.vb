Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
Public Enum SASLSearch
    ShipToPartyName
    SoldToPartyName
    CustomerOrderNo
    SebAsiaSalesOrder
    SebAsiaPoNo
    VendorCode
    VendorName
    CMMF
    ItemId
End Enum

Public Enum SASLStatus
    Failed = 0
    Success = 1
End Enum

Public Class FormSASLStatusComments
    Dim FieldSearch() As String = {"shiptopartyname", "soldtopartyname", "customerorderno", "sebasiasalesorder", "sebasiapono", "vendorcode", "vendorname", "cmmf", "itemid"}
    Dim FieldName() As String = {"c1.customername ", "c2.customername", "sd.customerorderno", "sd.sebasiasalesorder", "pd.sebasiapono", "e.vendorcode", "v.vendorname", "cm.cmmf", "cm.itemid"}
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim Dataset1 As DataSet
    Dim DS As DataSet
    Dim QueryThread As New Thread(AddressOf DoQuery)
    Private WithEvents datetimepicker1 As New DateTimePicker
    Private WithEvents datetimepicker2 As New DateTimePicker
    Private WithEvents btnRefresh As New ToolStripButton
    Private BS As New BindingSource
    Dim myFilter As String = String.Empty
    Dim SqlstrGrid As String = String.Empty

    Private Sub FormSASLStatusComments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        loaddata()
        ComboBox2.Text = ""
    End Sub

    Private Sub loaddata()
        If Not QueryThread.IsAlive Then
            QueryThread = New Thread(AddressOf DoQuery)
            Dim myparam = New objThread
            myparam.myfilter = "myfilter"
            myparam.mydate1 = datetimepicker1.Value.Date
            myparam.mydate2 = datetimepicker2.Value.Date
            QueryThread.Start(myparam)
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoQuery(ByVal myparam As objThread)
        ProgressReport(5, "Populating DataGridView. Please wait...")
        Dataset1 = New DataSet
        'Debug.Print(myparam.myfilter)
        '     Dim sqlstr As String = "select c1.customername as shiptopartyname,c2.customername as soldtopartyname,sd.customerorderno,sd.sebasiasalesorder,sd.solineno,pd.sebasiapono,pd.polineno,e.vendorcode,v.vendorname,cm.cmmf,cm.itemid,sp.shipdate,sd.currentinquiryetd,getsasl(e.vendorcode,pd.comments,pd.sebasiapono,pd.polineno,sp.shipdate,sd.currentinquiryetd) as sasl" & _
        '                            " FROM cxsebodtp od" &
        '                            " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
        '                            " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
        '" LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
        '" LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
        '" LEFT JOIN stdcmntdtl keyword ON keyword.stdcmntdtlid = pd.commentid" &
        '" LEFT JOIN convcomnt conversion ON conversion.stdcmntdtlid = keyword.stdcmntdtlid" &
        '" LEFT JOIN cmnttxdtl standard ON standard.cmnttxdtlid = conversion.cmnttxdtlid" &
        '" LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
        '" LEFT JOIN cxconf c ON c.sebodtpid = od.cxsebodtpid" &
        '" LEFT JOIN cxconfstatus cs ON cs.cxconfid = c.cxconfid" &
        '" LEFT JOIN cxconfother co ON co.sebodtpid = od.cxsebodtpid" &
        '" LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
        '" LEFT JOIN cxshipmentother so ON so.sebodtpid = od.cxsebodtpid" &
        '" LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty" &
        '" LEFT JOIN customer c2 ON c2.customercode = sh.soldtoparty" &
        '" LEFT JOIN cxpoconf pc ON pc.sebasiapono = pd.sebasiapono AND pc.polineno = pd.polineno" &
        '" LEFT JOIN cxpoconfother pco ON pco.sebasiapono = pd.sebasiapono AND pco.polineno = pd.polineno" &
        '" LEFT JOIN ekko e ON e.po = pd.sebasiapono" &
        '" LEFT JOIN vendor v ON v.vendorcode = e.vendorcode" &
        '" LEFT JOIN cmmf cm ON cm.cmmf = pd.cmmf" &
        '" LEFT JOIN saooplt sao ON sao.soldtoparty = sh.soldtoparty AND sao.shiptoparty = pd.shiptoparty" &
        '" LEFT JOIN activity ac ON ac.activitycode = cm.rir" &
        '" LEFT JOIN sbu s ON s.sbuid = ac.sbuidsp" &
        '" LEFT JOIN vendorspcomp vsc ON vsc.vendorcode = v.vendorcode" &
        '" LEFT JOIN purchasinggroup pg ON pg.purchasinggroup = e.purchasinggroup::bpchar" &
        ' " WHERE od.ordertype::text = 'Shipment'::text and  sp.shipdate >= " & DateFormatyyyyMMdd(myparam.mydate1) & " and sp.shipdate <= " & DateFormatyyyyMMdd(myparam.mydate2)
        SqlstrGrid = "select c1.customername::character varying as shiptopartyname,c2.customername::character varying as soldtopartyname,sd.customerorderno::character varying,sd.sebasiasalesorder::character varying,sd.solineno::character varying,pd.sebasiapono::character varying,pd.polineno::character varying,e.vendorcode::character varying,v.vendorname::character varying,cm.cmmf::character varying,cm.itemid::character varying,sp.shipdate,od.deliveredqty,sd.currentinquiryetd,getsasl(e.vendorcode,pd.comments,pd.sebasiapono,pd.polineno,sp.shipdate,sd.currentinquiryetd) as sasl,ps.comment,ps.latestupdate" &
                               " FROM cxsebodtp od " &
                               " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid " &
                               " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid " &
                               " LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder " &
                               " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid " &
                               " LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid " &
                               " LEFT JOIN customer c1 ON c1.customercode = pd.shiptoparty " &
                               " LEFT JOIN customer c2 ON c2.customercode = sh.soldtoparty " &
                               " LEFT JOIN ekko e ON e.po = pd.sebasiapono LEFT JOIN vendor v ON v.vendorcode = e.vendorcode " &
                               " LEFT JOIN cmmf cm ON cm.cmmf = pd.cmmf " &
                               " Left join posasl ps on ps.sebasiapono = pd.sebasiapono and ps.polineno = pd.polineno and ps.shipdate = sp.shipdate" & _
                               " WHERE od.ordertype::text = 'Shipment'::text and  sp.shipdate >= " & DateFormatyyyyMMdd(myparam.mydate1) & " and sp.shipdate <= " & DateFormatyyyyMMdd(myparam.mydate2)

        If DbAdapter1.TbgetDataSet(SqlstrGrid, Dataset1) Then
            Dataset1.Tables(0).TableName = "SASL"
            ProgressReport(1, "Populate DataGridView")

        Else
            ProgressReport(5, "Error while loading dataset 1")

        End If

        Dim sqlstr = "select sebasiapono::bigint,polineno::integer,cslstatus::boolean,comment::text,shipdate::date from posasl"
        DS = New DataSet
        If DbAdapter1.TbgetDataSet(sqlstr, DS) Then
            Try
                DS.Tables(0).TableName = "PoSASL"
                Dim idx(2) As DataColumn
                idx(0) = DS.Tables(0).Columns(0)
                idx(1) = DS.Tables(0).Columns(1)
                idx(2) = DS.Tables(0).Columns(4)

                DS.Tables(0).PrimaryKey = idx

                ProgressReport(2, "assign bs")
            Catch ex As Exception
                ProgressReport(5, ex.Message)
            End Try
            
        Else
            ProgressReport(5, "Error while loading dataset 2")

        End If

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.DataGridView1.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    BindingSource1.DataSource = Dataset1.Tables(0)
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = BindingSource1

                    BindingNavigator1.BindingSource = BindingSource1




                    ProgressReport(5, "Populating DataGridView. Done.")
                Case 2
                    BS.DataSource = DS.Tables(0)
                Case 5
                    Me.ToolStripStatusLabel1.Text = message
            End Select

        End If

    End Sub


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ToolStripComboBox1.ComboBox.DataSource = System.Enum.GetValues(GetType(SASLSearch))
        ComboBox2.DataSource = System.Enum.GetValues(GetType(SASLStatus))
        ' Add any initialization after the InitializeComponent() call.
        datetimepicker1.Format = DateTimePickerFormat.Custom
        datetimepicker2.Format = DateTimePickerFormat.Custom
        datetimepicker1.CustomFormat = "dd-MMM-yyyy"
        datetimepicker2.CustomFormat = "dd-MMM-yyyy"
        datetimepicker1.Size = New Point(100, 25)
        datetimepicker2.Size = New Point(100, 25)
        Dim datetimepickerHost1 As New ToolStripControlHost(datetimepicker1)
        Dim datetimepickerHost2 As New ToolStripControlHost(datetimepicker2)

        'BindingNavigator1.Items.Add(BindingNavigator1.Items.Count - 1, datetimepickerHost1)
        'BindingNavigator1.Items.Add(BindingNavigator1.Items.Count - 1, datetimepickerHost2)
        btnRefresh.Text = "Refresh"

        BindingNavigator1.Items.Add(datetimepickerHost1)
        BindingNavigator1.Items.Add(datetimepickerHost2)
        BindingNavigator1.Items.Add(btnRefresh)
    End Sub

    Private Sub Handle_LoadData(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        loaddata()
    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim myobj = CType(sender, System.Windows.Forms.ToolStripComboBox)
        If Not IsNothing(myobj.SelectedIndex) Then
            'MessageBox.Show(FieldSearch(myobj.SelectedIndex).ToString)
            'Select Case myobj.SelectedIndex
            '    Case SASLSearch.ShipToPartyName
            '        MessageBox.Show("ShiptoPartyname")
            '    Case SASLSearch.SoldToPartyName
            '        MessageBox.Show("SoldtoPartyname")
            '    Case SASLSearch.CustomerOrderNo
            '    Case SASLSearch.SebAsiaSalesOrder
            '    Case SASLSearch.SebAsiaPoNo
            '    Case SASLSearch.VendorCode
            '    Case SASLSearch.VendorName
            '    Case SASLSearch.CMMF
            '    Case SASLSearch.ItemId
            'End Select
        End If

    End Sub



    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim obj = CType(sender, ToolStripTextBox)
        myFilter = ""
        If obj.Text <> "" Then
            BindingSource1.Filter = (FieldSearch(ToolStripComboBox1.SelectedIndex).ToString) & " like '" & obj.Text & "'"
            myFilter = (FieldName(ToolStripComboBox1.SelectedIndex).ToString) & " like '" & Replace(obj.Text, "*", "%") & "'"
        Else
            BindingSource1.Filter = ""
        End If


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If validsasl Then
            For i = 0 To DataGridView1.SelectedRows.Count - 1
                Dim obj = CType(DataGridView1.SelectedRows(i).DataBoundItem, System.Data.DataRowView).Row
                ModifyDS(obj)
            Next

            If updateRecord() Then
                Try
                    For i = 0 To DataGridView1.SelectedRows.Count - 1

                        Dim obj = CType(DataGridView1.SelectedRows(i).DataBoundItem, System.Data.DataRowView).Row
                        obj.Item("sasl") = ComboBox2.SelectedIndex
                        obj.Item("comment") = TextBox2.Text
                        obj.Item("latestupdate") = Today
                    Next
                Catch ex As Exception
                    loaddata()
                End Try
                
            End If

        Else
            MessageBox.Show("Please select value from combobox.")
            ComboBox2.Select()
            ComboBox2.SelectAll()
        End If
        
    End Sub
    Private Sub ModifyDS(ByVal obj As DataRow)
        Dim myresult As DataRow
        Dim mypkey(2) As Object
        mypkey(0) = obj.Item("sebasiapono")
        mypkey(1) = obj.Item("polineno")
        mypkey(2) = obj.Item("shipdate")

        myresult = DS.Tables(0).Rows.Find(mypkey)
        'Select Case ComboBox2.SelectedIndex
        'Case SASLStatus.Failed
        '    'delete DS
        '    If Not IsNothing(myresult) Then
        '        myresult.Delete()
        '    End If
        'Case SASLStatus.Success
        'Add DS
        If IsNothing(myresult) Then
            If ComboBox2.SelectedIndex = SASLStatus.Success Then
                'Dim dr = BS.AddNew
                Dim dr = DS.Tables(0).NewRow
                dr.item("sebasiapono") = mypkey(0)
                dr.item("polineno") = mypkey(1)
                dr.item("cslstatus") = IIf(ComboBox2.SelectedIndex = SASLStatus.Success, True, False)
                dr.item("shipdate") = mypkey(2)
                dr.item("comment") = TextBox2.Text
                DS.Tables(0).Rows.Add(dr)
            End If
            
        Else
            myresult.Item("cslstatus") = IIf(ComboBox2.SelectedIndex = SASLStatus.Success, True, False)
            myresult.Item("comment") = TextBox2.Text
        End If
        'End Select
    End Sub
    Private Function validsasl() As Boolean
        Dim myret As Boolean = False

        If ComboBox2.SelectedIndex <> -1 AndAlso ComboBox2.Text <> "" Then
            myret = True
        End If
        Return myret
    End Function

    Private Function updateRecord() As Boolean
        Dim myret As Boolean = False
        'BS.EndEdit()
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            If DbAdapter1.AdapterSASLTx(Me, ds2) Then
                myret = True
            End If
            Dim addedrows = From row In DS.Tables(0)
            Where row.RowState = DataRowState.Added

            For Each row In addedrows.ToArray
                row.Delete()
            Next
            DS.Merge(ds2)
            DS.AcceptChanges()
        End If
        Return myret
    End Function



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim Filename As String = "SASL Status-" & Format(DateTime.Today, "yyyyMMdd") & ".xlsx"
        Dim criteria As String = IIf(BindingSource1.Filter = "", "", " and ") & myFilter
        Call ExcelStuff.ExportToExcelAskDirectory(Filename, SqlstrGrid & criteria, dbtools1, "A4", "\templates\ExcelTemplate.xltx")

    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub
End Class

Public Class objThread
    Public Property myfilter As String
    Public Property mydate1 As Date
    Public Property mydate2 As Date
End Class