Imports System.Threading
Imports Components.HelperClass
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
'Imports Components.ExportToExcel

Public Class FormCommentConversion
    Protected mypanel1 As UCSortTx
    Protected mypanel As UCFilterTx
    Dim QueryDelegate As New ThreadStart(AddressOf QueryWork)
    Dim WorkDelegate As New ThreadStart(AddressOf DoWork)
    Dim ConvertDelegate As New ThreadStart(AddressOf GoConvert)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim QueryThread As New System.Threading.Thread(QueryDelegate)
    Dim WorkThread As New System.Threading.Thread(WorkDelegate)
    Dim ConvertThread As New System.Threading.Thread(ConvertDelegate)

    Dim DS As DataSet
    Dim bindingsource1 As New BindingSource
    Dim UpdateCommentSB As New StringBuilder
    Dim UpdateCommentSB2 As New StringBuilder

    Public Enum SASLData
        SASL_Failed = 0
        SASL_Success = 1
        All_Data = 2
    End Enum


    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ToolStripComboBox1.ComboBox.DataSource = System.Enum.GetValues(GetType(SASLData))
        ToolStripComboBox1.ComboBox.SelectedIndex = 2
    End Sub

    Private Sub FormCommentConversion_Load1(ByVal sender As Object, ByVal e As System.EventArgs)

        LoadData()
        LoadToolstrip()

        'LoadDataDirect()
        'LoadToolstrip()
    End Sub

    Private Sub FormCommentConversion_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        Dim events As New List(Of ManualResetEvent)()
        'myevents.Add(New ManualResetEvent(False))
        Dim obj As New ThreadPoolManualResetEvent
        obj.ObjectID = 1
        obj.signal = New ManualResetEvent(False)
        events.Add(obj.signal)
        'LoadData()
        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf QueryWork), obj)
        WaitHandle.WaitAll(events.ToArray)
        bindingsource1.DataSource = DS.Tables(0)
        DataGridView1.DataSource = bindingsource1
        BindingNavigator1.BindingSource = bindingsource1
        LoadToolstrip()
    End Sub

    Private Sub LoadData()
        If QueryThread.IsAlive Then
            MessageBox.Show("Process still running. Please wait!")
        Else
            QueryThread = New System.Threading.Thread(QueryDelegate)
            QueryThread.Start()
        End If


    End Sub
    Private Sub LoadDataDirect()
        QueryWork()
    End Sub
    Sub QueryWork()
        DS = New DataSet
        Dim sqlstr = "select p.cxsebpodtlid,p.comments,keyword.stdcmntdtlcode::character varying ,cmnttxdtlname,commentid," &
                     " CASE" &
                     " WHEN abs(sp.shipdate - sd.currentinquiryetd) <= 7 THEN 1" &
                     " ELSE 0" &
                     " END AS sasl" &
                     " from cxsebodtp o" &
                     " left join cxshipment sp on sp.sebodtpid = o.cxsebodtpid" &
                     " left join cxrelsalesdocpo r on r.cxrelsalesdocpoid = o.relsalesdocpoid" &
                     " left join cxsebpodtl p on p.cxsebpodtlid = r.cxsebpodtlid" &
                     " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                     " left join stdcmntdtl keyword on keyword.stdcmntdtlid = p.commentid " &
                     " left join convcomnt conversion on conversion.stdcmntdtlid = keyword.stdcmntdtlid " &
                     " left join  cmnttxdtl standard on standard.cmnttxdtlid = conversion.cmnttxdtlid" &
                     " where o.ordertype = 'Shipment';" &
                     "select stdcmntdtlcode::character varying,stdcmntdtlid from stdcmntdtl where stdcmnthdid = 1;"

        Dim message As String = String.Empty
        ProgressReport(1, "Populating data. Please wait..")
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, message) Then
            ProgressReport(1, message)
        Else
            ProgressReport(1, "Populating data. Done")
            ProgressReport(3, "Assign Datagridview object")

        End If

    End Sub
    Sub QueryWork(ByVal obj As Object)
        DS = New DataSet
        Dim sqlstr = "select p.cxsebpodtlid,p.comments,keyword.stdcmntdtlcode::character varying ,cmnttxdtlname,commentid," &
                     " CASE" &
                     " WHEN abs(sp.shipdate - sd.currentinquiryetd) <= 7 THEN 1" &
                     " ELSE 0" &
                     " END AS sasl" &
                     " from cxsebodtp o" &
                     " left join cxshipment sp on sp.sebodtpid = o.cxsebodtpid" &
                     " left join cxrelsalesdocpo r on r.cxrelsalesdocpoid = o.relsalesdocpoid" &
                     " left join cxsebpodtl p on p.cxsebpodtlid = r.cxsebpodtlid" &
                     " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                     " left join stdcmntdtl keyword on keyword.stdcmntdtlid = p.commentid " &
                     " left join convcomnt conversion on conversion.stdcmntdtlid = keyword.stdcmntdtlid " &
                     " left join  cmnttxdtl standard on standard.cmnttxdtlid = conversion.cmnttxdtlid" &
                     " where o.ordertype = 'Shipment';" &
                     "select stdcmntdtlcode::character varying,stdcmntdtlid from stdcmntdtl where stdcmnthdid = 1;"

        Dim message As String = String.Empty
        ProgressReport1(1, "Populating data. Please wait..")
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, message) Then
            ProgressReport1(1, message)
        End If
        ProgressReport1(1, "Populating data. Done.")
        DirectCast(obj.signal, ManualResetEvent).Set()
    End Sub
    Sub DoWork()
        Throw New NotImplementedException
    End Sub
    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Dim myfilter = ""
        If CType(ToolStripComboBox1.SelectedItem, SASLData) < 2 Then
            myfilter = "sasl=" & CType(ToolStripComboBox1.SelectedItem, SASLData)
        End If
        bindingsource1.Filter = myfilter
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
                Case 3
                    bindingsource1.DataSource = DS.Tables(0)
                    DataGridView1.DataSource = bindingsource1
                    BindingNavigator1.BindingSource = bindingsource1
                    'DataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                    'LoadToolstrip()
                Case 4
                    LoadData()
            End Select

        End If

    End Sub

    Private Sub ProgressReport1(ByVal id As Integer, ByVal message As String)
        If Me.ToolStripContainer1.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.BeginInvoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
            End Select

        End If
       

    End Sub

    Private Sub LoadToolstrip()
        ' ToolStripCustom1 = New DJLib.ToolStripCustom()
        Dim myaction As HideToolbarDelegate = AddressOf toolstripvisible
        LoadToolstripFilterSort(myaction, DataGridView1, mypanel1, ToolStripCustom1, mypanel)
    End Sub
    Private Sub toolstripvisible(ByVal toolstripvisible As Boolean)
        ToolStripCustom1.Visible = Not (toolstripvisible)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Call toolstripvisible(ToolStripCustom1.Visible)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        LoadData()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not QueryThread.IsAlive Then


            If Not ConvertThread.IsAlive Then
                ConvertThread = New System.Threading.Thread(ConvertDelegate)
                ConvertThread.Start()
            Else
                MessageBox.Show("Converting still running. Please wait...")
            End If
            'GoConvert()
            'LoadData()
        Else
            MessageBox.Show("Please wait. Data still populating..")
        End If
    End Sub

    Private Sub GoConvert()
        'Clear conversion
        ProgressReport(1, "Start conversion...")
        Dim sqlstr As String = String.Empty
        Dim errmessage As String = String.Empty
        For Each dr As DataRow In DS.Tables(0).Rows
            dr.Item("commentid") = DBNull.Value
        Next

        'Find based on first character
        For Each dr As DataRow In DS.Tables(1).Rows

            If UpdateCommentSB.Length > 0 Then
                UpdateCommentSB.Append(",")
            End If

            UpdateCommentSB.Append(String.Format("['^{0}',{1}::character varying]", dr.Item("stdcmntdtlcode"), dr.Item("stdcmntdtlid")))
        Next

        For Each dr As DataRow In DS.Tables(1).Rows

            If UpdateCommentSB2.Length > 0 Then
                UpdateCommentSB2.Append(",")
            End If
            UpdateCommentSB2.Append(String.Format("['{0}',{1}::character varying]", dr.Item("stdcmntdtlcode").ToString.Trim, dr.Item("stdcmntdtlid")))


        Next

        'Execute 
        If UpdateCommentSB.Length > 0 Then
            ProgressReport(1, "Update at Beginning Entry")
            sqlstr = "update cxsebpodtl set commentid = foo.commentid::bigint from (select * from array_to_set2(Array[" & UpdateCommentSB.ToString & "]) as tb (id character varying,commentid character varying))foo where comments ~ foo.id;"
            Dim ra As Long

            If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
                MessageBox.Show(errmessage)
            End If
        End If

        If UpdateCommentSB2.Length > 0 Then
            ProgressReport(1, "Update Part of Entry")
            sqlstr = "update cxsebpodtl set commentid = foo.commentid::bigint from (select * from array_to_set2(Array[" & UpdateCommentSB2.ToString & "]) as tb (id character varying,commentid character varying))foo where cxsebpodtl.commentid is null and  comments ~ foo.id;"
            Dim ra As Long

            If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
                MessageBox.Show(errmessage)
            End If
        End If
        ProgressReport(1, "Finish conversion...")
        ProgressReport(4, "Refresh Record.")
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        ProgressReport(1, "Clear Conversion")
        Dim sqlstr = String.Empty
        Dim errmessage As String = String.Empty
        sqlstr = "update cxsebpodtl set commentid = Null"
        Dim ra As Long

        If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            MessageBox.Show(errmessage)
        End If
        ProgressReport(1, "Clear Conversion Done")
        LoadData()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim myform As New FormGenerateParetoChart
        Select Case ToolStripComboBox1.SelectedItem
            Case SASLData.All_Data
                myform.mySASL = ""
            Case SASLData.SASL_Failed
                myform.mySASL = " and abs(sp.shipdate - sd.currentinquiryetd) > 7"
            Case SASLData.SASL_Success
                myform.mySASL = " and abs(sp.shipdate - sd.currentinquiryetd) <= 7"
        End Select
        myform.Show()
    End Sub

    Private Sub ExtractKeywords_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtractKeywords.Click
        Dim sqlstr = "select keyword.stdcmntdtlcode as ""Keyword"", stdcmntdtldesc as ""Description"" from stdcmntdtl keyword" &
                     " left join stdcmnthd h on h.stdcmnthdid = keyword.stdcmnthdid" &
                     " where stdcmnthdname = 'SAP Keywords'" &
                     " order by ""Keyword"""
        Dim myfile As String = "Keyword.xlsx"
        ExportToExcel(myfile, sqlstr, DbAdapter1)
    End Sub

    Private Sub ExtractStandardComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExtractStandardComments.Click
        Dim sqlstr = "select d.stdcmntdtlcode as ""Standard Comments"",d.stdcmntdtldesc as ""Description"" from stdcmntdtl d" &
                     " left join stdcmnthd h on h.stdcmnthdid = d.stdcmnthdid" &
                     " where stdcmnthdname = 'SAP Keywords' order by ""Standard Comments"""
        Dim myfile As String = "StandardComments.xlsx"
        ExportToExcel(myfile, sqlstr, DbAdapter1)
    End Sub


End Class
