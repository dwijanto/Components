Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass

Public Class FormReportAccountingFG
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim combobs As BindingSource

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim myuser As String = String.Empty
        Dim sqlstr As String
        'If ComboBox1.Text <> "All" Then
        '    myuser = "'" & ComboBox1.Text & "'"
        '    sqlstr = "select tb.*,ph.deliverydate from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)" &
        '             " left join packinglisthd ph on ph.delivery = tb.delivery"
        'Else
        '    sqlstr = "select tb.*,ph.deliverydate from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)" &
        '             " left join packinglisthd ph on ph.delivery = tb.delivery"
        'End If
        'Dim sqlstr As String = "select * from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)"
        If ComboBox1.Text <> "All" Then
            myuser = "'" & ComboBox1.Text & "'"
            sqlstr = "select distinct tb.housebill as ""Bill of lading No"",tb.billingdoc as ""SEB Invoice No"",tb.delivery as ""Packing List No"",tb.reference as ""Supplier Invoice no."" from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"

        Else
            sqlstr = "select distinct tb.housebill as ""Bill of lading No"",tb.billingdoc as ""SEB Invoice No"",tb.delivery as ""Packing List No"",tb.reference as ""Supplier Invoice no."" from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"

        End If
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "LogBook" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub


    Private Sub FormReportAccountingFG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myThread = New System.Threading.Thread(myThreadDelegate)
        myThread.SetApartmentState(ApartmentState.MTA)
        myThread.Start()
    End Sub

    Sub DoQuery()
        'Get All user from PackingListDtl
        Dim sqlstr = "select 'All'::text as username union all (select distinct username from accountinghd order by username);"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            combobs = New BindingSource
            combobs.DataSource = DS.Tables(0)
            If DS.Tables(0).Rows.Count > 0 Then

                ProgressReport(4, "Fill Combo Datasource")
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
                    Me.ComboBox1.DataSource = combobs
                    Me.ComboBox1.DisplayMember = "username"
                    'Case (5)
                    '    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    'Case 6
                    '    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                    'Case 7
                    '    Dim myvalue = message.ToString.Split(",")
                    '    ToolStripProgressBar1.Minimum = 1
                    '    ToolStripProgressBar1.Value = myvalue(0)
                    '    ToolStripProgressBar1.Maximum = myvalue(1)

            End Select

        End If

    End Sub
End Class