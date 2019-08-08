Imports System.Threading
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Imports System.Text
Imports Components.PublicClass
Public Class FormReportAccountingComp2

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
        '    sqlstr = "select tb.*,ph.deliverydate from sp_getaccountingdatacomp(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(reference text,mironumber  text,amount  numeric,vendorname character varying,housebill text,delivery bigint,deliveryitem integer)" &
        '             " left join packinglisthd ph on ph.delivery = tb.delivery"
        'Else
        '    sqlstr = "select tb.*,ph.deliverydate from sp_getaccountingdatacomp(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(reference text,mironumber  text,amount  numeric,vendorname character varying,housebill text,delivery bigint,deliveryitem integer)" &
        '             " left join packinglisthd ph on ph.delivery = tb.delivery"
        'End If
        'Dim sqlstr As String = "select * from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(reference  text,amount  numeric,vendorname character varying,housebill text,docno character varying,delivery bigint,deliveryitem integer,billingdocument character varying)"

        If ComboBox1.Text <> "All" Then
            myuser = "'" & ComboBox1.Text & "'"
            'sqlstr = "select tb.* from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlcurr(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlsd(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlbill(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric, username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlfin(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric, username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'fin1 adding customername
            'sqlstr = "select tb.* from sp_getaccountingdatadtlfin1(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,soldtopartyname character varying,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric, username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'add shiptopary
            sqlstr = "select tb.* from sp_getaccountingdatadtlfin3stp(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date, " & myuser & ") as tb(postingdate date,soldtoparty bigint,soldtopartyname character varying,shiptoparty bigint,shiptopartyname character varying,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,containerno text,portofloading text,pod character(50),billingdoc bigint,billingitem integer,billtype character varying,netvalue numeric,curr text,valueinusd numeric, username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,billoflading text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"

        Else
            'sqlstr = "select tb.* from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlcurr(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlbill(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'sqlstr = "select tb.* from sp_getaccountingdatadtlfin(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'fin1 adding customername
            'sqlstr = "select tb.* from sp_getaccountingdatadtlfin1(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,soldtopartyname character varying, vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,housebill text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            'add shiptopary
            'sqlstr = "select tb.* from sp_getaccountingdatadtlfin3(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,soldtopartyname character varying,shiptoparty bigint,shiptopartyname character varying, vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,containerno text,portofloading text,pod character(50),billingdoc bigint,billingitem integer,netvalue numeric,curr text,valueinusd numeric,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,billoflading text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
            sqlstr = "select tb.* from sp_getaccountingdatadtlfin3(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,soldtoparty bigint,soldtopartyname character varying,shiptoparty bigint,shiptopartyname character varying, vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,customerpono character varying,sappo bigint,sappoitem integer,amount numeric,currency character varying,amountinusd numeric,qty numeric,purchasinggroup character varying,cmmf bigint,salesdoc bigint,salesdocitem integer,delivery bigint,item integer,deliverydate date,containerno text,portofloading text,pod character(50),billingdoc bigint,billingitem integer,billtype character varying,netvalue numeric,curr text,valueinusd numeric,username text,submitdate date,invreceiveddate date,fcrnumber text,remark text,billoflading text,blreceiveddate date,courier character varying,trackingno character varying,senddate date)"
        End If

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "LogbookDetail" '& GetCompanyName()
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
        Dim sqlstr = "select 'All'::text as username union all (select distinct userid from saoallocation order by userid);"
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