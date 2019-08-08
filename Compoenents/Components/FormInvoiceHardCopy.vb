Imports System.Threading
Imports Components.SharedClass
Imports Components.PublicClass
Imports Components.ExportToExcelFile
Imports System.Text
Imports System.IO
Imports Microsoft.Office.Interop
Public Class FormInvoiceHardCopy
    Dim myThreadDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myWorkDelegate As New ThreadStart(AddressOf DoWork)
    Dim myImportDelegate As New ThreadStart(AddressOf DoImport)
    Dim bsReference As BindingSource


    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myWork As New System.Threading.Thread(myWorkDelegate)
    Dim myImport As New System.Threading.Thread(myImportDelegate)


    Dim myImportDocumentReceivedDate As New System.Threading.Thread(AddressOf DoImportDocumentReceivedDate)
    Dim myLoadView As New System.Threading.Thread(AddressOf doLoadViewData)
    Dim myHousebilldata As New System.Threading.Thread(AddressOf doLoadHousebillData)
    Dim SenddateTmpDS As DataSet
    Dim SendDateBS As BindingSource
    Dim cmsenddate As CurrencyManager
    Dim TrackingNoDS As DataSet
    Dim TrackingNoBS As BindingSource
    Dim MyMarketName As Object
    Dim myCriteria As String
    Dim housebilltext As String
    Dim ViewHousebillBS As Object

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim combobs As BindingSource
    Dim combobsMiro As BindingSource
    Dim bs1 As BindingSource
    Dim MyDS As New DataSet
    Dim startdate As Date
    Dim username As String = String.Empty    
    Dim SInvDS As DataSet
    Dim SinvBS As BindingSource
    Dim ReceivedDS As DataSet
    Dim ReceivedBS As BindingSource
    Dim source As AutoCompleteStringCollection
    Dim sourcebill As AutoCompleteStringCollection
    Dim ReceivedTmpDS As DataSet
    Dim ReceivedTmpBS As BindingSource
    Dim ViewDS As DataSet
    Dim ViewBS As BindingSource
    Dim cm As CurrencyManager

    Dim ViewStartdate As Date
    Dim ViewLastDate As Date
    Dim ViewFieldname As String
    Dim ViewFieldvalue As String

    Sub DoQuery()
        'Get All user from PackingListDtl
        Dim sqlstr = "select ''::text as username union all (select distinct username from accountinghd order by username);" &
                     "select ''::text as miroyear union all (select distinct miroyear::text from accountinghd order by miroyear desc);"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            combobs = New BindingSource
            combobsMiro = New BindingSource
            combobs.DataSource = DS.Tables(0)
            combobsMiro.DataSource = DS.Tables(1)
            If DS.Tables(0).Rows.Count > 0 Then

                ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
            End If

        End If
        loadreceivedtmpds()
        loadsenddocumenttmpds()
        loadtrackingnumberds()

    End Sub

    Sub DoWork()
        Dim sqlstr = "select username,dateupload,invoicehardcopyhdid from invoicehardcopyhd order by username,dateupload desc;" &
                     " select username,max(dateupload) as dateupload,supplierinvoicenumber from invoicehardcopyhd ihd" &
                     " left join invoicehardcopydt idt on idt.invoicehardcopyhdid = ihd.invoicehardcopyhdid " &
                     " group by username,supplierinvoicenumber order by username,max(dateupload) ;" '&

        '" select reference from accountinghd;"

        MyDS = New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(6, "Marque")

        If Not DbAdapter1.TbgetDataSet(sqlstr, MyDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            Try
                Dim pk(1) As DataColumn
                pk(0) = MyDS.Tables(0).Columns(0)
                pk(1) = MyDS.Tables(0).Columns(1)
                MyDS.Tables(0).PrimaryKey = pk

                Dim pk2(1) As DataColumn
                pk2(0) = MyDS.Tables(1).Columns(0)
                pk2(1) = MyDS.Tables(1).Columns(2)
                MyDS.Tables(1).CaseSensitive = True
                MyDS.Tables(1).PrimaryKey = pk2

                bs1 = New BindingSource
                bs1.DataSource = MyDS.Tables(0)
                'bsReference = New BindingSource
                'bsReference.DataSource = MyDS.Tables(2)


                'Dim mySource = MyDS.Tables(2).AsEnumerable.Select(Of System.Data.DataRow, String)(Function(x) x.ToString("Reference")).toArray()
                'Dim mySource As String() = From p In bsReference.List
                '               Select p.row.item("reference")
                

                ProgressReport(8, "Fill DataGridView")
            Catch ex As Exception
                ProgressReport(5, ex.Message)
            End Try

            
            
        End If
        'loadreceivedtmpds()
        'loadsenddocumenttmpds()

        'ReceiveTMPDS dataset
        'ReceivedTmpDS = New DataSet
        'sqlstr = "select ''::character varying as supplierinvoicenumber,''::character varying as billoflading,''::character varying as fcrnumber,null::date as receiveddate from paramhd where paramname='hewllo'"
        'If Not DbAdapter1.TbgetDataSet(sqlstr, ReceivedTmpDS, mymessage) Then
        '    ProgressReport(2, mymessage)
        'Else
        '    ReceivedTmpBS = New BindingSource
        '    ReceivedTmpBS.DataSource = ReceivedTmpDS.Tables(0)
        '    ProgressReport(9, "Fill DataGrid")
        'End If
        'CM = CType(BindingContext(ReceivedTmpBS), CurrencyManager)

        'ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", MyDS.Tables(0).Rows(0).Item(0)))
        ProgressReport(5, "Continues")
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
                    ComboBox1.DataSource = combobs
                    ComboBox1.DisplayMember = "username"
                    ComboBox2.DataSource = combobs
                    ComboBox2.DisplayMember = "username"
                    ComboBox3.DataSource = combobs
                    ComboBox3.DisplayMember = "username"
                    ComboBox8.DataSource = combobsMiro
                    ComboBox8.DisplayMember = "miroyear"
                    ComboBox9.DataSource = combobsMiro
                    ComboBox9.DisplayMember = "miroyear"
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
                    DataGridView1.DataSource = bs1
                    bs1.Filter = ""

                    'Dim mysource = (From row In MyDS.Tables(2)
                    '                Select mycol = row(0).ToString).ToArray

                    'Dim mysource = (From p In bsReference.List
                    '               Select mycol = p.row.item("reference").ToString).ToArray

                    'Move to form.load
                    'source = New AutoCompleteStringCollection
                    ''source.AddRange(mysource)
                    'TextBox3.AutoCompleteCustomSource = source
                    'TextBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    'TextBox3.AutoCompleteSource = AutoCompleteSource.CustomSource



                    'TextBox4.AutoCompleteCustomSource = source
                    'TextBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    'TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
                Case 9
                    DataGridView4.AutoGenerateColumns = False
                    DataGridView4.DataSource = ReceivedTmpBS
                Case 10
                    DataGridView5.AutoGenerateColumns = False
                    DataGridView5.DataSource = SendDateBS
                Case 11
                    DataGridView6.AutoGenerateColumns = False
                    DataGridView6.DataSource = TrackingNoBS
                Case 12
                    DataGridView7.AutoGenerateColumns = False
                    DataGridView7.DataSource = ViewBS
                Case 13
                    DataGridView8.AutoGenerateColumns = False
                    DataGridView8.DataSource = ViewHousebillBS
            End Select

        End If

    End Sub

    Private Sub FormInvoiceHardCopy_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        source = New AutoCompleteStringCollection
        sourcebill = New AutoCompleteStringCollection
        'source.AddRange(mysource)
        TextBox3.AutoCompleteCustomSource = source
        TextBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TextBox3.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox4.AutoCompleteCustomSource = source
        TextBox4.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox5.AutoCompleteCustomSource = sourcebill
        TextBox5.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TextBox5.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox8.AutoCompleteCustomSource = sourcebill
        TextBox8.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        TextBox8.AutoCompleteSource = AutoCompleteSource.CustomSource
        If Not myWork.IsAlive Then
            myWork = New System.Threading.Thread(myWorkDelegate)
            myWork.Start()
        Else
            MessageBox.Show("Please wait, current process still running...")
        End If




    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myThread.Start()       
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged
        Dim cb As ComboBox = CType(sender, ComboBox)
        If Not IsNothing(cb.SelectedIndex) AndAlso Not IsNothing(bs1) Then

            bs1.Filter = ""
            If cb.Text <> "" Then
                bs1.Filter = "username = '" & cb.Text & "'"
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Start Thread
        If ComboBox1.Text = "" Then
            MessageBox.Show("Please select User Name.")
            Exit Sub
        End If
        If Not myImport.IsAlive Then
            startdate = DateTimePicker1.Value.Date
            username = ComboBox1.Text
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                myImport = New Thread(AddressOf DoImport)
                myImport.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoImport()
        Dim sw As New Stopwatch

        Dim ImportHardCopyHdSB As New System.Text.StringBuilder
        Dim ImportHardCopyDTSb As New System.Text.StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())
  
        Dim sqlstr As String = String.Empty
        Dim myID As Long
        Dim DS As New DataSet
        sw.Start()
        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Delete Existing Record
                ProgressReport(2, "Delete ..")
                ProgressReport(6, "Marque")

                'delete current hd
                '
                'sqlstr = "delete from invoicehardcopydt idt where idt.invoicehardcopyhdid in (select invoicehardcopyhdid from invoicehardcopyhd where username = '" & username & "' and dateupload = " & DateFormatyyyyMMdd(startdate) & ");" &
                '          " select setval('invoicehardcopydt_invoicehardcopydtid_seq',(select invoicehardcopydtid from invoicehardcopydt order by invoicehardcopydtid desc limit 1) + 1,false);" &
                '         " delete from invoicehardcopyhd ihd where ihd.dateupload = " & DateFormatyyyyMMdd(startdate) & " and ihd.username = '" & username & "';" &
                '        " select setval('invoicehardcopyhd_invoicehardcopyhdid_seq',(select invoicehardcopyhdid from invoicehardcopyhd order by invoicehardcopyhdid desc limit 1) + 1,false);"


                Dim mymessage As String = String.Empty
                'If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                '    ProgressReport(2, mymessage)
                '    Exit Sub
                'End If

                'Get IHD-ID
                ProgressReport(2, "Get Id")
                sqlstr = "select nextval('invoicehardcopyhd_invoicehardcopyhdid_seq');"


                mymessage = String.Empty

                If Not DbAdapter1.ExecuteScalar(sqlstr, myID, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                ProgressReport(2, "Read Text File...")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If i >= 0 And myrecord(0) <> "" Then
                        ImportHardCopyDTSb.Append(myID & vbTab &
                                                   validstr(myrecord(0)) & vbCrLf)

                    End If
                Next


            End With
        End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            'If ImportHardCopyHdSB.Length > 0 Then
            'ProgressReport(2, "Copy BillingHD")

            'sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "'," & DateFormatyyyyMMdd(startdate) & ");"
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'myret = DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmessage)
            'If Not myret Then
            '    ProgressReport(2, "Insert InvoiceHardCopy HD" & "::" & errmessage)
            '    Exit Sub
            'End If
            ' End If

            If ImportHardCopyDTSb.Length > 0 Then
                ProgressReport(2, "Copy Invoice Hard Copy Detail")
                'billingdocument bigint,item integer,billedqty numeric, salesunit text, requiredqty numeric,
                'netweight numeric, grossweight numeric, weightunit text, volume numeric,  vunit text,
                'pricingdate date,exrate numeric, netvalue numeric,curr,
                'salesdoc bigint, salesdocitem integer, material bigint, description text, shippingpoint integer,
                'plant integer, country text, cost numeric,curr2 subtotal(numeric)
                'sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "'," & DateFormatyyyyMMdd(startdate) & ");copy invoicehardcopydt(invoicehardcopyhdid,supplierinvoicenumber) from stdin with null as 'Null';"
                sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "','" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", Date.Now) & "');copy invoicehardcopydt(invoicehardcopyhdid,supplierinvoicenumber) from stdin with null as 'Null';"

                errmessage = DbAdapter1.copy(sqlstr, ImportHardCopyDTSb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Invoice Hard Copy Detail" & "::" & errmessage)
                    ProgressReport(5, "Continue")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        DoWork()
    End Sub

    Private Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dr As DataRow = CType(bs1.Current, DataRowView).Row
        Dim myuser = dr.Item("username")
        ExportToExcel(myuser, dr.Item("dateupload"), ComboBox9.Text)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim withbilloflading As Boolean = False
        If MessageBox.Show("Include bill of lading?", "Bill of lading", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            withbilloflading = True
        End If
        Dim dr As DataRow = CType(bs1.Current, DataRowView).Row
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim myuser As String = String.Empty
        Dim dateupload As Date
        Dim sqlstr As String
        Dim sqlstr1 As String = String.Empty

        'If ComboBox1.Text <> "" Then
        myuser = "'" & dr.Item("username") & "'"
        dateupload = dr.Item("dateupload")
        If withbilloflading Then
            sqlstr = "select tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost amount,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"

            sqlstr1 = " with foo as (select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice, invoicehardcopydtid from (select  tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
                     " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
                     " order by invoicehardcopydtid) " &
                     " select foo.supplierinvoice,foo.amount,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
                     " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"

        Else
            sqlstr = "select tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
            'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopynobill(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
            '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
            '         " order by invoicehardcopydtid "

            sqlstr1 = "with foo as (select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,invoicehardcopydtid   from (select  tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
                    " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
                    " order by invoicehardcopydtid) " &
                    " select foo.supplierinvoice,foo.amount,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
                    " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"
        End If
        'sqlstr = "select tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
        'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '         " order by invoicehardcopydtid "

        'Else
        '    'sqlstr = "select tb.*from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"
        '    MessageBox.Show("Please select User Name.")
        '    Exit Sub
        'End If

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "DocumentHardCopy" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "DETAIL",
                                                            .Sqlstr = sqlstr
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 2,
                                                            .SheetName = "Total",
                                                            .Sqlstr = sqlstr1
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

            myreport.Run(Me, e)

        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        If sender.name = "DETAIL" Then
            sender.columns("C:D").NumberFormat = "#,##0.00"
        ElseIf sender.name = "Total" Then
            sender.columns("B:C").NumberFormat = "#,##0.00"
        End If
            
        With sender.PageSetup
            .orientation = Excel.XlPageOrientation.xlLandscape
            .FitToPagesWide = False
            .FitToPagesTall = 1
            .Zoom = False
        End With

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.AutoValidate = Windows.Forms.AutoValidate.Disable
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ComboBox2.Text = "" Then
            MessageBox.Show("Please select user name from the list.")
            ComboBox2.Focus()
            Exit Sub
        End If
        Dim myobj(1) As Object
        myobj(0) = ComboBox1.Text
        myobj(1) = TextBox1.Text
        Dim mydr As DataRow = MyDS.Tables(1).Rows.Find(myobj)
        If IsNothing(mydr) Then
            MessageBox.Show("Record not found.")
            TextBox2.Text = ""
        Else
            TextBox2.Text = String.Format("{0:dd-MMM-yyyy}", mydr.Item("dateupload"))
        End If
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not myImportDocumentReceivedDate.IsAlive Then
            startdate = DateTimePicker1.Value.Date
            username = ComboBox1.Text
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myImportDocumentReceivedDate = New Thread(AddressOf DoImportDocumentReceivedDate)
                myImportDocumentReceivedDate.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoImportDocumentReceivedDate()
        Dim sw As New Stopwatch

        Dim forwarderhousebillSB As New StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())

        Dim sqlstr As String = String.Empty
        'Dim myID As Long
        Dim DS As New DataSet
        sw.Start()

        ProgressReport(1, "Read Text File...")
        ProgressReport(6, "Read Text File...")
        Dim mymessage As String = String.Empty

        sqlstr = "select supplierinvoicenumber,fcrnumber,remark,receiveddate,invoicehardcopyreceiveddateid from invoicehardcopyreceiveddate;" &
               " select housebill,receiveddate,courier,trackingno,senddate,housebilldocid from housebilldoc;"
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, message:=mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        DS.Tables(0).TableName = "InvoiceHardCopyReceivedDate"
        Dim idx0(0) As DataColumn
        idx0(0) = DS.Tables(0).Columns(0)
        DS.Tables(0).PrimaryKey = idx0

        DS.Tables(1).TableName = "HousebillDoc"
        Dim idx1(0) As DataColumn
        idx1(0) = DS.Tables(1).Columns(0)
        DS.Tables(1).PrimaryKey = idx1

        'convert from excel to csv
        Dim myexcel As New ExportToExcelFile(Me)

        Dim mycsvfile = myexcel.convertfile(OpenFileDialog1.FileName, mymessage, {"Receive", "Send"})
        If mymessage <> "" Then
            ProgressReport("1", mymessage)
            Exit Sub
        End If



        Using objTFParser = New FileIO.TextFieldParser(mycsvfile(0))
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                '.SetDelimiters(Chr(9))
                .SetDelimiters(",")
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0


                ProgressReport(2, "Read Text File...")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)

                    'If i >= 0 And myrecord(0) <> "" And myrecord(1) <> "" Then
                    If i >= 0 And myrecord(0) <> "" Then
                        Dim pkey0(0) As Object
                        pkey0(0) = myrecord(0)
                        Dim result = DS.Tables(0).Rows.Find(pkey0)
                        If IsNothing(result) Then 'PackingList Found
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(0)
                            dr.Item(1) = DbAdapter1.validstr(myrecord(1))                      'fcrnumber  
                            dr.Item(2) = DbAdapter1.validstr(myrecord(2)) 'remark
                            dr.Item(3) = myrecord(3) 'date

                            DS.Tables(0).Rows.Add(dr)
                        Else

                            'check fcr
                            If IsDBNull(result.Item(1)) Then
                                result.Item(1) = myrecord(1)
                            Else
                                If result.Item(1) <> myrecord(1) Then
                                    result.Item(1) = DbAdapter1.validstr(myrecord(1))
                                End If
                            End If
                            'check remark
                            If IsDBNull(result.Item(2)) Then
                                result.Item(2) = myrecord(2)
                            Else
                                If result.Item(2) <> myrecord(2) Then
                                    result.Item(2) = DbAdapter1.validstr(myrecord(2))
                                End If
                            End If
                            'Check date
                            If IsDBNull(result.Item(3)) Then
                                result.Item(3) = myrecord(3)
                            Else
                                If result.Item(3) <> myrecord(3) Then
                                    result.Item(3) = myrecord(3)
                                End If
                            End If
                        End If
                    End If
                    'Housebill
                    If i >= 0 And myrecord(4) <> "" And myrecord(5) <> "" Then
                        Dim pkey1(0) As Object
                        pkey1(0) = myrecord(4)
                        Dim result = DS.Tables(1).Rows.Find(pkey1)
                        If IsNothing(result) Then 'PackingList Found
                            Dim dr As DataRow = DS.Tables(1).NewRow
                            dr.Item(0) = myrecord(4)
                            dr.Item(1) = myrecord(5)                      'date

                            DS.Tables(1).Rows.Add(dr)
                        Else

                            'Check date
                            If IsDBNull(result.Item(1)) Then
                                result.Item(1) = myrecord(5)
                            Else
                                If result.Item(1) <> myrecord(5) Then
                                    result.Item(1) = myrecord(5)
                                End If
                            End If
                        End If
                    End If

                Next
            End With
        End Using

        Using objTFParser = New FileIO.TextFieldParser(mycsvfile(1))
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(",")
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                mylist = New List(Of String())

                ProgressReport(2, "Read Text File...")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)

                    'Housebill
                    '" select housebill,receiveddate,courier,trackingno,senddate,housebillid from housebilldoc;"
                    If i >= 0 And myrecord(0) <> "" Then
                        Dim pkey1(0) As Object
                        pkey1(0) = myrecord(0)
                        Dim result = DS.Tables(1).Rows.Find(pkey1)
                        If IsNothing(result) Then 'PackingList Found
                            Dim dr As DataRow = DS.Tables(1).NewRow
                            dr.Item(0) = myrecord(0)
                            dr.Item(2) = myrecord(1)                      'date
                            dr.Item(3) = DbAdapter1.validstr(myrecord(2))
                            dr.Item(4) = myrecord(3)
                            DS.Tables(1).Rows.Add(dr)
                        Else

                            'Check date
                            If IsDBNull(result.Item(4)) Then
                                result.Item(4) = myrecord(3)
                            Else
                                If result.Item(4) <> myrecord(3) Then
                                    result.Item(4) = myrecord(3)
                                End If
                            End If
                            'Check courier
                            If IsDBNull(result.Item(2)) Then
                                result.Item(2) = DbAdapter1.validstr(myrecord(1))
                            Else
                                If result.Item(2) <> myrecord(1) Then
                                    result.Item(2) = DbAdapter1.validstr(myrecord(1))
                                End If
                            End If
                            'Check trackingno
                            If IsDBNull(result.Item(3)) Then
                                result.Item(3) = DbAdapter1.validstr(myrecord(2))
                            Else
                                If result.Item(3) <> myrecord(2) Then
                                    result.Item(3) = DbAdapter1.validstr(myrecord(2))
                                End If
                            End If
                        End If
                    End If

                Next
            End With
        End Using

        For i = 0 To mycsvfile.Count - 1
            File.Delete(mycsvfile(i))
        Next

        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            'If ImportHardCopyHdSB.Length > 0 Then
            ProgressReport(2, "Update PackingListHousebill")

            'sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "'," & DateFormatyyyyMMdd(startdate) & ");"
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'myret = DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmessage)
            'If Not myret Then
            '    ProgressReport(2, "Insert InvoiceHardCopy HD" & "::" & errmessage)
            '    Exit Sub
            'End If
            ' End If

            Dim ds2 = DS.GetChanges()
            If Not IsNothing(ds2) Then
                'Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.UploadInvoiceReceivedDateTx(Me, mye) Then
                    ProgressReport(2, "Update InvoiceReceivedDate" & "::" & mye.message)
                    ProgressReport(5, "Continue")
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub




    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        'If e.KeyCode = System.Windows.Forms.Keys.Enter Then
        '    'MessageBox.Show("Click Enter keydown")
        'End If
    End Sub





    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress, TextBox4.KeyPress, TextBox5.KeyPress, TextBox6.KeyPress, DateTimePicker2.KeyPress, TextBox8.KeyPress, TextBox16.KeyPress
        If (Char.IsLower(e.KeyChar)) Then
            e.KeyChar = Char.ToUpper(e.KeyChar)
            'ElseIf e.KeyChar = Convert.ToChar(Keys.Enter) Then
            '    'MessageBox.Show("Click Enter")
        End If


    End Sub
    Private Sub TextBox4_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Leave
        docreceiveloaddata(TextBox4.Text)
        'TextBox4.Text = ""
    End Sub
    Private Sub TextBox3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyUp, TextBox4.KeyUp, TextBox5.KeyUp, TextBox6.KeyUp, DateTimePicker2.KeyUp, TextBox8.KeyUp, TextBox15.KeyUp, TextBox16.KeyUp, TextBox12.KeyUp, TextBox13.KeyUp, TextBox14.KeyUp
        'TextBox3.Text = TextBox3.Text.ToUpper
        'TextBox3.Select(TextBox3.Text.Length, 0)
        If e.KeyCode = System.Windows.Forms.Keys.Enter Then
            Select Case sender.name
                Case "TextBox3"
                    loaddata(TextBox3.Text)
                    source.Add(TextBox3.Text)
                    TextBox3.Text = ""
                Case "TextBox4", "TextBox5", "TextBox6", "DateTimePicker2", "TextBox15"
                    'assign to browser temp
                    'find detail done by on leave.

                    If Me.validate Then
                        source.Add(TextBox4.Text)
                        sourcebill.Add(TextBox5.Text)
                        'add new record 
                        '"select ''::character varying as supplierinvoicenumber,''::character varying as billoflading,''::character varying as fcrnumber,null::date as receiveddate from paramhd where paramname='hewllo'"
                        Dim dr As DataRowView = ReceivedTmpBS.AddNew()
                        dr.Row.Item("supplierinvoicenumber") = IIf(TextBox4.Text = "", DBNull.Value, TextBox4.Text)
                        dr.Row.Item("billoflading") = IIf(TextBox5.Text = "", DBNull.Value, TextBox5.Text)
                        dr.Row.Item("fcrnumber") = IIf(TextBox6.Text = "", DBNull.Value, TextBox6.Text)
                        dr.Row.Item("receiveddate") = DateTimePicker2.Value
                        dr.Row.Item("remarks") = IIf(TextBox15.Text = "", DBNull.Value, TextBox15.Text)
                        ReceivedTmpDS.Tables(0).Rows.Add(dr.Row)
                        ReceivedTmpBS.EndEdit()
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox15.Text = ""
                        docreceiveloaddata(TextBox4.Text)
                        DateTimePicker2.Value = Date.Today
                        TextBox4.Focus()
                    End If
                Case "TextBox8"
                    If Me.validatesenddoc Then
                        'Find receiveddate, Market Name
                        sourcebill.Add(TextBox8.Text)
                        Dim sqlstr = "select receiveddate from housebilldoc where housebill= '" & TextBox8.Text.Replace("'", "''") & "'"
                        Dim myreceiveddate As Object = DBNull.Value
                        Dim myhousebill As String = ""
                        If DbAdapter1.ExecuteScalar(sqlstr, myreceiveddate) Then

                        End If
                        sqlstr = "select c.customername::character varying from packinglisthd phd" &
                                 " left join packinglistdocument pld  on pld.typedoc = 1 and pld.delivery = phd.delivery" &
                                 " left join billinghd bh on bh.billingdocument = pld.docno" &
                                 " left join customer c on c.customercode = bh.soldtoparty" &
                                 " where housebill = '" & TextBox8.Text.Replace("'", "''") & "' limit 1"

                        sqlstr = "select c.customername::character varying from packinglisthd phd" &
                                 " left join packinglistdt pldt on pldt.delivery = phd.delivery" &                                 
                                 " left join cxsebpodtl pdt on pdt.sebasiapono = pldt.pohd and pdt.polineno = pldt.poitem" &
                                 " left join cxrelsalesdocpo r on r.cxsebpodtlid = pdt.cxsebpodtlid" &
                                 " left join cxsalesorderdtl sdt on sdt.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                                 " left join cxsalesorder sdh on sdh.sebasiasalesorder = sdt.sebasiasalesorder " &
                                 " left join customer c on c.customercode = sdh.soldtoparty" &
                                 " where housebill = '" & TextBox8.Text.Replace("'", "''") & "' limit 1"
                        sqlstr = "select c.customername::character varying from packinglisthd phd" &
                                 " left join packinglistdt pldt on pldt.delivery = phd.delivery" &
                                 " left join packinglisthousebill plhb on plhb.delivery = phd.delivery" &
                                 " left join cxsebpodtl pdt on pdt.sebasiapono = pldt.pohd and pdt.polineno = pldt.poitem" &
                                 " left join cxrelsalesdocpo r on r.cxsebpodtlid = pdt.cxsebpodtlid" &
                                 " left join cxsalesorderdtl sdt on sdt.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                                 " left join cxsalesorder sdh on sdh.sebasiasalesorder = sdt.sebasiasalesorder " &
                                 " left join customer c on c.customercode = sdh.soldtoparty" &
                                 " where plhb.housebill = '" & TextBox8.Text.Replace("'", "''") & "' limit 1"

                        If DbAdapter1.ExecuteScalar(sqlstr, MyMarketName) Then

                        End If
                        'Add record in browser
                        Dim drv As DataRowView = SendDateBS.AddNew


                        drv.Row.Item("billoflading") = TextBox8.Text
                        drv.Row.Item("receiveddate") = IIf(IsNothing(myreceiveddate), DBNull.Value, myreceiveddate)
                        drv.Row.Item("marketname") = IIf(IsNothing(MyMarketName), DBNull.Value, MyMarketName)

                        SenddateTmpDS.Tables(0).Rows.Add(drv.Row)
                        TextBox8.Text = ""
                    End If
                Case "TextBox16"
                    Button17.PerformClick()
                Case "TextBox12", "TextBox13", "TextBox14"
                    Button16.PerformClick()
            End Select
            ProgressReport(2, "")

        End If
    End Sub
    Private Function validatesenddoc() As Boolean
        Me.validate()
        Dim myret As Boolean = True
        If TextBox8.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox8, "Bill of lading cannot be blank!")
        Else
            ErrorProvider1.SetError(TextBox8, "")
        End If

        Return myret
    End Function
    Private Overloads Function validate() As Boolean
        MyBase.Validate()
        Dim myret As Boolean = True
        If TextBox6.Text <> "" And TextBox4.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox4, "Supplier Invoice Number cannot be blank!")
        ElseIf (TextBox15.Text <> "" And TextBox4.Text = "") And (TextBox15.Text <> "" And TextBox5.Text = "") Then
            myret = False
            If TextBox4.Text = "" Then
                ErrorProvider1.SetError(TextBox4, "Supplier Invoice Number cannot be blank!")
            End If

            If TextBox5.Text = "" Then
                ErrorProvider1.SetError(TextBox5, "Bill of lading cannot be blank!")
            End If

        ElseIf TextBox6.Text = "" And TextBox4.Text = "" And TextBox5.Text = "" Then
            myret = False
        Else
            ErrorProvider1.SetError(TextBox4, "")
            ErrorProvider1.SetError(TextBox5, "")
        End If

        Return myret
    End Function


    Private Sub loaddata(ByVal p1 As String)
        Dim sqlstr = "select supplierinvoice, sum(amount) as amount,sum(extracost) as extracost,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,receiveddate, invoicehardcopydtid from (" &
                     " select  tb.* from sp_gethardcopyextra('" & p1.Replace("'", "''") & "'::character varying) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying,receiveddate date))as foo" &
                     " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading,receiveddate"

        SInvDS = New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(6, "Marque")

        If Not DbAdapter1.TbgetDataSet(sqlstr, SInvDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            Try
                SinvBS = New BindingSource
                SinvBS.DataSource = SInvDS.Tables(0)

                DataGridView2.AutoGenerateColumns = False
                DataGridView2.DataSource = SinvBS
                If SInvDS.Tables(0).Rows.Count > 0 Then
                    'ListBox1.Items.Add(p1)
                    ListBox1.Items.Insert(0, p1)
                    showlabel()
                End If

            Catch ex As Exception
                ProgressReport(5, ex.Message)
            End Try

        End If
        'ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", MyDS.Tables(0).Rows(0).Item(0)))
        ProgressReport(5, "Continues")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        For i As Integer = 0 To ListBox1.SelectedIndices.Count - 1
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        Next
        showlabel()
    End Sub

    Private Sub showlabel()
        Label7.Text = "Total entries: " & ListBox1.Items.Count
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        'prepare detail
        If ComboBox3.Text = "" Then
            MessageBox.Show("Please select User Name.")
            Exit Sub
        End If
        username = ComboBox3.Text
        If ListBox1.Items.Count = 0 Then
            MessageBox.Show("Nothing to save.")
            Exit Sub
        End If
        Dim myID As Long
        Dim mymessage As String = String.Empty
        Dim errmessage As String = String.Empty
        Dim ImportHardCopyDTsb As New StringBuilder
        Dim myret As Boolean
        ProgressReport(2, "Get Id")
        Dim sqlstr As String = String.Empty
        sqlstr = "select nextval('invoicehardcopyhd_invoicehardcopyhdid_seq');"
        If Not DbAdapter1.ExecuteScalar(sqlstr, myID, mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        For i = 0 To ListBox1.Items.Count - 1
            'find the record in existing table.
            ImportHardCopyDTsb.Append(myID & vbTab &
                                       ListBox1.Items(i) & vbCrLf)

        Next
        'insert header
        'copy detail
        Dim mydate As DateTime = Date.Now
        If ImportHardCopyDTsb.Length > 0 Then
            ProgressReport(2, "Copy Invoice Hard Copy Detail")

            sqlstr = "insert into invoicehardcopyhd(invoicehardcopyhdid,username,dateupload) values(" & myID & ",'" & username & "','" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", mydate) & "');copy invoicehardcopydt(invoicehardcopyhdid,supplierinvoicenumber) from stdin with null as 'Null';"

            errmessage = DbAdapter1.copy(sqlstr, ImportHardCopyDTsb.ToString, myret)
            If Not myret Then
                ProgressReport(2, "Copy Invoice Hard Copy Detail" & "::" & errmessage)
                ProgressReport(5, "Continue")
                Exit Sub
            Else
                'print
                ExportToExcel(username, mydate, ComboBox8.Text)
                ListBox1.Items.Clear()
            End If
        End If
    End Sub

    Public Sub ExportToExcel(ByVal myuser As String, ByVal dateupload As DateTime, ByVal myyear As Object, Optional ByVal withbilloflading As Boolean = True)

        'Dim dr As DataRow = CType(bs1.Current, DataRowView).Row
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty



        Dim sqlstr As String
        Dim sqlstr1 As String = String.Empty

        'If ComboBox1.Text <> "" Then
        myuser = "'" & myuser & "'"

        'dateupload = dr.Item("dateupload")
        'If myyear = "" Then
        '    If withbilloflading Then
        '        sqlstr = "select tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"

        '        sqlstr1 = " with foo as (select supplierinvoice,sum(amount) as amount,sum(extracost) as extracost,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice, invoicehardcopydtid from (select  tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '                 " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '                 " order by invoicehardcopydtid) " &
        '                 " select foo.supplierinvoice,foo.amount,foo.extracost,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
        '                 " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"

        '    Else
        '        sqlstr = "select tb.* from sp_gethardcopynobill(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
        '        'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopynobill(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '        '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '        '         " order by invoicehardcopydtid "

        '        sqlstr1 = "with foo as (select supplierinvoice,sum(amount) as amount,sum(extracost) as extracost,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,invoicehardcopydtid   from (select  tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp ) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '                " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '                " order by invoicehardcopydtid) " &
        '                " select foo.supplierinvoice,foo.amount,foo.extracost,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
        '                " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"
        '    End If
        'Else
        '    If withbilloflading Then
        '        sqlstr = "select tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp," & myyear & ") as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"

        '        sqlstr1 = " with foo as (select supplierinvoice,sum(amount) as amount,sum(extracost) as extracost,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice, invoicehardcopydtid from (select  tb.* from sp_gethardcopyextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp," & myyear & ") as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '                 " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '                 " order by invoicehardcopydtid) " &
        '                 " select foo.supplierinvoice,foo.amount,foo.extracost,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
        '                 " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"

        '    Else
        '        sqlstr = "select tb.* from sp_gethardcopynobill(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp," & myyear & ") as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
        '        'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopynobill(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '        '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '        '         " order by invoicehardcopydtid "

        '        sqlstr1 = "with foo as (select supplierinvoice,sum(amount) as amount,sum(extracost) as extracost,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,invoicehardcopydtid   from (select  tb.* from sp_gethardcopynobillextra(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp ," & myyear & ") as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '                " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '                " order by invoicehardcopydtid) " &
        '                " select foo.supplierinvoice,foo.amount,foo.extracost,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
        '                " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"
        '    End If
        'End If
        
        'sqlstr = "select tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"
        'sqlstr1 = "select supplierinvoice,sum(amount) as amount,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice,'' as receivedate   from (select  tb.* from sp_gethardcopy(" & myuser & "::character varying," & DateFormatyyyyMMdd(dateupload) & "::date) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
        '         " group by invoicehardcopydtid,supplierinvoice,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
        '         " order by invoicehardcopydtid "

        'Else
        '    'sqlstr = "select tb.*from sp_getaccountingdata(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & "::date," & DateFormatyyyyMMdd(DateTimePicker2.Value) & "::date ) as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)"
        '    MessageBox.Show("Please select User Name.")
        '    Exit Sub
        'End If

        sqlstr = "select tb.* from sp_gethardcopyextravendor(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorcode bigint,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)"

        sqlstr1 = " with foo as (select supplierinvoice,sum(amount) as amount,sum(extracost) as extracost,vendorcode,vendorname,readydate,accountingdoc,billoflading,deliverydate,sebinvoice, invoicehardcopydtid from (select  tb.* from sp_gethardcopyextravendor(" & myuser & "::character varying,'" & String.Format("{0:yyyy-MM-dd HH:mm:ss.fff}", dateupload) & "'::timestamp) as tb(invoicehardcopydtid bigint,supplierinvoice character varying,amount numeric,extracost numeric,vendorcode bigint,vendorname character varying,readydate date,accountingdoc bigint,delivery bigint,item integer,pohd bigint,poitem integer,deliverydate date,sebinvoice bigint,billoflading character varying)) foo " &
                 " group by invoicehardcopydtid,supplierinvoice,vendorcode,vendorname,readydate,accountingdoc,delivery,deliverydate,sebinvoice,billoflading" &
                 " order by invoicehardcopydtid) " &
                 " select foo.supplierinvoice,foo.amount,foo.extracost,foo.vendorcode,foo.vendorname,foo.readydate,foo.accountingdoc,foo.billoflading,foo.deliverydate,foo.sebinvoice,i.receiveddate from foo " &
                 " left join invoicehardcopyreceiveddate i on i.supplierinvoicenumber = foo.supplierinvoice order by invoicehardcopydtid"


        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "DocumentHardCopy" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "DETAIL",
                                                            .Sqlstr = sqlstr
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            myqueryworksheet = New QueryWorksheet With {.DataSheet = 2,
                                                            .SheetName = "Total",
                                                            .Sqlstr = sqlstr1
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            'Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

            myreport.Run(Me, New System.EventArgs)

        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.FormInvoiceHardCopy_Load(Me, e)
    End Sub




    Private Sub docreceiveloaddata(ByVal p1 As String)
        'Dim sqlstr = "select pm.amount::numeric,pd.pohd,pd.polineno,v.vendorname::character varying,c.customername::character varying,ph.pono,ah.reference from accountinghd ah" &
        '              " left join miro m on m.mironumber = ah.miro" &
        '              " left join pomiro pm on pm.miroid = m.miroid" &
        '              " left join podtl pd on pd.podtlid = pm.podtlid" &
        '              " left join pohd ph on ph.pohd = pd.pohd" &
        '              " left join packinglistdocument pld on pld.docno = ah.docno and pld.pohd = pd.pohd and pld.poitem = pd.polineno" &
        '              " left join packinglistdocument pldb on pldb.delivery = pld.delivery and pldb.item =pld.item and pldb.typedoc = 1" &
        '              " left join billinghd bh on bh.billingdocument = pldb.docno" &
        '              " left join vendor v on v.vendorcode = m.vendorcode" &
        '              " left join customer c on c.customercode = bh.soldtoparty" &
        '              " where reference = '" & p1.Replace("'", "''") & "'"


        Dim sqlstr = "select pm.amount::numeric,abs((bd.subtotal - bd.netvalue)) * (bd.exrate / getcurrratio(bd.curr)) as extracost,pd.pohd,pd.polineno,v.vendorname::character varying,c.customername::character varying,ph.pono,ah.reference from accountinghd ah" &
                      " left join miro m on m.mironumber = ah.miro" &
                      " left join pomiro pm on pm.miroid = m.miroid" &
                      " left join podtl pd on pd.podtlid = pm.podtlid" &
                      " left join pohd ph on ph.pohd = pd.pohd" &
                      " left join packinglistdocument pld on pld.docno = ah.docno and pld.pohd = pd.pohd and pld.poitem = pd.polineno" &
                      " left join packinglistdocument pldb on pldb.delivery = pld.delivery and pldb.item =pld.item and pldb.typedoc = 1" &
                      " left join cxsebpodtl pdt on pdt.sebasiapono = ph.pohd and pdt.polineno = pd.polineno" &
                      " left join cxrelsalesdocpo r on r.cxsebpodtlid = pdt.cxsebpodtlid" &
                      " left join cxsalesorderdtl sdt on sdt.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                      " left join cxsalesorder sdh on sdh.sebasiasalesorder = sdt.sebasiasalesorder " &
                      " left join billingdtl bd on bd.billingdocument = pldb.docno and bd.salesdoc = sdt.sebasiasalesorder and bd.salesdocitem = sdt.solineno" &
                      " left join vendor v on v.vendorcode = m.vendorcode" &
                      " left join customer c on c.customercode = sdh.soldtoparty" &
                      " where reference = '" & p1.Replace("'", "''") & "'"

        ReceivedDS = New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(6, "Marque")

        If Not DbAdapter1.TbgetDataSet(sqlstr, ReceivedDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            Try
                ReceivedBS = New BindingSource
                ReceivedBS.DataSource = ReceivedDS.Tables(0)

                DataGridView3.AutoGenerateColumns = False
                DataGridView3.DataSource = ReceivedBS
                If ReceivedDS.Tables(0).Rows.Count > 0 Then
                    'calculate amount


                    'source.Add(TextBox4.Text)
                End If
                showTotalAmount(ReceivedBS)
            Catch ex As Exception
                ProgressReport(5, ex.Message)
            End Try

        End If
        'ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", MyDS.Tables(0).Rows(0).Item(0)))
        ProgressReport(5, "Continues")
    End Sub

    Private Sub showTotalAmount(ByVal ReceivedBS As BindingSource)
        'Dim myQ = From c In ReceivedBS.List
        '           Group By reference = c.row.item("reference") Into mygroup = Group
        '           Select TotalAmount = mygroup.Sum(Function(x) x.row.item("amount"))

        'For Each k As Double In myQ
        '    Debug.Print(k)
        '    TextBox7.Text = k
        'Next
        Dim mytotal As Double = 0
        Dim mytotalExtracost As Double = 0
        For Each r As DataRowView In ReceivedBS.List
            mytotal = mytotal + r.Row.Item("amount")
            mytotalExtracost = mytotalExtracost + r.Row.Item("extracost")
        Next
        TextBox7.Text = String.Format("{0:#,##0.00}", mytotal)
        TextBox17.Text = String.Format("{0:#,##0.00}", mytotalExtracost)
    End Sub



    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                'If DataGridView4.SelectedRows.Count = 0 Then
                '    ReceivedTmpBS.RemoveAt(CM.Position)
                'Else
                For Each a As DataGridViewRow In DataGridView4.SelectedRows
                    ReceivedTmpBS.RemoveAt(a.Index)
                Next
                'End If

            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim ds2 As DataSet
        ds2 = ReceivedTmpDS.GetChanges
        Dim mymessage As String = String.Empty
        If Not IsNothing(ds2) Then
            'Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.InvoiceReceivedDateTx(Me, mye) Then
                ProgressReport(2, "Update InvoiceReceivedDate" & "::" & mye.message)
                ProgressReport(5, "Continue")
                Exit Sub
            End If
            loadreceivedtmpds()
            ProgressReport(2, "Update Done!")
        End If
    End Sub

    Private Sub loadreceivedtmpds()
        ReceivedTmpDS = New DataSet
        Dim mymessage As String = String.Empty
        Dim sqlstr = "select ''::character varying as supplierinvoicenumber,''::character varying as billoflading,''::character varying as fcrnumber,null::date as receiveddate ,''::character varying as remarks from paramhd where paramname='hewllo'"
        If Not DbAdapter1.TbgetDataSet(sqlstr, ReceivedTmpDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            ReceivedTmpBS = New BindingSource
            ReceivedTmpBS.DataSource = ReceivedTmpDS.Tables(0)
            ProgressReport(9, "Fill DataGrid")
            'cm = CType(BindingContext(ReceivedTmpBS), CurrencyManager)
        End If

    End Sub

    Private Sub loadsenddocumenttmpds()
        SenddateTmpDS = New DataSet
        Dim mymessage As String = String.Empty
        Dim sqlstr = "select ''::character varying as billoflading,null::date as receiveddate,''::character varying as marketname from paramhd where paramname = 'donotusethisname';"
        If Not DbAdapter1.TbgetDataSet(sqlstr, SenddateTmpDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            SendDateBS = New BindingSource
            SendDateBS.DataSource = SenddateTmpDS.Tables(0)
            ProgressReport(10, "Fill DataGrid")
            'cmsenddate = CType(BindingContext(SendDateBS), CurrencyManager)
        End If
    End Sub

    Private Sub loadtrackingnumberds()
        TrackingNoDS = New DataSet
        Dim mymessage As String = String.Empty
        Dim sqlstr = "select ''::character varying as trackingnumber,''::character varying as marketname from paramhd where paramname = 'donotusethisname';"
        If Not DbAdapter1.TbgetDataSet(sqlstr, TrackingNoDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            TrackingNoBS = New BindingSource
            TrackingNoBS.DataSource = TrackingNoDS.Tables(0)
            ProgressReport(11, "Fill DataGrid")
            'cmsenddate = CType(BindingContext(SendDateBS), CurrencyManager)
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If MessageBox.Show("Delete selected record?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                'If DataGridView4.SelectedRows.Count = 0 Then
                '    ReceivedTmpBS.RemoveAt(CM.Position)
                'Else
                For Each a As DataGridViewRow In DataGridView5.SelectedRows
                    SendDateBS.RemoveAt(a.Index)
                Next
                'End If

            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If validatesenddate() Then
            Dim ds2 = SenddateTmpDS.GetChanges()
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.UploadHousebillsenddate(Me, mye) Then
                    ProgressReport(2, "Update InvoiceReceivedDate" & "::" & mye.message)
                    ProgressReport(5, "Continue")
                    Exit Sub
                End If

                Dim drv As DataRowView = TrackingNoBS.AddNew
                drv.Row.Item("trackingnumber") = TextBox10.Text
                drv.Row.Item("marketname") = MyMarketName
                TrackingNoDS.Tables(0).Rows.Add(drv.Row)

                TextBox8.Text = ""
                TextBox9.Text = ""
                TextBox10.Text = ""

                loadsenddocumenttmpds()
            End If
        End If
    End Sub

    Private Function validatesenddate() As Boolean
        Dim myret As Boolean = True
        Me.validate()

        If TextBox9.Text = "" Then
            ErrorProvider1.SetError(TextBox9, "Value cannot be blank.")
            myret = False
        Else
            ErrorProvider1.SetError(TextBox9, "")
        End If
        If TextBox10.Text = "" Then
            ErrorProvider1.SetError(TextBox10, "Value cannot be blank.")
            myret = False
        Else
            ErrorProvider1.SetError(TextBox10, "")
        End If

        Return myret

    End Function

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim myform As New FormModifyTrackingNo
        myform.ShowDialog()

    End Sub



    Private Sub DataGridView6_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellDoubleClick
        If Not IsNothing(TrackingNoBS.Current) Then
            Dim drv As DataRowView = TrackingNoBS.Current
            Dim myform As New FormModifyTrackingNo(drv.Row.Item("trackingnumber"))
            myform.ShowDialog()
            If myform.modified Then
                drv.Item("trackingnumber") = myform.newtrackingnumber
            End If

        End If

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        'Dim dr As DataRow = CType(bs1.Current, DataRowView).Row
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty



        Dim sqlstr As String
        Dim sqlstr1 As String = String.Empty

        sqlstr = "select distinct hbd.housebill,courier,trackingno,c.customername::character varying,senddate,remarks " &
                 " from housebilldoc hbd " &
                 " left join packinglisthd phd on phd.housebill = hbd.housebill" &
                 " left join packinglistdocument pld on pld.delivery = phd.delivery and pld.typedoc = 1" &
                 " left join billinghd bh on billingdocument = pld.docno" &
                 " left join customer c on c.customercode = bh.soldtoparty" &
                 " where senddate >= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker4.Value) & " and senddate <= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker5.Value) & " order by housebill "

        sqlstr = "select distinct hbd.housebill,courier,trackingno,c.customername::character varying,senddate,remarks " &
                 " from housebilldoc hbd " &
                 " left join packinglisthousebill plhb on plhb.housebill = hbd.housebill" &
                 " left join packinglisthd phd on phd.delivery = plhb.delivery" &
                 " left join packinglistdocument pld on pld.dselect distinct hbd.housebill,courier,trackingno,c.customername::character varying,senddate,remarks  from housebilldoc hbd  left join packinglisthousebill plhb on plhb.housebill = hbd.housebill left join packinglisthd phd on phd.delivery = plhb.delivery left join packinglistdocument pld on pld.delivery = phd.delivery and pld.typedoc = 1 left join billinghd bh on billingdocument = pld.docno left join customer c on c.customercode = bh.soldtoparty where senddate >= '2014-02-01' and senddate <= '2014-03-03' order by housebill elivery = phd.delivery and pld.typedoc = 1" &
                 " left join billinghd bh on billingdocument = pld.docno" &
                 " left join customer c on c.customercode = bh.soldtoparty" &
                 " where senddate >= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker4.Value) & " and senddate <= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker5.Value) & " order by housebill "

        sqlstr = "select distinct hbd.housebill,courier,trackingno,c.customername::character varying,senddate,remarks " &
                 " from housebilldoc hbd " &
                 " left join packinglisthousebill plhb on plhb.housebill = hbd.housebill" &
                 " left join packinglisthd phd on phd.delivery = plhb.delivery" &
                 " left join packinglistdocument pld on pld.delivery = phd.delivery and pld.typedoc = 1" &
                 " left join billinghd bh on billingdocument = pld.docno" &
                 " left join customer c on c.customercode = bh.soldtoparty" &
                 " where senddate >= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker4.Value) & " and senddate <= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker5.Value) & " order by housebill "

        sqlstr = "select distinct hbd.housebill,courier,trackingno,c.customername::character varying,senddate,remarks " &
                 " from housebilldoc hbd " &
                 " left join packinglisthousebill plhb on plhb.housebill = hbd.housebill" &
                 " left join packinglisthd phd on phd.delivery = plhb.delivery" &
                 " left join packinglistdt pldt on pldt.delivery = phd.delivery " &
                 " left join cxsebpodtl pdt on pdt.sebasiapono = pldt.pohd and pdt.polineno = pldt.poitem " &
                 " left join cxrelsalesdocpo r on r.cxsebpodtlid = pdt.cxsebpodtlid " &
                 " left join cxsalesorderdtl sdt on sdt.cxsalesorderdtlid = r.cxsalesorderdtlid " &
                 " left join cxsalesorder sdh on sdh.sebasiasalesorder = sdt.sebasiasalesorder " &
                 " left join customer c on c.customercode = sdh.soldtoparty" &
                  " where senddate >= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker4.Value) & " and senddate <= " & String.Format("'{0:yyyy-MM-dd}'", DateTimePicker5.Value) & " order by housebill "


        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "SendDocumentHardCopy" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "DETAIL",
                                                            .Sqlstr = sqlstr
                                                            }
            myQueryWorksheetList.Add(myqueryworksheet)

            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback)

            myreport.Run(Me, New System.EventArgs)

        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        ViewStartdate = DateTimePicker6.Value
        ViewLastDate = DateTimePicker7.Value
        ViewFieldname = ComboBox4.SelectedIndex
        ViewFieldvalue = TextBox11.Text
        loadviewdata()
    End Sub

    Private Sub loadviewdata()

        If Not myLoadView.IsAlive Then
            myLoadView = New System.Threading.Thread(AddressOf doLoadViewData)
            myLoadView.Start()
        Else
            MessageBox.Show("Please wait, current process still running...")
        End If

    End Sub

    Sub doLoadViewData()
        'Get All user from PackingListDtl
        '            SupplierInvoice()
        'B/L
        '            Miro#()
        '            Billing(Doc#)
        '            Document#()
        '            Supplier(Name)
        '            Market(Name)
        'Sinv 1st received date
        'Sinv Last received date
        'B/L First Received date
        'B/L Last Received date
        'Submit Date to A/C
        '            Courier(Name)
        ProgressReport(6, "Marque")
        Dim myfields() = {"", "ah.reference", "hbd.housebill", "m.mironumber::character varying", "pld.docno::character varying", "pld2.docno::character varying", "v.vendorname::character varying", "customername::character varying", "ihdrd.receiveddate::character varying", "ihdrd.lastreceiveddate::character varying", "hbd.receiveddate::character varying", "hbd.lastreceiveddate::character varying", "ihhd.dateupload::character varying", "hbd.courier", "hbd.trackingno", "hbd.senddate::character varying"}
        myCriteria = ""

        If Not (ViewFieldname < 1) Then
            myCriteria = " and " & myfields(ViewFieldname) & " = '" & ViewFieldvalue & "'"
        End If

        'Dim sqlstr = "select ah.reference,hbd.housebill,pm.amount,m.mironumber::character varying,pld.docno::character varying as docno1,pld2.docno::character varying as docno2,ah.postingdate::character varying,v.vendorname::character varying,c.customername::character varying,ihdrd.receiveddate::character varying as sinvreceiveddate,ihdrd.lastreceiveddate::character varying as sinvlastreceiveddate,hbd.receiveddate::character varying blreceiveddate,hbd.lastreceiveddate::character varying as bllastreceiveddate,ihhd.dateupload::character varying as submitdate,ihdrd.fcrnumber::character varying,hbd.courier,hbd.trackingno,hbd.senddate::character varying ,phd.deliverydate::character varying" &
        '             " from housebilldoc hbd" &
        '             " left join packinglisthd phd on phd.housebill = hbd.housebill" &
        '             " left join packinglistdocument pld on pld.delivery = phd.delivery and pld.typedoc = 1" &
        '             " left join packinglistdt pldt on pldt.delivery = phd.delivery and pldt.deliveryitem = pld.item" &
        '             " left join packinglistdocument pld2 on pld2.delivery = phd.delivery and pld2.item = pld.item and pld2.pohd = pld.pohd and pld2.poitem = pld.poitem and pld2.typedoc = 2" &
        '             " left join accountinghd ah on ah.docno = pld2.docno" &
        '             " left join miro m on m.mironumber = ah.miro" &
        '             " left join podtl pd on pd.pohd = pld.pohd and pd.polineno = pld.poitem" &
        '             " left join pomiro pm on pm.miroid = m.miroid and pm.podtlid = pd.podtlid" &
        '             " left join billinghd bh on billingdocument = pld.docno" &
        '             " left join billingdtl bd on bd.billingdocument = bh.billingdocument and bd.item = pld2.poitem" &
        '             " left join customer c on c.customercode = bh.soldtoparty" &
        '             " left join vendor v on v.vendorcode = m.vendorcode" &
        '             " left join invoicehardcopyreceiveddate ihdrd on ihdrd.supplierinvoicenumber = m.supplierinvoicenum" &
        '             " left join invoicehardcopydt ihdt on ihdt.supplierinvoicenumber = m.supplierinvoicenum" &
        '             " left join invoicehardcopyhd ihhd on ihhd.invoicehardcopyhdid = ihdt.invoicehardcopyhdid" &
        '             " where  ah.postingdate >=  '" & String.Format("{0:yyyy-MM-dd}", ViewStartdate) & "' and ah.postingdate <= '" & String.Format("{0:yyyy-MM-dd}", ViewLastDate) & "' " & myCriteria &
        '             " order by pld.delivery,pld.item"
        Dim sqlstr = "select ah.reference,phd.housebill,pm.amount,abs((bd.subtotal - bd.netvalue)) * (bd.exrate / getcurrratio(bd.curr)) as extracost, m.mironumber::character varying,pld.docno::character varying as docno1,pld2.docno::character varying as docno2,ah.postingdate::character varying,v.vendorname::character varying,c.customername::character varying,ihdrd.receiveddate::character varying as sinvreceiveddate,ihdrd.lastreceiveddate::character varying as sinvlastreceiveddate,hbd.receiveddate::character varying blreceiveddate,hbd.lastreceiveddate::character varying as bllastreceiveddate,ihhd.dateupload::character varying as submitdate,ihdrd.fcrnumber::character varying,ihdrd.lastreceiveddatefcr::character varying,hbd.courier,hbd.trackingno,hbd.senddate::character varying ,phd.deliverydate::character varying" &
                    " from packinglisthd phd" &
                    " left join housebilldoc hbd on hbd.housebill = phd.housebill " &
                    " left join packinglistdocument pld on pld.delivery = phd.delivery and pld.typedoc = 1" &
                    " left join packinglistdt pldt on pldt.delivery = phd.delivery and pldt.deliveryitem = pld.item" &
                    " left join packinglistdocument pld2 on pld2.delivery = phd.delivery and pld2.item = pld.item and pld2.pohd = pld.pohd and pld2.poitem = pld.poitem and pld2.typedoc = 2" &
                    " left join accountinghd ah on ah.docno = pld2.docno" &
                    " left join miro m on m.mironumber = ah.miro" &
                    " left join podtl pd on pd.pohd = pld.pohd and pd.polineno = pld.poitem" &
                    " left join cxsebpodtl pdt on pdt.sebasiapono = pd.pohd and pdt.polineno = pd.polineno" &
                    " left join cxrelsalesdocpo r on r.cxsebpodtlid = pdt.cxsebpodtlid" &
                    " left join cxsalesorderdtl sdt on sdt.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                    " left join cxsalesorder sdh on sdh.sebasiasalesorder = sdt.sebasiasalesorder " &
                    " left join billingdtl bd on bd.billingdocument = pld.docno and bd.salesdoc = sdt.sebasiasalesorder and bd.salesdocitem = sdt.solineno" &
                    " left join pomiro pm on pm.miroid = m.miroid and pm.podtlid = pd.podtlid" &
                    " left join billinghd bh on bh.billingdocument = pld.docno" &
                    " left join customer c on c.customercode = sdh.soldtoparty" &
                    " left join vendor v on v.vendorcode = m.vendorcode" &
                    " left join invoicehardcopyreceiveddate ihdrd on ihdrd.supplierinvoicenumber = m.supplierinvoicenum" &
                    " left join invoicehardcopydt ihdt on ihdt.supplierinvoicenumber = m.supplierinvoicenum" &
                    " left join invoicehardcopyhd ihhd on ihhd.invoicehardcopyhdid = ihdt.invoicehardcopyhdid" &
                    " where   ah.postingdate >=  '" & String.Format("{0:yyyy-MM-dd}", ViewStartdate) & "' and ah.postingdate <= '" & String.Format("{0:yyyy-MM-dd}", ViewLastDate) & "' " & myCriteria &
                    " order by pld.delivery,pld.item"

        Dim viewDS As New DataSet
        Dim mymessage As String = String.Empty
        ProgressReport(2, "")
        If Not DbAdapter1.TbgetDataSet(sqlstr, viewDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            Try
                ViewBS = New BindingSource
                ViewBS.DataSource = viewDS.Tables(0)
                ProgressReport(12, "Fill Datagridview View Data")
                ProgressReport(2, "Done")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If
        ProgressReport(5, "Continuous")
    End Sub



    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If Not myLoadView.IsAlive Then
            Me.validate()
            ViewBS.Filter = ""
            Dim myfilter As String = ""
            Dim myfields() = {"", "reference", "housebill", "mironumber", "docno1", "docno2", "vendorname", "customername", "sinvreceiveddate", "sinvlastreceiveddate", "blreceiveddate", "bllastreceiveddate", "submitdate", "courier", "trackingno", "senddate"}
            If ComboBox5.SelectedIndex > 1 Then
                If TextBox12.Text <> "" Then
                    myfilter = "[" & myfields(ComboBox5.SelectedIndex) & "] like '" & TextBox12.Text & "'"
                Else
                    myfilter = "[" & myfields(ComboBox5.SelectedIndex) & "] is null"
                End If
            End If
            If ComboBox6.SelectedIndex > -1 Then
                If TextBox13.Text <> "" Then
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox6.SelectedIndex) & "] like '" & TextBox13.Text & "'"

                Else
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox6.SelectedIndex) & "] is null"


                End If
            End If
            If ComboBox7.SelectedIndex > -1 Then
                If TextBox14.Text <> "" Then
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox7.SelectedIndex) & "] like '" & TextBox14.Text & "'"

                Else
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox7.SelectedIndex) & "] is null"


                End If
            End If
            Try
                ViewBS.Filter = myfilter
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

    

        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        ViewBS.Filter = ""
        ComboBox5.SelectedIndex = -1
        ComboBox6.SelectedIndex = -1
        ComboBox7.SelectedIndex = -1
        TextBox12.Text = ""
        TextBox13.Text = ""

        TextBox14.Text = ""
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Me.validate()
        housebilltext = TextBox16.Text

        If Not myHousebilldata.IsAlive Then
            myHousebilldata = New System.Threading.Thread(AddressOf doLoadHousebillData)
            myHousebilldata.Start()
        Else
            MessageBox.Show("Please wait, current process still running...")
        End If
    End Sub

    Sub doLoadHousebillData()

        'If housebilltext = "" Then
        '    ViewHousebillBS = New BindingSource
        '    DataGridView8.Invalidate()
        '    Exit Sub
        'End If
        ProgressReport(6, "Marque")
        Dim sqlstr = "select housebill,receiveddate,lastreceiveddate,courier,trackingno,senddate,remarks" &
                    " from housebilldoc " &
                    " where   housebill = '" & housebilltext.Replace("'", "''") & "'"

        Dim viewHousebillDS As New DataSet

        Dim mymessage As String = String.Empty
        ProgressReport(2, "")
        If Not DbAdapter1.TbgetDataSet(sqlstr, viewHousebillDS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            Try
                ViewHousebillBS = New BindingSource
                ViewHousebillBS.DataSource = viewHousebillDS.Tables(0)
                ProgressReport(13, "Fill Datagridview View Data")
                If viewHousebillDS.Tables(0).Rows.Count = 0 Then
                    ProgressReport(2, "Record not available.")
                Else
                    ProgressReport(2, "Done.")
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If
        ProgressReport(5, "Continuous")
    End Sub








End Class