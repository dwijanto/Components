Imports Components.PublicClass
Imports System.Threading
Imports Components.SharedClass
Imports Microsoft.Exchange.WebServices.Data
Public Class FormShowAllItemsCP
    Dim myShowAllItemsDelegate As New ThreadStart(AddressOf ShowAllItems)
    Dim myShowAllItems As New System.Threading.Thread(myShowAllItemsDelegate)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Public Property username As String
    Public Property password As String
    Public Property selecteduser As String
    Public Property startdate As Date
    Public Property enddate As Date

    Dim BS As BindingSource
    Dim DS As DataSet

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FormShowAllItems_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton1.Click
        LoadData()
    End Sub
    Public Sub LoadData()
        If Not myShowAllItems.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""

            myShowAllItems = New Thread(AddressOf ShowAllItems)
            myShowAllItems.Start()
        Else
            MessageBox.Show("Please wait until current process finished!")
        End If
    End Sub

    Sub ShowAllItems()
        ProgressReport(1, "Loading Data,Please wait...")
        ProgressReport(6, "Marque")
        Dim myuser As String = String.Empty
        myuser = IIf(selectedUser = "", "", ",'" & selectedUser & "'")

        'Dim sqlincomplete = "with bl as (" &
        '                    " select distinct docemailname,docemaildtname from docemailhd dh" &
        '                    " left join docemaildt dt on dt.docemailhdid = dh.docemailhdid" &
        '                    " where(docemailtype = 0 And mycontains(dt.docemaildtname, dh.docemailname))" &
        '                    " )," &
        '                    " tb as (" &
        '                    " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
        '                    " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
        '                    " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '                    " )" &
        '                    " select tb.billingdoc::character varying,tb.delivery::character varying,tb.reference,tb.housebill,getfilename(tb.housebill::character varying)as billoflading,  getfilename(tb.billingdoc::character varying)as invoice,getfilename(tb.delivery::character varying) as packinglist,draftcreateddate::character varying from tb" &
        '                    " left join docemailtx tx on tx.billoflading = tb.housebill" &
        '                    " left join bl on bl.docemailname = tb.housebill"
        'Dim sqlincomplete = "with tb as (" &
        '                   " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
        '                   " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
        '                   " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '                   " )" &
        '                   " select tb.billingdoc::character varying,tb.delivery::character varying,tb.reference,tb.housebill,getfilename(tb.housebill::character varying)as billoflading,  getfilename(tb.billingdoc::character varying)as invoice,getfilename(tb.delivery::character varying) as packinglist,draftcreateddate::character varying from tb" &
        '                   " left join docemailtx tx on tx.billoflading = tb.housebill"
        Dim sqlincomplete = "with tb as ( " &
                            " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
                            " from sp_getaccountingdata('2014-6-1'::date,'2014-6-13'::date ,'FWONG')  as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text) ) " &
                            " select tb.delivery::character varying,tb.reference,tb.housebill,tb.billingdoc::character varying,getfilenamecp(tb.housebill::character varying)as billoflading	" &
                            " from tb;"

        DS = New DataSet
        Dim mymessage As String = String.Empty

        If DbAdapter1.TbgetDataSet(sqlincomplete, DS, mymessage) Then

            Dim view As DataView = New DataView(DS.Tables(0))
            ProgressReport(4, "InitDataSource")



        Else
            MessageBox.Show(mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(5, "Continuous")
        ProgressReport(1, "Loading Data, Done.")
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(ByVal username As String, ByVal password As String, ByVal selecteduser As String, ByVal startdate As Date, ByVal enddate As Date)

        ' This call is required by the designer.
        InitializeComponent()
        Me.username = username
        Me.password = password
        Me.selecteduser = selecteduser
        Me.startdate = startdate
        Me.enddate = enddate
        ' Add any initialization after the InitializeComponent() call.

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
                    BS = New BindingSource
                    BS.DataSource = DS.Tables(0)

                    With DataGridView1
                        .AutoGenerateColumns = False
                        .DataSource = BS
                    End With
                    ' ToolStripComboBox1.ComboBox.Text = "Billing Doc"
                    ToolStripComboBox1.ComboBox.SelectedIndex = 0
                    'ToolStripComboBox2.ComboBox.Text = "Bill Of Lading"
                    'ToolStripComboBox2.ComboBox.SelectedIndex = 0
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


            End Select

        End If

    End Sub


    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged, ToolStripTextBox1.Click
        Dim myobj = CType(sender, ToolStripTextBox)
        Select Case myobj.Name
            Case "ToolStripTextBox1"
                Dim myfields() = {"delivery", "reference", "housebill", "billoflading", "billingdoc"}
                Try
                    BS.Filter = ""
                    If ToolStripTextBox1.Text <> "" Then
                        BS.Filter = "[" & myfields(ToolStripComboBox1.ComboBox.SelectedIndex) & "] like '" & ToolStripTextBox1.Text & "'"
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
        End Select
    End Sub


    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        ToolStripTextBox1.PerformClick()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myShowAllItems.IsAlive Then
            Me.Validate()
            BS.Filter = ""
            Dim myfilter As String = ""
            Dim myfields() = {"delivery", "reference", "housebill", "billoflading", "billingdoc"}
            If ComboBox1.SelectedIndex <> -1 Then
                If TextBox1.Text <> "" Then
                    myfilter = "[" & myfields(ComboBox1.SelectedIndex) & "] like '" & TextBox1.Text & "'"
                Else
                    myfilter = "[" & myfields(ComboBox1.SelectedIndex) & "] is null"
                End If
            End If
            If ComboBox2.SelectedIndex <> -1 Then
                If TextBox2.Text <> "" Then
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox2.SelectedIndex) & "] like '" & TextBox2.Text & "'"

                Else
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox2.SelectedIndex) & "] is null"


                End If
            End If

            If ComboBox3.SelectedIndex <> -1 Then
                If TextBox3.Text <> "" Then
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox3.SelectedIndex) & "] like '" & TextBox3.Text & "'"
                Else
                    myfilter = myfilter & IIf(myfilter = "", "", " and ") & "[" & myfields(ComboBox3.SelectedIndex) & "] is null"
                End If

            End If
            BS.Filter = myfilter
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Not myShowAllItems.IsAlive Then
            BS.Filter = ""
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
        End If
    End Sub
End Class