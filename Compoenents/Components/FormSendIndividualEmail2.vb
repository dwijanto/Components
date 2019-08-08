Imports Components.PublicClass
Imports System.Threading
Imports Components.SharedClass
Imports Microsoft.Exchange.WebServices.Data

Public Class FormSendIndividualEmail2

    Private _username As String
    Private _password As String
    Private _selectedUser As String
    Private _startdate As Date
    Private _enddate As Date
    Dim myDelegate As New ThreadStart(AddressOf doWork)
    Dim myThread As New System.Threading.Thread(myDelegate)
    Dim myProcess As New System.Threading.Thread(AddressOf doDraft)
    Dim emaildict As Dictionary(Of String, String)
    Dim emaillist As List(Of emailData)
    Dim attachmentdict As Dictionary(Of String, String)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)

    Dim BS As BindingSource
    Dim DS As DataSet
    Sub New(ByVal username As String, ByVal password As String, ByVal selectedUser As String, ByVal startdate As Date, ByVal enddate As Date)

        InitializeComponent()

        ' TODO: Complete member initialization 
        _username = username
        _password = password
        _selectedUser = selectedUser
        _startdate = startdate
        _enddate = enddate

    End Sub

    Sub doWork()
        ProgressReport(1, "Loading Data,Please wait...")
        ProgressReport(6, "Marque")
        Dim myuser As String = String.Empty
        myuser = IIf(_selectedUser = "", "", ",'" & _selectedUser & "'")
        Dim PrevSixMonth = _startdate.AddMonths(-6)
        'Dim sqlincomplete = "with tb as (" &
        '                   " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
        '                   " from sp_getaccountingdatastp(" & DateFormatyyyyMMdd(_startdate) & "::date," & DateFormatyyyyMMdd(_enddate) & "::date " & myuser & ") " &
        '                   " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '                   " )" &
        '                   " select false::boolean as toggle,tb.billingdoc::character varying,tb.delivery::character varying,tb.reference,tb.housebill,getfilename(tb.housebill::character varying)as billoflading,  getfilename(tb.billingdoc::character varying)as invoice,getfilename(tb.delivery::character varying) as packinglist,draftcreateddate::character varying,sender,soldtoparty from tb" &
        '                   " left join docemailtx tx on tx.billoflading = tb.housebill" &
        '                   " left join docemailhd dhd on dhd.docemailname = tb.housebill and docemailtype = 0" &
        '                   " left join billinghd bh on bh.billingdocument = tb.billingdoc;"
        Dim sqlincomplete = "with tb as (" &
                          " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
                          " from sp_getaccountingdatastp(" & DateFormatyyyyMMdd(_startdate) & "::date," & DateFormatyyyyMMdd(_enddate) & "::date " & myuser & ") " &
                          " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
                          " )" &
                          " select false::boolean as toggle,tb.billingdoc::character varying,tb.delivery::character varying,tb.reference,tb.housebill,getfilename(tb.housebill::character varying," & DateFormatyyyyMMdd(PrevSixMonth) & "::date)as billoflading,  getfilename(tb.billingdoc::character varying)as invoice,getfilename(tb.delivery::character varying) as packinglist,draftcreateddate::character varying,sender,soldtoparty from tb" &
                          " left join docemailtx tx on tx.billoflading = tb.housebill" &
                          " left join docemailhd dhd on dhd.docemailname = tb.housebill and docemailtype = 0" &
                          " left join billinghd bh on bh.billingdocument = tb.billingdoc;"
        Dim sqlstr = " select dt.* from paramdt dt" &
                     " left join paramhd hd on hd.paramhdid = dt.paramhdid" &
                     " where hd.paramname = 'logbook'" &
                     " order by dt.ivalue;"
        Dim sqldeliverybrand = "select * from sp_getdeliverybrand(" & DateFormatyyyyMMdd(_startdate) & "::date," & DateFormatyyyyMMdd(_enddate) & "::date " & myuser & ") as tb(delivery bigint,brand character varying);"
        Dim sqlcustomeremail = "select m.customercode,b.brandname::character varying,c.customername::character varying as displayname,name,email from marketemail m " &
                       " left join brand b on b.brandid = m.brandid  " &
                       " left join customer c on c.customercode = m.customercode " &
                       " where b.brandname isnull;" &
                       "select m.customercode,b.brandname::character varying,c.customername::character varying as displayname,name,email from marketemail m " &
                       " left join brand b on b.brandid = m.brandid  " &
                       " left join customer c on c.customercode = m.customercode " &
                       " where not b.brandname isnull;"

        DS = New DataSet
        Dim mymessage As String = String.Empty

        If DbAdapter1.TbgetDataSet(sqlincomplete & sqlstr & sqldeliverybrand & sqlcustomeremail, DS, mymessage) Then

            Dim view As DataView = New DataView(DS.Tables(0))
            ProgressReport(4, "InitDataSource")

            Dim pkey3(0) As DataColumn
            pkey3(0) = DS.Tables(3).Columns(0)
            DS.Tables(3).PrimaryKey = pkey3

            Dim pkey4(1) As DataColumn
            pkey4(0) = DS.Tables(4).Columns(0)
            pkey4(1) = DS.Tables(4).Columns(1)
            DS.Tables(4).PrimaryKey = pkey4

        Else
            MessageBox.Show(mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(5, "Continuous")
        ProgressReport(1, "Loading Data, Done.")
    End Sub

    Private Sub FormSendIndividualEmail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs)
        If myProcess.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormSendIndividualEmail_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then

            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""

            myThread = New Thread(AddressOf doWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until current process finished!")
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
                    BS = New BindingSource
                    BS.DataSource = DS.Tables(0)

                    With DataGridView1
                        .AutoGenerateColumns = False
                        .DataSource = BS
                    End With
                    ' ToolStripComboBox1.ComboBox.Text = "Billing Doc"
                    'ToolStripComboBox1.ComboBox.SelectedIndex = 0
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
                    MessageBox.Show("Done.")
            End Select

        End If

    End Sub



    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'For Each dsrow As DataGridViewRow In DataGridView1.SelectedRows
        '    'BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
        '    Dim row As DataRowView = DirectCast(dsrow.DataBoundItem, DataRowView)
        '    MessageBox.Show(dsrow.Cells(0).Value)
        'Next
        Me.Validate()

        If Not myProcess.IsAlive Then
            myProcess = New Thread(AddressOf doDraft)
            myProcess.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            Me.Validate()
            BS.Filter = ""
            Dim myfilter As String = ""
            Dim myfields() = {"billingdoc", "delivery", "reference", "housebill", "billoflading", "invoice", "packinglist", "draftcreateddate"}
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
        If Not myThread.IsAlive Then
            BS.Filter = ""
        End If
    End Sub
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Me.Validate()
        For i = 0 To DataGridView1.Rows.Count - 1
            'DataGridView1.Rows(i).Selected = True
            DataGridView1.Rows(i).Cells("Toggle").Value = True
        Next
        Me.Validate()
    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Me.Validate()
        For i = 0 To DataGridView1.Rows.Count - 1
            'DataGridView1.Rows(i).Selected = True
            DataGridView1.Rows(i).Cells("Toggle").Value = False
        Next
        Me.Validate()
    End Sub

    Private Function CreateDraft(ByVal d As Object) As Boolean
        Dim myret As Boolean = False
        Dim url As String = DS.Tables(1).Rows(1).Item("cvalue")  '"https://mail-eu.seb.com/ews/exchange.asmx"
        Dim service As ExchangeService
        Using myservice As New ClassEWS(url, _username, _password, "as", False)
            'service = myservice.CreateConnectionAutoDiscover()
            service = myservice.CreateConnection()
            Try
                'find ds.tables(0)
                'Dim myselect() As DataRow = d.mygroup 'DS.Tables(0).Select("housebill = '" & dr.Item(0) & "'")

                'If myselect(0).Item(4) <> 1 Then
                '    'Consolidation-> find in incomplete document. if found ->return false
                '    Dim recordfound() As DataRow = DS.Tables(1).Select("housebill = '" & dr.Item(0) & "'")
                '    If recordfound.Length <> 0 Then
                '        Return myret
                '    End If

                'End If

                Dim msg As EmailMessage = New EmailMessage(service)



                'ProgressReport(1, "Adding Attachment...")
                'create datatable to store email and displayname based on comma delimited
                'Dim emaildict As New Dictionary(Of String, String)
                'Dim emaillist As New List(Of emailData)

                emaildict = New Dictionary(Of String, String)
                emaillist = New List(Of emailData)
                attachmentdict = New Dictionary(Of String, String)

                Dim recepient As String = String.Empty
                Dim recepientname As String = String.Empty
                Dim marketname As String = String.Empty
                Dim mycc As String = String.Empty

                For Each mydr In d.mygroup

                    'if Soldtoparty US -> get brand from table delivery brand
                    '   For each brand -> get email from customer email
                    'add datatable based on comma delimited
                    'else
                    '   Find mydr.item("soldtoparty") name and email
                    '   add datatable name and email
                    'endif
                    Dim myresult As DataRow

                    mycc = ""
                    If Not IsDBNull(mydr.row.Item("sender")) Then
                        mycc = mydr.row.Item("sender")
                    End If
                    If Not IsDBNull(mydr.row.Item("soldtoparty")) Then


                        If mydr.row.Item("soldtoparty") = 99008400 Then 'US Market
                            Dim drow() As DataRow = DS.Tables(2).Select("delivery = '" & mydr.row.Item("delivery") & "'")
                            For Each row In drow
                                'Find 
                                Dim mykey(1) As Object
                                mykey(0) = 99008400
                                mykey(1) = row.Item("brand")
                                myresult = DS.Tables(4).Rows.Find(mykey)
                                If Not IsNothing(myresult) Then
                                    If Not IsDBNull(myresult.Item("email")) Then
                                        recepient = myresult.Item("email")
                                        recepientname = myresult.Item("name")
                                    Else
                                        recepient = ""
                                        recepientname = ""
                                    End If

                                    marketname = myresult.Item("displayname")
                                    addEmailList(recepient, recepientname)
                                End If
                            Next

                        Else
                            Dim mykey(0) As Object
                            mykey(0) = mydr.row.Item("soldtoparty")
                            myresult = DS.Tables(3).Rows.Find(mykey)
                            If Not IsNothing(myresult) Then
                                If Not IsDBNull(myresult.Item("email")) Then
                                    recepient = myresult.Item("email")
                                Else
                                    recepient = ""

                                End If

                                marketname = myresult.Item("displayname")
                                If Not IsDBNull(myresult.Item("name")) Then
                                    recepientname = myresult.Item("name")
                                Else
                                    recepientname = ""
                                End If


                                addEmailList(recepient, recepientname)
                                'Dim myrecepients() As String = recepient.Split(";")
                                'For i = 0 To myrecepients.Count - 1
                                '    If Not myrecepients(i).Length = 0 Then
                                '        If Not emaildict.ContainsKey(Trim(myrecepients(i))) Then
                                '            emaildict.Add(Trim(myrecepients(i)), recepientname)
                                '            'add email
                                '            emaillist.Add(New emailData With {.email = Trim(myrecepients(i)), .displayname = recepientname})
                                '        End If
                                '    End If

                                'Next

                            End If

                        End If
                    End If
                    Dim myfile As String = DS.Tables(1).Rows(4).Item("cvalue")
                    Dim forwarder As String = myfile & "\Forwarder\" & d.housebill & "\" & mydr.row.Item("billoflading")
                    Dim invoice As String = myfile & "\invoice\" & mydr.row.Item("invoice")
                    Dim packinglist As String = myfile & "\packinglist\" & mydr.row.Item("packinglist")
                    'mycc = mydr.Item("sender")

                    'msg.Attachments.AddFileAttachment(forwarder)
                    'msg.Attachments.AddFileAttachment(invoice)
                    'msg.Attachments.AddFileAttachment(packinglist)
                    If Not IsDBNull(mydr.row.Item("billoflading")) Then
                        addAttachmentList(forwarder, msg)
                    End If
                    If Not IsDBNull(mydr.row.Item("invoice")) Then
                        addAttachmentList(invoice, msg)
                    End If
                    If Not IsDBNull(mydr.row.Item("packinglist")) Then
                        addAttachmentList(packinglist, msg)
                    End If

                Next



                'for each data table 
                Dim recepientnamelist As String = String.Empty
                For Each mydata As emailData In emaillist
                    'If Not IsDBNull(mydata.email) Then
                    If Not mydata.email = "" Then
                        msg.ToRecipients.Add(New EmailAddress(mydata.email))
                    Else
                        If Not mycc = "" Then
                            msg.ToRecipients.Add(New EmailAddress(mycc))
                        End If
                    End If
                    recepientnamelist = recepientnamelist & IIf(recepientnamelist = "", "", ",") & mydata.displayname
                Next

                Dim mynamelist As String = IIf(recepientnamelist.Split(",").Count > 1, "All", recepientnamelist)

                'next
                If Not mycc = "" Then
                    msg.CcRecipients.Add(New EmailAddress(mycc))
                End If

                msg.Subject = String.Format("SEB Asia Document - {0} - {1} ", marketname, d.housebill)
                Dim bodymessage As String = "<!DOCTYPE html><html><head><meta charset=utf-8 /><style>p.normal{font-size:11.0pt;font-family:""Calibri"",""sans-serif"";}</style></head><body>" &
                                            "<p class=""normal"">Dear " & mynamelist & ",</p>" &
                                            "<p class=""normal"">Enclose please find the shipping document. Thanks.</p>" &
                                            "<p class=""normal"">Regards,<br>" &
                                              _selectedUser & "</p>" &
                                            "</body></html>"

                msg.Body = New MessageBody(BodyType.HTML, bodymessage)
                'msg.SendAndSaveCopy()


                msg.Save(WellKnownFolderName.Drafts)
                'MessageBox.Show("Sent.")
                myret = True
            Catch ex As Exception

                MessageBox.Show(ex.Message)

            End Try

        End Using
        Return myret
    End Function
    Private Sub addAttachmentList(ByVal filename As String, ByRef msg As EmailMessage)
        If Not attachmentdict.ContainsKey(filename) Then
            attachmentdict.Add(filename, filename)
            msg.Attachments.AddFileAttachment(filename)
        End If
    End Sub
    Private Sub addEmailList(ByVal recepient As String, ByVal recepientname As String)
        Dim myrecepients() As String = recepient.Split(";")
        For i = 0 To myrecepients.Count - 1
            If Not myrecepients(i).Length = 0 Then
                If Not emaildict.ContainsKey(Trim(myrecepients(i))) Then
                    emaildict.Add(Trim(myrecepients(i)), recepientname)
                    'add email
                    emaillist.Add(New emailData With {.email = Trim(myrecepients(i)), .displayname = recepientname})
                End If
            End If

        Next
    End Sub

    Sub doDraft()
        Dim qry = From p In BS.List
                  Where (p.row.item("Toggle") = True)
                  Group By housebill = p.row.item("housebill") Into mygroup = Group


        For Each d In qry
            'Debug.Print("hello")
            'For Each myData In d.mygroup
            'Debug.Print("inside")
            'Next

            'create email with files
            'd is group object with details information
            ProgressReport(1, "Create Draft : " & d.housebill & ". Please wait ....")
            If Not CreateDraft(d) Then
                ProgressReport(1, "Create Draft : " & d.housebill & ". Error found.")
                Exit Sub
            End If
        Next

        ProgressReport(1, "Create Draft : Done.")
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        LoadData()
    End Sub


    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.Validate()

        If Not myProcess.IsAlive Then
            myProcess = New Thread(AddressOf doDraft)
            myProcess.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
End Class