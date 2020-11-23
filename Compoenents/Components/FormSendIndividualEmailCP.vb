Imports Components.PublicClass
Imports System.Threading
Imports Components.SharedClass
Imports Microsoft.Exchange.WebServices.Data
Public Class FormSendIndividualEmailCP
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
        Dim sqldata = "with tb as (" &
                            " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference,tb.vendorcode,tb.vendorname::character varying" &
                            " from sp_getaccountingdata(" & DateFormatyyyyMMdd(_startdate) & "::date," & DateFormatyyyyMMdd(_enddate) & "::date " & myuser & ") " &
                            "  as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text) ) " &
                            " select distinct false::boolean as toggle,tb.delivery::character varying,tb.reference,ph.biloflading as containerno,tb.housebill,c.customercode as shiptopartycode,c.customername::text as shiptoparty,tb.vendorcode,tb.vendorname,getfilenamecp(tb.housebill::character varying)as billoflading,tb.billingdoc::character varying ,dtx.draftcreateddate,ph.deliverydate from tb " &
                            " left join packinglisthd ph on ph.delivery = tb.delivery" &
                            " left join packinglistdt pld on pld.delivery = tb.delivery" &
                            " left join cxsebpodtl cx on cx.sebasiapono = pld.pohd and cx.polineno = pld.poitem" &
                            " left join customer c on c.customercode = cx.shiptoparty" &
                            " left join docemailcptx dtx on dtx.delivery = tb.delivery" &
                            " order by delivery,reference,shiptoparty,containerno;"

        Dim sqlstr = " select dt.* from paramdt dt" &
                     " left join paramhd hd on hd.paramhdid = dt.paramhdid" &
                     " where hd.paramname = 'logbookcp'" &
                     " order by dt.ivalue;"
        Dim sqlemail = " select distinct shiptopartycode,vendorcode,name,email from marketemailcp where not  vendorcode isnull;" &
                       " select distinct shiptopartycode,name,email from marketemailcp where vendorcode isnull;"
        Dim draftdate = "select null::bigint as delivery,null::date as draftdate limit 0;"

        DS = New DataSet
        Dim mymessage As String = String.Empty

        'If DbAdapter1.TbgetDataSet(sqlincomplete & sqlstr & sqldeliverybrand & sqlcustomeremail, DS, mymessage) Then
        Try
            If DbAdapter1.TbgetDataSet(sqldata & sqlstr & sqlemail & draftdate, DS, mymessage) Then

                Dim view As DataView = New DataView(DS.Tables(0))

                ProgressReport(4, "InitDataSource")


                Dim pkey2(1) As DataColumn
                pkey2(0) = DS.Tables(2).Columns(0)
                pkey2(1) = DS.Tables(2).Columns(1)
                DS.Tables(2).PrimaryKey = pkey2
                DS.Tables(2).TableName = "ShipToPartyVendor"

                Dim pkey3(0) As DataColumn
                pkey3(0) = DS.Tables(3).Columns(0)
                DS.Tables(3).PrimaryKey = pkey3
                DS.Tables(3).TableName = "ShipToParty"

                DS.Tables(4).TableName = "DraftDate"

            Else
                MessageBox.Show(mymessage)
                ProgressReport(5, "Continuous")
                Exit Sub
            End If

            ProgressReport(5, "Continuous")
            ProgressReport(1, "Loading Data, Done.")
            If Not IsNothing(BS) Then
                ProgressReport(1, String.Format("Loading Data, Done. Record(s) Count: {0}", BS.Count))
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        

    End Sub

    Private Sub FormSendIndividualEmail_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If myProcess.IsAlive Then
            'MessageBox.Show("Please wait until the current process is finished.")
            If MessageBox.Show("Current Process still processing. Do you want to terminate this process?", "Active Process", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                myProcess.Abort()
            Else
                e.Cancel = True
            End If

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

                Case 5
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



    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
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
            Try

            
                Dim myfields() = {"delivery", "reference", "containerno", "housebill", "shiptoparty", "vendorname", "billoflading", "billingdoc"}
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

                If Not IsNothing(BS) Then
                    ProgressReport(1, String.Format("Loading Data, Done. Record(s) Count: {0}", BS.Count))
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Not myThread.IsAlive Then
            BS.Filter = ""
            If Not IsNothing(BS) Then
                ProgressReport(1, String.Format("Loading Data, Done. Record(s) Count: {0}", BS.Count))
            End If
        End If
    End Sub
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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



                Dim mydr As DataRowView = DirectCast(d, DataRowView)
                Dim myresult As DataRow


                '***************
                'Find Recepient
                '1. ShiptoParty + Vendorcode if not avail then
                '2. ShiptoParty Only
                'in table MarketEmailCP

                'Locate Attachment based on housebill folder , get all files inside the folder

                If Not IsDBNull(mydr.Row.Item("shiptopartycode")) Then
                    Dim mykey(1) As Object
                    mykey(0) = mydr.Row.Item("shiptopartycode")
                    mykey(1) = mydr.Row.Item("vendorcode")
                    myresult = DS.Tables(2).Rows.Find(mykey)
                    recepient = ""
                    recepientname = ""
                    If Not IsNothing(myresult) Then
                        If Not IsDBNull(myresult.Item("email")) Then
                            recepient = myresult.Item("email")
                        End If
                        If Not IsDBNull(myresult.Item("name")) Then
                            recepientname = myresult.Item("name")
                        End If

                        addEmailList(recepient, recepientname)
                    Else
                        Dim mykey1(0) As Object
                        mykey1(0) = mydr.Row.Item("shiptopartycode")
                        myresult = DS.Tables(3).Rows.Find(mykey1)
                        If Not IsNothing(myresult) Then
                            If Not IsDBNull(myresult.Item("email")) Then
                                recepient = myresult.Item("email")
                            End If
                            If Not IsDBNull(myresult.Item("name")) Then
                                recepientname = myresult.Item("name")
                            End If
                            addEmailList(recepient, recepientname)
                        End If
                    End If
                End If



                Dim myfile As String = DS.Tables(1).Rows(4).Item("cvalue")
                Dim MyFolder As String = myfile & "\" & mydr.Row.Item("billoflading")
                'Dim invoice As String = myfile & "\invoice\" & mydr.Row.Item("invoice")
                'Dim packinglist As String = myfile & "\packinglist\" & mydr.Row.Item("packinglist")
                'mycc = mydr.Item("sender")

                'msg.Attachments.AddFileAttachment(forwarder)
                'msg.Attachments.AddFileAttachment(invoice)
                'msg.Attachments.AddFileAttachment(packinglist)
                If Not IsDBNull(mydr.Row.Item("billoflading")) Then
                    addAttachmentList(MyFolder, msg)
                End If


                'for each data table 
                Dim recepientnamelist As String = String.Empty
                For Each mydata As emailData In emaillist
                    'If Not IsDBNull(mydata.email) Then
                    If Not mydata.email = "" Then
                        msg.ToRecipients.Add(New EmailAddress(mydata.email))
                    Else
                        'If Not mycc = "" Then
                        '    msg.ToRecipients.Add(New EmailAddress(mycc))
                        'End If
                    End If
                    recepientnamelist = recepientnamelist & IIf(recepientnamelist = "", "", ",") & mydata.displayname
                Next

                Dim mynamelist As String = IIf(recepientnamelist.Split(",").Count > 1, "All", recepientnamelist)

                'next
                'If Not mycc = "" Then
                '    msg.CcRecipients.Add(New EmailAddress(mycc))
                'End If

                msg.Subject = String.Format("{0}/SHIP DOC/{1}/{2}/{4:ddMMMyyyy}/{3}", mydr.Row.Item("shiptoparty"), mydr.Row.Item("containerno"), mydr.Row.Item("vendorname"), mydr.Row.Item("billoflading"), mydr.Row.Item("deliverydate")).ToUpper
                Dim bodymessage As String = "<!DOCTYPE html><html><head><meta charset=utf-8 /><style>p.normal{font-size:11.0pt;font-family:""Calibri"",""sans-serif"";}</style></head><body>" &
                                            "<p class=""normal"">Dear " & mynamelist & ",</p>" &
                                            "<p class=""normal"">Enclose please find the shipping document. <br>Thank you very much.</p>" &
                                            "<p class=""normal"">Best regards,<br>" &
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
        Dim di As New System.IO.DirectoryInfo(filename)
        Try
            For Each f As System.IO.FileInfo In di.GetFiles
                msg.Attachments.AddFileAttachment(f.FullName)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        'If Not attachmentdict.ContainsKey(filename) Then
        '    attachmentdict.Add(filename, filename)
        '    msg.Attachments.AddFileAttachment(filename)
        'End If
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
        'Dim qry = From p In BS.List
        '          Where (p.row.item("Toggle") = True)
        '          Group By housebill = p.row.item("housebill") Into mygroup = Group
        ProgressReport(6, "Marquee")
        Dim qry = From p In BS.List
                  Where (p.row.item("Toggle") = True)

        For Each d In qry
            'Debug.Print("hello")
            'For Each myData In d.mygroup
            'Debug.Print("inside")
            'Next

            'create email with files
            'd is group object with details information
            ProgressReport(1, "Create Draft : " & d.row.item("delivery") & ". Please wait ....")
            If CreateDraft(d) Then
                d.row.item("draftcreateddate") = Today.Date
                DataGridView1.Invalidate()
                Dim dr = DS.Tables(4).NewRow()
                dr.Item("delivery") = d.row.item("delivery")
                dr.Item("draftdate") = Today.Date
                DS.Tables(4).Rows.Add(dr)
            Else
                ProgressReport(5, "Continuous")
                ProgressReport(1, "Create Draft : Error Found.")
                Exit Sub
            End If

        Next
        'update draftdate
        Dim mymessage As String = String.Empty
        Dim ra As Integer
        Dim mye As New ContentBaseEventArgs(DS, True, mymessage, ra, True)
        DbAdapter1.DocEmailDraftCPTx(Me, mye)
        ProgressReport(5, "Continuous")
        ProgressReport(1, "Create Draft : Done.")
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        LoadData()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        For Each drv As DataRowView In BS.List
            drv.Item("toggle") = CheckBox1.Checked
        Next
        DataGridView1.Invalidate()
    End Sub
End Class