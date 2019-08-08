Imports Components.PublicClass
Imports System.Threading
Imports Components.SharedClass
Imports Microsoft.Exchange.WebServices.Data
Public Class FormEmailPreparationCP
    Dim myThreadDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myGenerateDelegate As New ThreadStart(AddressOf GenerateEmailDraft)
    Dim myGenerate As New System.Threading.Thread(myGenerateDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)




    Dim DS As DataSet
    Dim combobs As BindingSource
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim username As String
    Dim domain As String
    Dim password As String
    Dim mydata As String()
    Dim selectedUser As String
    Dim startdate As Date
    Dim enddate As Date
    Dim emaildict As Dictionary(Of String, String)
    Dim attachmentdict As Dictionary(Of String, String)
    Dim emaillist As List(Of emailData)
    Dim BS As BindingSource
    Dim overridedraft As Boolean

    Private Sub FormEmailPreparationCP_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If myGenerate.IsAlive Then
            If MessageBox.Show("Terminate the process?", "Terminate Process", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                myGenerate.Abort()
            Else
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub FormEmailPreparation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        mydata = Split(HelperClass1.UserId, "\")
        username = mydata(1)
        domain = mydata(0)
        TextBox1.Text = username
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

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myThread.Start()
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

    Sub GenerateEmailDraft()
        'Fill Dataset1 ->Detail document relations
        'Fill Dataset2 ->Distinct billoflading
        'With criteria avail billoflading,sebinvoice,packinglist
        ' not createddrafemail, selected SAO
        ProgressReport(1, "Finding Data,Please wait...")
        ProgressReport(6, "Marque")
        Dim myuser As String = String.Empty
        myuser = IIf(selectedUser = "", "", ",'" & selectedUser & "'")

        'Dim sqlstrComplete = "with bl as (" &
        '             " select distinct docemailname,docemaildtname from docemailhd dh" &
        '             " left join docemaildt dt on dt.docemailhdid = dh.docemailhdid" &
        '             " where(docemailtype = 0 And mycontains(dt.docemaildtname, dh.docemailname))" &
        '             " )," &
        '             " tb as (" &
        '             " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
        '             " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
        '             " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '             " )," &
        '             " consol as (" &
        '             " select count (tb.housebill) as mycount,tb.housebill from tb" &
        '             " group by tb.housebill" &
        '             " )" &
        '             " select tb.housebill,tb.billingdoc,tb.delivery,tb.reference,c.mycount,tx.billoflading,soldtoparty, docemaildtname,getfilename(tb.billingdoc::character varying)as invoice,getfilename(delivery::character varying) as packinglist from tb" &
        '             " left join consol c on c.housebill = tb.housebill" &
        '             " left join docemailtx tx on tx.billoflading = tb.housebill" &
        '             " left join bl on bl.docemailname = tb.housebill" &
        '             " left join billinghd bh on bh.billingdocument = tb.billingdoc" &
        '             " where not tb.housebill isnull and not tb.billingdoc isnull" &
        '             " and  tx.billoflading isnull" &
        '             " and ((not getfilename(tb.billingdoc::character varying) isnull) or (not getfilename(delivery::character varying) isnull) or (not bl.docemailname isnull));"
        Dim replacedraft As String = String.Empty
        If Not CheckBox1.Checked Then
            'If overridedraft Then
            replacedraft = " and dtx.draftcreateddate isnull"
        End If
        'Dim sqlstrComplete = "with bl as (" &
        '            " select distinct docemailname,docemaildtname,sender from docemailhd dh" &
        '            " left join docemaildt dt on dt.docemailhdid = dh.docemailhdid" &
        '            " where(docemailtype = 0 And mycontains(dt.docemaildtname, dh.docemailname))" &
        '            " )," &
        '            " tb as (" &
        '            " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference" &
        '            " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
        '            " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '            " )," &
        '            " consol as (" &
        '            " select count (tb.housebill) as mycount,tb.housebill from tb" &
        '            " group by tb.housebill" &
        '            " )" &
        '            " select tb.housebill,tb.billingdoc,tb.delivery,tb.reference,c.mycount,tx.billoflading,soldtoparty, docemaildtname,getfilename(tb.billingdoc::character varying)as invoice,getfilename(delivery::character varying) as packinglist,sender from tb" &
        '            " left join consol c on c.housebill = tb.housebill" &
        '            " left join docemailtx tx on tx.billoflading = tb.housebill" &
        '            " left join bl on bl.docemailname = tb.housebill" &
        '            " left join billinghd bh on bh.billingdocument = tb.billingdoc" &
        '            " where not tb.housebill isnull and not tb.billingdoc isnull" &
        '            replacedraft &
        '            " and ((not getfilename(tb.billingdoc::character varying) isnull) and (not getfilename(delivery::character varying) isnull) and (not bl.docemailname isnull));"

        'Dim sqlincomplete = "with bl as (" &
        '                    " select distinct docemailname,docemaildtname from docemailhd dh" &
        '                    " left join docemaildt dt on dt.docemailhdid = dh.docemailhdid" &
        '                    " where(docemailtype = 0 And mycontains(dt.docemaildtname, dh.docemailname))" &
        '                    " )," &
        '                    " tb as (" &
        '                    " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference " &
        '                    " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
        '                    " as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text)" &
        '                    " )" &
        '                    " select tb.housebill,tb.billingdoc,tb.delivery,tb.reference,tx.billoflading, getfilename(tb.billingdoc::character varying)as invoice,getfilename(delivery::character varying) as packinglist from tb" &
        '                    " left join docemailtx tx on tx.billoflading = tb.housebill" &
        '                    " left join bl on bl.docemailname = tb.housebill" &
        '                    " where " &
        '                    " not tb.billingdoc isnull" &
        '                    " and tx.billoflading isnull" &
        '                    " and ((getfilename(tb.billingdoc::character varying) isnull) or (getfilename(delivery::character varying) isnull) or (bl.docemailname isnull));"
        'Dim sqlstr = "select * from docemailtx limit 1;" &
        '             " select billoflading,draftcreateddate from docemailtx;" &
        '             " select dt.* from paramdt dt" &
        '             " left join paramhd hd on hd.paramhdid = dt.paramhdid" &
        '             " where hd.paramname = 'logbook'" &
        '             " order by dt.ivalue;"
        'Dim sqldeliverybrand = "select * from sp_getdeliverybrand(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") as tb(delivery bigint,brand character varying);"
        'Dim sqlcustomeremail = "select m.customercode,b.brandname::character varying,c.customername::character varying as displayname,name,email from marketemail m " &
        '                       " left join brand b on b.brandid = m.brandid  " &
        '                       " left join customer c on c.customercode = m.customercode " &
        '                       " where b.brandname isnull;" &
        '                       "select m.customercode,b.brandname::character varying,c.customername::character varying as displayname,name,email from marketemail m " &
        '                       " left join brand b on b.brandid = m.brandid  " &
        '                       " left join customer c on c.customercode = m.customercode " &
        '                       " where not b.brandname isnull;"

        Dim sqldata = "with tb as (" &
                    " select distinct tb.housebill ,tb.billingdoc ,tb.delivery ,tb.reference,tb.vendorcode,tb.vendorname::character varying" &
                    " from sp_getaccountingdata(" & DateFormatyyyyMMdd(startdate) & "::date," & DateFormatyyyyMMdd(enddate) & "::date " & myuser & ") " &
                    "  as tb(postingdate date,vendorcode bigint,vendorname character varying,reference text,accountingdoc bigint,miro bigint,pohd bigint,poitem integer,amount numeric,qty numeric,delivery bigint,item integer,deliverydate date,billingdoc bigint,housebill text,username text) ) " &
                    " select distinct false::boolean as toggle,tb.delivery::character varying,tb.reference,ph.biloflading as containerno,tb.housebill,c.customercode as shiptopartycode,c.customername::text as shiptoparty,tb.vendorcode,tb.vendorname,getfilenamecp(tb.housebill::character varying)as billoflading,tb.billingdoc::character varying ,dtx.draftcreateddate,ph.deliverydate from tb " &
                    " left join packinglisthd ph on ph.delivery = tb.delivery" &
                    " left join packinglistdt pld on pld.delivery = tb.delivery" &
                    " left join cxsebpodtl cx on cx.sebasiapono = pld.pohd and cx.polineno = pld.poitem" &
                    " left join customer c on c.customercode = cx.shiptoparty" &
                    " left join docemailcptx dtx on dtx.delivery = tb.delivery" &
                    " where not getfilenamecp(tb.housebill::character varying) isnull" &
                    replacedraft &
                    " order by delivery,reference,shiptoparty,containerno;"

        Dim sqlstr = " select dt.* from paramdt dt" &
                     " left join paramhd hd on hd.paramhdid = dt.paramhdid" &
                     " where hd.paramname = 'logbookcp'" &
                     " order by dt.ivalue;"
        Dim sqlemail = " select distinct shiptopartycode,vendorcode,name,email from marketemailcp where not  vendorcode isnull;" &
                       " select distinct shiptopartycode,name,email from marketemailcp where vendorcode isnull and not shiptopartycode isnull;"
        Dim draftdate = "select null::bigint as delivery,null::date as draftdate limit 0;"

        DS = New DataSet
        Dim mymessage As String = String.Empty
        'Dim distinctvalue As DataTable
        If DbAdapter1.TbgetDataSet(sqldata & sqlstr & sqlemail & draftdate, DS, mymessage) Then

            Dim view As DataView = New DataView(DS.Tables(0))

            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
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

            ProgressReport(5, "Continuous")
            ProgressReport(1, "Finding Data,Done.")
            MessageBox.Show(mymessage)
            Exit Sub
        End If

        If DS.Tables(0).Rows.Count = 0 Then

            ProgressReport(1, "Finding Data,Done.")
            ProgressReport(5, "Continuous")
            MessageBox.Show("Document not found.")
            Exit Sub
        End If

        'show message how many dataset2.record count, asking to continue create draft

        If MessageBox.Show(String.Format("Found {0} bill of lading(s) for {1}. Generate email draft?", DS.Tables(0).Rows.Count, IIf(selectedUser = "", "all SAO", selectedUser)), "Message", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            'if yes then create emaildraft one by one based on dataset2.rows

            Try
                ProgressReport(1, "Create email draft.")
                Dim i As Integer = 0

                'For Each dr As DataRow In DS.Tables(0).Rows
                For Each dr As DataRowView In BS.List
                    i += 1
                    ProgressReport(1, String.Format("Create email draft {0} of {1}", i, DS.Tables(0).Rows.Count))


                    If createemail(dr) Then
                        'Thread.Sleep(2500)                        
                        Dim newdr As DataRow = DS.Tables(4).NewRow
                        newdr.Item("delivery") = dr.Item("delivery")
                        newdr.Item("draftdate") = Date.Today
                        DS.Tables(4).Rows.Add(newdr)
                    Else
                        'Err.Raise("1")
                        ProgressReport(5, "Continuous")
                        If i = 1 Then
                            Exit Sub
                        End If
                    End If
                Next

                Dim ds2 As DataSet = DS.GetChanges
                If Not (IsNothing(ds2)) Then
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(DS, True, mymessage, ra, True)
                    If Not DbAdapter1.DocEmailDraftCPTx(Me, mye) Then
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If

                    'Save Docemailtx
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Debug.Print("hello")
            End Try



        Else

        End If
        ProgressReport(1, "Done.")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myGenerate.IsAlive Then
            If TextBox2.Text = "" Or TextBox1.Text = "" Then
                MessageBox.Show("Userid and Password cannot be blank!")
                Exit Sub
            End If
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            username = TextBox1.Text
            password = TextBox2.Text
            selectedUser = ComboBox1.Text
            startdate = DateTimePicker1.Value
            enddate = DateTimePicker2.Value

            myGenerate = New Thread(AddressOf GenerateEmailDraft)
            myGenerate.Start()
        Else
            MessageBox.Show("Please wait until current process finished!")
        End If
    End Sub
    Private Function createemail(ByVal dr As DataRowView) As Boolean
        Dim myret As Boolean = False
        Dim url As String = DS.Tables(1).Rows(1).Item("cvalue")  '"https://mail-eu.seb.com/ews/exchange.asmx"
        Dim service As ExchangeService
        'Using myservice As New ClassEWS(url, _username, _password, "as", False)
        Using myservice As New ClassEWS(url, username, password, domain, False)
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



                Dim mydr As DataRowView = DirectCast(dr, DataRowView)
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

                msg.Subject = String.Format("{0}/{1}/{2}/{3}/{4:ddMMMyyyy} ", mydr.Row.Item("shiptoparty"), mydr.Row.Item("containerno"), mydr.Row.Item("vendorname"), mydr.Row.Item("billoflading"), mydr.Row.Item("deliverydate")).ToUpper
                Dim bodymessage As String = "<!DOCTYPE html><html><head><meta charset=utf-8 /><style>p.normal{font-size:11.0pt;font-family:""Calibri"",""sans-serif"";}</style></head><body>" &
                                            "<p class=""normal"">Dear " & mynamelist & ",</p>" &
                                            "<p class=""normal"">Enclose please find the shipping document. <br>Thank you very much.</p>" &
                                            "<p class=""normal"">Best regards,<br>" &
                                              selectedUser & "</p>" &
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
    Private Function createemail1(ByVal dr As DataRow) As Boolean
        Dim myret As Boolean = False
        Dim url As String = DS.Tables(4).Rows(1).Item("cvalue")  '"https://mail-eu.seb.com/ews/exchange.asmx"
        Dim service As ExchangeService
        Using myservice As New ClassEWS(url, username, password, domain, False)
            'service = myservice.CreateConnectionAutoDiscover()
            service = myservice.CreateConnection()
            Try
                'find ds.tables(0)
                Dim myselect() As DataRow = DS.Tables(0).Select("housebill = '" & dr.Item(0) & "'")

                If myselect(0).Item(4) <> 1 Then
                    'Consolidation-> find in incomplete document. if found ->return false
                    Dim recordfound() As DataRow = DS.Tables(1).Select("housebill = '" & dr.Item(0) & "'")
                    If recordfound.Length <> 0 Then
                        Return myret
                    End If

                End If

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

                For Each mydr In myselect

                    'if Soldtoparty US -> get brand from table delivery brand
                    '   For each brand -> get email from customer email
                    'add datatable based on comma delimited
                    'else
                    '   Find mydr.item("soldtoparty") name and email
                    '   add datatable name and email
                    'endif
                    Dim myresult As DataRow

                    mycc = mydr.Item("sender")

                    If mydr.Item("soldtoparty") = 99008400 Then 'US Market
                        Dim drow() As DataRow = DS.Tables(5).Select("delivery = '" & mydr.Item("delivery") & "'")
                        For Each row In drow
                            'Find 
                            Dim mykey(1) As Object
                            mykey(0) = 99008400
                            mykey(1) = row.Item("brand")
                            myresult = DS.Tables(7).Rows.Find(mykey)
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
                        mykey(0) = mydr.Item("soldtoparty")
                        myresult = DS.Tables(6).Rows.Find(mykey)
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

                    Dim myfile As String = DS.Tables(4).Rows(4).Item("cvalue")
                    Dim forwarder As String = myfile & "\Forwarder\" & dr.Item(0) & "\" & mydr.Item("docemaildtname")
                    Dim invoice As String = myfile & "\invoice\" & mydr.Item("invoice")
                    Dim packinglist As String = myfile & "\packinglist\" & mydr.Item("packinglist")
                    'mycc = mydr.Item("sender")

                    'msg.Attachments.AddFileAttachment(forwarder)
                    'msg.Attachments.AddFileAttachment(invoice)
                    'msg.Attachments.AddFileAttachment(packinglist)
                    addAttachmentList(forwarder, msg)
                    addAttachmentList(invoice, msg)
                    addAttachmentList(packinglist, msg)
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

                msg.Subject = String.Format("SEB Asia Document - {0} - {1} ", marketname, dr.Item(0))
                Dim bodymessage As String = "<!DOCTYPE html><html><head><meta charset=utf-8 /><style>p.normal{font-size:11.0pt;font-family:""Calibri"",""sans-serif"";}</style></head><body>" &
                                            "<p class=""normal"">Dear " & mynamelist & ",</p>" &
                                            "<p class=""normal"">Enclose please find the shipping document. Thanks.</p>" &
                                            "<p class=""normal"">Regards,<br>" &
                                              selectedUser & "</p>" &
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
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        username = TextBox1.Text
        password = TextBox2.Text
        selectedUser = ComboBox1.Text
        startdate = DateTimePicker1.Value
        enddate = DateTimePicker2.Value
        Dim myShowAllItemForm As New FormShowAllItemsCP(username, password, selectedUser, startdate, enddate)
        myShowAllItemForm.Show()


    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Or TextBox1.Text = "" Then
            MessageBox.Show("Userid and Password cannot be blank!")
            Exit Sub
        End If
        username = TextBox1.Text
        password = TextBox2.Text
        selectedUser = ComboBox1.Text
        startdate = DateTimePicker1.Value
        enddate = DateTimePicker2.Value

        Dim mySendIndividualEmail As New FormSendIndividualEmailCP(username, password, selectedUser, startdate, enddate)
        mySendIndividualEmail.Show()

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        overridedraft = CheckBox1.Checked
    End Sub
End Class