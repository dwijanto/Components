﻿Imports Microsoft.Exchange.WebServices.Data
Imports System.IO
Imports System.Threading
Imports Components.PublicClass
Public Class FormGetSelectedEmailCP

    Dim service As ExchangeService

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim AutoTask As Boolean = True
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim url As String = "https://mail-eu.seb.com/ews/exchange.asmx"
    Dim username As String = "sebdoccomp" '"sebshipdoc"
    Dim password As String = "honspH@ndfree01" '"honscH@ndfrgz01"
    Dim mybasefolder As String = "c:\temp\documents"
    Dim bs As BindingSource
    Dim bs2 As BindingSource
    Dim selecteddate As DateTime
    Dim selecteddate2 As DateTime
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'selecteddate = DateTimePicker1.Value.AddMinutes(-1)
            'selecteddate = selecteddate.Date & " " & selecteddate.Hour & ":" & selecteddate.Minute & ":00"
            'selecteddate2 = DateTimePicker1.Value.AddMinutes(1)
            'selecteddate2 = selecteddate2.Date & " " & selecteddate2.Hour & ":" & selecteddate2.Minute & ":00"
            selecteddate = DateTimePicker1.Value.AddMinutes(-1)
            selecteddate = selecteddate.Date & " " & selecteddate.Hour & ":" & selecteddate.Minute & ":00"
            selecteddate2 = DateTimePicker2.Value.AddMinutes(1)
            selecteddate2 = selecteddate2.Date & " " & selecteddate2.Hour & ":" & selecteddate2.Minute & ":00"
            'Get file
            AutoTask = False
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoWork()

        If DbAdapter1.checkLockFile(Application.StartupPath & "\log\GetAttachment.lck") Then
            If Not AutoTask Then
                If Not MessageBox.Show("Process is running in different computer! Force to continue? ", "Question", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
        ProgressReport(6, "Marque")
        If GetFolder(0) Then
            ProgressReport(2, "Done")
            ProgressReport(1, "")
        End If
        ProgressReport(5, "Continuous")
        File.Delete(Application.StartupPath & "\log\GetAttachment.lck")
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
                    ' Me.Label4.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select

        End If

    End Sub

    Private Function GetFolder(ByVal offset As Integer) As Boolean
        'ProgressReport(1, "Get Folder")
        Dim mydoctype As Integer

        Dim savingfolder As String = String.Empty
        Dim myfilenamelog As String = String.Empty
        Dim myreturn As Boolean = False
        Dim ds As New DataSet

        Dim sqlstr = "select dt.* from paramdt dt" &
             " left join paramhd hd on hd.paramhdid = dt.paramhdid" &
             " where hd.paramname= 'logbookcp'" &
             " order by dt.ivalue;" &
            " select * from docemailhdcp;" &
            " select * from docemaildtcp;" &
            " select docemailname,docemailhdid from docemailhdcp;" &
            " select distinct docemaildtname,docemailhdid from docemaildtcp;"

        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            ProgressReport(1, mymessage)
            Logg(mymessage)
            Return myreturn
        End If
        Dim mylastdate As DateTime = ds.Tables(0).Rows(0).Item("ts")
        url = ds.Tables(0).Rows(1).Item("cvalue")
        username = ds.Tables(0).Rows(2).Item("cvalue")
        password = ds.Tables(0).Rows(3).Item("cvalue")
        mybasefolder = ds.Tables(0).Rows(4).Item("cvalue")
        'Dim mylastdateinvoice As DateTime = ds.Tables(0).Rows(5).Item("ts")
        'Dim mylastdatepackinglist As DateTime = ds.Tables(0).Rows(6).Item("ts")

        Try
            'ProgressReport(1, "After Get DataSet")
            'Header and Detail
            ds.Tables(1).TableName = "DocHeader"
            ds.Tables(1).CaseSensitive = True
            ds.Tables(2).TableName = "DocDtl"
            ds.Tables(2).CaseSensitive = True

            Dim idx1(0) As DataColumn
            idx1(0) = ds.Tables(1).Columns(0)
            ds.Tables(1).PrimaryKey = idx1
            ds.Tables(1).Columns(0).AutoIncrement = True
            ds.Tables(1).Columns(0).AutoIncrementSeed = -1
            ds.Tables(1).Columns(0).AutoIncrementStep = -1
            ds.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn
            idx2(0) = ds.Tables(2).Columns(0)
            ds.Tables(2).PrimaryKey = idx2
            ds.Tables(2).Columns(0).AutoIncrement = True
            ds.Tables(2).Columns(0).AutoIncrementSeed = -1
            ds.Tables(2).Columns(0).AutoIncrementStep = -1
            ds.Tables(2).PrimaryKey = idx2


            Dim rel As DataRelation
            Dim hcol As DataColumn
            Dim dcol As DataColumn

            hcol = ds.Tables(1).Columns(0) 'docemailhdid in table header
            dcol = ds.Tables(2).Columns(1) 'docemailhdid in table dtl
            rel = New DataRelation("hdrel", hcol, dcol)
            ds.Relations.Add(rel)

            bs = New BindingSource
            bs2 = New BindingSource
            bs.DataSource = ds.Tables(1)
            bs2.DataSource = bs
            bs2.DataMember = "hdrel"

            'Find Using Index For Header And Detail
            ds.Tables(3).TableName = "FindHD"
            Dim idx3(0) As DataColumn
            idx3(0) = ds.Tables(3).Columns(0) 'docemailname
            ds.Tables(3).PrimaryKey = idx3

            ds.Tables(4).TableName = "FindDT"
            Dim idx4(1) As DataColumn
            idx4(0) = ds.Tables(4).Columns(0) 'docemaildtname
            idx4(1) = ds.Tables(4).Columns(1) 'doceamilhdid
            ds.Tables(4).CaseSensitive = True
            ds.Tables(4).PrimaryKey = idx4


        Catch ex As Exception
            Logg(ex.Message)
            ProgressReport(1, ex.Message)
            Return myreturn
        End Try

        Dim totalview As Integer = Integer.MaxValue
        'totalview = IIf(IsNumeric(TextBox7.Text), CInt(TextBox7.Text), 0)
        'ProgressReport(1, "Using Service")
        If Not AutoTask Then
            ProgressReport(1, "Search Document. Please wait... ")
        End If
        Using myservice As New ClassEWS(url, username, password, "as", False)
            service = myservice.CreateConnection()
            Dim searchFilterCollection As New List(Of SearchFilter)

            'searchFilterCollection.Add(New SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, DateTime.Parse(mylastdate.ToString)))
            'Dim searchFilter As SearchFilter = New SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray)

            searchFilterCollection.Add(New SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.Parse(selecteddate.ToString)))
            searchFilterCollection.Add(New SearchFilter.IsLessThan(ItemSchema.DateTimeReceived, DateTime.Parse(selecteddate2.ToString)))
            
            Dim searchfilter As SearchFilter = New SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray)




            Dim view As New ItemView(totalview)
            view.PropertySet = New PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived)
            view.Traversal = FolderTraversal.Shallow


            'Dim view As FolderView = New FolderView(totalview)
            'view.PropertySet = New PropertySet(BasePropertySet.IdOnly)
            'view.PropertySet.Add(FolderSchema.DisplayName)
            'view.Offset = offset
            ''MessageBox.Show(view.Offset)

            'Dim searchFilter As SearchFilter = New SearchFilter.IsGreaterThan(FolderSchema.TotalCount, 0)
            'view.Traversal = FolderTraversal.Deep
            Try
                'Dim results As FindFoldersResults = service.FindFolders(WellKnownFolderName.Root, searchFilter, view)
                'Dim results As FindFoldersResults = service.FindFolders(WellKnownFolderName.Inbox, searchFilter, view)


                


                'Dim results As FindFoldersResults = service.FindFolders(WellKnownFolderName.Inbox, New FolderView(Integer.MaxValue) With {.Traversal = FolderTraversal.Deep})
                'Dim results As FindItemsResults(Of Item) = service.FindItems(WellKnownFolderName.Inbox, searchfilter, view)

                Dim userMailbox = New Mailbox("sebdoccomp@groupeseb.com")
                Dim folderId = New FolderId(WellKnownFolderName.Inbox, userMailbox)
                Dim results As FindItemsResults(Of Item) = service.FindItems(folderId, searchfilter, view)

                For Each Item As Item In results.Items
                    If TypeOf Item Is EmailMessage Then
                        'Debug.Print("Email Message: " & TryCast(Item, EmailMessage).Subject)
                        Dim myarray = Item.Subject.Split("/")

                        'If ds.Tables(0).Rows(0).Item("ts") < Item.DateTimeReceived Then
                        ' ds.Tables(0).Rows(0).Item("ts") = Item.DateTimeReceived
                        'End If
                        'Update parameter emaillastreceived for forwarder,INVOICE,PACKING LIST
                        'If Folder.DisplayName.Contains("Forwarder") Then
                        '    If ds.Tables(0).Rows(0).Item("ts") < Item.DateTimeReceived Then
                        '        ds.Tables(0).Rows(0).Item("ts") = Item.DateTimeReceived
                        '    End If
                        'ElseIf Folder.DisplayName.Contains("INVOICE") Then
                        '    If ds.Tables(0).Rows(5).Item("ts") < Item.DateTimeReceived Then
                        '        ds.Tables(0).Rows(5).Item("ts") = Item.DateTimeReceived
                        '    End If
                        'ElseIf Folder.DisplayName.Contains("PACKING LIST") Then
                        '    If ds.Tables(0).Rows(6).Item("ts") < Item.DateTimeReceived Then
                        '        ds.Tables(0).Rows(6).Item("ts") = Item.DateTimeReceived
                        '    End If
                        'End If

                        Dim myitems As List(Of Item) = New List(Of Item)
                        myitems.Add(Item)
                        service.LoadPropertiesForItems(myitems, PropertySet.FirstClassProperties)
                        Dim message As EmailMessage = EmailMessage.Bind(service, Item.Id, New PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Attachments))
                        'Debug.Print(message.From.Name & " " & message.From.Address)



                        If Item.HasAttachments Then
                            'save to db
                            'check header
                            Dim pkey1(0) As Object
                            'Replace any character contains ' (singlequote)
                            Dim mydocemailname = DbAdapter1.validfilename(myarray(myarray.Count - 1).Trim).Replace("'", "''")
                            If mydocemailname.Length = 0 Then
                                mydocemailname = "-BLANK-"
                            End If

                            'If Not AutoTask Then
                            '    ProgressReport(1, "Found Document: " & mydocemailname)
                            'End If

                            pkey1(0) = mydocemailname
                            Dim result As DataRow = ds.Tables(3).Rows.Find(pkey1)
                            Dim myid As Long
                            If IsNothing(result) Then
                                'create new record
                                Dim dr As DataRow = ds.Tables(1).NewRow
                                dr.Item("docemailname") = mydocemailname
                                dr.Item("docemailtype") = mydoctype
                                dr.Item("sender") = message.From.Address
                                dr.Item("sendername") = message.From.Name
                                dr.Item("receiveddate") = Item.DateTimeReceived
                                dr.Item("foldername") = "Inbox"
                                myid = dr.Item("docemailhdid")
                                ds.Tables(1).Rows.Add(dr)
                                Dim mydr As DataRow = ds.Tables(3).NewRow
                                mydr.Item(0) = mydocemailname
                                mydr.Item(1) = myid
                                ds.Tables(3).Rows.Add(mydr)


                            Else
                                myid = result.Item(1)
                                Dim pkey11(0) As Object
                                pkey11(0) = myid
                                Dim myresult As DataRow = ds.Tables(1).Rows.Find(pkey11)

                                myresult.Item("receiveddate") = Item.DateTimeReceived
                                myresult.Item("sender") = message.From.Address
                                myresult.Item("sendername") = message.From.Name
                                myresult.Item("foldername") = "Inbox"
                            End If

                            'Dim savingfolder As String = myfolder
                            savingfolder = mybasefolder
                            'If Folder.DisplayName.Contains("Forwarder") Then
                            savingfolder = savingfolder & "\" & mydocemailname 'DbAdapter1.validfilename(myarray(myarray.Count - 1).Trim)
                            If Not Directory.Exists(savingfolder) Then
                                Directory.CreateDirectory(savingfolder)
                            End If
                            'End If
                            For Each Attachment As Attachment In Item.Attachments


                                If TypeOf Attachment Is FileAttachment Then
                                    Dim fileattachment As FileAttachment = DirectCast(Attachment, FileAttachment)
                                    'fileattachment.Load() 'this one saving using original filename

                                    'save to db
                                    'check detail
                                    Dim pkey2(1) As Object
                                    pkey2(0) = fileattachment.Name
                                    pkey2(1) = myid
                                    result = ds.Tables(4).Rows.Find(pkey2)

                                    If IsNothing(result) Then
                                        'create new record
                                        Dim dr As DataRow = ds.Tables(2).NewRow
                                        dr.Item("docemailhdid") = myid
                                        dr.Item("docemaildtname") = fileattachment.Name
                                        ds.Tables(2).Rows.Add(dr)

                                        Dim mydr As DataRow = ds.Tables(4).NewRow
                                        mydr.Item(0) = fileattachment.Name
                                        mydr.Item(1) = myid
                                        ds.Tables(4).Rows.Add(mydr)
                                    End If

                                    If Not AutoTask Then
                                        ProgressReport(1, "Folder : " & mydocemailname & " , Attachment name: " & fileattachment.Name)
                                    End If
                                    'Debug.WriteLine("Attachment name: " & fileattachment.Name)
                                    'fileattachment.Load("c:\\temp\\" + fileattachment.Name)
                                    'Using thestream As FileStream = New FileStream("c:\\temp\\stream_" + fileattachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                                    myfilenamelog = savingfolder + "\" + fileattachment.Name
                                    Using thestream As FileStream = New FileStream(savingfolder + "\" + fileattachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                                        fileattachment.Load(thestream)
                                        thestream.Close()
                                        thestream.Dispose()
                                    End Using
                                End If
                            Next
                        End If

                    ElseIf TypeOf Item Is MeetingRequest Then
                        'Debug.Print("Metting Request: " & TryCast(Item, MeetingRequest).Subject)
                    Else

                    End If

                Next
                
            Catch ex As Exception
                Logg(ex.Message & " " & savingfolder & " :: " & myfilenamelog)
                ProgressReport(1, ex.Message)
                Return myreturn  'do not save the latest update
            End Try
            'No need to update latest date
            Dim ds2 As DataSet = ds.GetChanges

            If Not IsNothing(ds2) Then
                Dim ra As Integer

                'update table Header detail
                'If Not DbAdapter1.DocEmailTx() Then

                'End If
                mymessage = String.Empty

                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.DocEmailTx(Me, mye) Then
                    ProgressReport(2, "Error" & "::" & mye.message)
                    Logg(mye.message)
                    Return False
                End If
                
            End If
            Return True

        End Using
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        If AutoTask Then
            'HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter
        End If
    End Sub
    Public Sub New(ByVal AutoTask As Boolean)

        ' This call is required by the designer.
        InitializeComponent()
        Me.AutoTask = AutoTask
        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub Logg(ByVal mymessage As String)
        If AutoTask Then
            Logger.log(mymessage)
        End If
    End Sub

    Private Function getsqlstr(ByVal mydate As Date, ByVal paramname As String) As String
        Dim myvaliddate = "'" & mydate.Year & "-" & mydate.Month & "-" & mydate.Day & " " & mydate.Hour & ":" & mydate.Minute & ":" & mydate.Second & "'"
        Dim sqlstr = "update paramdt set ts = " & myvaliddate & " where paramdt.paramname = '" & paramname & "';"
        Return sqlstr
    End Function

    Private Sub FormGetEmailFromExServer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If myThread.IsAlive Then
            MessageBox.Show("Please wait until the current process is finished.")
            e.Cancel = True
        End If
    End Sub

    Private Sub FormGetEmailFromExServer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class