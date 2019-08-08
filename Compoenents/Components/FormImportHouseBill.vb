Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Xml
Imports System.IO

Public Class FormImportHouseBill

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    'Dim startdate As Date
    'Dim enddate As Date

    'Dim miroSeq As Long
    'Dim podtlseq As Long
    'Dim cmmfpriceseq As Long
    'Dim cmmfvendorpriceseq As Long

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            'startdate = getfirstdate(DateTimePicker1.Value)
            'enddate = getlastdate(DateTimePicker2.Value)
            'startdate = DateTimePicker1.Value
            'enddate = DateTimePicker2.Value
            'appendfile = RadioButton1.Checked

            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
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


    Sub DoWork()
        Dim sw As New Stopwatch
        'Dim AccountingHDSB As New System.Text.StringBuilder
        Dim HouseBillSB As New System.Text.StringBuilder


        'Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        'Dim billingdate As Date
        Dim DS As New DataSet
        sw.Start()
        'Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
        '    With objTFParser
        '        .TextFieldType = FileIO.FieldType.Delimited
        '        .SetDelimiters(Chr(9))
        '        .HasFieldsEnclosedInQuotes = True
        '        Dim count As Long = 0

        'Delete Existing Record
        ProgressReport(2, "Delete ..")
        ProgressReport(6, "Marque")

        'sqlstr = "delete from billingdtl bd where bd.billingdocument in (select billingdocument from billinghd bh where" &
        '         " bh.billingdate >= " & DateFormatyyyyMMdd(startdate) & " and bh.billingdate <= " & DateFormatyyyyMMdd(enddate) & ");" &
        '         " select setval('billingdtl_billingdtlid_seq',(select billingdtlid from billingdtl order by billingdtlid desc limit 1) + 1,false);" &
        '         " delete from billinghd bh where bh.billingdate >= " & DateFormatyyyyMMdd(startdate) & " and bh.billingdate <= " & DateFormatyyyyMMdd(enddate) & ";"

        'sqlstr = " delete from housebill h where h.housebilldate >= " & DateFormatyyyyMMdd(startdate) & " and h.housebilldate <= " & DateFormatyyyyMMdd(enddate) & ";" &
        '         " select setval('housebill_housebillid_seq',(select housebillid from housebill order by housebillid desc limit 1) + 1,false);"


        Dim mymessage As String = String.Empty
        If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        'Fill Header
        ProgressReport(2, "Initialize Table..")
        sqlstr = "select containerno,po,partno from housebill;"

        mymessage = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        DS.Tables(0).TableName = "Housebill"
        Dim idx0(2) As DataColumn
        idx0(0) = DS.Tables(0).Columns(0)
        idx0(1) = DS.Tables(0).Columns(1)
        idx0(2) = DS.Tables(0).Columns(2)
        DS.Tables(0).PrimaryKey = idx0


        ProgressReport(2, "Read Text File...")
        'Do Until .EndOfData
        '    myrecord = .ReadFields
        '    If count > 0 Then
        '        mylist.Add(myrecord)
        '    End If
        '    count += 1
        'Loop
        ProgressReport(2, "Build Record...")
        ProgressReport(5, "Continuous")
        'Dim result As DataRow
        'For i = 0 To mylist.Count - 1
        '    'find the record in existing table.
        '    ProgressReport(7, i + 1 & "," & mylist.Count)
        '    myrecord = mylist(i)
        '    If i >= 0 Then
        '        billingdate = DbAdapter1.dateformatdotdate(myrecord(2))
        '        'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
        '        If billingdate >= startdate.Date AndAlso billingdate <= enddate.Date Then

        '            Dim widthdict As New Dictionary(Of Integer, Integer)

        Dim xmldoc As New XmlDocument
        Dim xmldoc2 As New XmlDocument
        Dim xmlnodelist As XmlNodeList
        Dim i As Long
        Dim str As New StringBuilder
        Try
            Using fs As New FileStream(OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read)
                xmldoc.Load(fs)
            End Using

            'xmlnodelist = xmldoc.GetElementsByTagName("ss:Column")
            'For Each node As XmlNode In xmlnodelist
            '    Dim attribute = node.Attributes("ss:Width")
            '    If attribute IsNot Nothing Then
            '        Debug.Print(attribute.Value)
            '    End If
            'Next

            'Dim mykey As Integer = 0
            'Using reader As XmlReader = XmlReader.Create(OpenFileDialog1.FileName)
            '    While reader.Read()
            '        If reader.IsStartElement() Then
            '            If reader.Name = "ss:Column" Then
            '                Dim attribute As String = reader("ss:Width")
            '                If attribute IsNot Nothing Then
            '                    widthdict.Add(mykey, CInt(attribute))
            '                    mykey = mykey + 1
            '                End If

            '            End If
            '        End If
            '    End While
            'End Using


            'ListView1.View = View.Details
            'ListView1.FullRowSelect = True
            'ListView1.Items.Clear()
            'ListView1.Columns.Clear()
            'create column

            xmlnodelist = xmldoc.GetElementsByTagName("ss:Row")
            'For i = 0 To xmlnodelist(0).ChildNodes.Count - 1
            '    'ListView1.Columns.Add(xmlnodelist(0).ChildNodes.Item(i).InnerText.Trim, widthdict(i))
            'Next
            Dim myarray As Integer = xmlnodelist(0).ChildNodes.Count - 1
            For i = 1 To xmlnodelist.Count - 1

                'xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim()
                Dim result As Object
                Dim txdate As Date = CDate(xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString())
                Dim pkey0(2) As Object
                pkey0(0) = xmlnodelist(i).ChildNodes.Item(4).InnerText.Trim()
                pkey0(1) = xmlnodelist(i).ChildNodes.Item(5).InnerText.Trim()
                pkey0(2) = xmlnodelist(i).ChildNodes.Item(6).InnerText.Trim()
                result = DS.Tables(0).Rows.Find(pkey0)
                If IsNothing(result) Then
                    HouseBillSB.Append(xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim().ToString & vbTab &
                                         validstr(xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString) & vbTab &
                                             validstr(xmlnodelist(i).ChildNodes.Item(3).InnerText.Trim()) & vbTab &
                                             xmlnodelist(i).ChildNodes.Item(4).InnerText.Trim() & vbTab &
                                             xmlnodelist(i).ChildNodes.Item(5).InnerText.Trim() & vbTab &
                                             xmlnodelist(i).ChildNodes.Item(6).InnerText.Trim() & vbTab &
                                             xmlnodelist(i).ChildNodes.Item(7).InnerText.Trim() & vbTab &
                                             validstr(xmlnodelist(i).ChildNodes.Item(11).InnerText.Trim()) & vbTab &
                                             validstr(xmlnodelist(i).ChildNodes.Item(12).InnerText.Trim()) & vbCrLf)
                End If
                'If txdate >= startdate.Date Or txdate <= enddate.Date Then

                '    'xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString()
                '    HouseBillSB.Append(xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim().ToString & vbTab &
                '                          validstr(xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString) & vbTab &
                '                              validstr(xmlnodelist(i).ChildNodes.Item(3).InnerText.Trim()) & vbTab &
                '                              xmlnodelist(i).ChildNodes.Item(4).InnerText.Trim() & vbTab &
                '                              xmlnodelist(i).ChildNodes.Item(5).InnerText.Trim() & vbTab &
                '                              xmlnodelist(i).ChildNodes.Item(6).InnerText.Trim() & vbTab &
                '                              xmlnodelist(i).ChildNodes.Item(7).InnerText.Trim() & vbTab &
                '                              validstr(xmlnodelist(i).ChildNodes.Item(11).InnerText.Trim()) & vbTab &
                '                              validstr(xmlnodelist(i).ChildNodes.Item(12).InnerText.Trim()) & vbCrLf)
                'End If
                'Dim mytext As String() = {xmlnodelist(i).ChildNodes.Item(0).InnerText.Trim().ToString,
                '                          xmlnodelist(i).ChildNodes.Item(1).InnerText.Trim().ToString,
                '                          xmlnodelist(i).ChildNodes.Item(2).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(3).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(4).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(5).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(6).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(7).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(8).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(9).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(10).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(11).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(12).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(13).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(14).InnerText.Trim(),
                '                          xmlnodelist(i).ChildNodes.Item(15).InnerText.Trim()}
                'Dim mytext(myarray) As String
                'For j = 0 To myarray
                'mytext(j) = xmlnodelist(i).ChildNodes.Item(j).InnerText.Trim().ToString
                'Next

                'Dim item As New ListViewItem(mytext)
                'ListView1.Items.Add(item)



            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        '        End If
        '    End If
        'Next


        '    End With
        'End Using
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If HouseBillSB.Length > 0 Then
                ProgressReport(2, "Copy HouseBill")
                'housebillid bigserial NOT NULL,  seqno bigint,    housebilldate timestamp without time zone,  housebill text,
                'containerno text,  po bigint,  partno bigint,  qty numeric,  etddate date,  dep date,
                sqlstr = "copy housebill( seqno,housebilldate,housebill,containerno,po,partno,qty,etddate,dep) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, HouseBillSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy HouseBill" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            'If BillingDtlSB.Length > 0 Then
            '    ProgressReport(2, "Copy Billing DTL")
            '    'billingdocument bigint,item integer,billedqty numeric, salesunit text, requiredqty numeric,
            '    'netweight numeric, grossweight numeric, weightunit text, volume numeric,  vunit text,
            '    'pricingdate date,exrate numeric, netvalue numeric,curr,
            '    'salesdoc bigint, salesdocitem integer, material bigint, description text, shippingpoint integer,
            '    'plant integer, country text, cost numeric,curr2 subtotal(numeric)
            '    sqlstr = "copy billingdtl(billingdocument,item,billedqty,salesunit,requiredqty,netweight,grossweight,weightunit,volume,vunit,pricingdate,exrate,netvalue,curr,salesdoc,salesdocitem,material,description,shippingpoint,plant,country,cost,curr2,subtotal) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, BillingDtlSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy BillingDTL" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub
End Class