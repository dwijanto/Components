Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass

Public Class FormImportMaterialMaster

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf doWork)


    Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    Dim miroSeq As Long
    Dim podtlseq As Long
    Dim cmmfpriceseq As Long
    Dim cmmfvendorpriceseq As Long

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoQuery()
        'Get last MiroPostingDate
        Dim sqlstr = "select miropostingdate from miro m order by miropostingdate desc limit 1;"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            If DS.Tables(0).Rows.Count > 0 Then
                ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
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
                    'Me.Label4.Text = message
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

    Private Sub FormImportZZ0035_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        myQueryThread.Start()
    End Sub

    Sub DoWork()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mycheck As Long
        'Dim mystr As New StringBuilder
        'Dim MiroSB As New System.Text.StringBuilder
        'Dim POHDSB As New System.Text.StringBuilder
        'Dim PODtlSB As New System.Text.StringBuilder
        'Dim POMiroSB As New System.Text.StringBuilder
        Dim cmmfSB As New System.Text.StringBuilder
        'Dim cmmfpriceSB As New System.Text.StringBuilder
        'Dim cmmfvendorpriceSB As New System.Text.StringBuilder
        'Dim updatecmmfpricesb As New System.Text.StringBuilder
        'Dim updateCMMFvendorpriceLastsb As New System.Text.StringBuilder
        'Dim updateCMMFvendorpriceInitsb As New System.Text.StringBuilder
        'Dim vendorSB As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        'Dim postingdate As Date

        sw.Start()
        Try
            Dim mymessage As String = String.Empty
            Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0

                    ProgressReport(6, "Marque")



                    'FillData
                    'materialmaster
                    'family
                    'brand
                    'sbu
                    'range
                    '
                    ProgressReport(2, "Initialize Table..")
                    sqlstr = "select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range,vendorcode,cmmftype,owner from materialmaster;" &
                             "select familyid,familyname::character varying from family;" &
                             "select familylv2id,familylv2name from familylv2;" &
                             "select brandid,brandname::character varying from brand;" &
                             "select sbuid,sbuname from sbusap;" &
                             "select range,rangedesc from range;" &
                             "select vendorcode,vendorname::character varying from vendor;" &
                             "select owner,ownerdescription from owner;"


                    mymessage = String.Empty
                    If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                        ProgressReport(2, mymessage)
                        Exit Sub
                    End If

                    DS.Tables(0).TableName = "MM"
                    Dim idx0(0) As DataColumn
                    idx0(0) = DS.Tables(0).Columns(0)
                    DS.Tables(0).PrimaryKey = idx0

                    DS.Tables(1).TableName = "family"
                    Dim idx1(0) As DataColumn
                    idx1(0) = DS.Tables(1).Columns(0)
                    DS.Tables(1).PrimaryKey = idx1

                    DS.Tables(2).TableName = "familylv2"
                    Dim idx2(0) As DataColumn
                    idx2(0) = DS.Tables(2).Columns(0)
                    DS.Tables(2).PrimaryKey = idx2

                    DS.Tables(3).TableName = "brand"
                    Dim idx3(0) As DataColumn
                    idx3(0) = DS.Tables(3).Columns(0)
                    DS.Tables(3).PrimaryKey = idx3

                    DS.Tables(4).TableName = "sbusap"
                    Dim idx4(0) As DataColumn
                    idx4(0) = DS.Tables(4).Columns(0)
                    DS.Tables(4).PrimaryKey = idx4

                    DS.Tables(5).TableName = "range"
                    Dim idx5(0) As DataColumn
                    idx5(0) = DS.Tables(5).Columns(0)
                    DS.Tables(5).PrimaryKey = idx5

                    DS.Tables(6).TableName = "vendor"
                    Dim idx6(0) As DataColumn
                    idx6(0) = DS.Tables(6).Columns(0)
                    DS.Tables(6).PrimaryKey = idx6

                    DS.Tables(7).TableName = "owner"
                    Dim idx7(0) As DataColumn
                    idx7(0) = DS.Tables(7).Columns(0)
                    DS.Tables(7).PrimaryKey = idx7

                    ProgressReport(2, "Read Text File...")
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count > 2 Then
                            mylist.Add(myrecord)
                        End If
                        count += 1
                    Loop
                    ProgressReport(2, "Build Record...")
                    ProgressReport(5, "Continuous")
                    For i = 0 To mylist.Count - 1
                        'ProgressReport(7, i + 1 & "," & mylist.Count)
                        mycheck = i
                        If mycheck = 3643 Then
                            Debug.Print("debug")
                        End If
                        myrecord = mylist(i)
                        If CLng(myrecord(2)) = 2100101519 Then
                            Debug.Print("debug")
                        End If
                        If myrecord(0) = "3701" Then
                            'find cmmf, check for update
                            Dim pkey0(0) As Object
                            pkey0(0) = myrecord(2)
                            Dim result As DataRow
                            result = DS.Tables(0).Rows.Find(pkey0)
                            If IsNothing(result) Then
                                'select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster;
                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item("cmmf") = myrecord(2)
                                dr.Item("sorg") = myrecord(0)
                                dr.Item("plant") = myrecord(1)
                                dr.Item("materialdesc") = myrecord(3)
                                If (myrecord(4) <> "") Then
                                    dr.Item("commref") = myrecord(4)
                                End If
                                If (myrecord(5) <> "") Then
                                    dr.Item("familylv1") = myrecord(5)
                                End If
                                If (myrecord(7) <> "") Then
                                    dr.Item("familylv2") = myrecord(7)
                                End If
                                If (myrecord(9) <> "") Then
                                    dr.Item("sbu") = myrecord(9)
                                End If
                                If (myrecord(15) <> "") Then
                                    dr.Item("brandid") = myrecord(15)
                                End If
                                If (myrecord(17) <> "") Then
                                    dr.Item("rri") = myrecord(17)
                                End If
                                If (myrecord(18) <> "") Then
                                    dr.Item("range") = myrecord(18)
                                End If
                                If (myrecord(26) <> "#N/A") Then
                                    If myrecord(26) <> "" Then
                                        dr.Item("vendorcode") = myrecord(26)
                                    End If
                                End If
                                If (myrecord(34) <> "") Then
                                    dr.Item("cmmftype") = myrecord(34)
                                End If

                                If (myrecord(37) <> "") Then
                                    dr.Item("owner") = myrecord(37)
                                End If
                                DS.Tables(0).Rows.Add(dr)
                                'cmmfSB.Append(myrecord(4) & vbCrLf)
                            Else
                                'Update

                                If myrecord(3) <> result.Item("materialdesc") Then
                                    result.Item("materialdesc") = myrecord(3)
                                End If
                                If Not IsDBNull(result.Item("commref")) Then
                                    If myrecord(4) <> result.Item("commref") Then
                                        result.Item("commref") = myrecord(4)
                                    End If
                                End If
                                If Not IsDBNull(result.Item("familylv1")) Then
                                    If myrecord(5) = "" Then
                                        result.Item("familylv1") = DBNull.Value
                                    Else
                                        If myrecord(5) <> result.Item("familylv1") Then
                                            result.Item("familylv1") = myrecord(5)
                                        End If
                                    End If

                                End If

                                If Not IsDBNull(result.Item("familylv2")) Then
                                    If myrecord(7) = "" Then
                                        result.Item("familylv2") = DBNull.Value
                                    Else
                                        If myrecord(7) <> result.Item("familylv2") Then
                                            result.Item("familylv2") = myrecord(7)
                                        End If
                                    End If

                                End If
                                If Not IsDBNull(result.Item("sbu")) Then
                                    If myrecord(9) = "" Then
                                        result.Item("sbu") = DBNull.Value
                                    Else
                                        If myrecord(9) <> result.Item("sbu") Then
                                            result.Item("sbu") = myrecord(9)
                                        End If
                                    End If

                                End If

                                If Not IsDBNull(result.Item("brandid")) Then
                                    If myrecord(15) = "" Then
                                        result.Item("brandid") = DBNull.Value
                                    Else
                                        If myrecord(15) <> result.Item("brandid") Then
                                            result.Item("brandid") = myrecord(15)
                                        End If
                                    End If
                                End If

                                If Not IsDBNull(result.Item("rri")) Then
                                    If myrecord(17) = "" Then
                                        result.Item("rri") = DBNull.Value
                                    Else
                                        If myrecord(17) <> result.Item("rri") Then
                                            result.Item("rri") = myrecord(17)
                                        End If
                                    End If

                                End If


                                If Not IsDBNull(result.Item("range")) Then
                                    If myrecord(18) = "" Then
                                        result.Item("range") = DBNull.Value
                                    Else
                                        If myrecord(18) <> result.Item("range") Then
                                            result.Item("range") = myrecord(18)
                                        End If
                                    End If

                                End If

                                If Not IsDBNull(result.Item("vendorcode")) Then
                                    If myrecord(26) = "#N/A" Or myrecord(26) = "" Then
                                        result.Item("vendorcode") = DBNull.Value
                                    Else
                                        If myrecord(26) <> result.Item("vendorcode") Then
                                            result.Item("vendorcode") = myrecord(26)
                                        End If
                                    End If
                                Else
                                    If myrecord(26) <> "#N/A" And myrecord(26) <> "" Then
                                        result.Item("vendorcode") = myrecord(26)
                                    End If
                                End If
                                'ProductType
                                If Not IsDBNull(result.Item("cmmftype")) Then
                                    If myrecord(34) = "" Then
                                        result.Item("cmmftype") = DBNull.Value
                                    Else
                                        If myrecord(34) <> result.Item("cmmftype") Then
                                            result.Item("cmmftype") = myrecord(34)
                                        End If
                                    End If
                                Else
                                    result.Item("cmmftype") = myrecord(34)
                                End If

                                'Owner
                                If Not IsDBNull(result.Item("owner")) Then
                                    If myrecord(37) = "" Then
                                        result.Item("owner") = DBNull.Value
                                    Else
                                        If myrecord(37) <> result.Item("owner") Then
                                            result.Item("owner") = myrecord(37)
                                        End If
                                    End If
                                Else
                                    result.Item("owner") = myrecord(37)
                                End If
                            End If
                            

                            'familylv2
                            'brand
                            'sbu
                            'range

                            'sqlstr = "select cmmf,sorg,plant,materialdesc,commref,familylv1,familylv2,sbu,brandid,rri,range from materialmaster;" &
                            '"select familyid,familyname from family;" &
                            '"select familylv2,familylv2name from familylv2;" &
                            '"select brandid,brandname from brand;" &
                            '"select sbuid,sbuname from sbusap;" &
                            '"select range,rangedesc from range"


                            'family
                            If myrecord(5) <> "" Then
                                Dim pkey1(0) As Object
                                pkey1(0) = myrecord(5)
                                result = DS.Tables(1).Rows.Find(pkey1)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(1).NewRow
                                    dr.Item("familyid") = myrecord(5)
                                    dr.Item("familyname") = myrecord(6)
                                    DS.Tables(1).Rows.Add(dr)
                                Else
                                    'Check if not the same the update
                                    If myrecord(6) = "" Then
                                        result.Item("familyname") = DBNull.Value
                                    ElseIf IsDBNull(result.Item("familyname")) Then
                                        result.Item("familyname") = myrecord(6)
                                    Else
                                        If myrecord(6) <> result.Item("familyname") Then
                                            result.Item("familyname") = myrecord(6)
                                        End If
                                    End If
                                End If
                            End If

                            'familylv2
                            If myrecord(7) <> "" Then
                                Dim pkey2(0) As Object
                                pkey2(0) = myrecord(7)
                                result = DS.Tables(2).Rows.Find(pkey2)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(2).NewRow
                                    dr.Item("familylv2id") = myrecord(7)
                                    dr.Item("familylv2name") = myrecord(8)
                                    DS.Tables(2).Rows.Add(dr)
                                Else
                                    'Check if not the same the update
                                    If myrecord(8) = "" Then
                                        result.Item("familylv2name") = DBNull.Value
                                    ElseIf IsDBNull(result.Item("familylv2name")) Then
                                        result.Item("familylv2name") = myrecord(8)
                                    Else
                                        If myrecord(8) <> result.Item("familylv2name") Then
                                            result.Item("familylv2name") = myrecord(8)
                                        End If
                                    End If
                                End If
                            End If

                            'Brand
                            If myrecord(15) <> "" Then
                                Dim pkey3(0) As Object
                                pkey3(0) = myrecord(15)
                                result = DS.Tables(3).Rows.Find(pkey3)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(3).NewRow
                                    dr.Item("brandid") = myrecord(15)
                                    dr.Item("brandname") = myrecord(16)
                                    DS.Tables(3).Rows.Add(dr)
                                Else
                                    'Check if not the same the update
                                    If myrecord(16) = "" Then
                                        result.Item("brandname") = DBNull.Value
                                    Else
                                        If myrecord(16) <> result.Item("brandname") Then
                                            result.Item("brandname") = myrecord(16)
                                        End If
                                    End If
                                End If
                            End If

                            'SBU
                            If myrecord(9) <> "" Then
                                Dim pkey4(0) As Object
                                pkey4(0) = myrecord(9)
                                result = DS.Tables(4).Rows.Find(pkey4)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(4).NewRow
                                    dr.Item("sbuid") = myrecord(9)
                                    dr.Item("sbuname") = myrecord(10)
                                    DS.Tables(4).Rows.Add(dr)
                                Else
                                    'Check if not the same the update
                                    If myrecord(10) = "" Then
                                        result.Item("sbuname") = DBNull.Value
                                    Else
                                        If myrecord(10) <> result.Item("sbuname") Then
                                            result.Item("sbuname") = myrecord(10)
                                        End If
                                    End If
                                End If
                            End If

                            'Range
                            If myrecord(18) <> "" Then
                                Dim pkey5(0) As Object
                                pkey5(0) = myrecord(18)
                                result = DS.Tables(5).Rows.Find(pkey5)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(5).NewRow
                                    dr.Item("range") = myrecord(18)
                                    dr.Item("rangedesc") = myrecord(19)
                                    DS.Tables(5).Rows.Add(dr)
                                Else
                                    'Check if not the same the update
                                    If myrecord(19) = "" Then
                                        result.Item("rangedesc") = DBNull.Value
                                    Else
                                        If IsDBNull(result.Item("rangedesc")) Then
                                            result.Item("rangedesc") = myrecord(19)
                                        Else
                                            If myrecord(19) <> result.Item("rangedesc") Then
                                                result.Item("rangedesc") = myrecord(19)
                                            End If
                                        End If


                                    End If
                                End If
                            End If

                            'vendorocde
                            If myrecord(26) <> "#N/A" And myrecord(26) <> "" Then
                                Dim pkey6(0) As Object
                                pkey6(0) = myrecord(26)
                                result = DS.Tables(6).Rows.Find(pkey6)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(6).NewRow
                                    dr.Item("vendorcode") = myrecord(26)
                                    dr.Item("vendorname") = myrecord(27)
                                    DS.Tables(6).Rows.Add(dr)
                                Else
                                    'Check if not the same then update
                                    If myrecord(27) = "#N/A" Then
                                        result.Item("vendorname") = DBNull.Value
                                    Else
                                        If myrecord(27) <> result.Item("vendorname") Then
                                            'result.Item("vendorname") = myrecord(27)
                                        End If
                                    End If
                                End If
                            End If

                            'owner
                            If myrecord(37) <> "" Then
                                Dim pkey7(0) As Object
                                pkey7(0) = myrecord(37)
                                result = DS.Tables(7).Rows.Find(pkey7)
                                If IsNothing(result) Then
                                    'create
                                    Dim dr As DataRow = DS.Tables(7).NewRow
                                    dr.Item("owner") = myrecord(37)
                                    dr.Item("ownerdescription") = myrecord(38)
                                    DS.Tables(7).Rows.Add(dr)
                                Else
                                    'Check if not the same then update
                                    If myrecord(38) = "" Then
                                        result.Item("ownerdescription") = DBNull.Value
                                    Else
                                        If IsDBNull(result.Item("ownerdescription")) Then
                                            result.Item("ownerdescription") = myrecord(38)
                                        Else
                                            If myrecord(38) <> result.Item("ownerdescription") Then
                                                result.Item("ownerdescription") = myrecord(38)
                                            End If
                                        End If


                                    End If
                                End If
                            End If
                        End If

                    Next




                End With
            End Using
            'update record

            'Dim mymessage As String = String.Empty
            Dim ds2 As DataSet
            ds2 = DS.GetChanges
            If Not IsNothing(ds2) Then
                ProgressReport(2, "Update Record.. Please wait!")
                mymessage = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If Not DbAdapter1.MaterialMaster(Me, mye) Then
                    ProgressReport(2, "Error" & "::" & mye.message)
                    Exit Sub
                End If
            End If

            'Dim errmsg As String = String.Empty
            'ProgressReport(6, "Marque")
            'If vendorSB.Length > 0 Then
            '    ProgressReport(2, "Copy Vendor")
            '    'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
            '    sqlstr = "copy vendor(vendorcode,vendorname) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, vendorSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy Vendor" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If

            'If cmmfpriceSB.Length > 0 Then
            '    ProgressReport(2, "Copy CMMFPrice")
            '    'cmmf,myyear,initailtx,initialprice,incoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2
            '    sqlstr = "copy cmmfprice(cmmf,myyear,initialtx,initialprice,invoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, cmmfpriceSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy CMMFPrice" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If
            'If updatecmmfpricesb.Length > 0 Then
            '    ProgressReport(2, "Update CMMFPrice")
            '    'lasttx,lastprice,invoiceverificationnumber2
            '    sqlstr = "update cmmfprice set lasttx= foo.lasttx::date,lastprice = foo.lastprice::numeric,invoiceverificationnumber2 = foo.invoiceverificationnumber2::bigint from (select * from array_to_set4(Array[" & updatecmmfpricesb.ToString &
            '             "]) as tb (id character varying,lasttx character varying,lastprice character varying,invoiceverificationnumber2 character varying))foo where cpid = foo.id::bigint;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
            '        ProgressReport(2, "Copy CMMFVendorPrice" & "::" & errmsg)
            '        Exit Sub
            '    End If
            'End If

            'If cmmfvendorpriceSB.Length > 0 Then
            '    ProgressReport(2, "Copy CMMFVendorPrice")
            '    'cmmf,myyear,initailtx,initialprice,incoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2
            '    sqlstr = "copy cmmfvendorprice(cmmf,vendorcode,myyear,initialtx,initialprice,invoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, cmmfvendorpriceSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy CMMFVendorPrice" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If

            'If updateCMMFvendorpriceLastsb.Length > 0 Then
            '    ProgressReport(2, "Update CMMFVendorPrice LastTx")
            '    'lasttx,lastprice,invoiceverificationnumber2
            '    sqlstr = "update cmmfvendorprice set lasttx= foo.lasttx::date,lastprice = foo.lastprice::numeric,invoiceverificationnumber2 = foo.invoiceverificationnumber2::bigint,agv2 = 0 from (select * from array_to_set4(Array[" & updateCMMFvendorpriceLastsb.ToString &
            '             "]) as tb (id character varying,lasttx character varying,lastprice character varying,invoiceverificationnumber2 character varying))foo where cpid = foo.id::bigint;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
            '        ProgressReport(2, "Copy CMMFVendorPrice LastTx" & "::" & errmsg)
            '        Exit Sub
            '    End If
            'End If

            'If updateCMMFvendorpriceInitsb.Length > 0 Then
            '    ProgressReport(2, "Update CMMFVendorPrice InitTx")
            '    'lasttx,lastprice,invoiceverificationnumber2
            '    sqlstr = "update cmmfvendorprice set inittx= foo.inittx::date,initialprice = foo.initialprice::numeric,invoiceverificationnumber = foo.invoiceverificationnumber::bigint from (select * from array_to_set4(Array[" & updateCMMFvendorpriceInitsb.ToString &
            '             "]) as tb (id character varying,inittx character varying,initialprice character varying,invoiceverificationnumber character varying))foo where cpid = foo.id::bigint;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
            '        ProgressReport(2, "Copy CMMFVendorPrice InitTx" & "::" & errmsg)
            '        Exit Sub
            '    End If
            'End If




            'If MiroSB.Length > 0 Then
            '    ProgressReport(2, "Copy Miro")
            '    'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
            '    sqlstr = "copy miro(mironumber,miropostingdate,supplierinvoicenum ,vendorcode) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, MiroSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy Miro" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If

            'If POHDSB.Length > 0 Then
            '    ProgressReport(2, "Copy POHD")
            '    'pohd bigint, pono character varying,purchasinggroup character varying, payt character varying
            '    sqlstr = "copy pohd(pohd,pono,purchasinggroup) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, POHDSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy POHD" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If
            'If PODtlSB.Length > 0 Then
            '    ProgressReport(2, "Copy PODTL")
            '    'pohd bigint, polineno character varying,cmmf bigint,oun character varying
            '    sqlstr = "copy podtl(pohd,polineno,cmmf,oun) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, PODtlSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy PODTL" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If
            'If POMiroSB.Length > 0 Then
            '    ProgressReport(2, "Copy POMIRO")
            '    'podtlid bigint,miroid bigint,amount numeric,qty numeric,crcy charcter varying,unitprice
            '    sqlstr = "copy pomiro(podtlid,miroid,amount,qty,crcy,unitprice) from stdin with null as 'Null';"
            '    Dim errmessage As String = String.Empty
            '    Dim myret As Boolean = False
            '    errmessage = DbAdapter1.copy(sqlstr, POMiroSB.ToString, myret)
            '    If Not myret Then
            '        ProgressReport(2, "Copy POMiro" & "::" & errmessage)
            '        Exit Sub
            '    End If
            'End If

        Catch ex As Exception
            ProgressReport(1, String.Format("Error found {0} {1} ", mylist(mycheck), ex.Message))

        End Try

        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub
End Class