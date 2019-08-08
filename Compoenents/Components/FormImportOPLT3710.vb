Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass

Public Class FormImportOPLT3710
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
   
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)



    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim ReadFileStatus As Boolean = False
    Dim Dataset1 As DataSet
    Dim sb As StringBuilder
 
    Dim Source As String

    Dim OpenFIleDialog1 As New OpenFileDialog
    Dim mySelectedPath As String
    Dim startdate As Date


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""
        startdate = DateTimePicker1.Value.Date
       

        If Not myThread.IsAlive Then
            With OpenFIleDialog1
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .FileName
                    Try
                        myThread = New System.Threading.Thread(myThreadDelegate)
                        myThread.SetApartmentState(ApartmentState.MTA)
                        myThread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            End With           
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()

        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()

        If DbAdapter1.getproglock("FImOPLT", HelperClass1.UserInfo.DisplayName, 1) Then
            ProgressReport(2, "This Program is being used by other person")
        Else

            ReadFileStatus = ImportTextFile(mySelectedPath, errMsg)
            If ReadFileStatus Then
                sw.Stop()
                DbAdapter1.getproglock("FImOPLT", HelperClass1.UserInfo.DisplayName, 0)
                ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                ProgressReport(5, "Set to continuous mode again")
            Else
                errSB.Append(errMsg & vbCrLf)
                ProgressReport(3, errSB.ToString)
                ProgressReport(5, "Set to continuous mode again")
            End If

        End If

        sw.Stop()
    End Sub

  
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3
                    ToolStripStatusLabel2.Text = message
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


    Private Function ImportTextFile(ByVal FileName As String, ByRef errMsg As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of myData)

        ProgressReport(3, "Open Text File...")
        Dim i As Long
        Try
            Dim myrecord() As String
            Dim tcount As Long = 0
            Dim docdate As Date
            ProgressReport(6, "Set To Marque")
            'For Each fi As IO.FileInfo In arrFI
            ProgressReport(3, String.Format("Read Text File...{0}", FileName))
            Using objTFParser = New FileIO.TextFieldParser(FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = False
                    Dim count As Long = 0

                    Do Until .EndOfData

                        myrecord = .ReadFields

                        If count > 1 Then
                            docdate = CDate(dateformatdotyyyymmddstring(myrecord(4)))
                            If docdate >= startdate Then
                                Dim mydata As New myData(FileName, count + 1, myrecord)
                                myList.Add(mydata)                               
                            End If
                        End If
                        count += 1
                        tcount += 1
                    Loop
                End With
            End Using

            If myList.Count = 0 Then
                errMsg = "Nothing to process."
                Return myret
            End If
            'get dataset
            Dim DS As New DataSet




            'get initial keys from Database fro related table
            ProgressReport(3, String.Format("Delete rows ..........."))
            DbAdapter1.deleteOPLT(startdate)
            ProgressReport(3, String.Format("Fill Data Set..........."))



            If Not FillDataset(DS, errMsg) Then
                Return False
            End If

            'Create object for handleing row creation

            Dim OPLT As New OPLTNew3710(DS)

            ProgressReport(3, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")
            For i = 0 To myList.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & myList.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                If Not OPLT.buildSB(errMsg, myList(i)) Then
                    Return False
                End If

                'End If
            Next
            ProgressReport(6, "Set To Marque")
            ProgressReport(3, String.Format("Copy To Db"))
            If Not OPLT.copyToDb(errMsg, Me) Then
                Return False
            End If
            myret = True

        Catch ex As Exception
            errMsg = String.Format("Row : {0} ", i) & ex.Message
        End Try
        'copy

        myret = True
        'ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(SW.Elapsed.Minutes, "00"), Format(SW.Elapsed.Seconds, "00"), SW.Elapsed.Milliseconds.ToString))
        Return myret
    End Function

    Private Function validchar(ByVal strvalue As String) As Object
        If strvalue = "" Then
            Return ""
        Else
            'Return "'" & Trim(strvalue.Replace("'", "''").Replace("""", "")) & "'"
            Return Trim(strvalue.Replace("'", "''").Replace("""", ""))
        End If
    End Function
    Private Function validint(ByVal intvalue As String) As Object
        If intvalue = "" Then
            Return "NULL"
        Else
            Return CInt(intvalue.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Private Function validdec(ByVal decvalue As String) As Object
        If decvalue = "" Then
            Return "NULL"
        Else
            Return CDec(decvalue.Replace(",", "").Replace("""", ""))
        End If
    End Function

    Private Function validdate(ByVal datevalue As String) As Object
        Dim mydata() As String
        If datevalue.Contains(".") Then
            mydata = datevalue.Split(".")
        Else
            mydata = datevalue.Split("/")
        End If

        If mydata.Length > 1 Then
            Return "'" & mydata(2) & "-" & mydata(1) & "-" & mydata(0) & "'"
        End If
        Return "NULL"
    End Function

    Private Function validboolean(ByVal booleanvalue As String) As String
        If booleanvalue = "Y" Then
            Return "True"
        Else
            Return "False"
        End If
    End Function

    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As String) As Boolean
        Dim myret As Boolean = False
        Dim sb As New StringBuilder
        'sb.Append(String.Format("select pd.* from purchdoc pd" &
        '                        " left join  purchdocdtl dt on dt.purchdoc = pd.purchdoc " &
        '                        " left join opltdtl d on d.purchdocdtlid = dt.purchdocid " &
        '                        " left join oplthd h on h.salesdoc = d.salesdoc" &
        '                        " where h.sdhddocdate >= '{0:yyyy-MM-dd}';", startdate))
        'sb.Append(String.Format("select dt.* from purchdocdtl dt " &
        '                        " left join opltdtl d on d.purchdocdtlid = dt.purchdocid" &
        '                        " left join oplthd h on h.salesdoc = d.salesdoc " &
        '                        " where h.sdhddocdate >='{0:yyyy-MM-dd}';", startdate))
        'sb.Append(String.Format("select hd.* from oplthd hd where hd.sdhddocdate >= '{0:yyyy-MM-dd}';", startdate))
        'sb.Append(String.Format("select odt.* from opltdtl odt left join oplthd hd on hd.salesdoc = odt.salesdoc" &
        '                       " where hd.sdhddocdate >='{0:yyyy-MM-dd}';", startdate))
        'sb.Append(" select opltdtlid  from opltdtl order by opltdtlid desc limit 1;select purchdocid  from purchdocdtl order by purchdocid desc limit 1 ;;")
        sb.Append(String.Format("select pd.purchdoc from purchdoc pd;")) 'POHD
        sb.Append(String.Format("select dt.purchdocid,dt.purchdoc,purchdocitem from purchdocdtl dt;")) 'PODTL
        sb.Append(String.Format("select hd.salesdoc from oplthd hd; ")) 'OPLTHD
        sb.Append(String.Format("select odt.purchdocdtlid,odt.salesdoc,odt.item from opltdtl odt;")) 'OPLTDTL
        sb.Append(" select opltdtlid from opltdtl  order by opltdtlid desc limit 1;select purchdocid  from purchdocdtl order by purchdocid desc limit 1 ;")


        If DbAdapter1.TbgetDataSet(sb.ToString, DS, errmessage) Then
            DS.Tables(0).TableName = "purchdoc"
            DS.Tables(1).TableName = "purchdocdtl"
            DS.Tables(2).TableName = "oplthd"
            DS.Tables(3).TableName = "opltdtl"
            DS.Tables(4).TableName = "seqsalesdtl"
            DS.Tables(5).TableName = "seqpodtl"


            Dim idx0(0) As DataColumn
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(1) As DataColumn
            idx1(0) = DS.Tables(1).Columns(1)
            idx1(1) = DS.Tables(1).Columns(2)
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn
            idx2(0) = DS.Tables(2).Columns(0)
            DS.Tables(2).PrimaryKey = idx2

            'Dim idx3(1) As DataColumn
            'idx3(0) = DS.Tables(3).Columns(0)
            'idx3(1) = DS.Tables(3).Columns(1)
            'DS.Tables(3).PrimaryKey = idx3


        Else
            Return False
        End If
        myret = True
        Return myret
    End Function
End Class

Public Class OPLTNew3710
    Public Property ds As DataSet
    Public Property cxopltsaleshd As New StringBuilder
    Public Property cxopltsalesdtl As New StringBuilder
    Public Property cxopltpohd As New StringBuilder
    Public Property cxopltpodtl As New StringBuilder
    Public Property cxvendor As New StringBuilder
    Public Property cxpovendor As New StringBuilder
    Public Property cxopltsb As New StringBuilder

    Dim seqsalesdtl As Long
    Dim seqpodtl As Long


    Dim salesdtlid As Long
    Dim podtlid As Long


    Public Sub New(ByVal ds As DataSet)
        Me.ds = ds
        seqsalesdtl = ds.Tables(4).Rows(0).Item(0)
        seqpodtl = ds.Tables(5).Rows(0).Item(0)
    End Sub

    Public Function buildSB(ByRef message As String, ByVal mydata As myData) As Boolean
        Dim myret As Boolean = False

        Dim myprogress As String = String.Empty
        Dim data = mydata.data
        Dim comments As String = String.Empty
        Dim result As DataRow
        Try

            myprogress = "OPLTHD" 'OPLTHD
            Dim pkey2(0) As Object
            pkey2(0) = data(1)
            result = ds.Tables(2).Rows.Find(pkey2)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(2).NewRow
                dr.Item(0) = data(1)
                ds.Tables(2).Rows.Add(dr)
                'Sqlstr = "Insert into oplthd(salesdoc,ponumber,sdhddocdate,soldtoparty,shiptoparty) Values (" & myrecord(1) & "," & escapeString(myrecord(3)) & "," & DateFormatDot(myrecord(4)) & "," & validNum(myrecord(7)) & "," & validNum(myrecord(9)) & ");"
                'sqlstr = "copy oplthd(salesdoc,ponumber,sdhddocdate,soldtoparty) from stdin with null as 'Null';"
                cxopltsaleshd.Append(data(1) & vbTab &
                                    validstr(data(3)) & vbTab &
                                    dateformatdotyyyymmdd(data(4)) & vbTab &
                                    data(7) & vbTab &
                                    data(9) & vbCrLf)
            End If

          

          

            myprogress = "PO HD"

            Dim pkey0(0) As Object
            pkey0(0) = data(11)
            result = ds.Tables(0).Rows.Find(pkey0)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(0).NewRow
                dr.Item(0) = data(11)
                ds.Tables(0).Rows.Add(dr)
                'Sqlstr = "Insert into purchdoc(purchdoc,vendorcode,pddocdate) values (" & myrecord(11) & "," & validNum(myrecord(16)) & "," & DateFormatDot(myrecord(13)) & ")"
                'sqlstr = "copy purchdoc(purchdoc,vendorcode,podocdate) from stdin with null as 'Null';"
                cxopltpohd.Append(data(11) & vbTab &
                                  data(16) & vbTab &
                                    dateformatdotyyyymmdd(data(13)) & vbCrLf)
            End If

            myprogress = "PO DTL"
            Dim pkey3(1) As Object
            pkey3(0) = data(11)
            pkey3(1) = data(12)
            result = ds.Tables(1).Rows.Find(pkey3)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(1).NewRow
                seqpodtl += 1
                podtlid = seqpodtl
                dr.Item(0) = podtlid
                dr.Item(1) = data(11)
                dr.Item(2) = data(12)
                ds.Tables(1).Rows.Add(dr)
                'Sqlstr = "Insert into purchdocdtl(purchdoc,purchdocitem,createdon,leadtime,material,quantity,oun) values
                ' (" & myrecord(11) & "," & validNum(myrecord(12)) & "," & DateFormatDot(myrecord(14)) & "," & validNum(myrecord(15))
                ' & "," & validNum(myrecord(18)) & "," & validNum(myrecord(20)) & "," & escapeString(myrecord(21)) & ")"
                'copy purchdocdtl(pohd,poitem,createdon,leadtime,material,quantity,oun,ms) from stdin with null as 'Null';"
                If data.length - 1 = 21 Then
                    cxopltpodtl.Append(data(11) & vbTab &
                                      data(12) & vbTab &
                                      dateformatdotyyyymmdd(data(14)) & vbTab &
                                      validint(data(15)) & vbTab &
                                      validstr(data(18)) & vbTab &
                                      validint(data(20)) & vbTab &
                                      validstr(data(21)) & vbTab &
                                      validstr("") & vbCrLf)
                Else
                    cxopltpodtl.Append(data(11) & vbTab &
                                      data(12) & vbTab &
                                      dateformatdotyyyymmdd(data(14)) & vbTab &
                                      validint(data(15)) & vbTab &
                                      validstr(data(18)) & vbTab &
                                      validint(data(20)) & vbTab &
                                      validstr(data(21)) & vbTab &
                                      validstr(data(22)) & vbCrLf)
                End If

            Else
                podtlid = result.Item(0)
            End If

            myprogress = "OPLT DTL"

            'Dim pkey1(1) As Object
            'pkey1(0) = data(1)
            'pkey1(1) = data(2)
            'result = ds.Tables(1).Rows.Find(pkey1)
            'If IsNothing(result) Then
            '    Dim dr As DataRow = ds.Tables(1).NewRow
            '    seqsalesdtl += 1
            '    salesdtlid = seqsalesdtl
            '    dr.Item(0) = salesdtlid
            '    dr.Item(1) = data(1)
            '    dr.Item(2) = data(2)
            '    ds.Tables(1).Rows.Add(dr)
            '    'Sqlstr = "Insert into opltdtl(salesdoc,item,createdby,rj,purchdocdtlid) values (" & myrecord(1) & "," & validNum(myrecord(2)) & "," & escapeString(myrecord(5)) & "," & escapeString(myrecord(6)) & "," & myPurchdocid & ")"
            '    ' sqlstr = "copy opltdtl(salesdoc,item,createdby,rj) from stdin with null as 'Null';"
            '    cxopltsalesdtl.Append(data(1) & vbTab &
            '                        data(2) & vbTab &
            '                        data(5) & vbTab &
            '                        validstr(data(6)) & vbTab &
            '                        podtlid & vbCrLf)
            'Else
            '    salesdtlid = result.Item(0)
            'End If
            cxopltsalesdtl.Append(data(1) & vbTab &
                                    data(2) & vbTab &
                                    data(5) & vbTab &
                                    validstr(data(6)) & vbTab &
                                    podtlid & vbCrLf)
            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2},Row Num {3}", myprogress, ex.Message, mydata.filename, mydata.rownumber)
        End Try
        Return myret
    End Function

    Public Function copyToDb(ByRef errMsg As String, ByVal myform As FormImportOPLT3710) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String
        Try

            If cxopltsaleshd.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy SalesHD"))
                sqlstr = "copy oplthd(salesdoc,ponumber,sdhddocdate,soldtoparty,shiptoparty) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxopltsaleshd.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            

            If cxopltpohd.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy POHD"))
                sqlstr = "copy purchdoc(purchdoc,vendorcode,pddocdate) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxopltpohd.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            If cxopltpodtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy PODTL"))
                sqlstr = "copy purchdocdtl(purchdoc,purchdocitem,createdon,leadtime,material,quantity,oun,ms) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxopltpodtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If

            If cxopltsalesdtl.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Sales Dtl"))
                sqlstr = "copy opltdtl(salesdoc,item,createdby,rj,purchdocdtlid) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxopltsalesdtl.ToString, myret)
                If Not myret Then
                    Return myret
                End If
            End If
            myret = True
        Catch ex As Exception
            errMsg = ex.Message
        End Try
        Return myret

    End Function

End Class