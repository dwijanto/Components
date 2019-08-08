Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Public Class FormImportAccountingHD
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)


    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
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
            'Get file
            'startdate = getfirstdate(DateTimePicker1.Value)
            'enddate = getlastdate(DateTimePicker2.Value)
            startdate = DateTimePicker1.Value.Date
            enddate = DateTimePicker2.Value.Date
            'appendfile = RadioButton1.Checked

            If openfiledialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    'Sub DoQuery()
    '    'Get last MiroPostingDate
    '    Dim sqlstr = "select miropostingdate from miro m order by miropostingdate desc limit 1;"
    '    Dim DS As New DataSet
    '    Dim mymessage As String = String.Empty
    '    If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
    '        ProgressReport(2, mymessage)
    '    Else
    '        ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
    '    End If
    'End Sub

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

    Private Sub FormImportZZ0035_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        'myQueryThread.Start()
    End Sub

    Sub DoWork()
        Dim sw As New Stopwatch
        Dim AccountingHDSB As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim miroid As Long
        'Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        Dim postingdate As Date

        sw.Start()
        Try
            Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0

                    'Delete Existing Record
                    ProgressReport(2, "Delete ..")
                    ProgressReport(6, "Marque")

                    sqlstr = "delete from accountinghd where postingdate >= " & DateFormatyyyyMMdd(startdate) & " and postingdate <= " & DateFormatyyyyMMdd(enddate) & ";" &
                             " select setval('accountinghd_accountinghdid_seq',(select accountinghdid from accountinghd order by accountinghdid desc limit 1) + 1,false);"

                    Dim mymessage As String = String.Empty
                    If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
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
                        If i >= 0 Then
                            postingdate = DbAdapter1.dateformatdotdate(myrecord(5))
                            'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
                            If postingdate >= startdate.Date AndAlso postingdate <= enddate.Date Then
                                Dim miro As String = (myrecord(13).Substring(0, 10))
                                Dim myyear As String = myrecord(13).Substring(10, 4)
                                AccountingHDSB.Append(validstr(myrecord(0)) & vbTab &
                                                    validlong(myrecord(1)) & vbTab &
                                                    validint(myrecord(2)) & vbTab &
                                                    validstr(myrecord(3)) & vbTab &
                                                    dateformatdotyyyymmdd(myrecord(4)) & vbTab &
                                                    dateformatdotyyyymmdd(myrecord(5)) & vbTab &
                                                    validstr(myrecord(6)) & vbTab &
                                                    validstr(myrecord(7)) & vbTab &
                                                    validstr(myrecord(8)) & vbTab &
                                                    validstr(myrecord(9)) & vbTab &
                                                    validreal(myrecord(10)) & vbTab &
                                                    validstr(myrecord(11)) & vbTab &
                                                    validreal(myrecord(12)) & vbTab &
                                                    validstr(miro) & vbTab &
                                                    validstr(myyear) & vbCrLf)
                            End If
                        End If
                    Next


                End With
            End Using
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
        
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If AccountingHDSB.Length > 0 Then
                ProgressReport(2, "Copy AccountingHeader")
                'podtlid bigint,miroid bigint,amount numeric,qty numeric,crcy charcter varying,unitprice
                sqlstr = "copy accountinghd(cocd,docno,myyear,doctype,docdate,postingdate,username,tcode,reference,crcy,exrate,lcur,exrate2,miro,miroyear) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, AccountingHDSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy AccountingHeader" & "::" & errmessage)
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
End Class