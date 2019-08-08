Imports System.Threading
Public Class FormUpdateCMMFFamily
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)

    Dim myThreadDelegate As New ThreadStart(AddressOf doWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim CMMFAdapter1 As CMMFAdapter
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
    Sub doWork()
        Dim sw As New Stopwatch
        sw.Start()
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        ProgressReport(2, "Processing...")
        ProgressReport(6, "Marque")
        CMMFAdapter1 = New CMMFAdapter
        CMMFAdapter1.LoadCMMF()
        Using Parser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With Parser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

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
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If myrecord(1) = "3701" Then
                        Dim CMMFModel1 As CMMFModel = New CMMFModel With {
                                                                         .sorg = myrecord(1),
                                                                         .plnt = myrecord(2),
                                                                         .cmmf = myrecord(4),
                                                                         .materialdesc = myrecord(5),
                                                                         .commercialref = myrecord(6),
                                                                         .modelcode = myrecord(7),
                                                                         .cmmftype = myrecord(8),
                                                                         .sbu = myrecord(13),
                                                                         .brandid = IIf(myrecord(17) = "", 0, myrecord(17)),
                                                                         .rir = myrecord(19),
                                                                         .activitycode = myrecord(19),
                                                                         .comfam = IIf(myrecord(21) = "", 0, myrecord(21)),
                                                                         .range = myrecord(26),
                        .createon = CDate(myrecord(31).Substring(6, 4) & "-" & myrecord(31).Substring(3, 2) & "-" & myrecord(31).Substring(0, 2))
                                                                         }


                        CMMFAdapter1.ValidateCMMF(CMMFModel1)


                    End If
                Next
            End With
        End Using

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
        'Update
        'Range
        ProgressReport(6, "Continue")
        If CMMFAdapter1.AddRangeSB.Length > 0 Then
            ProgressReport(2, "Copy Range")

            Dim sqlstr = "copy range(rangeid,range) from stdin with null as 'Null';"
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            errmessage = CMMFAdapter1.dbadapter1.copy(sqlstr, CMMFAdapter1.AddRangeSB.ToString, myret)
            If Not myret Then
                ProgressReport(2, "Copy Range" & "::" & errmessage)
                Exit Sub
            End If
        End If

        If CMMFAdapter1.AddCMMFSB.Length > 0 Then
            ProgressReport(2, "Copy CMMF")
            'vbTab, vbCrLf, cmmf.activitycode, cmmf.brandid, cmmf.cmmf, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, 
            'cmmf.createon, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rangeid, cmmf.rir, cmmf.sbu, cmmf.sorg))
            Dim sqlstr = "copy cmmf(activitycode, brandid, cmmf, cmmftype,comfam, commercialref,createon,materialdesc,modelcode,plnt, rangeid, rir,sbu, sorg) from stdin with null as 'Null';"
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            errmessage = CMMFAdapter1.dbadapter1.copy(sqlstr, CMMFAdapter1.AddCMMFSB.ToString, myret)
            If Not myret Then
                ProgressReport(2, "Copy CMMF" & "::" & errmessage)
                Exit Sub
            End If
        End If

        If CMMFAdapter1.UpdCMMFSB.Length > 0 Then
            ProgressReport(2, "Update CMMF")
            'cmmf.activitycode, cmmf.brandid, cmmf.cmmftype, cmmf.comfam, cmmf.commercialref, cmmf.materialdesc, cmmf.modelcode, cmmf.plnt, cmmf.rir, cmmf.sbu, cmmf.sorg
            Dim sqlstr = "update cmmf set activitycode= foo.activitycode,brandid = foo.brandid::integer,cmmftype= foo.cmmftype, " &
                        " comfam=foo.comfam::integer,commercialref = foo.commercialref,materialdesc = foo.materialdesc,modelcode = foo.modelcode," &
                        " plnt = foo.plnt::integer,rir = foo.rir,sbu = foo.sbu,sorg = foo.sorg::integer" &
                        " from (select * from array_to_set12(Array[" & CMMFAdapter1.UpdCMMFSB.ToString &
                     "]) as tb (cmmf character varying,activitycode character varying, brandid character varying, cmmftype character varying, comfam character varying, commercialref character varying, materialdesc character varying, modelcode character varying, plnt character varying, rir character varying, sbu character varying, sorg character varying))foo where cmmf.cmmf = foo.cmmf::bigint;"
            Dim ra As Long
            Dim errmsg As String = String.Empty
            If Not CMMFAdapter1.dbadapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                ProgressReport(2, "Update CMMF" & "::" & errmsg)
                Exit Sub
            End If
        End If

        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))


    End Sub

End Class