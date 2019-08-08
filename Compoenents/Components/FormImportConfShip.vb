Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
Public Class FormImportConfShip

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            'Get file
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf doWork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                ProgressReport(1, "Get Reference Data")
                ProgressReport(2, "Read Data")

                'Get po with supplierinvoicenum
                Dim ds As New DataSet
                Dim sqlstr = "select pohd,polineno,supplierinvoicenum from miro m" &
                         " left join pomiro pm on pm.miroid = m.miroid" &
                         " left join podtl pd on pd.podtlid = pm.podtlid" &
                         " where miropostingdate >= " & DateFormatyyyyMMdd(DateTimePicker1.Value) &
                         " order by pohd,polineno,miropostingdate;" &
                         "select pohd,polineno,supplierinvoicenum from miro m" &
                         " left join pomiro pm on pm.miroid = m.miroid" &
                         " left join podtl pd on pd.podtlid = pm.podtlid" &
                         " where miropostingdate = '2000-01-01'" &
                         " order by pohd,polineno,miropostingdate;"

                If DbAdapter1.TbgetDataSet(sqlstr, ds) Then
                    Dim idx1(1) As DataColumn               '
                    idx1(0) = ds.Tables(1).Columns(0)       'po    
                    idx1(1) = ds.Tables(1).Columns(1)       'polineno
                    ds.Tables(1).PrimaryKey = idx1

                    'fill ds.tables(1) with Po + polineno
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If Not IsDBNull(ds.Tables(0).Rows(i).Item(2)) Then
                            Dim pkey1(1) As Object
                            pkey1(0) = ds.Tables(0).Rows(i).Item(0)
                            pkey1(1) = ds.Tables(0).Rows(i).Item(1)
                            Dim result = ds.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                Dim dr As DataRow = ds.Tables(1).NewRow
                                dr.Item(0) = pkey1(0)
                                dr.Item(1) = pkey1(1)
                                dr.Item(2) = ds.Tables(0).Rows(i).Item(2)
                                If Not IsDBNull(dr.Item(0)) Then
                                    ds.Tables(1).Rows.Add(dr)
                                End If

                            End If
                        Else
                            Dim dr As DataRow = ds.Tables(1).NewRow
                            dr.Item(0) = ds.Tables(0).Rows(i).Item(0)
                            dr.Item(1) = ds.Tables(0).Rows(i).Item(1)
                            Try
                                ds.Tables(1).Rows.Add(dr)
                            Catch ex As Exception
                                'Debug.Print("dup")
                            End Try

                            'Debug.Print("item2(blank) {0} {1}", ds.Tables(0).Rows(i).Item(0), ds.Tables(0).Rows(i).Item(1))
                        End If
                    Next
                End If

                ProgressReport(1, "Read Data")
                ProgressReport(2, "Read Data")
                Do Until .EndOfData
                    myrecord = .ReadFields

                    If count > 2 Then
                        If IsNumeric(myrecord(1)) Then


                            If CDate(myrecord(5).Substring(6, 4) & "-" & myrecord(5).Substring(3, 2) & "-" & myrecord(5).Substring(0, 2)) >= DateTimePicker1.Value Then

                                'If cc = LA And reference = "" Then find(reference)
                                If myrecord(7) = "LA" And myrecord(12) = "" Then
                                    Dim pkey1(1) As Object
                                    pkey1(0) = myrecord(2)
                                    pkey1(1) = CInt(myrecord(4))
                                    Dim dr = ds.Tables(1).Rows.Find(pkey1)
                                    If Not IsNothing(dr) Then
                                        Dim mystring As String = String.Empty

                                        If Not IsDBNull(dr.Item(2)) Then
                                            myrecord(12) = dr.Item(2)
                                        End If

                                    End If
                                End If

                                ' cocd character varying,  po bigint, poitem integer,  createdon date,  seqno integer,  cc character varying,  deliverydate date,  docdate date,  qty numeric,  reference character varying,  delivery bigint,  deliveryitem integer,
                                myInsert.Append(validstr(myrecord(1)) & vbTab &
                                                validlong(myrecord(2)) & vbTab &
                                                validint(myrecord(4)) & vbTab &
                                                dateformatdotyyyymmdd(myrecord(5)) & vbTab &
                                                validint(myrecord(6)) & vbTab &
                                                validstr(myrecord(7)) & vbTab &
                                                dateformatdotyyyymmdd(myrecord(8)) & vbTab &
                                                dateformatdotyyyymmdd(myrecord(9)) & vbTab &
                                                validreal(myrecord(10)) & vbTab &
                                                validstr(myrecord(12)) & vbTab &
                                                validlong(myrecord(13)) & vbTab &
                                                validint(myrecord(14)) & vbCrLf)
                            End If
                        End If
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            ' mystr.Append("delete from cxconfship where createdon >= " & DateFormatyyyyMMdd(DateTimePicker1.Value) & ";")
            mystr.Append("select deletecxconfship(" & DateFormatyyyyMMdd(DateTimePicker1.Value) & ");")
            Dim sqlstr As String = "copy cxconfship(cocd,po,poitem,createdon,seqno,cc,deliverydate,docdate,qty,reference,delivery,deliveryitem) from stdin with null as 'Null';"
            'mystr.Append(sqlstr)
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try
                If RadioButton1.Checked Then
                    ProgressReport(1, "Replace Record Please wait!")
                    ra = DbAdapter1.ExNonQuery(mystr.ToString)
                End If
                ProgressReport(1, "Add Record Please wait!")
                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    ProgressReport(1, "Add Records Done.")
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)

            End Try

        End If
        ProgressReport(3, "Set Continuous Again")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select

        End If

    End Sub
    Private Function validstr(ByVal data As Object) As Object
        If IsDBNull(data) Then
            Return "Null"
        ElseIf data = "" Then
            Return "Null"
        End If
        Return data
    End Function

End Class