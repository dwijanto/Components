Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Xml
Imports System.IO
Public Class FormImportPanalpina
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim mySelectedPath As String = String.Empty
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            With FolderBrowserDialog1
                .RootFolder = Environment.SpecialFolder.Desktop
                .SelectedPath = "c:\"
                .Description = "Select the source directory"
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .SelectedPath
                    Try

                        myThread = New Thread(AddressOf DoWork)
                        myThread.Start()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End With

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
        Dim HouseBillSB As New System.Text.StringBuilder
        Dim mylist As New List(Of String())
        Dim sqlstr As String = String.Empty

        Dim DS As New DataSet
        sw.Start()

        Dim mymessage As String = String.Empty

        'Fill Header
        ProgressReport(2, "Initialize Table..")
        sqlstr = "select housebill,containerno from forwarderhousebill where forwarder = 'PANALPINA';"

        mymessage = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If

        DS.Tables(0).TableName = "ForwarderHousebill"
        Dim idx0(1) As DataColumn
        idx0(0) = DS.Tables(0).Columns(0)
        idx0(1) = DS.Tables(0).Columns(1)
        DS.Tables(0).PrimaryKey = idx0

        ProgressReport(2, "Build Record...")
        ProgressReport(5, "Continuous")

        Dim dir As New IO.DirectoryInfo(mySelectedPath)
        Dim arrFI As IO.FileInfo() = dir.GetFiles("*.csv")
        Dim count As Long

        For Each fi As IO.FileInfo In arrFI
            Dim mycsvfile = fi.FullName
            ProgressReport(2, String.Format("Read Text File...{0}", fi.FullName))
            Try
                Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(",")
                        .HasFieldsEnclosedInQuotes = True
                        count = 0

                        Do Until .EndOfData
                            Dim myrecord = .ReadFields
                            If count = 209 Then
                                Debug.Print("hello")
                            End If
                            If count > 0 And myrecord(8) <> "-" And myrecord(21) <> "-" Then
                                Dim result As Object
                                Dim pkey0(1) As Object
                                pkey0(0) = myrecord(8)
                                pkey0(1) = myrecord(21)

                                result = DS.Tables(0).Rows.Find(pkey0)
                                If IsNothing(result) Then

                                    Dim dr As DataRow = DS.Tables(0).NewRow
                                    dr.Item(0) = myrecord(8)
                                    dr.Item(1) = myrecord(21)
                                    DS.Tables(0).Rows.Add(dr)
                                    HouseBillSB.Append(validstr(myrecord(8).Replace("""", "")) & vbTab &
                                                               myrecord(21).Replace("""", "") & vbTab &
                                                               "PANALPINA" & vbCrLf)
                                End If
                            End If
                            count += 1
                        Loop
                    End With
                End Using
            Catch ex As Exception
                mymessage = mymessage & fi.FullName & " " & ex.Message & vbCrLf
            End Try
            
        Next
        If mymessage.Length > 0 Then
            ProgressReport(2, "Done with Error:" & "::" & mymessage)
            Exit Sub
        End If
        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If HouseBillSB.Length > 0 Then
                ProgressReport(2, "Copy ForwarderHouseBill")
                sqlstr = "copy forwarderhousebill(housebill,containerno,forwarder) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, HouseBillSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy ForwarderHouseBill" & "::" & errmessage)
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