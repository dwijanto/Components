Imports System.Threading
Imports System.Text
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Xml
Imports System.IO
Imports Microsoft.Office.Interop
Public Class FormImportTEUCMMFVolume

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim mySelectedPath As String = String.Empty
    Dim myfilename As String
    Dim mysb As StringBuilder
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            With OpenFileDialog1
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    myfilename = OpenFileDialog1.FileName
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

        'Open Excel, saveas text file

        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim csvfile As String = String.Empty
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            oWb = oXl.Workbooks.Open(myfilename)
            csvfile = Application.StartupPath & "\teu.txt"
            'change format to general to get the original value. not formated one!
            oSheet = oWb.Worksheets(1)
            oSheet.Columns(38).select()
            oSheet.Application.Selection.NumberFormat = "General"
            oSheet.Columns(34).select()
            oSheet.Application.Selection.NumberFormat = "General"
            oWb.SaveAs(Filename:=csvfile, FileFormat:=Excel.XlPivotFieldDataType.xlText, CreateBackup:=False)
        Catch ex As Exception
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch exe As Exception
            End Try
            ProgressReport(2, String.Format("Error:: {0}", ex.Message))
            Exit Sub     
        End Try
        oXl.Quit()
        releaseComObject(oSheet)
        releaseComObject(oWb)
        releaseComObject(oXl)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Try
            'to make sure excel is no longer in memory
            EndTask(hwnd, True, True)
        Catch ex As Exception
        End Try

        'For Each fi As IO.FileInfo In arrFI
        'Dim mycsvfile = Application.StartupPath & "\teu.txt"
        ProgressReport(2, String.Format("Read Text File...{0}", csvfile))
        Dim count As Integer = 0
        Dim mysb = New StringBuilder
        Try
            Using objTFParser = New FileIO.TextFieldParser(csvfile)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True



                    Dim dt As New DataTable
                    Dim dt2 As New DataTable
                    dt.Columns.Add(New DataColumn("cmmf", System.Type.GetType("System.Int64")))
                    dt.Columns.Add(New DataColumn("value", System.Type.GetType("System.Decimal")))
                    dt2.Columns.Add(New DataColumn("cmmf", System.Type.GetType("System.Int64")))
                    dt2.Columns.Add(New DataColumn("value", System.Type.GetType("System.Decimal")))
                    DS.Tables.Add(dt)
                    DS.Tables.Add(dt2)

                    Do Until .EndOfData
                        Dim myrecord = .ReadFields
                        If count >= 4 And myrecord(28) <> "" Then
                            Dim dr = dt.NewRow
                            dr.Item("cmmf") = myrecord(28)
                            dr.Item("value") = (CDbl(myrecord(37)) / CDbl(myrecord(33))) '* 1000000
                            'If myrecord(28) = 1102270096 Then
                            '    Debug.Print("hello")
                            'End If

                            DS.Tables(0).Rows.Add(dr)
                        End If
                        count += 1
                    Loop
                End With
            End Using


            'Dim mydata = From p In DS.Tables(0).AsEnumerable
            '             Group p By cmmf = p.Item("cmmf") Into Group
            '             Select cmmf, average = Group.Average(Function(p) p.Item("value"))

            'use 
            Dim mydata = From p In DS.Tables(0).AsEnumerable
             Group p By cmmf = p.Field(Of Int64)("cmmf") Into Group
             Order By cmmf
             Select cmmf, average = Group.Average(Function(p) p.Field(Of Decimal)("value"))

            For Each dr In mydata
                mysb.Append(dr.cmmf & vbTab & dr.average & vbCrLf)
            Next
            Debug.Print(mysb.ToString)
        Catch ex As Exception
            mymessage = mymessage & csvfile & " " & ex.Message & vbCrLf
        End Try

        If mymessage.Length > 0 Then
            ProgressReport(2, "Done with Error:" & "::" & mymessage & " Line number : " & count)
            Exit Sub
        End If

        'update record
        Try
            ProgressReport(6, "Marque")
            Dim errmsg As String = String.Empty

            If mysb.Length > 0 Then
                ProgressReport(2, "Copy ForwarderHouseBill")
                sqlstr = "delete from cmmfvolume;"
                DbAdapter1.ExecuteNonQuery(sqlstr, message:=errmsg)
                sqlstr = "copy cmmfvolume(cmmf,avgvolume) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, mysb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy CmmfValue" & "::" & errmessage)
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