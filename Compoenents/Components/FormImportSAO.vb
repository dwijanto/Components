Imports System.Threading
Imports Components.PublicClass
Imports System.Text

Public Class FormImportSAO

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

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        myInsert.Append(myrecord(0) & vbTab &
                                        myrecord(2) & vbTab &
                                        validstr(myrecord(4)) & vbTab &
                                        validstr(myrecord(5)) & vbTab &
                                        validstr(myrecord(6)) & vbTab &
                                        validdate(myrecord(7)) & vbTab &
                                        validstr(myrecord(8)) & vbTab &
                                        validdate(myrecord(9)) & vbCrLf)
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            mystr.Append("delete from saooplt;")
            mystr.Append("select setval('saooplt_saoopltid_seq',1,false);")
            Dim sqlstr As String = "copy saooplt(soldtoparty,shiptoparty,saoname,saost,saofg,saofgstartdate,saocp,saocpstartdate) from stdin with null as 'Null';"
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
                ra = DbAdapter1.ExNonQuery(mystr.ToString)
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
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
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

    Private Sub StatusStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Function validdate(ByVal myrecord As String) As Object
        If myrecord = "" Then
            Return "Null"
        Else
            Return String.Format("'{0:yyyy-MM-dd}'", CDate(myrecord))
        End If
    End Function

End Class