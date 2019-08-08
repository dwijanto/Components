Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
Public Class FormImportMasterContactCP

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Public Property groupid As Long = 0
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
        Dim myvendor As String = String.Empty
        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                ProgressReport(1, "Read Data")
                ProgressReport(2, "Read Data")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count >= 1 Then
                        If IsNumeric(myrecord(0)) Then
                            'vendorcode,groupid
                            myvendor = myrecord(2)
                            If myvendor = "" Then
                                myvendor = "Null"
                            End If
                            myInsert.Append(myrecord(0) & vbTab &
                                            myvendor & vbTab &
                                            myrecord(4) & vbTab &
                                            myrecord(5) & vbCrLf)

                        End If
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            mystr.Append("delete from marketemailcp; ")
            Dim sqlstr As String = "copy marketemailcp(shiptopartycode,vendorcode,name,email) from stdin with null as 'Null';"
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False

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
            End Select

        End If

    End Sub
End Class