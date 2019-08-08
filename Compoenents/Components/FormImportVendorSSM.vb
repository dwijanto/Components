Imports System.Threading
Imports Components.PublicClass
Public Class FormImportVendorSSM
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
        'Read File
        'update Vendor
        Dim myupdate As New System.Text.StringBuilder
        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = False
                Dim count As Long = 0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        If myupdate.Length > 0 Then
                            myupdate.Append(",")
                        End If
                        myupdate.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", myrecord(0), myrecord(1)))
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        If myupdate.Length > 0 Then
            ProgressReport(1, "Start Update")
            Dim sqlstr As String = "update vendor v set officerid = foo.officerid from (select * from array_to_set2(Array[" & myupdate.ToString & "]) as tb (vendorcode character varying,officerid character varying))foo where v.vendorcode = foo.vendorcode::bigint;"
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try
                ra = DbAdapter1.ExNonQuery(sqlstr)
                ProgressReport(1, ra & "Record(s) affected. Update Done.")
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
End Class