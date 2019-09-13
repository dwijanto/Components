Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass

Public Class FormImportEkko

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
                ProgressReport(1, "Read Data")
                ProgressReport(2, "Read Data")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 2 Then
                        If IsNumeric(myrecord(1)) Then


                            If CDate(myrecord(5).Substring(6, 4) & "-" & myrecord(5).Substring(3, 2) & "-" & myrecord(5).Substring(0, 2)) >= DateTimePicker1.Value.Date Then
                                Try
                                    'po,companycode,createdon,createdby,vendorcode,termsofpayment,purchgroup,currency,docdate,incoterm1,incoterm2
                                    myInsert.Append(myrecord(1) & vbTab &
                                                    validstr(myrecord(3)) & vbTab &
                                                    validint(myrecord(4)) & vbTab &
                                                    dateformatdotyyyymmdd(myrecord(5)) & vbTab &
                                                    validstr(myrecord(6)) & vbTab &
                                                    validlong(myrecord(7)) & vbTab &
                                                    validstr(myrecord(8)) & vbTab &
                                                    validstr(myrecord(9)) & vbTab &
                                                    validstr(myrecord(10)) & vbTab &
                                                    dateformatdotyyyymmdd(myrecord(11)) & vbTab &
                                                    validstr(myrecord(12)) & vbTab &
                                                    validstr(myrecord(13)) & vbCrLf)

                                Catch ex As Exception
                                    ProgressReport(1, String.Format("PO# {0} - Err: {1}", myrecord(1), ex.Message))
                                    ProgressReport(3, "Set Continuous Again")
                                    Exit Sub
                                End Try
                                
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
            mystr.Append("delete from ekko where createdon >= " & DateFormatyyyyMMdd(DateTimePicker1.Value.Date) & ";")
            Dim sqlstr As String = "copy ekko(po,companycode,purchasingorg,createdon,createdby,vendorcode,termsofpayment,purchasinggroup,currency,docdate,incoterms,incoterms2) from stdin with null as 'Null';"
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
                Case 8
                    Label2.Text = message
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
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

    End Sub
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub FormImportEkko_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        showEkkoLastCreationDate()
    End Sub

    Private Sub showEkkoLastCreationDate()
        If Not mythread.IsAlive Then
            mythread = New Thread(AddressOf doQuery)
            mythread.Start()        
        Else
        MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doQuery()
        Dim myresult As Date
        If DbAdapter1.ExecuteScalar("select getekkolastcreatedon();", myresult) Then
            ProgressReport(8, String.Format("Latest Ekko Creation Date : {0:dd-MMM-yyyy} ", myresult))
        End If
    End Sub

End Class