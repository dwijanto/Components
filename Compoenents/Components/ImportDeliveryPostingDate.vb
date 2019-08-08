Imports System.Threading
Imports System.ComponentModel
Imports Components.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports Components.SharedClass
Public Class ImportDeliveryPostingDate
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)


    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim ReadFileStatus As Boolean = False
    Dim Dataset1 As DataSet
    Dim sb As StringBuilder
    Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    Dim aprocess As Process = Nothing
    Dim Source As String
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim mySelectedPath As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        If Not myThread.IsAlive Then

            With FolderBrowserDialog1
                .RootFolder = Environment.SpecialFolder.Desktop
                '.SelectedPath = "c:\"
                .Description = "Select the source directory"
                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    mySelectedPath = .SelectedPath

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
        ProgressReport(2, TextBox2.Text & "Read Folder..")

        ReadFileStatus = ImportTextFile(mySelectedPath, errMsg)
        If ReadFileStatus Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(2, TextBox2.Text & "Done.")
            ProgressReport(5, "Set to continuous mode again")
        Else
            errSB.Append(errMsg & vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If

        sw.Stop()
    End Sub

    Private Shared Function WaitForAll(ByVal events As ManualResetEvent()) As Boolean
        Dim result As Boolean = False
        Try
            If events IsNot Nothing Then
                For i As Integer = 0 To events.Length - 1
                    events(i).WaitOne()
                Next
                result = True
            End If
        Catch
            result = False
        End Try
        Return result
    End Function
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.TextBox1.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    TextBox2.Text = message
                Case 3
                    TextBox3.Text = message
                Case 4
                    TextBox1.Text = message
                Case 5
                    ProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ProgressBar1.Minimum = 1
                    ProgressBar1.Value = myvalue(0)
                    ProgressBar1.Maximum = myvalue(1)
            End Select

        End If

    End Sub

    Private Sub FormImportData_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Load the query in background

        'myQueryThread.Start()
    End Sub

    Private Function ImportTextFile(ByVal FileName As String, ByRef errMsg As String) As Boolean
        Dim sb As New StringBuilder
        Dim myret As Boolean = False

        Dim list As New List(Of String)
        Dim myList As New List(Of myData)

        ProgressReport(2, "Scanning Text File...")
        ProgressReport(3, "Open Text File...")
        Dim i As Long
        Try
            'Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)

            Dim dir As New IO.DirectoryInfo(mySelectedPath)
            Dim arrFI As IO.FileInfo() = dir.GetFiles("*.XLS")
            'Dim objTFParser As FileIO.TextFieldParser
            Dim myrecord() As String
            Dim tcount As Long = 0
            ProgressReport(6, "Set To Marque")
            For Each fi As IO.FileInfo In arrFI
                ProgressReport(3, String.Format("Read Text File...{0}", fi.FullName))
                Using objTFParser = New FileIO.TextFieldParser(fi.FullName)
                    With objTFParser
                        .TextFieldType = FileIO.FieldType.Delimited
                        .SetDelimiters(Chr(9))
                        .HasFieldsEnclosedInQuotes = False
                        Dim count As Long = 0

                        Do Until .EndOfData
                            'If count > 0 Then
                            myrecord = .ReadFields
                            If count > 2 Then
                                If myrecord(1) <> "" And myrecord(1) <> "Delivery" Then
                                    Dim mydata As New myData(fi.FullName, count + 1, myrecord)
                                    myList.Add(mydata)
                                End If
                            End If

                            tcount += 1
                            'End If
                            count += 1

                        Loop
                    End With
                End Using
            Next
            If myList.Count = 0 Then
                errMsg = "Nothing to process."
                Return myret
            End If
            'get dataset
            Dim DS As New DataSet

            'get initial keys from Database fro related table
          
            If Not FillDataset(DS, errMsg) Then
                Return False
            End If

            'Create object for handleing row creation
            Dim DP As New DeliveryPostingDate(DS)

            ProgressReport(3, String.Format("Build Data row..........."))
            ProgressReport(5, "Set To Continuous")
            For i = 0 To myList.Count - 1
                'If i > 4 Then
                ProgressReport(7, i + 1 & "," & myList.Count)
                'ProgressReport(3, String.Format("Build Data row ....{0} of {1}", i, myList.Count - 1))
                If Not DP.buildSB(errMsg, myList(i)) Then
                    Return False
                End If

                'End If
            Next
            ProgressReport(6, "Set To Marque")
            ProgressReport(3, String.Format("Copy To Db"))
            If Not DP.copyToDb(errMsg, Me) Then
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

        Dim Sqlstr As String = " select delivery from cxdeliveryposting;" 


        If DbAdapter1.TbgetDataSet(Sqlstr, DS, errmessage) Then
            DS.Tables(0).TableName = "cxdeliveryposting"


            Dim idx0(0) As DataColumn               'cxdeliveryposting
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0
        Else
            Return False
        End If
        myret = True
        Return myret
    End Function
End Class

Public Class DeliveryPostingDate
    Public Property ds As DataSet
    Public Property cxdeliveryposting As New StringBuilder
   
    Public Sub New(ByVal ds As DataSet)
        Me.ds = ds
    End Sub

    Public Function buildSB(ByRef message As String, ByVal mydata As myData) As Boolean
        Dim myret As Boolean = False

        Dim myprogress As String = String.Empty
        Dim data = mydata.data
        Dim comments As String = String.Empty
        Dim result As DataRow
        Try

            myprogress = "DeliveryPosting"

            'Delivery Posting
            Dim pkey(0) As Object
            pkey(0) = data(1)
            result = ds.Tables(0).Rows.Find(pkey)
            If IsNothing(result) Then
                Dim dr As DataRow = ds.Tables(0).NewRow
                dr.Item(0) = data(1)
                ds.Tables(0).Rows.Add(dr)
                cxdeliveryposting.Append(data(1) & vbTab &
                                    dateformatdotyyyymmdd(data(16)) & vbCrLf)
            End If

            myret = True
        Catch ex As Exception
            message = String.Format("Progess {0} Errormessage {1} Filename {2},Row Num {3}", myprogress, ex.Message, mydata.filename, mydata.rownumber)
        End Try
        Return myret
    End Function

    Public Function copyToDb(ByRef errMsg As String, ByVal myform As ImportDeliveryPostingDate) As Boolean
        Dim myret As Boolean = False
        Dim sqlstr As String
        Try
            If cxdeliveryposting.ToString <> "" Then
                myform.ProgressReport(2, String.Format("Copy Delivery Posting Date"))
                sqlstr = "copy cxdeliveryposting(delivery,postingdate) from stdin with null as 'Null';"
                errMsg = DbAdapter1.copy(sqlstr, cxdeliveryposting.ToString, myret)
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