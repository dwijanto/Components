Imports System.Threading
Imports Components.PublicClass
Public Class FormAssignPOSASLShipdate
    Dim mythread As New Thread(AddressOf doWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim ds As New DataSet

    Dim myposasldict As New Dictionary(Of Integer, DataRow)
    Dim mylist As New List(Of DataRow)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not mythread.IsAlive Then
            mythread = New Thread(AddressOf doWork)
            mythread.Start()
        Else
            MessageBox.Show("Please wait until the current process finished.")

        End If
    End Sub

    Private Sub doWork()
        ProgressReport(1, "Populating data...")
        'get dataset1
        Dim sqlstr = "select sebasiapono,polineno,cslstatus,shipdate from posasl;" &
                     "select distinct ph.sebasiapono,pd.polineno,sp.shipdate" &
                     " FROM cxsebodtp od" &
                     " LEFT JOIN cxrelsalesdocpo r ON r.cxrelsalesdocpoid = od.relsalesdocpoid" &
                     " LEFT JOIN cxsalesorderdtl sd ON sd.cxsalesorderdtlid = r.cxsalesorderdtlid" &
                     " LEFT JOIN cxsalesorder sh ON sh.sebasiasalesorder = sd.sebasiasalesorder" &
                     " LEFT JOIN cxsebpodtl pd ON pd.cxsebpodtlid = r.cxsebpodtlid" &
                     " LEFT JOIN cxsebpo ph ON ph.sebasiapono = pd.sebasiapono" &
                     " LEFT JOIN cxshipment sp ON sp.sebodtpid = od.cxsebodtpid" &
                     " where not sp.shipdate isnull"
        DS = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "POSASL"
            DS.Tables(1).TableName = "Query"

            Dim idx(2) As DataColumn
            idx(0) = DS.Tables(0).Columns(0)
            idx(1) = DS.Tables(0).Columns(1)
            idx(2) = DS.Tables(0).Columns(3)
            DS.Tables(0).PrimaryKey = idx

            Dim idx1(2) As DataColumn
            idx1(0) = DS.Tables(1).Columns(0)
            idx1(1) = DS.Tables(1).Columns(1)
            idx1(2) = ds.Tables(1).Columns(2)
            DS.Tables(1).PrimaryKey = idx1

            ProgressReport(3, "Update Bindingsource")

            ProgressReport(1, "Update Shipdate....")
            Dim i = 0
            Dim myidx(2) As Object

            Dim mycount As Long = 0
            myposasldict = New Dictionary(Of Integer, DataRow)
            mylist = New List(Of DataRow)
            For Each dr As DataRow In ds.Tables(0).Rows
                i += 1
                ProgressReport(2, i & "," & ds.Tables(0).Rows.Count)

                'find based on sebasiapono + polineno + shipdate 
                'if not available then
                'select based on sebasiapono and polineno



                myidx(0) = dr.Item("sebasiapono")
                myidx(1) = dr.Item("polineno")
                myidx(2) = dr.Item("shipdate")

                Dim myresult As DataRow
                myresult = ds.Tables(1).Rows.Find(myidx)
                If IsNothing(myresult) Then
                    Dim strExpression = "sebasiapono = " & dr.Item("sebasiapono") & " and polineno = " & dr.Item("polineno")
                    Dim foundRow = ds.Tables(1).Select(strExpression)
                    Dim j As Integer = 0

                    For k = 1 To foundRow.Length
                        If k > 1 Then
                            Dim drn = ds.Tables(0).NewRow
                            drn.Item(0) = foundRow(k - 1).Item(0)
                            drn.Item(1) = foundRow(k - 1).Item(1)
                            drn.Item(3) = foundRow(k - 1).Item(2)
                            drn.Item(2) = dr.Item(2)

                            mylist.Add(drn)
                            mycount += 1
                        Else
                            dr.Item("shipdate") = foundRow(k - 1).Item(2)
                        End If
                    Next
                    
                End If
            Next
            For Each row In mylist
                ds.Tables(0).Rows.Add(row)
            Next
            Dim ds2 = ds.GetChanges
            If Not IsNothing(ds2) Then
                If Not DbAdapter1.AdapterSASLTx(Me, ds, message:=mymessage) Then
                    ProgressReport(1, mymessage)
                End If
            End If
           



            'Save
        Else
            ProgressReport(1, mymessage)
        End If



        ProgressReport(1, "Update Done")

        
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
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)            
            End Select

        End If

    End Sub

End Class