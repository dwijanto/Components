Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass
Public Class FormConvertFamilySBU

    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim DS As DataSet
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread

        If Not mythread.IsAlive Then
            If MessageBox.Show("Execute the process?", "Question", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                ToolStripStatusLabel1.Text = ""

                mythread = New Thread(AddressOf doWork)
                mythread.Start()

            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub doWork()
        ProgressReport(2, "Marque")
        Dim sw As New Stopwatch
        sw.Start()
        'Dataset has 2 Table
        '1.Fill Dataset from turnoverhistory  where familysbu not isnull and sbuid isnull
        '2.prepare query family-sbu
        '
        'Process
        'Find family-sbu from turnoverhistory
        'if not isnull sbu id then
        'assign sbuid to turnoverhistory
        'get only changes record
        'update dataset
        'End

        DS = New DataSet

        Dim sqlstr = "select * from turnoverhistory where not familysbu isnull;" &
                     "select familylv1,sbuid from (select distinct m.familylv1,max(m.sbu) as sbu from materialmaster m" &
                     " where(Not m.familylv1 Is null)" &
                     " group by familylv1" &
                     " order by familylv1) as foo" &
                     " left join sbusap s on s.sbuid = foo.sbu" &
                     " order by familylv1;"
        Dim mymessage As String = String.Empty
        ProgressReport(1, "Query Data...")
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(1, mymessage)
            Exit Sub
        End If

        'Index for table familysbu
        DS.Tables(1).TableName = "FamilySBu"
        Dim idx1(0) As DataColumn
        idx1(0) = DS.Tables(1).Columns(0)
        DS.Tables(1).PrimaryKey = idx1

        ProgressReport(1, "Search Family SBU...")
        For i = 0 To DS.Tables(0).Rows.Count - 1
            'find table familysbu
            Dim dr As DataRow = DS.Tables(0).Rows(i)

            'Find FamilySBU
            Dim pkey0(0) As Object
            pkey0(0) = dr.Item("familysbu")

            Dim result As DataRow
            result = DS.Tables(1).Rows.Find(pkey0)

            If Not IsNothing(result) Then
                If IsDBNull(dr.Item("sbuid")) Then
                    dr.Item("sbuid") = result.Item("sbuid")
                ElseIf IsDBNull(result.Item("sbuid")) Then
                    dr.Item("sbuid") = DBNull.Value
                ElseIf dr.Item("sbuid") <> result.Item("sbuid") Then
                    dr.Item("sbuid") = result.Item("sbuid")
                End If                
            End If
        Next

        Dim ds2 As DataSet
        ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            ProgressReport(1, "Update Record.. Please wait!")
            mymessage = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If Not DbAdapter1.ConvertFamilySBU(Me, mye) Then
                ProgressReport(1, "Error" & "::" & mye.message)
                Exit Sub
            End If
        End If
        ProgressReport(3, "Continuous")
        sw.Stop()
        ProgressReport(1, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
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