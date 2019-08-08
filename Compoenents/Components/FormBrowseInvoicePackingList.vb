Imports System.Threading
Imports Components.PublicClass
Imports Components.SharedClass
Imports System.Text
Public Class FormBrowseInvoicePackingList
    Dim myThreadDelegate As New ThreadStart(AddressOf doLoad)
    Dim myQueryDelegate As New ThreadStart(AddressOf doQuery)
    Dim myThread As New Thread(myThreadDelegate)
    Dim myQuery As New Thread(myQueryDelegate)


    Dim ds As DataSet

    Dim BS As BindingSource
    Dim cbbs As BindingSource
    Dim startdate As DateTime
    Dim enddate As DateTime
    Dim startdateDTP As New DateTimePicker
    Dim enddateDTP As New DateTimePicker
    Dim myfilter As String = String.Empty
    Dim mybasefolder As String
    Dim pdfplus As String

    Dim mytextfilter As String = String.Empty
    Private Sub FormBrowseFolder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripButton6.Click
        mytextfilter = ToolStripComboBox1.Text
        LoadData()
        'ToolStripComboBox1.ComboBox.Text = mytextfilter
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            myThread = New Thread(AddressOf doLoad)
            startdate = startdateDTP.Value.Date
            enddate = enddateDTP.Value.Date.AddDays(1)
            myThread.Start()
        End If
    End Sub

    Private Sub doLoad()
        ProgressReport(6, "Marquee")

        Dim sqlstr As String = "select dt.docemaildtname, hd.foldername,hd.receiveddate from docemaildt dt" &
                                " left join docemailhd hd on hd.docemailhdid = dt.docemailhdid" &
                                " where docemailtype <> 0 and receiveddate >= " & DateFormatyyyyMMdd(startdate) & " and receiveddate <=  " & DateFormatyyyyMMdd(enddate) & " order by receiveddate;" &
                                " select ''::character varying as foldername union all (select distinct foldername from docemailhd where docemailtype <> 0" &
                                " order by foldername);" &
                                " select * from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'logbook' order by pd.ivalue"

        ds = New DataSet
        BS = New BindingSource
        cbbs = New BindingSource

        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            Dim idx0(0) As DataColumn
            'idx0(0) = ds.Tables(0).Columns("docemailhdid")
            'ds.Tables(0).PrimaryKey = idx0



            'Dim docemailhdidxU As UniqueConstraint = New UniqueConstraint(New DataColumn() {ds.Tables(0).Columns("docemaildtname")})
            'ds.Tables(0).Constraints.Add(docemailhdidxU)

            ds.Tables(0).TableName = "DocEmail"
            BS.DataSource = ds.Tables(0)

            cbbs.DataSource = ds.Tables(1)
            ProgressReport(1, "Assign DataGridView DataSource")
            mybasefolder = ds.Tables(2).Rows(4).Item("cvalue")
            pdfplus = ds.Tables(2).Rows(7).Item("cvalue")
        Else
            ProgressReport(2, mymessage)
        End If
        ProgressReport(5, "Continuous")
        ProgressReport(8, "Update Combobox")
    End Sub


    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    BS.DataSource = ds.Tables(0)
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = BS


                    ToolStripComboBox1.ComboBox.DisplayMember = "foldername"
                    ToolStripComboBox1.ComboBox.DataSource = cbbs
                    'ToolStripComboBox1.ComboBox.Text = myfilter
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                Case 8
                    ToolStripComboBox1.ComboBox.Text = mytextfilter
            End Select

        End If

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Dim myrow As DataRowView = BS.Current
        Dim myfilename = myrow.Row.Item("docemaildtname").ToString
        Dim myfolder = IIf(myfilename.Contains("PACKING LIST"), "packinglist", "invoice")
        'Process.Start("explorer.exe", "/select," & "C:\temp\Documents\Forwarder\""" & DbAdapter1.validfilename(myfolder) & """")
        Process.Start("explorer.exe", "/select," & mybasefolder & "\" & myfolder & "\" & """" & myfilename & """")
    End Sub

    Sub doQuery()

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        'If ToolStripComboBox1.ComboBox.Text = "" Then
        '    BS.Filter = ""
        'Else
        '    BS.Filter = "[foldername] = '" & ToolStripComboBox1.ComboBox.Text & "'"
        '    'myfilter = ToolStripComboBox1.ComboBox.Text
        'End If
        ToolStripTextBox1_TextChanged(Me, e)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        With startdateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        With enddateDTP
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "dd-MMM-yyyy"
            .Width = 120
        End With
        Dim host1 = New ToolStripControlHost(startdateDTP)
        Dim host2 = New ToolStripControlHost(enddateDTP)
        ToolStrip1.Items.Insert(6, host1)
        ToolStrip1.Items.Insert(8, host2)
    End Sub


  

    

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim sb As New StringBuilder
        Dim myfilter As String = String.Empty
        Dim userfolder As String = String.Empty
        BS.Filter = ""
        If Not ToolStripComboBox1.Text = "" Then
            userfolder = "[foldername] = '" & ToolStripComboBox1.Text & "'"
            sb.Append(userfolder)
        End If
        If ToolStripTextBox1.Text <> "" Then
            myfilter = "[docemaildtname] like '" & ToolStripTextBox1.Text & "'"
        End If

        If sb.Length > 0 And myfilter <> "" Then
            sb.Append(" and ")
        End If
        sb.Append(myfilter)
        BS.Filter = sb.ToString
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim myrow As DataRowView = BS.Current
        Dim myfilename = myrow.Row.Item("docemaildtname").ToString
        Dim myfolder = IIf(myfilename.Contains("PACKING LIST"), "packinglist", "invoice")
        Dim p As New System.Diagnostics.Process
       
        p.StartInfo.FileName = pdfplus '"C:\Program Files\Nuance\PDF Professional 7\bin\PDFPlus.exe"
        p.StartInfo.Arguments = String.Format("""{0}\{1}\{2}""", mybasefolder, myfolder, myfilename)
        'p.StartInfo.UseShellExecute = False
        p.Start()
        'p.WaitForExit()

    End Sub
End Class