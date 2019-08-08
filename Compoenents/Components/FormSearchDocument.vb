Imports System.Threading
Imports Components.PublicClass
Imports System.Text
Imports Components.SharedClass

Public Class FormSearchDocument
    Dim sqlstr As String = String.Empty
    Dim myCriteria As String = String.Empty
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim DoSearchThread As New Thread(AddressOf doSearch)
    Dim DS As DataSet
    Dim mymessage As String = String.Empty
    Dim bs As BindingSource
    Dim mybasefolder As String = String.Empty
    Dim selectedfile As String = String.Empty
    Dim myfoldertx As String = String.Empty

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        '    0    Bill Of Lading
        '    1    SEB(Invoice)
        '    2    SEB Packing List
        '    3    Supplier(Invoice)
        '    4    Container(Number)
        '    5    SAP(PO)
        '    6    Client(PO)

        If Not DoSearchThread.IsAlive Then
            'Get file
            Dim mysearch As String = ToolStripTextBox1.Text.Replace("*", "%")

            Select Case ToolStripComboBox1.SelectedIndex
                'Case 0
                '    myCriteria = " where plh.housebill = '" & validstr(ToolStripTextBox1.Text) & "';"
                'Case 1
                '    myCriteria = " where  pd.docno = " & validstr(ToolStripTextBox1.Text) & ";"
                'Case 2
                '    myCriteria = " where  pld.delivery = " & validstr(ToolStripTextBox1.Text) & ";"
                'Case 3
                '    myCriteria = " where  ah.reference = " & validstr(ToolStripTextBox1.Text) & ";"
                'Case 4
                '    myCriteria = " where  plh.biloflading = '" & validstr(ToolStripTextBox1.Text) & "';"
                'Case 5
                '    myCriteria = " where  pdtl.pohd = " & validstr(ToolStripTextBox1.Text) & ";"
                'Case 6
                '    myCriteria = " where  ph.pono = " & validstr(ToolStripTextBox1.Text) & ";"
                'Case 0
                '    myCriteria = " where plh.housebill like '" & validstr(mysearch) & "'"
                'Case 1
                '    myCriteria = " where  pd.docno::text like '" & validstr(mysearch) & "'"
                'Case 2
                '    myCriteria = " where  pld.delivery::text like '" & validstr(mysearch) & "'"
                'Case 3
                '    myCriteria = " where  ah.reference like '" & validstr(mysearch) & "'"
                'Case 4
                '    myCriteria = " where  plh.biloflading like '" & validstr(mysearch) & "'"
                'Case 5
                '    myCriteria = " where  pdtl.pohd::text like '" & validstr(mysearch) & "'"
                'Case 6
                '    myCriteria = " where  ph.pono::text like '" & validstr(mysearch) & "'"
                Case 0
                    myCriteria = " where plh.housebill = '" & validstr(ToolStripTextBox1.Text) & "'"
                Case 1
                    myCriteria = " where  pd.docno = " & validstr(ToolStripTextBox1.Text)
                Case 2
                    myCriteria = " where  pld.delivery = " & validstr(ToolStripTextBox1.Text)
                Case 3
                    myCriteria = " where  ah.reference = '" & validstr(ToolStripTextBox1.Text) & "'"
                Case 4
                    myCriteria = " where  plh.biloflading = '" & validstr(ToolStripTextBox1.Text) & "'"
                Case 5
                    myCriteria = " where  pdtl.pohd = " & validstr(ToolStripTextBox1.Text)
                Case 6
                    myCriteria = " where  ph.pono = '" & validstr(ToolStripTextBox1.Text) & "'"
            End Select
            sqlstr = "select distinct plh.housebill,pd.docno, plh.delivery,ah.reference, plh.biloflading,pdtl.pohd,ph.pono, getfilename(housebill) as housebillpdf,getfilename(pld.delivery::character varying) as deliverypdf,getfilename(pd.docno::character varying) as docnopdf,plh.vendorcode,v.vendorname::character varying,plh.deliverydate,dtx.draftcreateddate from packinglisthd plh" &
                     " left join packinglistdt pld on plh.delivery = pld.delivery" &
                     " left join packinglistdocument pd on pd.delivery = plh.delivery and pd.typedoc = 1" &
                     " left join packinglistdocument pda on pda.delivery = plh.delivery and pda.typedoc = 2 " &
                     " left join accountinghd ah on ah.docno = pda.docno" &
                     " left join miro m on m.mironumber = ah.miro" &
                     " left join pomiro pm on pm.miroid = m.miroid" &
                     " left join podtl pdtl on pdtl.podtlid = pm.podtlid" &
                     " left join pohd ph  on ph.pohd = pdtl.pohd" &
                     " left join vendor v on v.vendorcode = plh.vendorcode" &
                     " left join docemailtx dtx on dtx.billoflading = housebill " & myCriteria & ";" &
                     " select * from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'logbook' order by pd.ivalue;"

            sqlstr = "select distinct plh.housebill,pd.docno, plh.delivery,ah.reference, plh.biloflading,pdtl.pohd,ph.pono, getfilename(housebill) as housebillpdf,getfilename(pld.delivery::character varying) as deliverypdf,getfilename(pd.docno::character varying) as docnopdf,plh.vendorcode,v.vendorname::character varying,plh.deliverydate,dtx.draftcreateddate from packinglisthd plh" &
                     " left join packinglistdt pld on plh.delivery = pld.delivery" &
                     " left join packinglistdocument pd on pd.delivery = plh.delivery and pd.typedoc = 1" &
                     " left join packinglistdocument pda on pda.delivery = plh.delivery and pda.typedoc = 2 " &
                     " left join accountinghd ah on ah.docno = pda.docno" &
                     " left join miro m on m.mironumber = ah.miro" &
                     " left join pomiro pm on pm.miroid = m.miroid" &
                     " left join podtl pdtl on pdtl.podtlid = pm.podtlid" &
                     " left join pohd ph  on ph.pohd = pdtl.pohd" &
                     " left join vendor v on v.vendorcode = plh.vendorcode" &
                     " left join docemailtx dtx on dtx.billoflading = housebill " & myCriteria & ";" &
                     " select * from paramdt pd left join paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'logbook' order by pd.ivalue;"


            DoSearchThread = New Thread(AddressOf doSearch)
            DoSearchThread.Start()

        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If

        


    End Sub

    Sub doSearch()
        DS = New DataSet
        ProgressReport(2, "Marque")
        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(4, "Init DataGridView")
        Else
            ProgressReport(1, mymessage)

        End If
        ProgressReport(3, "Continuous")
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
                    Me.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 3
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 4

                    bs = New BindingSource
                    bs.DataSource = DS.Tables(0)
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Invalidate()

                    With DataGridView1
                        .AutoGenerateColumns = False
                        .DataSource = bs
                    End With

                    mybasefolder = DS.Tables(1).Rows(4).Item("cvalue")
                    If DS.Tables(0).Rows.Count = 0 Then
                        Me.ToolStripStatusLabel1.Text = "Data not found!"
                    Else
                        Me.ToolStripStatusLabel1.Text = ""
                    End If
            End Select

        End If

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If DS.Tables(0).Rows.Count > 0 Then
            If Not selectedfile = "None" Then


                If selectedfile = "" Then
                    selectedfile = "\FORWARDER\" & CType(bs.Current, DataRowView).Row.Item("housebill") & "\" & DataGridView1.Rows(0).Cells(0).Value
                    myfoldertx = mybasefolder & "\" & CType(bs.Current, DataRowView).Row.Item("housebill")
                End If
                Dim myrow As DataRowView = bs.Current
                Dim myfolder = mybasefolder & selectedfile
                Dim p As New System.Diagnostics.Process
                p.StartInfo.FileName = "explorer.exe"
                p.StartInfo.Arguments = String.Format("{0},""{1}{2}""", "/select", mybasefolder, selectedfile)
                p.Start()
            End If
        End If


    End Sub



    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If e.RowIndex >= 0 Then


            Select Case e.ColumnIndex
                Case 0
                    selectedfile = "\FORWARDER\" & CType(bs.Current, DataRowView).Row.Item("housebill") & "\" & DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Case 1
                    selectedfile = "\Invoice\" & DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Case 2
                    selectedfile = "\PackingList\" & DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                Case Else                    
                    Exit Sub
            End Select
            Dim myrow As DataRowView = bs.Current
            Dim myfolder = mybasefolder & selectedfile
            Dim p As New System.Diagnostics.Process
            p.StartInfo.FileName = "explorer.exe"
            p.StartInfo.Arguments = String.Format("{0},""{1}{2}""", "/select", mybasefolder, selectedfile)
            p.Start()


        End If
    End Sub

    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If DoSearchThread.IsAlive Then
            DoSearchThread.Abort()
            ProgressReport(3, "Continuous")
        End If
    End Sub
End Class