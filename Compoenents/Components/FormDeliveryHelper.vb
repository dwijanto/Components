Imports System.Threading
Imports Components.SharedClass
Imports Components.PublicClass
Public Class FormDeliveryHelper
    Dim myWorkDelegate As New ThreadStart(AddressOf DoWork)
    Dim myWork As New System.Threading.Thread(myWorkDelegate)
    Public Property bs As BindingSource
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)

    Private pohd As Long
    Private poitem As Integer
    Public Sub New(ByVal pohd As Long, ByVal poitem As Integer)

        ' This call is required by the designer.
        InitializeComponent()
        Me.pohd = pohd
        Me.poitem = poitem
        ' Add any initialization after the InitializeComponent() call.
        myWork.Start()

    End Sub

    Sub DoWork()
        'Dim sqlstr = "select plh.delivery,pld.deliveryitem,plh.reference,plh.deliverydate,pohd,poitem,e.vendorcode,v.vendorname::character varying,cmmf,description,deliveredqty,biloflading as containernumber,meansoftransid,meansoftranstype,housebill,plh.createdby " &
        '             " from packinglistdt pld" &
        '             " left join packinglisthd plh on plh.delivery = pld.delivery " &
        '             " left join housebill h on h.po = pld.pohd and h.containerno = plh.biloflading " &
        '             " left join ekko e on e.po = pld.pohd" &
        '             " left join vendor v on v.vendorcode = e.vendorcode" &
        '            " where pohd = " & pohd & " and poitem = " & poitem
        Dim sqlstr = "select plh.delivery,pld.deliveryitem,plh.reference,plh.deliverydate,pohd,poitem,e.vendorcode,v.vendorname::character varying,cmmf,description,deliveredqty,biloflading as containernumber,meansoftransid,meansoftranstype,housebill,plh.createdby " &
                     " from packinglistdt pld" &
                     " left join packinglisthd plh on plh.delivery = pld.delivery " &                    
                     " left join ekko e on e.po = pld.pohd" &
                     " left join vendor v on v.vendorcode = e.vendorcode" &
                    " where pohd = " & pohd & " and poitem = " & poitem
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            bs = New BindingSource
            bs.DataSource = DS.Tables(0)
            Try
                ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
            Catch ex As Exception

            End Try

        End If
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    'Me.ToolStripStatusLabel1.Text = message
                Case 2
                    'Me.ToolStripStatusLabel2.Text = message
                Case 4
                    Me.DataGridView1.AutoGenerateColumns = False
                    Me.DataGridView1.DataSource = bs
                Case (5)
                    'ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    'ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                  
                Case 8
                    'Fill DataGridView
                    'Label4.Text = "Record" & IIf(CType(bs1.DataSource, DataTable).Rows.Count > 1, "s", "") & " Found :" & CType(bs1.DataSource, DataTable).Rows.Count.ToString
            End Select

        End If

    End Sub

    Private Sub FormDeliveryHelper_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub


    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Button1.PerformClick()
    End Sub
End Class