Imports System.Threading
Imports Components.SharedClass
Imports Components.PublicClass
Public Class FormDeliveryHelper2
    Public Enum POTYPE
        Billing = 1
        Accounting = 2
    End Enum
    Public myPOTYPE As POTYPE = POTYPE.Accounting
    Dim myWorkDelegate As New ThreadStart(AddressOf DoWork)
    Dim myWork As New System.Threading.Thread(myWorkDelegate)
    Public Property bs As BindingSource
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)

    Private pohd As Long
    Private poitem As Integer
    Dim dt As DataTable
    Public Sub New(ByVal pohd As Long, ByVal poitem As Integer, ByVal bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.pohd = pohd
        Me.poitem = poitem
        dt = DirectCast(bs.DataSource, DataTable).Copy
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
        'Dim sqlstr = "select plh.delivery,pld.deliveryitem,plh.reference,plh.deliverydate,pohd,poitem,e.vendorcode,v.vendorname::character varying,cmmf,description,deliveredqty,biloflading as containernumber,meansoftransid,meansoftranstype,housebill,plh.createdby " &
        '             " from packinglistdt pld" &
        '             " left join packinglisthd plh on plh.delivery = pld.delivery " &
        '             " left join ekko e on e.po = pld.pohd" &
        '             " left join vendor v on v.vendorcode = e.vendorcode" &
        '            " where pohd = " & pohd & " and poitem = " & poitem
        Dim sqlstr = "with de as (select pld.delivery,pld.deliveryitem from packinglistdt pld" &
                     " where pld.pohd = " & pohd & " And pld.poitem = " & poitem &
                     " except " &
                     " select delivery,item from packinglistdocument where typedoc = " & myPOTYPE & ")" &
                     " select plh.delivery,pld.deliveryitem,plh.reference,plh.deliverydate,pohd,poitem,e.vendorcode,v.vendorname::character varying,cmmf,description,deliveredqty,biloflading as containernumber,meansoftransid,meansoftranstype,housebill,plh.createdby " &
                     " from packinglistdt pld" &
                     " left join packinglisthd plh on plh.delivery = pld.delivery " &
                     " inner join de on de.delivery = pld.delivery and de.deliveryitem = pld.deliveryitem" &
                     " left join ekko e on e.po = pld.pohd" &
                     " left join vendor v on v.vendorcode = e.vendorcode" &
                    " where pohd = " & pohd & " and poitem = " & poitem
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then

            ProgressReport(2, mymessage)
        Else
            Dim idx(1) As DataColumn
            idx(0) = DS.Tables(0).Columns("delivery")
            idx(1) = DS.Tables(0).Columns("deliveryitem")
            DS.Tables(0).PrimaryKey = idx

            bs = New BindingSource
            bs.DataSource = DS.Tables(0)

            'LINQ
            Dim Q As Object
            If myPOTYPE = POTYPE.Accounting Then
                Q = From row1 In dt
                   Where row1!pohd = pohd
                   Select row1
            Else
                Q = From row In dt
                   Where row!sebasiapono = pohd
                   Select row
            End If
           
            For Each x In q
                Debug.Print(String.Format("{0} {1}", x!delivery, x!deliveryitem))
                If Not IsDBNull(x!delivery) Then

                    Dim obj(1) As Object
                    obj(0) = x!delivery
                    obj(1) = x!deliveryitem
                    Dim result = DS.Tables(0).Rows.Find(obj)
                    If Not IsNothing(result) Then
                        DS.Tables(0).Rows.Remove(result)
                    End If                
                End If
            Next
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