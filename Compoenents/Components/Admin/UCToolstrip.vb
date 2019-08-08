Public Class UCToolstrip
    Private maxHeight As Integer
    Private dofade As Boolean = True
    Dim myLoc As Point
    Dim _PanelSize As Point
    Private accelerator As Integer
    Private Sub ToolStripButton1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Timer1.Start()
    End Sub

    Public Property setPanelSize As Point
        Get
            Return _PanelSize
        End Get
        Set(ByVal value As Point)
            myLoc = value
            maxHeight = myLoc.Y
        End Set
    End Property


    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If dofade Then
            If myLoc.Y > 0 Then                
                Panel1.Size = myLoc
                'myLoc.Offset(0, -10)
                'myLoc.Offset(0, -1 * (maxHeight / 10))
                myLoc.Offset(0, (2 + accelerator) * -1)
            Else
                Timer1.Stop()
                dofade = False
                Panel1.Height = 0
                myLoc.Y = 0
                ToolStripButton1.Image = My.Resources.ieframe_587 'My.Resources.go_bottom
                ToolStripButton1.ToolTipText = "Expand"
                accelerator = 0
            End If
        Else
            If myLoc.Y < maxHeight Then                
                Panel1.Size = myLoc
                myLoc.Offset(0, maxHeight / 1)
            Else
                Timer1.Stop()
                ToolStripButton1.Image = My.Resources.ieframe_584 'My.Resources.go_top
                Panel1.Height = maxHeight
                myLoc = Panel1.Size
                dofade = True
                ToolStripButton1.ToolTipText = "Collapse"
                accelerator = 0                
            End If
        End If
        accelerator += 20
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        myLoc = Panel1.Size
        maxHeight = Panel1.Height
        ' Add any initialization after the InitializeComponent() call.

        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        UpdateStyles()
    End Sub

End Class
