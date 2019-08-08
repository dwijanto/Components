Public Class UCDoubleBufferPanel
    Inherits Panel

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.        
        Me.DoubleBuffered = True        
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.DoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        UpdateStyles()
    End Sub

End Class
