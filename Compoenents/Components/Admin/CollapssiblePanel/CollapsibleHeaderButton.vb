Imports System.Windows.Forms.VisualStyles

Public Class CollapsibleHeaderButton
    Inherits StateButtonBase

    Public Sub New()
        Me.SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.SupportsTransparentBackColor, True)
        Me.DoubleBuffered = True
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        If Application.RenderWithVisualStyles Then
            Dim renderer As VisualStyleRenderer
            InvokePaintBackground(Me, New PaintEventArgs(e.Graphics, ClientRectangle))
            renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.HeaderBackground.Normal)
            renderer.DrawBackground(e.Graphics, New Rectangle(0, 0, e.ClipRectangle.Width, 25))

            'Draw Text
            Dim fontRect As Rectangle = New Rectangle(17, 6, Me.Width - 17 - 24, Me.Height)
            If (state And StateButtonState.Pressed) <> 0 Then
                TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, fontRect, SystemColors.InactiveCaptionText, TextFormatFlags.Top Or TextFormatFlags.Left)
            ElseIf (state And StateButtonState.MouseHover) <> 0 Then
                TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, fontRect, SystemColors.ControlDarkDark, TextFormatFlags.Top Or TextFormatFlags.Left)
            ElseIf Not Enabled Then
                TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, fontRect, SystemColors.GrayText, TextFormatFlags.Top Or TextFormatFlags.Left)
            Else
                TextRenderer.DrawText(e.Graphics, Me.Text, Me.Font, fontRect, SystemColors.ActiveCaptionText, TextFormatFlags.Top Or TextFormatFlags.Left)
            End If


            'Draw button

            If (state And StateButtonState.Pressed) <> 0 Then
                renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.HeaderClose.Pressed)
            ElseIf (state And StateButtonState.MouseHover) <> 0 Then
                renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.HeaderClose.Hot)
            ElseIf Not Enabled Then
                renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.HeaderClose.Normal)
            Else
                renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.HeaderClose.Normal)
            End If
            renderer.DrawBackground(e.Graphics, New Rectangle(Me.Width - 22, 2, 20, 20))
        End If        
        MyBase.OnPaint(e)
    End Sub

End Class
