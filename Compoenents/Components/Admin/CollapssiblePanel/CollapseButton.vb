Imports System.Windows.Forms.VisualStyles

Public Class CollapseButton
    Inherits StateButtonBase

    Dim _Collapsed As Boolean

    Public Sub New()
        Me.SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.SupportsTransparentBackColor, True)
        Me.DoubleBuffered = True

    End Sub

    Public Property Collapsed As Boolean
        Get
            Return _Collapsed
        End Get
        Set(ByVal value As Boolean)
            If value <> _Collapsed Then
                _Collapsed = value
                Invalidate()
            End If
        End Set
    End Property

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        Invalidate()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        If Application.RenderWithVisualStyles Then
            Dim renderer As VisualStyleRenderer

            'Paint Parent Background
            InvokePaintBackground(Me, New PaintEventArgs(e.Graphics, ClientRectangle))
            renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupHead.Normal)
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

            'Draw Button
            If Not Collapsed Then
                'if pressed
                If (state And StateButtonState.Pressed) <> 0 Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupCollapse.Pressed)
                ElseIf (state And StateButtonState.MouseHover) <> 0 Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupCollapse.Hot)
                ElseIf Not Enabled Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupCollapse.Normal)
                Else
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupCollapse.Normal)
                End If
            Else
                'if pressed
                If (state And StateButtonState.Pressed) <> 0 Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupExpand.Pressed)
                ElseIf (state And StateButtonState.MouseHover) <> 0 Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupExpand.Hot)
                ElseIf Not Enabled Then
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupExpand.Normal)
                Else
                    renderer = New VisualStyleRenderer(VisualStyleElement.ExplorerBar.NormalGroupExpand.Normal)
                End If
            End If
            renderer.DrawBackground(e.Graphics, New Rectangle(Me.Width - 22, 3, 20, 20))
        Else

        End If
        

        MyBase.OnPaint(e)
    End Sub



End Class
