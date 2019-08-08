Imports System.ComponentModel

<Docking(DockingBehavior.Never)>
<Designer(GetType(CollapsiblePanelDesigner))>
Public Class CollapsiblePanel
    Inherits Panel

    Dim collapsing As Boolean
    Dim _collapsed As Boolean
    Dim button As CollapseButton
    Dim Timer As Timer
    Dim accelerator As Integer
    Dim oldHeight As Integer

    Public Property Collapsed As Boolean
        Get
            Return _collapsed
        End Get
        Set(ByVal value As Boolean)
            If _collapsed <> value Then
                _collapsed = value
                If _collapsed Then
                    PerformCollapsed()
                Else
                    PerformExpand()
                End If

            End If
        End Set
    End Property

    <Browsable(True)>
    Public Overrides Property Text As String
        Get
            Return button.Text
        End Get
        Set(ByVal value As String)
            button.Text = value
        End Set
    End Property

    Protected Overrides Sub SetBoundsCore(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal specified As System.Windows.Forms.BoundsSpecified)
        MyBase.SetBoundsCore(x, y, width, height, specified)
        If (Not (Timer.Enabled) And Not (Collapsed)) Then
            oldHeight = height
        End If
    End Sub
    Public Sub New()
        Me.SetStyle(ControlStyles.ResizeRedraw Or ControlStyles.SupportsTransparentBackColor, True)
        Me.DoubleBuffered = True

        'Setup Button
        button = New CollapseButton
        button.Size = New Size(Me.Width, 25)
        button.Location = New Point(0, 0)
        button.Font = New Font("Tahoma", 8.0F, FontStyle.Bold)
        button.Dock = DockStyle.Top
        AddHandler button.Click, AddressOf button_click
        Me.Controls.Add(button)
        Timer = New Timer
        timer.Interval = 25
        AddHandler timer.Tick, AddressOf timmer_tick
    End Sub

    Private Sub button_click(ByVal sender As Object, ByVal e As EventArgs)
        If Not collapsing Then
            collapsed = True
        Else
            Collapsed = False
        End If
    End Sub

    Private Sub PerformCollapsed()
        collapsing = True
        SuspendLayout()
        Timer.Enabled = True
    End Sub

    Private Sub PerformExpand()
        collapsing = False
        SuspendLayout()
        Timer.Enabled = True
    End Sub

    Private Sub timmer_tick(ByVal sender As Object, ByVal e As EventArgs)
        If collapsing Then
            Me.Size = New Size(Me.Width, Me.Height - 2 - accelerator)
            If (Me.Height <= 25) Then
                Me.Size = New Size(Me.Width, 25)
                Timer.Enabled = False
                button.Collapsed = True
                accelerator = 0
                ResumeLayout()
            End If
        Else
            Me.Size = New Size(Me.Width, Me.Height + 2 + accelerator)
            If Me.Height >= oldHeight Then
                Me.Size = New Size(Me.Width, oldHeight)
                Timer.Enabled = False
                button.Collapsed = False
                accelerator = 0
                ResumeLayout()
            End If
        End If
        accelerator += 1
    End Sub

End Class
