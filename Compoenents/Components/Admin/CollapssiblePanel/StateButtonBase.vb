Imports System.Windows.Forms
Imports System.Drawing.Drawing2D
Public MustInherit Class StateButtonBase
    Inherits Control
    Implements IButtonControl

    Dim _dialogresult As DialogResult = DialogResult.None
    Dim _state As StateButtonState

    Public Sub New()
        SetStyle(ControlStyles.StandardClick Or ControlStyles.StandardDoubleClick, False)
    End Sub

    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(75, 23)
        End Get

    End Property

    Public Property DialogResult As System.Windows.Forms.DialogResult Implements System.Windows.Forms.IButtonControl.DialogResult
        Get
            Return _dialogresult

        End Get
        Set(ByVal value As System.Windows.Forms.DialogResult)
            _dialogresult = value
        End Set
    End Property

    Protected ReadOnly Property state As StateButtonState
        Get
            Return _state
        End Get        
    End Property

    Private Sub SetState(ByVal SetState As StateButtonState, ByVal [Set] As Boolean)
        Dim newstate As StateButtonState = Me.state
        If [Set] Then
            newstate = newstate Or SetState
        Else
            newstate = newstate And Not SetState
        End If
        If Me.state <> newstate Then
            Me._state = newstate
            Invalidate()
        End If
    End Sub

    Public Sub NotifyDefault(ByVal value As Boolean) Implements System.Windows.Forms.IButtonControl.NotifyDefault
        SetState(StateButtonState.Default, value)
    End Sub

    Public Sub PerformClick() Implements System.Windows.Forms.IButtonControl.PerformClick
        OnClick(EventArgs.Empty)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        SetState(StateButtonState.Pressed, True)
        MyBase.OnMouseDown(e)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim click As Boolean = False
        If (state And StateButtonState.Pressed) <> 0 Then
            click = True
        End If
        SetState(StateButtonState.Pressed, False)
        MyBase.OnMouseUp(e)

        If click Then
            Update()
            PerformClick()
        End If
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If Capture Then
            SetState(StateButtonState.Pressed, ClientRectangle.Contains(e.X, e.Y))
        End If
        MyBase.OnMouseMove(e)
    End Sub

    Protected Overrides Sub OnMouseEnter(ByVal e As System.EventArgs)
        SetState(StateButtonState.MouseHover, True)
        MyBase.OnMouseEnter(e)
    End Sub

    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        SetState(StateButtonState.MouseHover, False)
        MyBase.OnMouseLeave(e)
    End Sub

    <Flags()>
    Public Enum StateButtonState
        none = &H0
        Pressed = &H1
        MouseHover = &H2
        [Default] = &H4
        Disabled = &H8
    End Enum
End Class
