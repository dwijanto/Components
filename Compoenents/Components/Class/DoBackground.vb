Imports System.Threading
Public Delegate Sub ProgressReportCallback(ByRef sender As Object, ByRef e As EventArgs)
Public Class DoBackground

    Dim Parent As Object
    Private _progressreportCallback As ProgressReportCallback
    Private _myThread As Thread

    Public Shared Event Callback(ByVal sender As Object, ByVal e As EventArgs)

    Public Property myThread As Thread
        Get
            Return _myThread
        End Get
        Set(ByVal value As Thread)
            _myThread = value
        End Set
    End Property

    Public Property ProgressReportCallback As ProgressReportCallback
        Get
            Return _progressreportCallback
        End Get
        Set(ByVal value As ProgressReportCallback)
            _progressreportCallback = value
        End Set
    End Property

    Public Sub New(ByVal Parent As Object, ByVal myCallBack As ProgressReportCallback)
        Me.Parent = Parent
        Me.ProgressReportCallback = myCallBack
    End Sub

    Public Sub New(ByVal parent As Object)
        Me.Parent = parent
    End Sub
    Public Sub doThread(ByVal start As System.Threading.ThreadStart)
        If IsNothing(myThread) Then
            run(start)
        Else
            If Not myThread.IsAlive Then
                run(start)
            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
        End If
    End Sub

    Sub run(ByVal start As System.Threading.ThreadStart)
        myThread = New Thread(start)
        myThread.SetApartmentState(ApartmentState.STA)
        myThread.Start()
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Parent.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Parent.ToolStripStatusLabel1.Text = message
                Case 2
                    Parent.ToolStripStatusLabel2.Text = message
                Case 4
                    If IsNothing(ProgressReportCallback) Then
                        MessageBox.Show("Error found :: ProgressReportCallback is not assigned")
                    Else
                        ProgressReportCallback.Invoke(Me, New System.EventArgs)
                        RaiseEvent Callback(4, New EventArgs)
                    End If
                                       
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 7
                Case 8
                    RaiseEvent Callback(8, New EventArgs)

            End Select

        End If

    End Sub

End Class
