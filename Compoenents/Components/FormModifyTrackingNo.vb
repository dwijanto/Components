Imports Components.PublicClass
Public Class FormModifyTrackingNo
    Private Property oldtrackingnumber As String
    Public Property newtrackingnumber As String
    Public Property modified As Boolean = False
    Private Sub FormModifyTrackingNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = oldtrackingnumber
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal oldtrackingnumber As String)

        ' This call is required by the designer.
        InitializeComponent()
        Me.oldtrackingnumber = oldtrackingnumber
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Me.validate Then
            Dim sqlstr As String
            sqlstr = "update housebilldoc set trackingno = '" & TextBox2.Text.Replace("'", "''") & "' where trackingno = '" & TextBox1.Text.Replace("'", "''") & "'"


            Dim mymessage = String.Empty
            Dim myrecordaffected As Long
            If Not DbAdapter1.ExecuteNonQuery(sqlstr, myrecordaffected, mymessage) Then
                MessageBox.Show(mymessage)
                Exit Sub
            Else
                MessageBox.Show("Record affected: " & myrecordaffected)
                newtrackingnumber = TextBox2.Text
                modified = True
                If Not IsNothing(oldtrackingnumber) Then
                    Me.Close()
                End If
            End If
        End If
    End Sub

    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()
        If TextBox1.Text = "" Then
            myret = False
            ErrorProvider1.SetError(TextBox1, "Value cannot be blank")
        Else
            ErrorProvider1.SetError(TextBox1, "")
        End If

        Return myret
    End Function

End Class