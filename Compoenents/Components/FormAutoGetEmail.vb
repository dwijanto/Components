Public Class FormAutoGetEmail

    Private Sub FormAutoGetEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Logger.log("----Start ---")
        Logger.log("----FinishedGood ---")
        Dim myform As New FormGetEmailFromExServer
        myform.DoWork()
        Logger.log("----Component ---")
        Dim myform2 As New FormGetEmailFromExServerCP

        myform2.DoWork()
        Me.Close()
        Logger.log("----End ---")
    End Sub
End Class