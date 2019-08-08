Imports Components.PublicClass
Imports Components.SharedClass
Public Class FormChangePackingList

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        If MessageBox.Show("Are you sure to run this process?", "Execute?", MessageBoxButtons.OKCancel) = DialogResult.OK Then
            Dim sqlstr = "update packinglistdocument set delivery = " & TextBox2.Text & " where delivery = " & TextBox1.Text
            Dim mymessage As String = String.Empty
            Dim ra As Integer = 0
            If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                MessageBox.Show("Upadate failed. " & mymessage)
            Else
                If ra = 0 Then
                    MessageBox.Show("Nothing to update.")
                Else
                    MessageBox.Show("Update succeeded. " & ra & " records affected.")
                End If

            End If
        End If
    End Sub
End Class