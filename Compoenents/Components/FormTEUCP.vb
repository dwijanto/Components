Public Class FormTEUCP

    Dim mycontroller As TEUCPController
    Public StartDate As Date
    Public EndDate As Date
    Public ForecastStart As Date
    Public ForecastEnd As Date

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ComboBox1.SelectedIndex = 6

        mycontroller = New TEUCPController(Me)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            StartDate = CDate(String.Format("{0:yyyy}-{0:MM}-1", DateTimePicker1.Value.Date))
            EndDate = CDate(String.Format("{0:yyyy-MM}-1", DateTimePicker1.Value.AddMonths(ComboBox1.Items(ComboBox1.SelectedIndex)))).AddDays(-1)            
            mycontroller.run()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
End Class