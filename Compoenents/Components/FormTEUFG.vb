Public Class FormTEUFG
    Dim mycontroller As TEUFGController
    Public StartDate As Date
    Public EndDate As Date    
    Public ForecastStart As Date
    Public ForecastEnd As Date

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ComboBox1.SelectedIndex = 6
        ComboBox2.SelectedIndex = 2
        mycontroller = New TEUFGController(Me)
        Label4.Text = ""
        If mycontroller.getLastUpdate Then
            Label4.Text = String.Format("Last update on {0:dd-MMM-yyyy} ", mycontroller.LastUpdate)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            StartDate = CDate(String.Format("{0:yyyy}-{0:MM}-1", DateTimePicker1.Value.Date))
            EndDate = CDate(String.Format("{0:yyyy-MM}-1", DateTimePicker1.Value.AddMonths(ComboBox1.Items(ComboBox1.SelectedIndex)))).AddDays(-1)
            ForecastStart = CDate(String.Format("{0:yyyy-MM}-1", DateTimePicker1.Value.AddMonths(ComboBox2.SelectedIndex)))
            mycontroller.run()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub



End Class