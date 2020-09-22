Imports Components.SharedClass
Public Class FormScoreboardNet
    Dim myController As ScoreboardController = New ScoreboardController(Me)
    Public FullNameDirectory As String
    Private Sub FormScoreboardNet_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        CheckedListBox_SelectedIndexChanged(sender, e)
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DateTimePickerStartDate.Value.Year <> DateTimePickerEndDate.Value.Year Then
            MsgBox("Year selection date 1 and date 2 not the same")
            DateTimePickerStartDate.Focus()
            Exit Sub
        End If

        If Not ValidateFSLDate() Then
            Exit Sub
        End If
        If RadioButtonFinishedGoods.Checked Then
            myController.Department = DepartmentEnum.FinishedGoods
        Else
            myController.Department = DepartmentEnum.Components
        End If

        If RadioButtonExcludeSIS.Checked Then
            myController.ExcludeSiS = True
            myController.OnlySIS = False
            myController.OnlyWMF = False
        ElseIf RadioButtonOnlySIS.Checked Then
            myController.ExcludeSiS = False
            myController.OnlySIS = True
            myController.OnlyWMF = False
        ElseIf RadioButtonOnlyWMF.Checked Then
            myController.ExcludeSiS = False
            myController.OnlySIS = False
            myController.OnlyWMF = True
        End If

        myController.startdate = DateTimePickerStartDate.Value.Date
        myController.currentmonth = DateTimePickerCurrentMonth.Value.Date
        myController.enddate = DateTimePickerEndDate.Value.Date
        myController.fslstartdate = DateTimerPickerFSLFSSLStartDate.Value.Date

        'Validate FSSL Date First
        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
        DirectoryBrowser.Description = "Which directory do you want to use?"

        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            FullNameDirectory = DirectoryBrowser.SelectedPath
            myController.GenerateReport()
        End If


    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButtonFinishedGoods.CheckedChanged
        If RadioButtonFinishedGoods.Checked Then
            myController.Department = DepartmentEnum.FinishedGoods
        Else
            myController.Department = DepartmentEnum.Components
        End If

        myController.GetInitialData()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxWMF.CheckedChanged
        myController.WMF = CheckBoxWMF.Checked
        myController.GetInitialData()
    End Sub

    Private Function ValidateFSLDate() As Boolean
        ValidateFSLDate = True
        If Not Weekday(DateTimerPickerFSLFSSLStartDate.Value, vbUseSystemDayOfWeek) = 2 Then
            ValidateFSLDate = False
            MsgBox("Please select date on Monday for FSL / FSSL Start Date!")
        End If
    End Function
End Class