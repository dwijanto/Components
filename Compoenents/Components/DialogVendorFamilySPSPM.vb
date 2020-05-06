Imports System.Windows.Forms

Public Class DialogVendorFamilySPSPM
    Private DRV As DataRowView
    Public Sub New(ByVal drv As DataRowView)
        InitializeComponent()
        Me.DRV = drv
        initData()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub initData()
        Throw New NotImplementedException
    End Sub

End Class
