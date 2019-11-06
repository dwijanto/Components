Imports System.Reflection
Imports Components.PublicClass
Public Class FormMenu

    Public Function GetMenuDesc() As String
        Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId

    End Function
    Public Sub LoadMe() Handles Me.Load

        Me.FormMenu_Load(Me, New EventArgs)

        AddHandler ChangeUserToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ScoreBoardReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SASLSSLToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler WeeklyFigureToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CommentConversionToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportVendorSSMToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportSAOToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportSPComponentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

        AddHandler ImportWeeklyReportComponentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportIPLTComponentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportOPLTComponentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportDeliveryPostingToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler IPLTToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportWeeklyReportFGToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportEKKOToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ABCSupplierToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportConfirmationShipmentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportSPFGToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SASLStatusCommentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler AssignPOSASLShipdateToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportOPLTZZA013ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ForecastComponentsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ReadXmlToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportMiroToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportAccountingHeadersq01ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportPackingListToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportBillingDocumentToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        'AddHandler ImportForwarderHousebillToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler LogBookToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler LogBookV2ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

        AddHandler TableToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ChartToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CommentCodeTxToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportPO39ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler WORToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler DSVToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler PANALPINAToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportMaterialMastercsvToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler TurnoverHistoryConvertFamilySBUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ChartCriteriaToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportVendorSISSAOToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ShipmentFreightReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click



        AddHandler ImportTEUCMMFVolumeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ConversionCustSASSEBAsiaToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SPManagerToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler OPLTToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler OPLT720ToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click

        AddHandler CommentMappingToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler UpdateCMMFFamilysq01F037FGPuToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CMMFVolumeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterSBUToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterActivityToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
    End Sub
    Private Sub FormMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter
            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
            Try
                loglogin(DbAdapter1.userid)
            Catch ex As Exception
            End Try
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try

    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Lg Quick Upload"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Select Case ctrl.Name
            Case "ChartCriteriaToolStripMenuItem", "ShipmentFreightReportToolStripMenuItem"
                If DbAdapter1.checkproglock("FImportweeklyreportFGComp") Then
                    MessageBox.Show("WOR Data is being imported by another person. Please wait, do it later or contact Admin")
                    Exit Sub
                End If
                'Case "OPLTToolStripMenuItem"
                '    If DbAdapter1.checkproglock("FImOPLT") Then
                '        MessageBox.Show("OPLT Data is being imported by another person. Please wait, do it later or contact Admin")
                '        Exit Sub
                '    End If
        End Select


        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormMenu))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Protected Friend Sub setBubbleMessage(ByVal title As String, ByVal message As String)
        NotifyIcon1.BalloonTipText = message
        NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
        NotifyIcon1.BalloonTipTitle = title
        NotifyIcon1.Visible = True
        NotifyIcon1.ShowBalloonTip(200)
        'ShowballonWindow(200)
    End Sub





    'Private Sub ImportTEUCMMFVolumeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportTEUCMMFVolumeToolStripMenuItem.Click

    'End Sub


    'Private Sub WORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WORToolStripMenuItem.Click

    'End Sub

    'Private Sub LogBookToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogBookToolStripMenuItem.Click

    'End Sub

    'Private Sub ImportWeeklyReportComponentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportWeeklyReportComponentsToolStripMenuItem.Click

    'End Sub

    'Private Sub ConversionCustSASSEBAsiaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConversionCustSASSEBAsiaToolStripMenuItem.Click

    'End Sub

    'Private Sub SPManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPManagerToolStripMenuItem.Click

    'End Sub


    'Private Sub SASLStatusCommentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SASLStatusCommentsToolStripMenuItem.Click

    'End Sub




    Private Sub LogBookV2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogBookV2ToolStripMenuItem.Click

    End Sub

    Private Sub CommentMappingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CommentMappingToolStripMenuItem.Click

    End Sub

    Private Sub ConversionCustSASSEBAsiaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConversionCustSASSEBAsiaToolStripMenuItem.Click

    End Sub

    Private Sub UpdateCMMFFamilysq01F037FGPuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateCMMFFamilysq01F037FGPuToolStripMenuItem.Click

    End Sub

    Private Sub ImportSPFGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportSPFGToolStripMenuItem.Click

    End Sub

    Private Sub WORToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WORToolStripMenuItem.Click

    End Sub

    Private Sub CMMFBlankLoadingCodeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMMFBlankLoadingCodeToolStripMenuItem.Click
        Dim myform = New FormCMMFMissingLoadingCode
        myform.ShowDialog()

    End Sub

    Private Sub ImportOPLTCommentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportOPLTCommentsToolStripMenuItem.Click
        Dim myform = New FormImportOPLTComments
        myform.ShowDialog()
    End Sub

    Private Sub ImportMiroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportMiroToolStripMenuItem.Click

    End Sub

    Private Sub CustomerFlowToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CustomerFlowToolStripMenuItem.Click
        Dim myform = New FormCustomerFlow
        myform.ShowDialog()
    End Sub

    Private Sub FGToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FGToolStripMenuItem.Click
        Dim myform = New FormTEUFG
        myform.ShowDialog()
    End Sub

    Private Sub CPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CPToolStripMenuItem.Click
        Dim myform = New FormTEUCP
        myform.ShowDialog()
    End Sub

    Private Sub CommentCodeTxToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CommentCodeTxToolStripMenuItem.Click

    End Sub

    Private Sub TurnoverReprotToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TurnoverReprotToolStripMenuItem.Click
        Dim myform As New FormTurnoverReport
        myform.Show()
    End Sub

    Private Sub TurnoverCCToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TurnoverCCToolStripMenuItem.Click
        Dim myform As New FormRTurnoverExtendYearCur
        myform.Show()
    End Sub

    Private Sub ImportMaterialMastercsvToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportMaterialMastercsvToolStripMenuItem.Click

    End Sub

    Private Sub CMMFVolumeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMMFVolumeToolStripMenuItem.Click

    End Sub

    Private Sub SPManagerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SPManagerToolStripMenuItem.Click

    End Sub

    Private Sub ImportPackingListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportPackingListToolStripMenuItem.Click

    End Sub

    Private Sub MasterSBUToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterSBUToolStripMenuItem.Click

    End Sub

    Private Sub MasterActivityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterActivityToolStripMenuItem.Click

    End Sub
End Class