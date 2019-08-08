Imports System.Windows.Forms.Design
Imports System.ComponentModel.Design
Imports System.Windows.Forms.Design.Behavior

Public Class CollapsiblePanelDesigner
    Inherits ParentControlDesigner

    Dim actions As System.ComponentModel.Design.DesignerActionListCollection

    Public Overrides ReadOnly Property ActionLists As System.ComponentModel.Design.DesignerActionListCollection
        Get
            'Return MyBase.ActionLists
            Return actions
        End Get
    End Property

    Public Overrides ReadOnly Property SnapLines As System.Collections.IList
        Get
            Dim _snaplines As Collections.IList = MyBase.SnapLines
            _snaplines.Add(New SnapLine(SnapLineType.Baseline, 6 + Control.Font.Height))
            Return _snaplines
        End Get
    End Property
End Class
