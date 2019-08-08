Public Class CollapsiblePanelColorTable
    Inherits ProfessionalColorTable

    Public Overrides ReadOnly Property ButtonPressedBorder As System.Drawing.Color
        Get
            'Return MyBase.ButtonPressedBorder
            Return Color.Red
        End Get
    End Property
End Class
