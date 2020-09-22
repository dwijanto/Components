<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormScoreboardNet
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.RadioButtonOnlyWMF = New System.Windows.Forms.RadioButton()
        Me.RadioButtonOnlySIS = New System.Windows.Forms.RadioButton()
        Me.RadioButtonExcludeSIS = New System.Windows.Forms.RadioButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DateTimerPickerFSLFSSLStartDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePickerCurrentMonth = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTimePickerStartDate = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePickerEndDate = New System.Windows.Forms.DateTimePicker()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.CheckBoxWMF = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButtonComponents = New System.Windows.Forms.RadioButton()
        Me.RadioButtonFinishedGoods = New System.Windows.Forms.RadioButton()
        Me.ToolStripContainer1.BottomToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.BottomToolStripPanel
        '
        Me.ToolStripContainer1.BottomToolStripPanel.Controls.Add(Me.StatusStrip1)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label6)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label5)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Button1)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.RadioButtonOnlyWMF)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.RadioButtonOnlySIS)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.RadioButtonExcludeSIS)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label2)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.DateTimerPickerFSLFSSLStartDate)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label1)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.DateTimePickerCurrentMonth)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label4)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.Label3)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.DateTimePickerStartDate)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.DateTimePickerEndDate)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.CheckedListBox1)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.CheckBoxWMF)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.GroupBox1)
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(785, 428)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(785, 475)
        Me.ToolStripContainer1.TabIndex = 0
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(785, 22)
        Me.StatusStrip1.TabIndex = 0
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(668, 17)
        Me.ToolStripStatusLabel2.Spring = True
        Me.ToolStripStatusLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(356, 298)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(415, 13)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Import packing list (SQ01 0037): LgQuick Upload -> Upload Data -> Import PackingL" & _
            "ist"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(356, 273)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(356, 13)
        Me.Label5.TabIndex = 40
        Me.Label5.Text = "Import ZZA013: Lg Quick Upload -> Upload Data -> Import OPLT ZZA013"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(344, 386)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(81, 31)
        Me.Button1.TabIndex = 39
        Me.Button1.Text = "Export"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RadioButtonOnlyWMF
        '
        Me.RadioButtonOnlyWMF.AutoSize = True
        Me.RadioButtonOnlyWMF.Location = New System.Drawing.Point(184, 364)
        Me.RadioButtonOnlyWMF.Name = "RadioButtonOnlyWMF"
        Me.RadioButtonOnlyWMF.Size = New System.Drawing.Size(75, 17)
        Me.RadioButtonOnlyWMF.TabIndex = 38
        Me.RadioButtonOnlyWMF.Text = "Only WMF"
        Me.RadioButtonOnlyWMF.UseVisualStyleBackColor = True
        '
        'RadioButtonOnlySIS
        '
        Me.RadioButtonOnlySIS.AutoSize = True
        Me.RadioButtonOnlySIS.Location = New System.Drawing.Point(184, 341)
        Me.RadioButtonOnlySIS.Name = "RadioButtonOnlySIS"
        Me.RadioButtonOnlySIS.Size = New System.Drawing.Size(66, 17)
        Me.RadioButtonOnlySIS.TabIndex = 37
        Me.RadioButtonOnlySIS.Text = "Only SIS"
        Me.RadioButtonOnlySIS.UseVisualStyleBackColor = True
        '
        'RadioButtonExcludeSIS
        '
        Me.RadioButtonExcludeSIS.AutoSize = True
        Me.RadioButtonExcludeSIS.Checked = True
        Me.RadioButtonExcludeSIS.Location = New System.Drawing.Point(184, 318)
        Me.RadioButtonExcludeSIS.Name = "RadioButtonExcludeSIS"
        Me.RadioButtonExcludeSIS.Size = New System.Drawing.Size(83, 17)
        Me.RadioButtonExcludeSIS.TabIndex = 36
        Me.RadioButtonExcludeSIS.TabStop = True
        Me.RadioButtonExcludeSIS.Text = "Exclude SIS"
        Me.RadioButtonExcludeSIS.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(64, 298)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(114, 13)
        Me.Label2.TabIndex = 35
        Me.Label2.Text = "FSL / FSSL Start Date"
        '
        'DateTimerPickerFSLFSSLStartDate
        '
        Me.DateTimerPickerFSLFSSLStartDate.CustomFormat = "dd-MMM-yyyy"
        Me.DateTimerPickerFSLFSSLStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimerPickerFSLFSSLStartDate.Location = New System.Drawing.Point(184, 292)
        Me.DateTimerPickerFSLFSSLStartDate.Name = "DateTimerPickerFSLFSSLStartDate"
        Me.DateTimerPickerFSLFSSLStartDate.Size = New System.Drawing.Size(125, 20)
        Me.DateTimerPickerFSLFSSLStartDate.TabIndex = 34
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(103, 272)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Current Month"
        '
        'DateTimePickerCurrentMonth
        '
        Me.DateTimePickerCurrentMonth.CustomFormat = "MMM-yyyy"
        Me.DateTimePickerCurrentMonth.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerCurrentMonth.Location = New System.Drawing.Point(183, 266)
        Me.DateTimePickerCurrentMonth.Name = "DateTimePickerCurrentMonth"
        Me.DateTimePickerCurrentMonth.Size = New System.Drawing.Size(125, 20)
        Me.DateTimePickerCurrentMonth.TabIndex = 32
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(121, 246)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Date From"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(315, 244)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 30
        Me.Label3.Text = "To"
        '
        'DateTimePickerStartDate
        '
        Me.DateTimePickerStartDate.CustomFormat = "dd-MMM-yyyy"
        Me.DateTimePickerStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerStartDate.Location = New System.Drawing.Point(184, 240)
        Me.DateTimePickerStartDate.Name = "DateTimePickerStartDate"
        Me.DateTimePickerStartDate.Size = New System.Drawing.Size(125, 20)
        Me.DateTimePickerStartDate.TabIndex = 29
        '
        'DateTimePickerEndDate
        '
        Me.DateTimePickerEndDate.CustomFormat = "dd-MMM-yyyy"
        Me.DateTimePickerEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerEndDate.Location = New System.Drawing.Point(341, 240)
        Me.DateTimePickerEndDate.Name = "DateTimePickerEndDate"
        Me.DateTimePickerEndDate.Size = New System.Drawing.Size(118, 20)
        Me.DateTimePickerEndDate.TabIndex = 28
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.CheckOnClick = True
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(184, 95)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(449, 124)
        Me.CheckedListBox1.TabIndex = 4
        '
        'CheckBoxWMF
        '
        Me.CheckBoxWMF.AutoSize = True
        Me.CheckBoxWMF.Location = New System.Drawing.Point(185, 71)
        Me.CheckBoxWMF.Name = "CheckBoxWMF"
        Me.CheckBoxWMF.Size = New System.Drawing.Size(52, 17)
        Me.CheckBoxWMF.TabIndex = 2
        Me.CheckBoxWMF.Text = "WMF"
        Me.CheckBoxWMF.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButtonComponents)
        Me.GroupBox1.Controls.Add(Me.RadioButtonFinishedGoods)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(761, 53)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'RadioButtonComponents
        '
        Me.RadioButtonComponents.AutoSize = True
        Me.RadioButtonComponents.Location = New System.Drawing.Point(329, 19)
        Me.RadioButtonComponents.Name = "RadioButtonComponents"
        Me.RadioButtonComponents.Size = New System.Drawing.Size(84, 17)
        Me.RadioButtonComponents.TabIndex = 1
        Me.RadioButtonComponents.Text = "Components"
        Me.RadioButtonComponents.UseVisualStyleBackColor = True
        '
        'RadioButtonFinishedGoods
        '
        Me.RadioButtonFinishedGoods.AutoSize = True
        Me.RadioButtonFinishedGoods.Checked = True
        Me.RadioButtonFinishedGoods.Location = New System.Drawing.Point(172, 20)
        Me.RadioButtonFinishedGoods.Name = "RadioButtonFinishedGoods"
        Me.RadioButtonFinishedGoods.Size = New System.Drawing.Size(98, 17)
        Me.RadioButtonFinishedGoods.TabIndex = 0
        Me.RadioButtonFinishedGoods.TabStop = True
        Me.RadioButtonFinishedGoods.Text = "Finished Goods"
        Me.RadioButtonFinishedGoods.UseVisualStyleBackColor = True
        '
        'FormScoreboardNet
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(785, 475)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Name = "FormScoreboardNet"
        Me.Text = "FormScoreboardNet"
        Me.ToolStripContainer1.BottomToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.BottomToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.ContentPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents CheckBoxWMF As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButtonComponents As System.Windows.Forms.RadioButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RadioButtonOnlyWMF As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonOnlySIS As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButtonExcludeSIS As System.Windows.Forms.RadioButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DateTimerPickerFSLFSSLStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerCurrentMonth As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePickerEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Public WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Public WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Public WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Public WithEvents RadioButtonFinishedGoods As System.Windows.Forms.RadioButton
End Class
