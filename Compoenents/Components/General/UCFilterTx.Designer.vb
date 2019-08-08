<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UCFilterTx
    Inherits Components.UCToolstrip

    'UserControl overrides dispose to clean up the component list.
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
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FilterTextBox = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.OperatorComboBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FieldComboBox = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.CheckBox1)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.FilterTextBox)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.OperatorComboBox)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.FieldComboBox)
        Me.Panel1.Size = New System.Drawing.Size(220, 228)
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(8, 122)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(70, 17)
        Me.CheckBox1.TabIndex = 45
        Me.CheckBox1.Text = "Join Filter"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(5, 197)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(212, 23)
        Me.Button2.TabIndex = 44
        Me.Button2.Text = "Reset Filter"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(4, 168)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(213, 23)
        Me.Button1.TabIndex = 43
        Me.Button1.Text = "Execute"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Enabled = False
        Me.RadioButton2.Location = New System.Drawing.Point(77, 145)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(41, 17)
        Me.RadioButton2.TabIndex = 42
        Me.RadioButton2.Text = "OR"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Enabled = False
        Me.RadioButton1.Location = New System.Drawing.Point(27, 145)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(48, 17)
        Me.RadioButton1.TabIndex = 41
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "AND"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 40
        Me.Label3.Text = "Filter Text"
        '
        'FilterTextBox
        '
        Me.FilterTextBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.FilterTextBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource
        Me.FilterTextBox.Location = New System.Drawing.Point(8, 96)
        Me.FilterTextBox.Name = "FilterTextBox"
        Me.FilterTextBox.Size = New System.Drawing.Size(209, 20)
        Me.FilterTextBox.TabIndex = 39
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 38
        Me.Label2.Text = "Operator"
        '
        'OperatorComboBox
        '
        Me.OperatorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.OperatorComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.OperatorComboBox.FormattingEnabled = True
        Me.OperatorComboBox.Location = New System.Drawing.Point(8, 56)
        Me.OperatorComboBox.Name = "OperatorComboBox"
        Me.OperatorComboBox.Size = New System.Drawing.Size(209, 21)
        Me.OperatorComboBox.TabIndex = 37
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(54, 13)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Filter Field"
        '
        'FieldComboBox
        '
        Me.FieldComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.FieldComboBox.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FieldComboBox.FormattingEnabled = True
        Me.FieldComboBox.Location = New System.Drawing.Point(8, 16)
        Me.FieldComboBox.Name = "FieldComboBox"
        Me.FieldComboBox.Size = New System.Drawing.Size(209, 21)
        Me.FieldComboBox.TabIndex = 35
        '
        'UCFilterTx
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Name = "UCFilterTx"
        Me.Size = New System.Drawing.Size(225, 258)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents FilterTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OperatorComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents FieldComboBox As System.Windows.Forms.ComboBox

End Class
