<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Configuration
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
        Me.ComboBoxEMail = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ComboBoxContact = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ComboBoxEMail
        '
        Me.ComboBoxEMail.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxEMail.FormattingEnabled = True
        Me.ComboBoxEMail.Items.AddRange(New Object() {"Вручную", "Выбранный из CRM контакт"})
        Me.ComboBoxEMail.Location = New System.Drawing.Point(288, 9)
        Me.ComboBoxEMail.Name = "ComboBoxEMail"
        Me.ComboBoxEMail.Size = New System.Drawing.Size(209, 21)
        Me.ComboBoxEMail.TabIndex = 41
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(270, 19)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "Откуда брать Email для уведомления"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ComboBoxContact
        '
        Me.ComboBoxContact.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxContact.FormattingEnabled = True
        Me.ComboBoxContact.Items.AddRange(New Object() {"Вручную", "Из Scala"})
        Me.ComboBoxContact.Location = New System.Drawing.Point(288, 36)
        Me.ComboBoxContact.Name = "ComboBoxContact"
        Me.ComboBoxContact.Size = New System.Drawing.Size(209, 21)
        Me.ComboBoxContact.TabIndex = 43
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(270, 19)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "Заполнение поля контакт"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(13, 74)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(149, 29)
        Me.Button1.TabIndex = 44
        Me.Button1.Text = "Отмена"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button2.Location = New System.Drawing.Point(348, 74)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(149, 29)
        Me.Button2.TabIndex = 45
        Me.Button2.Text = "Записать"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Configuration
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 111)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ComboBoxContact)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBoxEMail)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Configuration"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Конфигурация"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ComboBoxEMail As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBoxContact As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
