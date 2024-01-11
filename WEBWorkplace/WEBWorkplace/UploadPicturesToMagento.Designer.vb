<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UploadPicturesToMagento
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
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 94)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 21
        Me.Button3.Text = "Закрыть"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(365, 94)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 20
        Me.Button2.Text = "Загрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label2.Location = New System.Drawing.Point(7, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 15)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "0"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(235, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(23, 18)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "из"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightGray
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Обработано картинок"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Location = New System.Drawing.Point(264, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(216, 15)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "0"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(174, 12)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox1.TabIndex = 23
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(35, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(135, 24)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Для поставщика:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'UploadPicturesToMagento
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(509, 124)
        Me.ControlBox = False
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UploadPicturesToMagento"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Загрузка картинок на web Magento"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
