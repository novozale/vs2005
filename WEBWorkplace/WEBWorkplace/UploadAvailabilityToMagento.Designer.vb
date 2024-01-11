<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UploadAvailabilityToMagento
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.LightGray
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Обработано товаров"
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
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 67)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 26
        Me.Button3.Text = "Закрыть"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(365, 67)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 25
        Me.Button2.Text = "Загрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'UploadAvailabilityToMagento
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(511, 97)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "UploadAvailabilityToMagento"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Загрузка информации о доступности товаров на WEB Magento"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
