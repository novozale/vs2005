<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProductGroup
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
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(8, 83)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(128, 26)
        Me.Button2.TabIndex = 22
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(405, 86)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 26)
        Me.Button1.TabIndex = 21
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(219, 60)
        Me.TextBox3.MaxLength = 200
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(314, 20)
        Me.TextBox3.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(5, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(212, 21)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Название  группы продуктов для WEB"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(219, 34)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(314, 20)
        Me.TextBox2.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(5, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(208, 21)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Название  группы продуктов"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(219, 8)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(314, 20)
        Me.TextBox1.TabIndex = 14
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(5, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 21)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Код группы продуктов"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ProductGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(541, 114)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "ProductGroup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Информация о группах продуктов"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
