<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ProductSubGroup
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
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(173, 12)
        Me.TextBox1.MaxLength = 100
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(393, 20)
        Me.TextBox1.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(2, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(170, 21)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Код группы продуктов"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(173, 38)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(393, 20)
        Me.TextBox2.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(2, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(170, 21)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Код подгруппы продуктов"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(173, 64)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(393, 20)
        Me.TextBox3.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(2, 62)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(170, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Название подгруппы продуктов"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(173, 90)
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(393, 141)
        Me.TextBox4.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(2, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(170, 21)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Описание подгруппы продуктов"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(173, 240)
        Me.TextBox5.MaxLength = 100
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(393, 20)
        Me.TextBox5.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(2, 238)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(170, 21)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Резервное поле"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(7, 267)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(128, 26)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(437, 268)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 26)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ProductSubGroup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(570, 299)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "ProductSubGroup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Информация о подгруппе продуктов"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
