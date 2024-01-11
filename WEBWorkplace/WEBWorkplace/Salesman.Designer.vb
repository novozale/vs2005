<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Salesman
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
        Me.Label4 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(180, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(362, 20)
        Me.TextBox1.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 21)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Код продавца"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(180, 38)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(362, 20)
        Me.TextBox2.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(162, 21)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Имя продавца"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(180, 64)
        Me.TextBox3.MaxLength = 50
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(362, 20)
        Me.TextBox3.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 62)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(162, 21)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Электронная почта"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(162, 21)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Город продавца"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(180, 91)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(362, 21)
        Me.ComboBox1.TabIndex = 10
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(1, 119)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox1.Size = New System.Drawing.Size(193, 17)
        Me.CheckBox1.TabIndex = 11
        Me.CheckBox1.Text = "ответственный за WEB в городе"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(56, 142)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox2.Size = New System.Drawing.Size(138, 17)
        Me.CheckBox2.TabIndex = 12
        Me.CheckBox2.Text = """дежурный"" продавец"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(12, 165)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox3.Size = New System.Drawing.Size(183, 17)
        Me.CheckBox3.TabIndex = 13
        Me.CheckBox3.Text = """активный"" (выгрузка на сайт)"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(180, 188)
        Me.TextBox4.MaxLength = 50
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(362, 20)
        Me.TextBox4.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(12, 186)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(162, 21)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Резервное поле1"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(180, 214)
        Me.TextBox5.MaxLength = 50
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(362, 20)
        Me.TextBox5.TabIndex = 17
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(12, 212)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(162, 21)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Резервное поле2"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(14, 241)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(128, 26)
        Me.Button2.TabIndex = 19
        Me.Button2.Text = "Отмена"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(414, 241)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 26)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Сохранить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Salesman
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 272)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Salesman"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Иинформация о продавце"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
