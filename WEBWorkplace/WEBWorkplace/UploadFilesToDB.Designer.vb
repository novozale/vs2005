<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UploadFilesToDB
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboBox4 = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ComboBox1)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.ComboBox2)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.GroupBox1)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(509, 195)
        Me.GroupBox2.TabIndex = 20
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Загрузка по поставщикам"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(189, 52)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox1.TabIndex = 29
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(11, 49)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(174, 24)
        Me.Label6.TabIndex = 28
        Me.Label6.Text = "Для поставщика:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Не перезаписывать", "Перезаписать несвязанные с кодом Scala", "Перезаписать все"})
        Me.ComboBox2.Location = New System.Drawing.Point(190, 79)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(306, 21)
        Me.ComboBox2.TabIndex = 27
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(15, 77)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(171, 24)
        Me.Label5.TabIndex = 26
        Me.Label5.Text = "Перезапись картинок в БД"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 110)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Обработано"
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
        Me.Button3.Location = New System.Drawing.Point(11, 162)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 24
        Me.Button3.Text = "Отмена"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(365, 162)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 23
        Me.Button2.Text = "Загрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(475, 25)
        Me.Button1.Margin = New System.Windows.Forms.Padding(0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(23, 24)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = ">"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(190, 26)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(274, 20)
        Me.TextBox1.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(10, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(170, 24)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "каталог с картинками"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ComboBox3)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.ComboBox4)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.Controls.Add(Me.Button5)
        Me.GroupBox3.Controls.Add(Me.Button6)
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(12, 224)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(509, 195)
        Me.GroupBox3.TabIndex = 21
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Загрузка по производителям"
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(189, 52)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox3.TabIndex = 29
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(13, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(172, 24)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "Для производителя:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox4
        '
        Me.ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Items.AddRange(New Object() {"Не перезаписывать", "Перезаписать несвязанные с кодом Scala", "Перезаписать все"})
        Me.ComboBox4.Location = New System.Drawing.Point(190, 79)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(306, 21)
        Me.ComboBox4.TabIndex = 27
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(15, 77)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(171, 24)
        Me.Label8.TabIndex = 26
        Me.Label8.Text = "Перезапись картинок в БД"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Location = New System.Drawing.Point(10, 110)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox4.TabIndex = 25
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Обработано"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(235, 13)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(23, 18)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "из"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.Location = New System.Drawing.Point(264, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(216, 15)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "0"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label11.Location = New System.Drawing.Point(7, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(216, 15)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "0"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(11, 162)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(132, 22)
        Me.Button4.TabIndex = 24
        Me.Button4.Text = "Отмена"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(365, 162)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(132, 22)
        Me.Button5.TabIndex = 23
        Me.Button5.Text = "Загрузить"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button6.Location = New System.Drawing.Point(475, 25)
        Me.Button6.Margin = New System.Windows.Forms.Padding(0)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(23, 24)
        Me.Button6.TabIndex = 22
        Me.Button6.Text = ">"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button6.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(190, 26)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(274, 20)
        Me.TextBox2.TabIndex = 21
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(11, 23)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(169, 24)
        Me.Label12.TabIndex = 20
        Me.Label12.Text = "каталог с картинками"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'UploadFilesToDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(527, 426)
        Me.ControlBox = False
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "UploadFilesToDB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Загрузка картинок в БД"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
End Class
