<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UploadFilesToCatalog
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboBox4 = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.ComboBox5 = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(42, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(135, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Группы товаров"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(42, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(135, 24)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Поставщики"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(42, 175)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(135, 24)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "каталог для выгрузки"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(466, 178)
        Me.Button1.Margin = New System.Windows.Forms.Padding(0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(23, 24)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = ">"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(179, 179)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(274, 20)
        Me.TextBox1.TabIndex = 4
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(181, 11)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox1.TabIndex = 6
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(181, 36)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox2.TabIndex = 7
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 204)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox1.TabIndex = 8
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
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label5.Location = New System.Drawing.Point(264, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(216, 15)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "0"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Location = New System.Drawing.Point(7, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(216, 15)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "0"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(4, 252)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "Отмена"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(358, 252)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "Выгрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"Коды Scala", "Коды товара поставщика"})
        Me.ComboBox3.Location = New System.Drawing.Point(181, 62)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox3.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(42, 59)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(135, 24)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Коды выгрузки"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox4
        '
        Me.ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Items.AddRange(New Object() {"Все товары", "Только товары из прайс листа"})
        Me.ComboBox4.Location = New System.Drawing.Point(181, 88)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox4.TabIndex = 14
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(42, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(135, 24)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Что выгружается"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(250, 140)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox1.Size = New System.Drawing.Size(239, 17)
        Me.CheckBox1.TabIndex = 26
        Me.CheckBox1.Text = "Создавать подкаталоги для поставщиков"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'ComboBox5
        '
        Me.ComboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox5.FormattingEnabled = True
        Me.ComboBox5.Items.AddRange(New Object() {"Поставщики -> группы товаров", "Группы товаров -> поставщики"})
        Me.ComboBox5.Location = New System.Drawing.Point(181, 114)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(308, 21)
        Me.ComboBox5.TabIndex = 28
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(42, 111)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(135, 24)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Группировка"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(246, 161)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox2.Size = New System.Drawing.Size(243, 17)
        Me.CheckBox2.TabIndex = 29
        Me.CheckBox2.Text = "Создавать подкаталоги для групп товаров"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'UploadFilesToCatalog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(495, 278)
        Me.ControlBox = False
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.ComboBox5)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "UploadFilesToCatalog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Выгрузка картинок из БД"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox5 As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
End Class
