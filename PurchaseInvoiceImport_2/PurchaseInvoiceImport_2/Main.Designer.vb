<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
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
        Me.label5 = New System.Windows.Forms.Label
        Me.label4 = New System.Windows.Forms.Label
        Me.label3 = New System.Windows.Forms.Label
        Me.progressBar1 = New System.Windows.Forms.ProgressBar
        Me.label6 = New System.Windows.Forms.Label
        Me.button2 = New System.Windows.Forms.Button
        Me.groupBox2 = New System.Windows.Forms.GroupBox
        Me.textBox5 = New System.Windows.Forms.TextBox
        Me.button3 = New System.Windows.Forms.Button
        Me.textBox4 = New System.Windows.Forms.TextBox
        Me.textBox3 = New System.Windows.Forms.TextBox
        Me.groupBox1 = New System.Windows.Forms.GroupBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.button1 = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.groupBox2.SuspendLayout()
        Me.groupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'label5
        '
        Me.label5.Location = New System.Drawing.Point(8, 95)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(124, 18)
        Me.label5.TabIndex = 4
        Me.label5.Text = "Валюта СФ"
        Me.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'label4
        '
        Me.label4.Location = New System.Drawing.Point(8, 41)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(124, 18)
        Me.label4.TabIndex = 3
        Me.label4.Text = "N СФ поставщика"
        Me.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'label3
        '
        Me.label3.Location = New System.Drawing.Point(6, 67)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(124, 18)
        Me.label3.TabIndex = 2
        Me.label3.Text = "Дата СФ"
        Me.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'progressBar1
        '
        Me.progressBar1.Location = New System.Drawing.Point(9, 16)
        Me.progressBar1.Name = "progressBar1"
        Me.progressBar1.Size = New System.Drawing.Size(453, 31)
        Me.progressBar1.TabIndex = 0
        '
        'label6
        '
        Me.label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.label6.ForeColor = System.Drawing.Color.Red
        Me.label6.Location = New System.Drawing.Point(8, 122)
        Me.label6.Name = "label6"
        Me.label6.Size = New System.Drawing.Size(451, 20)
        Me.label6.TabIndex = 10
        Me.label6.Text = "СФ уже загружена в Scala"
        '
        'button2
        '
        Me.button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.button2.Location = New System.Drawing.Point(7, 251)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(166, 28)
        Me.button2.TabIndex = 13
        Me.button2.Text = "Открыть файл с СФ"
        Me.button2.UseVisualStyleBackColor = True
        '
        'groupBox2
        '
        Me.groupBox2.Controls.Add(Me.progressBar1)
        Me.groupBox2.Location = New System.Drawing.Point(7, 190)
        Me.groupBox2.Name = "groupBox2"
        Me.groupBox2.Size = New System.Drawing.Size(469, 55)
        Me.groupBox2.TabIndex = 11
        Me.groupBox2.TabStop = False
        Me.groupBox2.Text = "Процесс загрузки СФ"
        '
        'textBox5
        '
        Me.textBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.textBox5.Location = New System.Drawing.Point(138, 95)
        Me.textBox5.Name = "textBox5"
        Me.textBox5.ReadOnly = True
        Me.textBox5.Size = New System.Drawing.Size(321, 20)
        Me.textBox5.TabIndex = 9
        '
        'button3
        '
        Me.button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.button3.ForeColor = System.Drawing.Color.Green
        Me.button3.Location = New System.Drawing.Point(181, 251)
        Me.button3.Name = "button3"
        Me.button3.Size = New System.Drawing.Size(153, 28)
        Me.button3.TabIndex = 14
        Me.button3.Text = "Загрузить СФ в Scala"
        Me.button3.UseVisualStyleBackColor = True
        '
        'textBox4
        '
        Me.textBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.textBox4.Location = New System.Drawing.Point(138, 67)
        Me.textBox4.Name = "textBox4"
        Me.textBox4.ReadOnly = True
        Me.textBox4.Size = New System.Drawing.Size(321, 20)
        Me.textBox4.TabIndex = 8
        '
        'textBox3
        '
        Me.textBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.textBox3.Location = New System.Drawing.Point(138, 41)
        Me.textBox3.Name = "textBox3"
        Me.textBox3.ReadOnly = True
        Me.textBox3.Size = New System.Drawing.Size(321, 20)
        Me.textBox3.TabIndex = 7
        '
        'groupBox1
        '
        Me.groupBox1.Controls.Add(Me.TextBox1)
        Me.groupBox1.Controls.Add(Me.Label1)
        Me.groupBox1.Controls.Add(Me.label6)
        Me.groupBox1.Controls.Add(Me.textBox5)
        Me.groupBox1.Controls.Add(Me.textBox4)
        Me.groupBox1.Controls.Add(Me.textBox3)
        Me.groupBox1.Controls.Add(Me.label5)
        Me.groupBox1.Controls.Add(Me.label4)
        Me.groupBox1.Controls.Add(Me.label3)
        Me.groupBox1.Location = New System.Drawing.Point(7, 38)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(469, 146)
        Me.groupBox1.TabIndex = 10
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Информация по СФ"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(138, 14)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(321, 20)
        Me.TextBox1.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(7, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 18)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Код поставщика"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.button1.Location = New System.Drawing.Point(361, 252)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(113, 28)
        Me.button1.TabIndex = 12
        Me.button1.Text = "Выход"
        Me.button1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(15, 7)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(124, 18)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Поставщик"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"АББ", "OBO Betterman"})
        Me.ComboBox1.Location = New System.Drawing.Point(144, 8)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(330, 21)
        Me.ComboBox1.TabIndex = 16
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(482, 285)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.groupBox2)
        Me.Controls.Add(Me.button3)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Импорт СФ на закупку"
        Me.groupBox2.ResumeLayout(False)
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents label5 As System.Windows.Forms.Label
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents label3 As System.Windows.Forms.Label
    Public WithEvents progressBar1 As System.Windows.Forms.ProgressBar
    Public WithEvents label6 As System.Windows.Forms.Label
    Private WithEvents button2 As System.Windows.Forms.Button
    Private WithEvents groupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents textBox5 As System.Windows.Forms.TextBox
    Public WithEvents button3 As System.Windows.Forms.Button
    Public WithEvents textBox4 As System.Windows.Forms.TextBox
    Public WithEvents textBox3 As System.Windows.Forms.TextBox
    Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents button1 As System.Windows.Forms.Button
    Private WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Public WithEvents TextBox1 As System.Windows.Forms.TextBox
    Private WithEvents Label1 As System.Windows.Forms.Label

End Class
