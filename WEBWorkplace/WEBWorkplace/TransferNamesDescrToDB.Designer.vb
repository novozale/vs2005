<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TransferNamesDescrToDB
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
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Все товары", "Все товары с картинками", "Все товары с картинками и описаниями", "Согласованный ассортимент", "Согласованный ассортимент с картинками", "Согласованный ассортимент с картинками и описаниями"})
        Me.ComboBox1.Location = New System.Drawing.Point(209, 12)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(301, 21)
        Me.ComboBox1.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(10, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(193, 29)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Компания - поставщик названий и описаний"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Названия", "Описания"})
        Me.ComboBox2.Location = New System.Drawing.Point(209, 39)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(301, 21)
        Me.ComboBox2.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(10, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(193, 29)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Что переносится"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(31, 67)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox1.Size = New System.Drawing.Size(191, 17)
        Me.CheckBox1.TabIndex = 25
        Me.CheckBox1.Text = "перезаписывать существующие"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(13, 90)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 27
        Me.Button3.Text = "Отмена"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(367, 90)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 26
        Me.Button2.Text = "Загрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TransferNamesDescrToDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(513, 119)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "TransferNamesDescrToDB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Перенос данных из WEB в БД"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
