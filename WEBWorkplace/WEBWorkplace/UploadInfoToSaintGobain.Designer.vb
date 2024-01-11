<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UploadInfoToSaintGobain
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button1.Location = New System.Drawing.Point(465, 37)
        Me.Button1.Margin = New System.Windows.Forms.Padding(0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(23, 24)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = ">"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(187, 38)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(274, 20)
        Me.TextBox1.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(3, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(178, 24)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "каталог для выгрузки картинок"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Все товары", "Все товары с картинками", "Все товары с картинками и описаниями", "Согласованный ассортимент", "Согласованный ассортимент с картинками", "Согласованный ассортимент с картинками и описаниями"})
        Me.ComboBox1.Location = New System.Drawing.Point(187, 9)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(301, 21)
        Me.ComboBox1.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(171, 24)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Выгружаемый ассортимент"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 92)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(486, 41)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Выгружено картинок"
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
        Me.Button3.Location = New System.Drawing.Point(1, 140)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 13
        Me.Button3.Text = "Закрыть"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(355, 140)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Выгрузить"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'UploadInfoToSaintGobain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 165)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "UploadInfoToSaintGobain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Выгрузка информации для Сен Гобен"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
End Class
