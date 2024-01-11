<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SalesComissionStarter
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label5 = New System.Windows.Forms.Label
        Me.MyCatalog = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.MyFinDate = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.MyStartDate = New System.Windows.Forms.DateTimePicker
        Me.MyFormat = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.MyCatalog)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.MyFinDate)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.MyStartDate)
        Me.TabPage1.Controls.Add(Me.MyFormat)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(493, 302)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Формирование отчетов"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Location = New System.Drawing.Point(23, 205)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(444, 53)
        Me.GroupBox2.TabIndex = 80
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Процесс формирования отчетов"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 16)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(431, 25)
        Me.ProgressBar1.Step = 1
        Me.ProgressBar1.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(41, 164)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(162, 13)
        Me.Label5.TabIndex = 78
        Me.Label5.Text = "Каталог для выгрузки отчетов"
        '
        'MyCatalog
        '
        Me.MyCatalog.Location = New System.Drawing.Point(209, 161)
        Me.MyCatalog.Name = "MyCatalog"
        Me.MyCatalog.Size = New System.Drawing.Size(223, 20)
        Me.MyCatalog.TabIndex = 77
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(70, 121)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 13)
        Me.Label4.TabIndex = 76
        Me.Label4.Text = "Формат выгрузки отчета"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(78, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(125, 13)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Дата окончания отчета"
        '
        'MyFinDate
        '
        Me.MyFinDate.Location = New System.Drawing.Point(209, 82)
        Me.MyFinDate.Name = "MyFinDate"
        Me.MyFinDate.Size = New System.Drawing.Size(239, 20)
        Me.MyFinDate.TabIndex = 72
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(96, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 13)
        Me.Label1.TabIndex = 71
        Me.Label1.Text = "Дата начала отчета"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(210, 270)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(239, 23)
        Me.Button1.TabIndex = 70
        Me.Button1.Text = "Сформировать"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'MyStartDate
        '
        Me.MyStartDate.Location = New System.Drawing.Point(209, 46)
        Me.MyStartDate.Name = "MyStartDate"
        Me.MyStartDate.Size = New System.Drawing.Size(239, 20)
        Me.MyStartDate.TabIndex = 69
        '
        'MyFormat
        '
        Me.MyFormat.FormattingEnabled = True
        Me.MyFormat.Items.AddRange(New Object() {"EXCEL", "MHTML", "IMAGE", "PDF"})
        Me.MyFormat.Location = New System.Drawing.Point(209, 118)
        Me.MyFormat.Name = "MyFormat"
        Me.MyFormat.Size = New System.Drawing.Size(239, 21)
        Me.MyFormat.TabIndex = 68
        Me.MyFormat.Text = "EXCEL"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Location = New System.Drawing.Point(23, 15)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(445, 180)
        Me.GroupBox1.TabIndex = 79
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Параметры отчетов"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(408, 146)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(18, 20)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = ">"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(2, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(501, 328)
        Me.TabControl1.TabIndex = 68
        '
        'SalesComissionStarter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(499, 330)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "SalesComissionStarter"
        Me.Text = "Запуск отчетов по расчету комиссии продавцов"
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents MyCatalog As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MyFinDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MyStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents MyFormat As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl

End Class
