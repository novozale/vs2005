<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Shipment
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
        Me.LblCustomerCode = New System.Windows.Forms.Label
        Me.LblCustomerINN = New System.Windows.Forms.Label
        Me.LblSalesmanCode = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.LblCustomerName = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LblCustomerLegalAddress = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.LblWHCode = New System.Windows.Forms.Label
        Me.LblSalesmanName = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'LblCustomerCode
        '
        Me.LblCustomerCode.BackColor = System.Drawing.Color.White
        Me.LblCustomerCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblCustomerCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblCustomerCode.Location = New System.Drawing.Point(6, 16)
        Me.LblCustomerCode.Name = "LblCustomerCode"
        Me.LblCustomerCode.Size = New System.Drawing.Size(130, 18)
        Me.LblCustomerCode.TabIndex = 1
        Me.LblCustomerCode.Text = "00"
        '
        'LblCustomerINN
        '
        Me.LblCustomerINN.BackColor = System.Drawing.Color.White
        Me.LblCustomerINN.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblCustomerINN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblCustomerINN.Location = New System.Drawing.Point(142, 41)
        Me.LblCustomerINN.Name = "LblCustomerINN"
        Me.LblCustomerINN.Size = New System.Drawing.Size(519, 18)
        Me.LblCustomerINN.TabIndex = 3
        Me.LblCustomerINN.Text = "ИНН"
        '
        'LblSalesmanCode
        '
        Me.LblSalesmanCode.BackColor = System.Drawing.Color.White
        Me.LblSalesmanCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblSalesmanCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblSalesmanCode.Location = New System.Drawing.Point(142, 16)
        Me.LblSalesmanCode.Name = "LblSalesmanCode"
        Me.LblSalesmanCode.Size = New System.Drawing.Size(122, 18)
        Me.LblSalesmanCode.TabIndex = 7
        Me.LblSalesmanCode.Text = "S01"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(36, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ИНН"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblCustomerName
        '
        Me.LblCustomerName.BackColor = System.Drawing.Color.White
        Me.LblCustomerName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblCustomerName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblCustomerName.Location = New System.Drawing.Point(142, 16)
        Me.LblCustomerName.Name = "LblCustomerName"
        Me.LblCustomerName.Size = New System.Drawing.Size(519, 18)
        Me.LblCustomerName.TabIndex = 2
        Me.LblCustomerName.Text = "Рога и копыта"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LblCustomerLegalAddress)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.LblCustomerINN)
        Me.GroupBox1.Controls.Add(Me.LblCustomerName)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.LblCustomerCode)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(671, 120)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Покупатель"
        '
        'LblCustomerLegalAddress
        '
        Me.LblCustomerLegalAddress.BackColor = System.Drawing.Color.White
        Me.LblCustomerLegalAddress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblCustomerLegalAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblCustomerLegalAddress.Location = New System.Drawing.Point(142, 68)
        Me.LblCustomerLegalAddress.Name = "LblCustomerLegalAddress"
        Me.LblCustomerLegalAddress.Size = New System.Drawing.Size(519, 41)
        Me.LblCustomerLegalAddress.TabIndex = 5
        Me.LblCustomerLegalAddress.Text = "Юридический адрес"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(130, 18)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Юридический адрес"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LblSalesmanCode)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.LblWHCode)
        Me.GroupBox2.Controls.Add(Me.LblSalesmanName)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 136)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(671, 70)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Дополнительная информация"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label12.Location = New System.Drawing.Point(36, 16)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 18)
        Me.Label12.TabIndex = 6
        Me.Label12.Text = "Продавец"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblWHCode
        '
        Me.LblWHCode.BackColor = System.Drawing.Color.White
        Me.LblWHCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblWHCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblWHCode.Location = New System.Drawing.Point(142, 41)
        Me.LblWHCode.Name = "LblWHCode"
        Me.LblWHCode.Size = New System.Drawing.Size(519, 18)
        Me.LblWHCode.TabIndex = 3
        Me.LblWHCode.Text = "01"
        '
        'LblSalesmanName
        '
        Me.LblSalesmanName.BackColor = System.Drawing.Color.White
        Me.LblSalesmanName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblSalesmanName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.LblSalesmanName.Location = New System.Drawing.Point(270, 16)
        Me.LblSalesmanName.Name = "LblSalesmanName"
        Me.LblSalesmanName.Size = New System.Drawing.Size(391, 18)
        Me.LblSalesmanName.TabIndex = 2
        Me.LblSalesmanName.Text = "Иванов"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label11.Location = New System.Drawing.Point(36, 39)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 18)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Склад"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox6)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.TextBox5)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.CheckBox3)
        Me.GroupBox3.Controls.Add(Me.CheckBox2)
        Me.GroupBox3.Controls.Add(Me.Button6)
        Me.GroupBox3.Controls.Add(Me.Button5)
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.TextBox4)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.CheckBox1)
        Me.GroupBox3.Controls.Add(Me.DateTimePicker1)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.TextBox3)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Controls.Add(Me.ComboBox1)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label16)
        Me.GroupBox3.Location = New System.Drawing.Point(9, 221)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(671, 310)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Информация по отгрузке"
        '
        'CheckBox3
        '
        Me.CheckBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(474, 224)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox3.Size = New System.Drawing.Size(181, 33)
        Me.CheckBox3.TabIndex = 23
        Me.CheckBox3.Text = "Печатать полный счет" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(восстановленный)"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(250, 224)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox2.Size = New System.Drawing.Size(181, 33)
        Me.CheckBox2.TabIndex = 22
        Me.CheckBox2.Text = "Печатать справку - счет " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "к Счет Фактуре"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button6.Location = New System.Drawing.Point(613, 280)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(45, 19)
        Me.Button6.TabIndex = 21
        Me.Button6.Text = "X"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Button5.Location = New System.Drawing.Point(613, 262)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(45, 19)
        Me.Button5.TabIndex = 20
        Me.Button5.Text = ">>>"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(613, 82)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(45, 36)
        Me.Button4.TabIndex = 19
        Me.Button4.Text = "ERP>"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(614, 41)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(45, 36)
        Me.Button3.TabIndex = 18
        Me.Button3.Text = "CRM>"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(177, 261)
        Me.TextBox4.MaxLength = 255
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(432, 38)
        Me.TextBox4.TabIndex = 17
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 260)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 39)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Путь к присоединяемому файлу"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CheckBox1
        '
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(25, 224)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.CheckBox1.Size = New System.Drawing.Size(166, 33)
        Me.CheckBox1.TabIndex = 15
        Me.CheckBox1.Text = "Печатать счет " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "к Счет Фактуре"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(179, 202)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(480, 20)
        Me.DateTimePicker1.TabIndex = 14
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label15.Location = New System.Drawing.Point(9, 204)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(161, 18)
        Me.Label15.TabIndex = 13
        Me.Label15.Text = "Дата отгрузки"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(221, 124)
        Me.TextBox3.MaxLength = 255
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(434, 17)
        Me.TextBox3.TabIndex = 12
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label14.Location = New System.Drawing.Point(8, 123)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(207, 18)
        Me.Label14.TabIndex = 11
        Me.Label14.Text = "Комментарий для склада"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(178, 82)
        Me.TextBox2.MaxLength = 255
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(430, 38)
        Me.TextBox2.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 81)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(161, 18)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Адрес доставки"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(179, 40)
        Me.TextBox1.MaxLength = 255
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(430, 38)
        Me.TextBox1.TabIndex = 8
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Доставка", "Самовывоз", "Доставка с оплатой клиентом"})
        Me.ComboBox1.Location = New System.Drawing.Point(179, 15)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(482, 21)
        Me.ComboBox1.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label8.Location = New System.Drawing.Point(9, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(164, 18)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Доставка или самовывоз"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label16.Location = New System.Drawing.Point(9, 39)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(161, 18)
        Me.Label16.TabIndex = 0
        Me.Label16.Text = "Контактная информация"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(9, 537)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(139, 29)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Отмена"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(541, 537)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(139, 29)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Записать"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(221, 146)
        Me.TextBox5.MaxLength = 255
        Me.TextBox5.Multiline = True
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(434, 17)
        Me.TextBox5.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(0, 145)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(219, 18)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Комментарий по транспорту"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(221, 168)
        Me.TextBox6.MaxLength = 255
        Me.TextBox6.Multiline = True
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(434, 17)
        Me.TextBox6.TabIndex = 27
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label4.Location = New System.Drawing.Point(9, 167)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(207, 18)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Комментарий по документам"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Shipment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 571)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Shipment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Отгрузка"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LblCustomerName As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LblCustomerLegalAddress As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents LblWHCode As System.Windows.Forms.Label
    Friend WithEvents LblSalesmanName As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents LblCustomerCode As System.Windows.Forms.Label
    Friend WithEvents LblCustomerINN As System.Windows.Forms.Label
    Friend WithEvents LblSalesmanCode As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
