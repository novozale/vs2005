<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MatchPictAndScalaCode
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
        Me.DataGridView12 = New System.Windows.Forms.DataGridView
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.DataGridView12, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView12
        '
        Me.DataGridView12.AllowUserToAddRows = False
        Me.DataGridView12.AllowUserToDeleteRows = False
        Me.DataGridView12.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.DataGridView12.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView12.Location = New System.Drawing.Point(2, 30)
        Me.DataGridView12.MultiSelect = False
        Me.DataGridView12.Name = "DataGridView12"
        Me.DataGridView12.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView12.Size = New System.Drawing.Size(1613, 554)
        Me.DataGridView12.TabIndex = 19
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(2, 590)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(132, 22)
        Me.Button3.TabIndex = 21
        Me.Button3.Text = "Закрыть окно"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(1483, 590)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 22)
        Me.Button2.TabIndex = 20
        Me.Button2.Text = "Выполнить связку"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(271, 2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(290, 22)
        Me.Button4.TabIndex = 28
        Me.Button4.Text = "Выставить все выделения для связывания картинок"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(2, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(263, 22)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "Снять все выделения для связывания картинок"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'MatchPictAndScalaCode
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1619, 613)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DataGridView12)
        Me.Name = "MatchPictAndScalaCode"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Связывание картинки с кодом Scala"
        CType(Me.DataGridView12, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView12 As System.Windows.Forms.DataGridView
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
