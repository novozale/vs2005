Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Drawing.Printing

Public Class ErrorForm
    Public MyMsg As String

    Private PrtSetupDB As New PrintDialog
    Private WithEvents PrtDocument As New System.Drawing.Printing.PrintDocument
    Private PageSetupDB As New PageSetupDialog
    Private PrintPreviewDB As New PrintPreviewDialog
    Private PrinterSettings As New System.Drawing.Printing.PrinterSettings

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Печать текста сообщения
        '// 
        '////////////////////////////////////////////////////////////////////////////////////

        If PrtSetupDB.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrtDocument.Print()
        End If
    End Sub

    Private Sub ErrorForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При загрузке окна вывод сообщения об ошибке
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        RichTextBox1.Text = MyMsg

        PageSetupDB.MinMargins.Left = 30
        PageSetupDB.MinMargins.Right = 30
        PageSetupDB.MinMargins.Top = 30
        PageSetupDB.MinMargins.Bottom = 30
        PrtSetupDB.Document = PrtDocument
        PrtSetupDB.PrinterSettings = PrinterSettings
        PrtSetupDB.AllowSelection = True
        PrtSetupDB.AllowSomePages = True
        PrtSetupDB.UseEXDialog = True
    End Sub

    Private Sub PrtDocument_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs) Handles PrtDocument.PrintPage
        ev.Graphics.DrawString(RichTextBox1.Text, RichTextBox1.Font, Brushes.Black, 0, 0)
    End Sub
End Class