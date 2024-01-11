Imports System.Xml

Public Class Main

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление ранее созданных объектов при выходе из приложения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Main_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список предложений данного пользователя
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyDs As New DataSet                       '

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            declarations.ScalaDate = CDate(Scala.ActiveFrame.Parent.ScalaDate)


            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
                MsgBox("Не найден код продавца, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Application.Exit()
            Else
                declarations.SalesmanCode = declarations.MyRec.Fields("SC").Value
                declarations.SalesmanName = declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
            End If
        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---значения
        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        CheckButtonState()
    End Sub

    Private Sub CheckButtonState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка и выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If ComboBox1.SelectedItem = "" Then
            button2.Enabled = False
        Else
            button2.Enabled = True
        End If
        button3.Enabled = False
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        CheckButtonState()
    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие Invoice - файла, определение основных параметров СФ
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        button3.Enabled = False
        progressBar1.Value = 0
        OpenInvoiceFile()
    End Sub

    Private Sub button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка Invoice - файла в Scala
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckDate() = False Then  '---дата в Scala не совпадает с текущей датой на компьютере
            MsgBox("Системная дата в Scala не совпадает с текущей датой на компьютере. Выставьте в Scala текущую дату и только после этого произведите импорт.", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        End If

        UploadInvoiceFile()
    End Sub

    Private Function CheckDate() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка системной даты в Scala - совпадает ли с компьютерной
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Math.Abs(DateDiff(DateInterval.Day, declarations.ScalaDate, Now())) >= 1 Then
            CheckDate = False
        Else
            CheckDate = True
        End If
    End Function
End Class
