Public Class MainForm
    Public MyLoadFlag = 1

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список параметров 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As DataSet

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode

        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---Годы----------------------------
        MySQLStr = "SELECT CASE WHEN RIGHT(name, 2) < '50' THEN '20' + RIGHT(name, 2) ELSE '19' + RIGHT(name, 2) END AS MyYear "
        MySQLStr = MySQLStr & "FROM sys.sysobjects  WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (xtype = 'U') AND (name LIKE N'GL0603%') "
        MySQLStr = MySQLStr & "AND (CASE WHEN RIGHT(name, 2) < '50' THEN '20' + RIGHT(name, 2) ELSE '19' + RIGHT(name, 2) END > 2006) "
        MySQLStr = MySQLStr & "ORDER BY MyYear "
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyDs = New DataSet
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "MyYear" 'Это то что будет отображаться
            ComboBox1.ValueMember = "MyYear"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---Продавцы-------------------------
        MySQLStr = "SELECT ST01001 AS Code, ST01001 + ' ' + ST01002 AS Name "
        MySQLStr = MySQLStr & "FROM ST010300  WITH(NOLOCK)"
        MySQLStr = MySQLStr & "ORDER BY ST01002 "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyDs = New DataSet
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox2.ValueMember = "Code"   'это то что будет храниться
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        MyLoadFlag = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка данных в Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            UploadToLO()
        Else
            UploadToExcel()
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена вида бюджетирования
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ChangeCombobox2Data()
    End Sub

    Private Sub ChangeCombobox2Data()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена вида бюджетирования - смена данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As DataSet

        If MyLoadFlag = 0 Then
            If RadioButton1.Checked = True Then     '---Бюджетирование по продавцам
                Label1.Text = "Продавец"
                MySQLStr = "SELECT ST01001 AS Code, ST01001 + ' ' + ST01002 AS Name "
                MySQLStr = MySQLStr & "FROM ST010300  WITH(NOLOCK)"
                MySQLStr = MySQLStr & "ORDER BY ST01002 "
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyDs = New DataSet
                    MyAdapter.Fill(MyDs)
                    ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
                    ComboBox2.ValueMember = "Code"   'это то что будет храниться
                    ComboBox2.DataSource = MyDs.Tables(0).DefaultView
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Else                                    '---Бюджетирование по кост центрам
                Label1.Text = "Кост центр"
                MySQLStr = "SELECT DISTINCT View_5.GL03002 AS Code, View_5.GL03002 + ' ' + View_5.GL03003 AS Name "
                MySQLStr = MySQLStr & "FROM ST010300 INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT GL03002, GL03003 "
                MySQLStr = MySQLStr & "FROM GL0303" & Microsoft.VisualBasic.Right(ComboBox1.SelectedValue.ToString, 2) & " "
                MySQLStr = MySQLStr & "WHERE (GL03001 = N'B')) AS View_5 ON SUBSTRING(ST010300.ST01021, 7, 3) = View_5.GL03002 "
                MySQLStr = MySQLStr & "Order By View_5.GL03002 "
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyDs = New DataSet
                    MyAdapter.Fill(MyDs)
                    ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
                    ComboBox2.ValueMember = "Code"   'это то что будет храниться
                    ComboBox2.DataSource = MyDs.Tables(0).DefaultView
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If
        End If
    End Sub

    Private Sub ComboBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Validated
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена года бюджетирования - смена данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        ChangeCombobox2Data()
    End Sub
End Class
