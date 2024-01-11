Public Class Form1

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список предложений данного пользователя
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "Veselkov"

            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Не найден код продавца, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Application.Exit()
            Else
                Declarations.SalesmanCode = Declarations.MyRec.Fields("SC").Value
                Declarations.SalesmanName = Declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
                Label1.Text = "Список предложений, сформированных продавцом " & Declarations.SalesmanCode & " " & Declarations.SalesmanName
            End If
        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '----Проверяем, чтобы был продавец, соответствующий пользователю.
        Try

        Catch ex As Exception
            MsgBox("В Scala нет продавца, соответствующего пользователю ", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---Является ли пользователь членом группы CRMManaregs или CRMDirector или 
        CheckRights(Declarations.UserCode, "CRMManagers")
        CheckRights1(Declarations.UserCode, "CRMDirector")
        CheckRights2(Declarations.UserCode, "ProposalManager")


        '---Передача действий другому продавцу
        If Declarations.MyCCPermission = True Or Declarations.MyPermission = True Or Declarations.MyCPPermission = True Then
            Button7.Enabled = True
        Else
            Button7.Enabled = False
        End If



        DateTimePicker1.Value = DateAdd(DateInterval.Quarter, -2, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
        DateTimePicker2.Value = CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))
        '---Вывод данных в окно
        'MySQLStr = "SELECT View_1.OR01001, "
        'MySQLStr = MySQLStr & "View_1.OR01015, "
        'MySQLStr = MySQLStr & "View_1.ExpirationDate AS ExpirationDate, "
        'MySQLStr = MySQLStr & "View_1.OR01003, "
        'MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, N''))) = '' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, N'') END AS CName, "
        'MySQLStr = MySQLStr & "ISNULL(View_1.AgentName,'') AS AgentName,  "
        'MySQLStr = MySQLStr & "View_1.OrderN "
        'MySQLStr = MySQLStr & "FROM SL010300 WITH (NOLOCK) RIGHT OUTER JOIN "
        'MySQLStr = MySQLStr & "(SELECT tbl_OR010300.* "
        'MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (OR01096 = N'" & Declarations.SalesmanCode & "') AND "
        'MySQLStr = MySQLStr & "(OR01015 >= CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)) AND "
        'MySQLStr = MySQLStr & "(OR01015 <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 103))) AS View_1 "
        'MySQLStr = MySQLStr & "ON SL010300.SL01001 = View_1.OR01003 "
        'MySQLStr = MySQLStr & "ORDER BY View_1.OR01001 "

        MySQLStr = "SELECT View_1.OR01001, View_1.OR01015, View_1.ExpirationDate, View_1.OR01003, CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') "
        MySQLStr = MySQLStr & "+ ' ' + ISNULL(SL010300.SL01003, N''))) = '' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, "
        MySQLStr = MySQLStr & "N'') END AS CName, ISNULL(View_1.AgentName, N'') AS AgentName, View_1.OrderN, ISNULL(View_3.Comment, N'') AS Comment, View_1.CPState "
        MySQLStr = MySQLStr & "FROM (SELECT OR17001, LTRIM(RTRIM(LTRIM(RTRIM(OR17005)) + ' ' + LTRIM(RTRIM(OR17006)))) AS Comment "
        MySQLStr = MySQLStr & "FROM  tbl_OR170300) AS View_3 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR01001, OR01003, OR01015, OrderN, CName, ExpirationDate, AgentName, CPState "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR01096 = N'" & Declarations.SalesmanCode & "') "
        MySQLStr = MySQLStr & "AND (OR01015 >= CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)) "
        MySQLStr = MySQLStr & "AND (OR01015 <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', "
        MySQLStr = MySQLStr & "103))) AS View_1 ON View_3.OR17001 = View_1.OR01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 WITH (NOLOCK) ON View_1.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "ORDER BY View_1.OR01001 desc "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "Номер предложения"
        DataGridView1.Columns(0).Width = 120
        DataGridView1.Columns(1).HeaderText = "Дата создания"
        DataGridView1.Columns(1).Width = 120
        DataGridView1.Columns(1).ReadOnly = True
        DataGridView1.Columns(2).HeaderText = "Действительно до"
        DataGridView1.Columns(2).Width = 120
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).HeaderText = "Код покупателя"
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(3).ReadOnly = True
        DataGridView1.Columns(4).HeaderText = "Имя покупателя"
        DataGridView1.Columns(4).ReadOnly = True
        DataGridView1.Columns(5).HeaderText = "Имя агента"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(5).ReadOnly = True
        If My.Settings.ShowAgentColumn = 0 Then
            DataGridView1.Columns(5).Visible = False
        Else
            DataGridView1.Columns(5).Visible = True
        End If
        DataGridView1.Columns(6).HeaderText = "Перенесено в заказ на продажу номер"
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(6).ReadOnly = True
        DataGridView1.Columns(7).HeaderText = "Комментарий"
        DataGridView1.Columns(7).Width = 200
        DataGridView1.Columns(7).ReadOnly = True
        DataGridView1.Columns(8).HeaderText = "Состояние КП"
        DataGridView1.Columns(8).Width = 150
        DataGridView1.Columns(8).ReadOnly = True

        CheckButtons()
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    End Sub

    Private Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При изменении даты "с" и "по" перезагружаем данные
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ReloadData()
        CheckButtons()

    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button6.Enabled = False
        Else
            Button4.Enabled = True
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString) = "" And Trim(DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString) = "" Then
                Button2.Enabled = True
            Else
                Button2.Enabled = False
            End If
            If Trim(DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString) = "" Then
                Button6.Enabled = True
                Button3.Enabled = True
            Else
                Button6.Enabled = False
                Button3.Enabled = False
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// копирование содержимого ячейки в буфер при начале редактирования
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        My.Computer.Clipboard.Clear()
        My.Computer.Clipboard.SetText(DataGridView1.CurrentCell.Value)
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка разных типов предложений
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)

        If Trim(row.Cells(6).Value.ToString) <> "" Then
            '---переведенные в 0 тип
            row.DefaultCellStyle.BackColor = Color.LightGreen
        Else
            If Trim(row.Cells(5).Value.ToString) <> "" Then
                row.DefaultCellStyle.BackColor = Color.LightYellow
            Else
                If Trim(row.Cells(2).Value) < Now Then
                    '---устаревшие - просроченная дата
                    row.DefaultCellStyle.BackColor = Color.LightGray
                Else
                    '--рабочие
                    row.DefaultCellStyle.BackColor = Color.White
                End If
            End If
        End If

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие заказа на редактирование
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button3.Enabled = True Then
            EditString()
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub ReloadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(TextBox1.Text) = "" Then
            MySQLStr = "SELECT View_1.OR01001, View_1.OR01015, View_1.ExpirationDate, View_1.OR01003, CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') "
            MySQLStr = MySQLStr & "+ ' ' + ISNULL(SL010300.SL01003, N''))) = '' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, "
            MySQLStr = MySQLStr & "N'') END AS CName, ISNULL(View_1.AgentName, N'') AS AgentName, View_1.OrderN, ISNULL(View_3.Comment, N'') AS Comment, View_1.CPState "
            MySQLStr = MySQLStr & "FROM (SELECT OR17001, LTRIM(RTRIM(LTRIM(RTRIM(OR17005)) + ' ' + LTRIM(RTRIM(OR17006)))) AS Comment "
            MySQLStr = MySQLStr & "FROM  tbl_OR170300) AS View_3 RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT OR01001, OR01003, OR01015, OrderN, CName, ExpirationDate, AgentName, CPState "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01096 = N'" & Declarations.SalesmanCode & "') "
            MySQLStr = MySQLStr & "AND (OR01015 >= CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)) "
            MySQLStr = MySQLStr & "AND (OR01015 <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', "
            MySQLStr = MySQLStr & "103))) AS View_1 ON View_3.OR17001 = View_1.OR01001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SL010300 WITH (NOLOCK) ON View_1.OR01003 = SL010300.SL01001 "
            MySQLStr = MySQLStr & "ORDER BY View_1.OR01001 Desc "
        Else
            MySQLStr = "SELECT View_1.OR01001, View_1.OR01015, View_1.ExpirationDate, View_1.OR01003, CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') "
            MySQLStr = MySQLStr & "+ ' ' + ISNULL(SL010300.SL01003, N''))) = '' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, "
            MySQLStr = MySQLStr & "N'') END AS CName, ISNULL(View_1.AgentName, N'') AS AgentName, View_1.OrderN, ISNULL(View_3.Comment, N'') AS Comment, View_1.CPState "
            MySQLStr = MySQLStr & "FROM (SELECT OR17001, LTRIM(RTRIM(LTRIM(RTRIM(OR17005)) + ' ' + LTRIM(RTRIM(OR17006)))) AS Comment "
            MySQLStr = MySQLStr & "FROM  tbl_OR170300) AS View_3 RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT OR01001, OR01003, OR01015, OrderN, CName, ExpirationDate, AgentName, CPState "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01096 = N'" & Declarations.SalesmanCode & "') "
            MySQLStr = MySQLStr & "AND (OR01015 >= CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)) "
            MySQLStr = MySQLStr & "AND (OR01015 <= CONVERT(DATETIME, '" & DateTimePicker2.Value & "', 103)) "
            MySQLStr = MySQLStr & ") AS View_1 ON View_3.OR17001 = View_1.OR01001 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "SL010300 WITH (NOLOCK) ON View_1.OR01003 = SL010300.SL01001 "
            '---
            MySQLStr = MySQLStr & "WHERE ((UPPER(View_1.OR01001) LIKE N'%" & UCase(Trim(TextBox1.Text)) & "%') "
            MySQLStr = MySQLStr & "OR (UPPER(View_1.OrderN) LIKE N'%" & UCase(Trim(TextBox1.Text)) & "%') "
            MySQLStr = MySQLStr & "OR (UPPER(CASE WHEN Ltrim(Rtrim(ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, N''))) = "
            MySQLStr = MySQLStr & "'' THEN ISNULL(View_1.CName, '') ELSE ISNULL(SL010300.SL01002, N'') + ' ' + ISNULL(SL010300.SL01003, N'') "
            MySQLStr = MySQLStr & "END) LIKE N'%" & UCase(Trim(TextBox1.Text)) & "%') "
            MySQLStr = MySQLStr & "OR (UPPER(View_3.Comment) LIKE N'%" & UCase(Trim(TextBox1.Text)) & "%')) "
            '---
            MySQLStr = MySQLStr & "ORDER BY View_1.OR01001 Desc "
        End If
        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выбранного предложения 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyOrder As String
        Dim MySQLStr As String                        'рабочая строка

        MyOrder = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
        MySQLStr = "DELETE FROM tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "DELETE FROM tbl_OR010300 "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(MyOrder) & "')"
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "DELETE FROM tbl_OR170300 "
        MySQLStr = MySQLStr & "WHERE (OR17001 = N'" & Trim(MyOrder) & "') "
        Declarations.MyConn.Execute(MySQLStr)

        ReloadData()
        CheckButtons()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Копирование выбранного предложения 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyOLDID As String                           'ID старого заказа (из которого копируют)
        Dim MyNEWID As Double                           'ID нового заказа (куда копируют)
        Dim MySQLStr As String                          'рабочая строка

        MyOLDID = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
        MyNEWID = GetNewID()

        '----заголовок
        MySQLStr = "INSERT INTO tbl_OR010300 "
        MySQLStr = MySQLStr & "SELECT CONVERT(nchar(10),'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "'), OR01002, OR01003, OR01004, OR01005, OR01006, OR01007, OR01008, OR01009, OR01010, OR01011, OR01012, OR01013, OR01014, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
        MySQLStr = MySQLStr & "OR01017, OR01018,    OR01096,    OR01020, OR01021, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "OR01023, OR01024, OR01025, OR01026, OR01027, OR01028, "
        MySQLStr = MySQLStr & "OR01029, OR01030, OR01031, OR01032, OR01033, OR01034, OR01035, OR01036, OR01037, OR01038, OR01039, OR01040, OR01041, OR01042, "
        MySQLStr = MySQLStr & "OR01043, OR01044, OR01045, OR01046, OR01047, OR01048, OR01049, OR01050, OR01051, OR01052, OR01053, OR01054, OR01055, OR01056, "
        MySQLStr = MySQLStr & "OR01057, OR01058, OR01059, OR01060, OR01061, OR01062, OR01063, OR01064, OR01065, OR01066, OR01067, OR01068, OR01069, OR01070, "
        MySQLStr = MySQLStr & "OR01071, OR01072, OR01073, OR01074, OR01075, OR01076, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "OR01078, OR01079, OR01080, OR01081, OR01082, OR01083, OR01084, "
        MySQLStr = MySQLStr & "OR01085, OR01086, OR01087, OR01088, OR01089, OR01090, OR01091, OR01092, OR01093, OR01094, OR01095, OR01096, OR01097, OR01098, "
        MySQLStr = MySQLStr & "OR01099, OR01100, OR01101, OR01102, OR01103, OR01104, OR01105, OR01106, OR01107, OR01108, OR01109, OR01110, OR01111, OR01112, "
        MySQLStr = MySQLStr & "OR01113, OR01114, OR01115, OR01116, OR01117, OR01118, OR01119, OR01120, OR01121, OR01122, OR01123, OR01124, OR01125, OR01126, "
        MySQLStr = MySQLStr & "OR01127, OR01128, OR01129, OR01130, OR01131, OR01132, OR01133, OR01134, OR01135, OR01136, OR01137, OR01138, OR01139, OR01140, "
        MySQLStr = MySQLStr & "OR01141, OR01142, OR01143, OR01144, OR01145, OR01146, OR01147, OR01148, OR01149, OR01150, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "OR01152, OR01153, OR01154, "
        MySQLStr = MySQLStr & "OR01155, OR01156, OR01157, OR01158, OR01159, OR01160, OR01161, OR01162, OR01163, OR01164, OR01165, OR01166, OR01167, OR01168, "
        MySQLStr = MySQLStr & "OR01169, OR01170, OR01171, OR01172, OR01173, OR01174, OR01175, OR01176, OR01177, OR01178, OR01179, OR01180, OR01181, OR01182, "
        MySQLStr = MySQLStr & "OR01183, OR01184, OR01185, OR01186, OR01187, OR01188, OR01189, OR01190, OR01191, OR01192, OR01193, OR01194, OR01195, OR01196, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "OR01198, OR01199, OR01200, OR01201, OR01202, OR01203, OR01204, OR01205, "
        MySQLStr = MySQLStr & "'', '', '', '', CONVERT(DATETIME, '1900-01-01 00:00:00', 102), '', CONVERT(DATETIME, '1900-01-01 00:00:00', 102), '', CONVERT(DATETIME, '1900-01-01 00:00:00', 102), 0, NULL, NULL  "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 AS tbl_OR010300_1 "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyOLDID & "')"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----строки
        MySQLStr = "INSERT INTO tbl_OR030300 "
        MySQLStr = MySQLStr & "SELECT CONVERT(nchar(10),'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "'), OR03002, OR03003, OR03004, OR03005, OR03006, OR03007, OR03008, OR03009, OR03010, OR03011, OR03012, OR03013, OR03014, "
        MySQLStr = MySQLStr & "OR03015, OR03016, OR03017, OR03018, "
        MySQLStr = MySQLStr & "DATEADD(day, Ceiling(CASE WHEN ISNULL(WeekQTY,0) < 1 THEN 1 ELSE ISNULL(WeekQTY,0) END * 7 - ISNULL(DelWeekQTY,0) * 7), CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
        MySQLStr = MySQLStr & "OR03020, OR03021, OR03022, OR03023, OR03024, OR03025, OR03026, OR03027, OR03028, "
        MySQLStr = MySQLStr & "OR03029, OR03030, OR03031, OR03032, OR03033, OR03034, OR03035, OR03036, "
        MySQLStr = MySQLStr & "DATEADD(day, Ceiling(CASE WHEN ISNULL(WeekQTY,0) < 1 THEN 1 ELSE ISNULL(WeekQTY,0) END * 7 - ISNULL(DelWeekQTY,0) * 7), CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
        MySQLStr = MySQLStr & "OR03038, OR03039, OR03040, OR03041, OR03042, "
        MySQLStr = MySQLStr & "OR03043, OR03044, OR03045, OR03046, OR03047, OR03048, OR03049, OR03050, OR03051, OR03052, OR03053, OR03054, OR03055, OR03056, "
        MySQLStr = MySQLStr & "OR03057, OR03058, OR03059, OR03060, OR03061, OR03062, OR03063, OR03064, OR03065, OR03066, OR03067, OR03068, OR03069, OR03070, "
        MySQLStr = MySQLStr & "OR03071, OR03072, OR03073, OR03074, OR03075, OR03076, OR03077, OR03078, OR03079, OR03080, OR03081, OR03082, OR03083, OR03084, "
        MySQLStr = MySQLStr & "OR03085, OR03086, OR03087, OR03088, OR03089, OR03090, OR03091, OR03092, OR03093, OR03094, OR03095, OR03096, OR03097, OR03098, "
        MySQLStr = MySQLStr & "OR03099, OR03100, OR03101, OR03102, OR03103, OR03104, OR03105, OR03106, OR03107, OR03108, OR03109, OR03110, OR03111, OR03112, "
        MySQLStr = MySQLStr & "OR03113, "
        MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
        MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
        MySQLStr = MySQLStr & "OR03116, OR03117, OR03118, OR03119, OR03120, "
        MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
        MySQLStr = MySQLStr & "OR03122, OR03123, OR03124, WeekQTY, DelWeekQTY, SuppItemCode, SuppCode, SuppName "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 AS tbl_OR030300_1 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOLDID & "') "
        Declarations.MyConn.Execute(MySQLStr)

        '----Корректировка даты отгрузки в заголовке ---
        MySQLStr = "Update tbl_OR010300 "
        MySQLStr = MySQLStr & "Set OR01016 = View_1.CC, "
        MySQLStr = MySQLStr & "ReadyDate = View_1.CC "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
        MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON tbl_OR010300.OR01001 = View_1.OR03001 "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----Корректировка НДС - выставляем новый (7)
        MySQLStr = "Update tbl_OR010300 "
        MySQLStr = MySQLStr & "SET OR01093 = 7, OR01094 = 7, OR01095 = N'7', OR01118 = 7 "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
        MySQLStr = MySQLStr & "AND (OR01093 = 2) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "Update tbl_OR030300 "
        MySQLStr = MySQLStr & "SET OR03061 = 7 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
        MySQLStr = MySQLStr & "AND (OR03061 = 2) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        ReloadData()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перенос выбранного предложения в заказ 0 типа 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyNEWID As Double                           'ID нового заказа (куда копируют)
        Dim MyPRID As String                            'ID предложения, которое будет перенесено в 0 тип
        Dim MySQLStr As String                          'рабочая строка
        Dim MyCurrCode As Integer                       'код валюты заказа
        Dim MyExRate As Double                          'курс валюты заказа
        Dim MyFvdCurrCode As Integer                    'код валюты валютной оговорки
        Dim MyFvdExRate As Double                       'курс валюты валютной оговорки
        Dim MyYear As String                            'текущий год
        Dim MyShCost As Double                          'стоимость доставки

        MyPRID = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
        If CheckDataInProposal(MyPRID) = True Then
            '---все проверки OK - копируем в заказ 0 типа
            MyNEWID = GetNewPRDID()

            '----получение курсов валют на текущую дату
            MyYear = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.DatePart(DateInterval.Year, Now), 2)
            '----Код валюты валютной оговорки
            MySQLStr = "SELECT OR05021 "
            MySQLStr = MySQLStr & "FROM OR050300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR05001 = N'" & MyYear & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyFvdCurrCode = 12
            Else
                MyFvdCurrCode = Declarations.MyRec.Fields("OR05021").Value
            End If
            trycloseMyRec()

            '----курс валюты валютной оговорки
            MySQLStr = "SELECT SYCH006 "
            MySQLStr = MySQLStr & "FROM SYCH0100 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SYCH001 = " & CStr(MyFvdCurrCode) & ") AND "
            MySQLStr = MySQLStr & "(SYCH004 <= CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)) AND "
            MySQLStr = MySQLStr & "(SYCH005 > CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyFvdExRate = 1
            Else
                MyFvdExRate = Declarations.MyRec.Fields("SYCH006").Value
            End If
            trycloseMyRec()

            '----код валюты заказа
            MySQLStr = "SELECT OR01028 "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyPRID & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyCurrCode = 0
            Else
                MyCurrCode = Declarations.MyRec.Fields("OR01028").Value
            End If
            trycloseMyRec()

            '----курс валюты заказа
            MySQLStr = "SELECT SYCH006 "
            MySQLStr = MySQLStr & "FROM SYCH0100 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SYCH001 = " & CStr(MyCurrCode) & ") AND "
            MySQLStr = MySQLStr & "(SYCH004 <= CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)) AND "
            MySQLStr = MySQLStr & "(SYCH005 > CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MyExRate = 1
            Else
                MyExRate = Declarations.MyRec.Fields("SYCH006").Value
            End If
            trycloseMyRec()


            '----перенос заголовка предложения
            MySQLStr = "INSERT INTO OR010300 "
            MySQLStr = MySQLStr & "SELECT N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "', OR01002, OR01003, OR01004, OR01005, OR01006, OR01007, OR01008, OR01009, OR01010, OR01011, OR01012, OR01013, OR01014, "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            'MySQLStr = MySQLStr & "OR01016, "
            MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "OR01017, OR01018, OR01019, OR01020, OR01021, "
            'MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            MySQLStr = MySQLStr & "ExpirationDate, "
            MySQLStr = MySQLStr & "OR01023, OR01024, OR01025, OR01026, OR01027, OR01028, "
            MySQLStr = MySQLStr & "OR01029, OR01030, OR01031, OR01032, "
            MySQLStr = MySQLStr & CStr(MyFvdCurrCode) & ", "
            MySQLStr = MySQLStr & "OR01034, OR01035, "
            MySQLStr = MySQLStr & Replace(CStr(MyFvdExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(MyFvdExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "OR01038, OR01039, OR01040, OR01041, OR01042, "
            MySQLStr = MySQLStr & "OR01043, OR01044, OR01045, OR01046, OR01047, OR01048, OR01049, OR01050, OR01051, OR01052, OR01053, OR01054, OR01055, OR01056, "
            MySQLStr = MySQLStr & "OR01057, OR01058, OR01059, OR01060, OR01061, OR01062, OR01063, OR01064, OR01065, OR01066, "
            MySQLStr = MySQLStr & Replace(CStr(MyExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "OR01068, OR01069, OR01070, "
            MySQLStr = MySQLStr & "OR01071, OR01072, OR01073, OR01074, OR01075, OR01076, "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            MySQLStr = MySQLStr & "OR01078, OR01079, OR01080, OR01081, OR01082, OR01083, OR01084, "
            MySQLStr = MySQLStr & "OR01085, OR01086, OR01087, OR01088, OR01089, OR01090, OR01091, OR01092, OR01093, OR01094, OR01095, OR01096, OR01097, OR01098, "
            MySQLStr = MySQLStr & "OR01099, OR01100, OR01101, OR01102, OR01103, OR01104, OR01105, OR01106, OR01107, OR01108, OR01109, OR01110, OR01111, OR01112, "
            MySQLStr = MySQLStr & "OR01113, OR01114, OR01115, OR01116, OR01117, OR01118, "
            MySQLStr = MySQLStr & "0, "
            MySQLStr = MySQLStr & "0, "
            MySQLStr = MySQLStr & "OR01121, OR01122, OR01123, OR01124, OR01125, OR01126, "
            MySQLStr = MySQLStr & "OR01127, OR01128, OR01129, OR01130, OR01131, OR01132, OR01133, OR01134, OR01135, OR01136, OR01137, OR01138, OR01139, OR01140, "
            MySQLStr = MySQLStr & "OR01141, OR01142, OR01143, OR01144, OR01145, OR01146, OR01147, OR01148, OR01149, OR01150, "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            MySQLStr = MySQLStr & "OR01152, OR01153, OR01154, "
            MySQLStr = MySQLStr & "OR01155, OR01156, OR01157, OR01158, OR01159, OR01160, OR01161, OR01162, OR01163, OR01164, OR01165, OR01166, OR01167, OR01168, "
            MySQLStr = MySQLStr & "OR01169, OR01170, OR01171, OR01172, OR01173, OR01174, OR01175, OR01176, OR01177, OR01178, OR01179, OR01180, OR01181, OR01182, "
            MySQLStr = MySQLStr & "OR01183, OR01184, OR01185, OR01186, OR01187, OR01188, OR01189, OR01190, OR01191, OR01192, OR01193, OR01194, OR01195, OR01196, "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            MySQLStr = MySQLStr & "OR01198, OR01199, OR01200, OR01201, OR01202, OR01203, OR01204, OR01205 "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 AS tbl_OR010300_1 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyPRID & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '----перенос строк предложения
            MySQLStr = "INSERT INTO OR030300 "
            MySQLStr = MySQLStr & "SELECT N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "', OR03002, OR03003, OR03004, Ltrim(Rtrim(OR03005)) AS OR03005, OR03006, OR03007, OR03008, OR03009, OR03010, OR03011, OR03012, OR03013, OR03014, "
            MySQLStr = MySQLStr & "OR03015, OR03016, OR03017, OR03018, "
            'MySQLStr = MySQLStr & "OR03019, "
            MySQLStr = MySQLStr & "DATEADD(day, Ceiling(CASE WHEN ISNULL(WeekQTY,0) < 1 THEN 1 ELSE ISNULL(WeekQTY,0) END * 7 - ISNULL(DelWeekQTY,0) * 7), CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "OR03020, OR03021, OR03022, OR03023, "
            MySQLStr = MySQLStr & Replace(CStr(MyExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "OR03025, OR03026, OR03027, OR03028, "
            MySQLStr = MySQLStr & "OR03029, OR03030, OR03031, OR03032, OR03033, OR03034, OR03035, OR03036, "
            'MySQLStr = MySQLStr & "OR03037, "
            MySQLStr = MySQLStr & "DATEADD(day, Ceiling(CASE WHEN ISNULL(WeekQTY,0) < 1 THEN 1 ELSE ISNULL(WeekQTY,0) END * 7 - ISNULL(DelWeekQTY,0) * 7), CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "OR03038, OR03039, "
            MySQLStr = MySQLStr & Replace(CStr(MyExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & Replace(CStr(MyExRate), ",", ".") & ", "
            MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "OR03043, OR03044, OR03045, OR03046, OR03047, OR03048, OR03049, OR03050, OR03051, OR03052, OR03053, OR03054, OR03055, OR03056, "
            MySQLStr = MySQLStr & "OR03057, OR03058, OR03059, OR03060, OR03061, OR03062, OR03063, OR03064, OR03065, OR03066, OR03067, OR03068, OR03069, OR03070, "
            MySQLStr = MySQLStr & "OR03071, OR03072, OR03073, OR03074, OR03075, OR03076, OR03077, OR03078, OR03079, OR03080, OR03081, OR03082, OR03083, OR03084, "
            MySQLStr = MySQLStr & "OR03085, OR03086, OR03087, OR03088, OR03089, OR03090, OR03091, OR03092, OR03093, OR03094, OR03095, OR03096, OR03097, OR03098, "
            MySQLStr = MySQLStr & "OR03099, OR03100, OR03101, OR03102, OR03103, OR03104, OR03105, OR03106, OR03107, OR03108, OR03109, OR03110, OR03111, OR03112, "
            MySQLStr = MySQLStr & "OR03113, "
            MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "DATEADD(day,5,CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103)), "
            MySQLStr = MySQLStr & "OR03116, OR03117, OR03118, OR03119, OR03120, "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(MM, GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy, GETDATE())), 103), "
            MySQLStr = MySQLStr & "OR03122, OR03123, OR03124 "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 AS tbl_OR030300_1 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyPRID & "') "
            Declarations.MyConn.Execute(MySQLStr)

            '----Перенос комментариев предложения (0 строки)
            MySQLStr = "DELETE FROM OR170300 "
            MySQLStr = MySQLStr & "WHERE (OR17001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO OR170300 "
            MySQLStr = MySQLStr & "(OR17001, OR17002, OR17003, OR17004, OR17005, OR17006, OR17007, OR17008, OR17009, OR17010, OR17011, OR17012, OR17013, OR17014, "
            MySQLStr = MySQLStr & "OR17015, OR17016, OR17017, OR17018, OR17019, OR17020, OR17021, OR17022, OR17023, OR17024) "
            MySQLStr = MySQLStr & "SELECT N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "', OR17002, OR17003, OR17004, OR17005, OR17006, OR17007, OR17008, OR17009, OR17010, OR17011, OR17012, OR17013, OR17014, "
            MySQLStr = MySQLStr & "OR17015, OR17016, OR17017, OR17018, OR17019, OR17020, OR17021, OR17022, OR17023, OR17024 "
            MySQLStr = MySQLStr & "FROM tbl_OR170300 "
            MySQLStr = MySQLStr & "WHERE (OR17001 = N'" & MyPRID & "') "
            Declarations.MyConn.Execute(MySQLStr)

            '----Добавление в доп. табличку информации о поставщиках
            MySQLStr = "DELETE FROM tbl_SupplierInSalesOrder0300 "
            MySQLStr = MySQLStr & "WHERE (OR001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO tbl_SupplierInSalesOrder0300 "
            MySQLStr = MySQLStr & "SELECT NEWID() AS Expr1, "
            MySQLStr = MySQLStr & "OR030300.OR03001, "
            MySQLStr = MySQLStr & "OR030300.OR03002, "
            MySQLStr = MySQLStr & "OR030300.OR03003, "
            MySQLStr = MySQLStr & "SC010300.SC01058 "
            MySQLStr = MySQLStr & "FROM OR030300 INNER JOIN "
            MySQLStr = MySQLStr & "SC010300 ON OR030300.OR03005 = SC010300.SC01001 "
            MySQLStr = MySQLStr & "WHERE (OR030300.OR03001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "')"
            Declarations.MyConn.Execute(MySQLStr)

            '----Перенос информации о стоимости доставки
            MySQLStr = "DELETE FROM tbl_SalesHdr_AddInfo "
            MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "SELECT DeliverySum "
            MySQLStr = MySQLStr & "FROM tbl_SW4SalesHdr_AddInfo WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OrderID = N'" & MyPRID & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MyShCost = 0
            Else
                Declarations.MyRec.MoveFirst()
                MyShCost = CStr(Declarations.MyRec.Fields("DeliverySum").Value)
            End If
            trycloseMyRec()

            If MyShCost <> 0 Then
                MySQLStr = "INSERT INTO tbl_SalesHdr_AddInfo "
                MySQLStr = MySQLStr & "(ID, OrderID, DeliverySum) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                MySQLStr = MySQLStr & "N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "', "
                MySQLStr = MySQLStr & Replace(CStr(MyShCost), ",", ".") & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '----Занесение инфо, в какой номер заказа перенесено предложение
            MySQLStr = "UPDATE tbl_OR010300 "
            MySQLStr = MySQLStr & "SET OrderN = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "' "
            MySQLStr = MySQLStr & "WHERE (OR01001  = N'" & MyPRID & "')"
            Declarations.MyConn.Execute(MySQLStr)

            '----Корректировка даты отгрузки в заголовке
            MySQLStr = "Update tbl_OR010300 "
            MySQLStr = MySQLStr & "Set OR01016 = View_1.CC, "
            MySQLStr = MySQLStr & "ReadyDate = View_1.CC "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
            MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON tbl_OR010300.OR01001 = View_1.OR03001 "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-----Перенос информации для ЭДО
            MySQLStr = "INSERT INTO tbl_SalesHdr_EDOInfo "
            MySQLStr = MySQLStr & "SELECT N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "', CustomerPONum, CustomerAgreementNum, CustomerManagerName, DeliveryAddress, GovermentID, InternalComment, "
            MySQLStr = MySQLStr & "CustomerAgreementDateStart, CustomerAgreementDateFin "
            MySQLStr = MySQLStr & "FROM tbl_SalesHdrCP_EDOInfo "
            MySQLStr = MySQLStr & "WHERE (OrderID = N'" & MyPRID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-----Обновление информации для ЭДО
            MySQLStr = "Update tbl_SalesHdr_EDOInfo "
            MySQLStr = MySQLStr & "SET CustomerAgreementNum = View_7.AgreementN, CustomerAgreementDateStart = View_7.DateFrom, CustomerAgreementDateFin = View_7.DateTo "
            MySQLStr = MySQLStr & "FROM tbl_SalesHdr_EDOInfo INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT OR010300.OR01001, CASE WHEN ISNULL(tbl_CustomerCard0300.AgreementN, OR010300.OR01001) "
            MySQLStr = MySQLStr & "= '' THEN OR010300.OR01001 ELSE ISNULL(tbl_CustomerCard0300.AgreementN, OR010300.OR01001) END AS AgreementN, "
            MySQLStr = MySQLStr & "CASE WHEN ISNULL(tbl_CustomerCard0300.DataFrom, CONVERT(datetime, '01/01/1900', 103)) = CONVERT(datetime, '01/01/1900', 103) "
            MySQLStr = MySQLStr & "THEN OR010300.OR01015 ELSE ISNULL(tbl_CustomerCard0300.DataFrom, CONVERT(datetime, '01/01/1900', 103)) END AS DateFrom, "
            MySQLStr = MySQLStr & "CASE WHEN ISNULL(tbl_CustomerCard0300.DataTo, CONVERT(datetime, '01/01/1900', 103)) = CONVERT(datetime, '01/01/1900', 103) "
            MySQLStr = MySQLStr & "THEN CONVERT(datetime, '01/01/1900', 103) ELSE ISNULL(tbl_CustomerCard0300.DataTo, CONVERT(datetime, '01/01/1900', 103)) "
            MySQLStr = MySQLStr & "END AS DateTo "
            MySQLStr = MySQLStr & "FROM OR010300 LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "(SELECT SL01001, AgreementN, DataFrom, DataTo "
            MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 AS tbl_CustomerCard0300_1 "
            MySQLStr = MySQLStr & "WHERE (DataFrom <= GETDATE()) AND (DataTo >= GETDATE() OR "
            MySQLStr = MySQLStr & "DataTo <> DataFrom AND DataTo = CONVERT(datetime, '01/01/1900', 103))) AS tbl_CustomerCard0300 ON "
            MySQLStr = MySQLStr & "OR010300.OR01003 = tbl_CustomerCard0300.SL01001 "
            MySQLStr = MySQLStr & "WHERE (OR010300.OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "')) AS View_7 ON tbl_SalesHdr_EDOInfo.OrderID = View_7.OR01001 "
            MySQLStr = MySQLStr & "WHERE (tbl_SalesHdr_EDOInfo.CustomerAgreementNum = N'') OR"
            MySQLStr = MySQLStr & "(tbl_SalesHdr_EDOInfo.CustomerAgreementNum = N'" & MyPRID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-----Обновление информации о заказе с WEB сайта (если есть)
            MySQLStr = "UPDATE tbl_WEB_OrderNum "
            MySQLStr = MySQLStr & "SET ScaOrderNUm = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "' "
            MySQLStr = MySQLStr & "WHERE (ProposalNum = N'" & MyPRID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-----Корректировка НДС - выставляем новый (7)
            MySQLStr = "Update OR010300 "
            MySQLStr = MySQLStr & "SET OR01093 = 7, OR01094 = 7, OR01095 = N'7', OR01118 = 7 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
            MySQLStr = MySQLStr & "AND (OR01093 = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "Update OR030300 "
            MySQLStr = MySQLStr & "SET OR03061 = 7 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyNEWID), 10) & "') "
            MySQLStr = MySQLStr & "AND (OR03061 = 2) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            ReloadData()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If Trim(DataGridView1.Item(0, i).Value.ToString) = Trim(MyPRID) Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Exit For
                End If
            Next
            CheckButtons()

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditHeader = New EditHeader
        MyEditHeader.StartParam = "Create"
        MyEditHeader.ShowDialog()
        ReloadData()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Trim(Declarations.MyOrderNum) Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие заказа на редактирование
        '//
        '////////////////////////////////////////////////////////////////////////////////

        EditString()
    End Sub

    Private Sub EditString()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие заказа на редактирование
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditHeader = New EditHeader
        MyEditHeader.StartParam = "Edit"
        MyEditHeader.ShowDialog()
        ReloadData()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Trim(Declarations.MyOrderNum) Then
                DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передать коммерческое предложение другому продавцу группы / вернуть
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySendReturnProposal = New SendReturnProposal
        MySendReturnProposal.ShowDialog()
        ReloadData()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрос на поиск поставщика товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MySearchSupplier Is Nothing Then
            MySearchSupplier = New SearchSupplier
            MySearchSupplier.Show()
        Else
            'MySearchSupplier.BringToFront()
            MySearchSupplier.Close()
            MySearchSupplier = New SearchSupplier
            MySearchSupplier.Show()
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки фильтра
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ReloadData()
        CheckButtons()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки снятие фильтра
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        ReloadData()
        CheckButtons()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
