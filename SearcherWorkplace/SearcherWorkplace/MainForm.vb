Imports Microsoft.Office

Public Class MainForm
    Public LoadFlag As Integer
    Public FullInfoFlag As Integer

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// после чего выводим список предложений данного пользователя
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        FullInfoFlag = 0
        LoadFlag = 1
        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "Gurda"

            MySQLStr = "SELECT tbl_SupplSearch_Searchers.PurchID, View_16.SYPD003, tbl_SupplSearch_Searchers.IsLeader "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearch_Searchers INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD003 "
            MySQLStr = MySQLStr & "FROM SYPD0300 "
            MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_16 ON tbl_SupplSearch_Searchers.PurchID = View_16.SYPD001 INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON View_16.SYPD003 = ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "

            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("Не найден код закупщика или код поисковика, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                trycloseMyRec()
                Application.Exit()
            Else
                Declarations.PurchCode = Declarations.MyRec.Fields("PurchID").Value
                Declarations.PurchName = Declarations.MyRec.Fields("SYPD003").Value
                Declarations.IsManager = Declarations.MyRec.Fields("IsLeader").Value
                trycloseMyRec()
                Label9.Text = Declarations.PurchCode & " " & Declarations.PurchName
            End If
        Catch ex As Exception
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---содержимое комбобокса - что выводить
        If Declarations.IsManager = 1 Then
            ComboBoxAct.Items.Clear()
            ComboBoxAct.Items.Add("Только активные для поисковика")
            ComboBoxAct.Items.Add("Только активные")
            ComboBoxAct.Items.Add("Приостановленные")
            ComboBoxAct.Items.Add("Все запросы")
            ComboBoxAct.Items.Add("Нераспределенные")
            ComboBoxAct.Items.Add("Только активные всех поисковиков для поисковиков")
            ComboBoxAct.Items.Add("Только активные всех поисковиков")
            ComboBoxAct.Items.Add("Приостановленные всех поисковиков")
            ComboBoxAct.Items.Add("Все запросы всех поисковиков")
            ComboBoxAct.SelectedIndex = 4
        Else
            ComboBoxAct.Items.Clear()
            ComboBoxAct.Items.Add("Только активные для поисковика")
            ComboBoxAct.Items.Add("Только активные")
            ComboBoxAct.Items.Add("Приостановленные")
            ComboBoxAct.Items.Add("Все запросы")
            ComboBoxAct.SelectedIndex = 0
        End If


        LoadFlag = 0
        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        SortColumnNum = Nothing
        SortColOrder = System.ComponentModel.ListSortDirection.Ascending
        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub LoadRequests()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных по запросам на поиск
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            If ComboBoxAct.Text = "Только активные для поисковика" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-1 "
            ElseIf ComboBoxAct.Text = "Только активные" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "0 "
            ElseIf ComboBoxAct.Text = "Приостановленные" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "41 "
            ElseIf ComboBoxAct.Text = "Все запросы" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "1 "
            ElseIf ComboBoxAct.Text = "Нераспределенные" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "2 "
            ElseIf ComboBoxAct.Text = "Только активные всех поисковиков для поисковиков" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-3 "
            ElseIf ComboBoxAct.Text = "Только активные всех поисковиков" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "3 "
            ElseIf ComboBoxAct.Text = "Приостановленные всех поисковиков" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "411 "
            ElseIf ComboBoxAct.Text = "Все запросы всех поисковиков" Then
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "4 "
            Else
                MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo 0, 0 "
            End If

            MySQLStr = MySQLStr + ", N'" + Trim(TextBox1.Text) + "' "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView1.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView1.Columns(0).HeaderText = "N за проса"
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(0).ReadOnly = True
            DataGridView1.Columns(1).HeaderText = "Дата за проса"
            DataGridView1.Columns(1).Width = 100
            DataGridView1.Columns(1).DefaultCellStyle.Format = "dd/MM/yyyy HH:mm"
            DataGridView1.Columns(1).ReadOnly = True
            DataGridView1.Columns(2).HeaderText = "Клиент"
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(2).ReadOnly = True
            DataGridView1.Columns(3).HeaderText = "Контактное лицо"
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(3).ReadOnly = True
            DataGridView1.Columns(4).HeaderText = "Телефон"
            DataGridView1.Columns(4).Width = 150
            DataGridView1.Columns(4).ReadOnly = True
            DataGridView1.Columns(5).HeaderText = "Email"
            DataGridView1.Columns(5).Width = 150
            DataGridView1.Columns(5).ReadOnly = True
            DataGridView1.Columns(6).HeaderText = "Срок представления КП"
            DataGridView1.Columns(6).Width = 100
            DataGridView1.Columns(6).DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridView1.Columns(6).ReadOnly = True
            DataGridView1.Columns(7).HeaderText = "Продавец"
            DataGridView1.Columns(7).Width = 150
            DataGridView1.Columns(7).ReadOnly = True
            DataGridView1.Columns(8).HeaderText = "ID Статус продавца"
            DataGridView1.Columns(8).Visible = False
            DataGridView1.Columns(8).ReadOnly = True
            DataGridView1.Columns(9).HeaderText = "Статус продавца"
            DataGridView1.Columns(9).Width = 150
            DataGridView1.Columns(9).ReadOnly = True
            DataGridView1.Columns(10).HeaderText = "Комментарий продавца"
            DataGridView1.Columns(10).Width = 250
            DataGridView1.Columns(10).DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            DataGridView1.Columns(11).HeaderText = "Поисковик"
            DataGridView1.Columns(11).Width = 150
            DataGridView1.Columns(11).ReadOnly = True
            DataGridView1.Columns(12).HeaderText = "ID Статус поисковика"
            DataGridView1.Columns(12).Visible = False
            DataGridView1.Columns(12).ReadOnly = True
            DataGridView1.Columns(13).HeaderText = "Статус поисковика"
            DataGridView1.Columns(13).Width = 150
            DataGridView1.Columns(13).ReadOnly = True
            DataGridView1.Columns(14).HeaderText = "Комментарий поисковика"
            DataGridView1.Columns(14).Width = 250
            DataGridView1.Columns(14).DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            DataGridView1.Columns(15).HeaderText = "Документы продавца"
            DataGridView1.Columns(15).Width = 100
            DataGridView1.Columns(15).ReadOnly = True
            DataGridView1.Columns(16).HeaderText = "Документы поисковика"
            DataGridView1.Columns(16).Width = 100
            DataGridView1.Columns(16).ReadOnly = True
            DataGridView1.Columns(17).HeaderText = "Сумма найденного РУБ"
            DataGridView1.Columns(17).Width = 150
            DataGridView1.Columns(17).ReadOnly = True
            DataGridView1.Columns(17).DefaultCellStyle.Format = "n2"
            DataGridView1.Columns(18).HeaderText = "N запроса от клиента"
            DataGridView1.Columns(18).Width = 200
            DataGridView1.Columns(18).ReadOnly = True
            DataGridView1.Columns(19).HeaderText = "Состояние КП продавца"
            DataGridView1.Columns(19).Width = 200
            DataGridView1.Columns(19).ReadOnly = True
            DataGridView1.Columns(20).HeaderText = "Причина отказа"
            DataGridView1.Columns(20).Width = 200
            DataGridView1.Columns(20).ReadOnly = True
            DataGridView1.Columns(21).HeaderText = "Условия оплаты клиентом"
            DataGridView1.Columns(21).Width = 200
            DataGridView1.Columns(21).ReadOnly = True


            FormatDataGridView1()
            ChangeColumnsVisibility()
        End If
    End Sub

    Private Sub ChangeColumnsVisibility()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// изменение видимости колонок в зависимости от флага FullInfoFlag
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If FullInfoFlag = 0 Then
            DataGridView1.Columns(2).Visible = False
            DataGridView1.Columns(3).Visible = False
            DataGridView1.Columns(4).Visible = False
            DataGridView1.Columns(5).Visible = False
        Else
            DataGridView1.Columns(2).Visible = True
            DataGridView1.Columns(3).Visible = True
            DataGridView1.Columns(4).Visible = True
            DataGridView1.Columns(5).Visible = True
        End If
    End Sub

    Private Sub FormatDataGridView1()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по запросам на поиск
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(8).Value = -1 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightSkyBlue
                '-----поля комментарии 
                DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(236, 244, 250)
                DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(236, 244, 250)
            ElseIf DataGridView1.Rows(i).Cells(8).Value = 0 Then
                '-----Создано продавцом
                If DataGridView1.Rows(i).Cells(12).Value = 0 Then
                    '-----еще не в работе поисковиком
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.White
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.White
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 1 Then
                    '-----в работе поисковиком
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(170, 255, 143)
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(217, 255, 205)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(217, 255, 205)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(246, 255, 140)
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(252, 255, 213)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(252, 255, 213)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells(6).Value).AddDays(1) < Now() Then
                    '-----Просрочено
                    DataGridView1.Rows(i).Cells(6).Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells(8).Value = 1 Then
                '-----Продавец подтвердил предложение
                If DataGridView1.Rows(i).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(65, 255, 5)
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(170, 255, 143)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(170, 255, 143)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 3 Then
                    '-----поисковик закрыл поиск
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(255, 255, 185)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(255, 255, 185)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells(6).Value).AddDays(1) < Now() Then
                    '-----Просрочено
                    DataGridView1.Rows(i).Cells(6).Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells(8).Value = 2 Then
                '-----Продавец не подтвердил предложение
                If DataGridView1.Rows(i).Cells(12).Value = 1 Then
                    '-----в работе поисковиком
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LimeGreen
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(163, 255, 163)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(163, 255, 163)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Orange
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(255, 255, 117)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(255, 255, 117)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 3 Then
                    '-----поисковик закрыл поиск
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(245, 137, 47)
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(250, 191, 142)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(250, 191, 142)
                ElseIf DataGridView1.Rows(i).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '-----поля комментарии 
                    DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(255, 179, 179)
                    DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(255, 179, 179)
                End If
                If CDate(DataGridView1.Rows(i).Cells(6).Value).AddDays(1) < Now() Then
                    '-----Просрочено
                    DataGridView1.Rows(i).Cells(6).Style.BackColor = Color.Red
                End If
            ElseIf DataGridView1.Rows(i).Cells(8).Value = 4 Then
                '-----Продавец приостановил запрос (поставил на паузу) (3)
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(179, 129, 217)
                '-----поля комментарии 
                DataGridView1.Rows(i).Cells(10).Style.BackColor = Color.FromArgb(216, 190, 236)
                DataGridView1.Rows(i).Cells(14).Style.BackColor = Color.FromArgb(216, 190, 236)
            Else
                '-----Продавец полностью закрыл запрос (3)
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
            If Not IsDBNull(DataGridView1.Rows(i).Cells(10).Value) Then
                DataGridView1.Rows(i).Cells(0).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(1).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(2).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(3).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(4).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(5).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(6).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(7).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(8).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(9).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(11).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(12).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(13).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(15).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(16).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
                DataGridView1.Rows(i).Cells(17).ToolTipText = DataGridView1.Rows(i).Cells(10).Value
            End If
        Next
    End Sub

    Private Sub CheckRequestButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// состояние кнопок относящихся к запросу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button19.Enabled = False
            Button4.Enabled = False
            Button17.Enabled = False
            Button18.Enabled = False
            Button12.Enabled = False
            Button14.Enabled = False
        Else
            Button12.Enabled = True
            Button14.Enabled = True

            If DataGridView1.SelectedRows.Item(0).Cells(8).Value = -1 Then
                Button19.Enabled = True
                Button17.Enabled = False
                Button18.Enabled = False
                Button4.Enabled = False
            ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 0 Then
                '-----Создано продавцом
                If DataGridView1.SelectedRows.Item(0).Cells(12).Value = -1 Then
                    '-----поисковик назначен, не в работе
                    Button15.Enabled = True
                    Button19.Enabled = False
                    Button4.Enabled = True
                    Button17.Enabled = False
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 0 Then
                    '-----еще не в работе поисковиком
                    Button15.Enabled = False
                    Button19.Enabled = True
                    Button4.Enabled = True
                    Button17.Enabled = False
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                    '-----в работе поисковиком
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = True
                    Button17.Enabled = True
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                End If


            ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 1 Then
                '-----Продавец подтвердил предложение
                If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = True
                    Button17.Enabled = True
                    Button18.Enabled = True
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 3 Then
                    '-----поисковик закрыл поиск
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                End If

            ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 2 Then
                '-----Продавец не подтвердил предложение
                If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                    '-----в работе поисковиком
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = True
                    Button17.Enabled = True
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                    '-----поисковик сформировал предложение
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                    '-----поисковик отказал
                    Button15.Enabled = False
                    Button19.Enabled = False
                    Button4.Enabled = False
                    Button17.Enabled = False
                    Button18.Enabled = False
                End If

            Else
                '-----Продавец полностью закрыл запрос (3)
                Button15.Enabled = False
                Button19.Enabled = False
                Button17.Enabled = False
                Button18.Enabled = False
                Button4.Enabled = False
            End If

            'Button19.Enabled = True
        End If
    End Sub

    Private Sub LoadItems()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информациии по товарам
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_SearchItemsInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            Else
                MySQLStr = "Exec spp_SupplSearch_SearchItemsInfo 0 "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView2.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView2.Columns(0).HeaderText = "ID"
            DataGridView2.Columns(0).Width = 40
            DataGridView2.Columns(0).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(1).HeaderText = "Код товара в Скала"
            DataGridView2.Columns(1).Width = 100
            DataGridView2.Columns(1).Visible = False
            DataGridView2.Columns(2).HeaderText = "Код товара производителя"
            DataGridView2.Columns(2).Width = 100
            DataGridView2.Columns(2).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(3).HeaderText = "Название товара"
            DataGridView2.Columns(3).Width = 230
            DataGridView2.Columns(3).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(4).HeaderText = "Ед изме рения"
            DataGridView2.Columns(4).Width = 40
            DataGridView2.Columns(4).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(5).HeaderText = "Кол-во"
            DataGridView2.Columns(5).Width = 70
            DataGridView2.Columns(5).HeaderCell.Style.BackColor = Color.LightBlue
            DataGridView2.Columns(5).DefaultCellStyle.Format = "n3"
            DataGridView2.Columns(6).HeaderText = "Срок постав ки (нед)"
            DataGridView2.Columns(6).Width = 50
            DataGridView2.Columns(6).DefaultCellStyle.Format = "n2"
            DataGridView2.Columns(6).Visible = False
            DataGridView2.Columns(7).HeaderText = "Комментарии"
            DataGridView2.Columns(7).Width = 250

            FormatDataGridView2()
        End If
    End Sub

    Private Sub FormatDataGridView2()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по товарам в поиске
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView2.Rows.Count - 1

            If Not IsDBNull(DataGridView2.Rows(i).Cells(7).Value) Then
                DataGridView2.Rows(i).Cells(0).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(1).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(2).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(3).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(4).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(5).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
                DataGridView2.Rows(i).Cells(6).ToolTipText = DataGridView2.Rows(i).Cells(7).Value
            End If

        Next
    End Sub

    Private Sub CheckItemButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// состояние кнопок относящихся к товару
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button5.Enabled = False
        Else
            Button5.Enabled = True
        End If
    End Sub

    Private Sub LoadSuppliers()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка поставщиков, по которым идет поиск
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_GetSupplInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
            Else
                MySQLStr = "Exec spp_SupplSearch_GetSupplInfo 0"
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView3.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView3.Columns(0).HeaderText = "ID"
            DataGridView3.Columns(0).Width = 0
            DataGridView3.Columns(0).Visible = False
            DataGridView3.Columns(1).HeaderText = "Код поставщика в Скала"
            DataGridView3.Columns(1).Width = 80
            DataGridView3.Columns(2).HeaderText = "Поставщик"
            DataGridView3.Columns(2).Width = 190
            DataGridView3.Columns(3).HeaderText = "Коэффициент, назначенный поставщику (%)"
            DataGridView3.Columns(3).Width = 100
            DataGridView3.Columns(3).DefaultCellStyle.Format = "n2"
        End If
    End Sub

    Private Sub CheckSupplierButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// состояние кнопок поставщиков
        '//
        '////////////////////////////////////////////////////////////////////////////////


        If DataGridView1.SelectedRows.Count = 0 Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button6.Enabled = False
        Else
            If DataGridView3.SelectedRows.Count = 0 Then
                Button3.Enabled = False
                Button6.Enabled = False
                If DataGridView1.SelectedRows.Item(0).Cells(8).Value = -1 Then
                    Button1.Enabled = False
                    Button2.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 0 Then
                    '-----Создано продавцом
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 0 Then
                        '-----еще не в работе поисковиком
                        Button1.Enabled = False
                        Button2.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button1.Enabled = True
                        Button2.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = False
                        Button2.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 1 Then
                    '-----Продавец подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = True
                        Button2.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 3 Then
                        '-----поисковик закрыл поиск
                        Button1.Enabled = False
                        Button2.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 2 Then
                    '-----Продавец не подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button1.Enabled = True
                        Button2.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = False
                        Button2.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                    End If

                Else
                    '-----Продавец полностью закрыл запрос (3)
                    Button1.Enabled = False
                    Button2.Enabled = False
                End If

            Else    '/////////////////////
                If DataGridView1.SelectedRows.Item(0).Cells(8).Value = -1 Then
                    Button1.Enabled = False
                    Button2.Enabled = False
                    Button3.Enabled = False
                    Button6.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 0 Then
                    '-----Создано продавцом
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 0 Then
                        '-----еще не в работе поисковиком
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button1.Enabled = True
                        Button2.Enabled = True
                        Button3.Enabled = True
                        Button6.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 1 Then
                    '-----Продавец подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = True
                        Button2.Enabled = True
                        Button3.Enabled = True
                        Button6.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 3 Then
                        '-----поисковик закрыл поиск
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 2 Then
                    '-----Продавец не подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button1.Enabled = True
                        Button2.Enabled = True
                        Button3.Enabled = True
                        Button6.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button1.Enabled = False
                        Button2.Enabled = False
                        Button3.Enabled = False
                        Button6.Enabled = False
                    End If

                Else
                    '-----Продавец полностью закрыл запрос (3)
                    Button1.Enabled = False
                    Button2.Enabled = False
                    Button3.Enabled = False
                    Button6.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub LoadProposal()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка результатов предложения
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка поставщиков
        Dim MyDs As New DataSet                       '

        If LoadFlag = 0 Then
            If DataGridView1.SelectedRows.Count <> 0 Then
                MySQLStr = "Exec spp_SupplSearch_GetProposalInfo " & Trim(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & ", 0 "
            Else
                MySQLStr = "Exec spp_SupplSearch_GetProposalInfo 0, 0 "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                DataGridView4.DataSource = MyDs.Tables(0)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---заголовки
            DataGridView4.Columns("ID").HeaderText = "ID"
            DataGridView4.Columns("ID").Width = 40
            DataGridView4.Columns("ID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("IsSelected").HeaderText = "Пре дло жен"
            DataGridView4.Columns("IsSelected").Width = 30
            DataGridView4.Columns("IsSelected").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("SelectedBySalesman").HeaderText = "Выб ран"
            DataGridView4.Columns("SelectedBySalesman").Width = 30
            DataGridView4.Columns("SelectedBySalesman").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("ItemID").HeaderText = "ItemID"
            DataGridView4.Columns("ItemID").Width = 0
            DataGridView4.Columns("ItemID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("ItemID").Visible = False
            DataGridView4.Columns("ItemCode").HeaderText = "Код товара в Скала"
            DataGridView4.Columns("ItemCode").Width = 100
            DataGridView4.Columns("ItemCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("ItemSuppCode").HeaderText = "Код товара поставщика"
            DataGridView4.Columns("ItemSuppCode").Width = 100
            DataGridView4.Columns("ItemSuppCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("ItemName").HeaderText = "Название товара"
            DataGridView4.Columns("ItemName").Width = 180
            DataGridView4.Columns("ItemName").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("UOM").HeaderText = "Ед изме рения"
            DataGridView4.Columns("UOM").Width = 40
            DataGridView4.Columns("UOM").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("QTY").HeaderText = "Кол-во"
            DataGridView4.Columns("QTY").Width = 70
            DataGridView4.Columns("QTY").DefaultCellStyle.Format = "n3"
            DataGridView4.Columns("QTY").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("Price").HeaderText = "Закуп Цена без НДС"
            DataGridView4.Columns("Price").Width = 70
            DataGridView4.Columns("Price").DefaultCellStyle.Format = "n2"
            DataGridView4.Columns("Price").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("PriCost").HeaderText = "Расчетная с.стоим без таможни и доп расходов"
            DataGridView4.Columns("PriCost").Width = 100
            DataGridView4.Columns("PriCost").DefaultCellStyle.Format = "n2"
            DataGridView4.Columns("PriCost").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("Curr").HeaderText = "Валюта"
            DataGridView4.Columns("Curr").Width = 50
            DataGridView4.Columns("Curr").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("LeadTimeWeek").HeaderText = "Срок постав ки (нед)"
            DataGridView4.Columns("LeadTimeWeek").Width = 50
            DataGridView4.Columns("LeadTimeWeek").DefaultCellStyle.Format = "n2"
            DataGridView4.Columns("LeadTimeWeek").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("SupplierID").HeaderText = "SupplierID"
            DataGridView4.Columns("SupplierID").Width = 0
            DataGridView4.Columns("SupplierID").Visible = False
            DataGridView4.Columns("SupplierID").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("SupplierCode").HeaderText = "Код постав щика в Скала"
            DataGridView4.Columns("SupplierCode").Width = 80
            DataGridView4.Columns("SupplierCode").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("SupplierName").HeaderText = "Поставщик"
            DataGridView4.Columns("SupplierName").Width = 120
            DataGridView4.Columns("SupplierName").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("ShippingCost").HeaderText = "Коэфф поставщика"
            DataGridView4.Columns("ShippingCost").Width = 80
            DataGridView4.Columns("ShippingCost").DefaultCellStyle.Format = "n2"
            DataGridView4.Columns("ShippingCost").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("Comments").HeaderText = "Комментарий поисковика"
            DataGridView4.Columns("Comments").Width = 150
            DataGridView4.Columns("Comments").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("SalesmanComments").HeaderText = "Комментарий продавца"
            DataGridView4.Columns("SalesmanComments").Width = 150
            DataGridView4.Columns("SalesmanComments").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("AlternateTo").HeaderText = "Альтернатива товару"
            DataGridView4.Columns("AlternateTo").Width = 150
            DataGridView4.Columns("AlternateTo").SortMode = DataGridViewColumnSortMode.Programmatic
            DataGridView4.Columns("DueDate").HeaderText = "Действ. до"
            DataGridView4.Columns("DueDate").Width = 100
            DataGridView4.Columns("DueDate").SortMode = DataGridViewColumnSortMode.Programmatic

            '---сотрировка (ручная)
            SetSorting()

            FormatDataGridView4()
        End If

    End Sub

    Private Sub FormatDataGridView4()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка информации по товарам в предложении
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        For i = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Rows(i).Cells("ShippingCost").Value <> 0 Then
                DataGridView4.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                DataGridView4.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 189)
            End If

            If DataGridView4.Rows(i).Cells("IsSelected").Value = False And DataGridView4.Rows(i).Cells("SelectedBySalesman").Value = True Then
                'DataGridView4.Rows(i).DefaultCellStyle.BackColor = Color.FromArgb(255, 179, 179)
                DataGridView4.Rows(i).Cells("IsSelected").Style.BackColor = Color.FromArgb(255, 179, 179)
            Else
                'DataGridView4.Rows(i).DefaultCellStyle.BackColor = Color.White
                DataGridView4.Rows(i).Cells("IsSelected").Style.BackColor = Color.White
            End If

            If Not IsDBNull(DataGridView4.Rows(i).Cells("ShippingCost").Value) Then
                DataGridView4.Rows(i).Cells("ID").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("IsSelected").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("SelectedBySalesman").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("ItemID").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("ItemCode").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("ItemSuppCode").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("ItemName").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("UOM").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("QTY").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("Price").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("PriCost").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("Curr").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("LeadTimeWeek").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("SupplierID").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("SupplierCode").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
                DataGridView4.Rows(i).Cells("AlternateTo").ToolTipText = DataGridView4.Rows(i).Cells("SupplierName").Value
            End If

        Next
    End Sub

    Private Sub CheckProposalButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка кнопок результатов предложения
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If DataGridView1.SelectedRows.Count = 0 Then
            Button7.Enabled = False
            Button9.Enabled = False
        Else
            If DataGridView4.SelectedRows.Count = 0 Then
                Button7.Enabled = False
                Button9.Enabled = False

            Else    '/////////////////////
                If DataGridView1.SelectedRows.Item(0).Cells(8).Value = -1 Then
                    Button7.Enabled = False
                    Button9.Enabled = False
                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 0 Then
                    '-----Создано продавцом
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 0 Then
                        '-----еще не в работе поисковиком
                        Button7.Enabled = False
                        Button9.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button7.Enabled = True
                        Button9.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button7.Enabled = False
                        Button9.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button7.Enabled = False
                        Button9.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 1 Then
                    '-----Продавец подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button7.Enabled = True
                        Button9.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 3 Then
                        '-----поисковик закрыл поиск
                        Button7.Enabled = False
                        Button9.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button7.Enabled = False
                        Button9.Enabled = False
                    End If

                ElseIf DataGridView1.SelectedRows.Item(0).Cells(8).Value = 2 Then
                    '-----Продавец не подтвердил предложение
                    If DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1 Then
                        '-----в работе поисковиком
                        Button7.Enabled = True
                        Button9.Enabled = True
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 2 Then
                        '-----поисковик сформировал предложение
                        Button7.Enabled = False
                        Button9.Enabled = False
                    ElseIf DataGridView1.SelectedRows.Item(0).Cells(12).Value = 4 Then
                        '-----поисковик отказал
                        Button7.Enabled = False
                        Button9.Enabled = False
                    End If

                Else
                    '-----Продавец полностью закрыл запрос (3)
                    Button7.Enabled = False
                    Button9.Enabled = False
                End If
            End If
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// обновление информации в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub ComboBoxAct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBoxAct.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора активности запросов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню комментария
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Declarations.MyRowIndex = e.RowIndex
            Declarations.MyRequestNum = DataGridView1.Rows(Declarations.MyRowIndex).Cells(0).Value
            If DataGridView1.Rows(Declarations.MyRowIndex).Cells(8).Value = 0 Or _
                DataGridView1.Rows(Declarations.MyRowIndex).Cells(8).Value = 1 Or _
                DataGridView1.Rows(Declarations.MyRowIndex).Cells(8).Value = 2 Then
                ContextMenuStrip1.Show(MousePosition.X, MousePosition.Y)
            End If
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора запроса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadItems()
            LoadProposal()
            LoadSuppliers()
            CheckRequestButtons()
            CheckItemButtons()
            CheckSupplierButtons()
            CheckProposalButtons()
            If DataGridView1.SelectedRows.Count <> 0 Then
                Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells(0).Value
            Else
                Declarations.MyRequestNum = 0
            End If
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub DataGridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Sorted
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена сортировки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FormatDataGridView1()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запрошенных товаров в спецификацию
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If My.Settings.UseOffice = "LibreOffice" Then
            If DataGridView2.SortedColumn Is Nothing Then
                ExportToSpecToLO(0, 1)
            Else
                ExportToSpecToLO(DataGridView2.SortedColumn.Index, DataGridView2.SortOrder)
            End If

        Else
            If DataGridView2.SortedColumn Is Nothing Then
                ExportToSpecToExcel(0, 1)
            Else
                ExportToSpecToExcel(DataGridView2.SortedColumn.Index, DataGridView2.SortOrder)
            End If
        End If
    End Sub

    Private Sub ExportToSpecToExcel(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запрошенных товаров в спецификацию в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""Скандика"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "Адрес:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "Спецификация поставки"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("F1") = "Запрос N "
        With MyWRKBook.ActiveSheet.Range("F1").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("H1") = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        With MyWRKBook.ActiveSheet.Range("H1").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        '--заголовок таблицы
        MyWRKBook.ActiveSheet.Range("A10") = "N п/п"
        MyWRKBook.ActiveSheet.Range("B10") = "Код товара Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("D10") = "Наименование товара"
        MyWRKBook.ActiveSheet.Range("E10") = "Ед измерения"
        MyWRKBook.ActiveSheet.Range("F10") = "Кол-во"
        MyWRKBook.ActiveSheet.Range("G10") = "Цена без НДС"
        MyWRKBook.ActiveSheet.Range("H10") = "Сумма без НДС"
        MyWRKBook.ActiveSheet.Range("I10") = "Срок поставки (нед)"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:I10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:I10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:I10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:I10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12

        '--Вывод строк спецификации
        MySQLStr = "SELECT ISNULL(tbl_SupplSearchItems.ItemID, N'') AS ItemCode, ISNULL(tbl_SupplSearchItems.ItemSuppID, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearchItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'шт' THEN 'pcs(шт.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'кг' THEN 'kg (кг)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'л' THEN 'litre (литр)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'компл' THEN 'set (Компл.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems.QTY, '' AS Price, ISNULL(tbl_SupplSearchItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearchItems.Comments, '') AS Comments "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearchItems.ItemID = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearchItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearchItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
            Case 2
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 3
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("QTY").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value
                MyWRKBook.ActiveSheet.Range("J" & CStr(i + 11)) = Declarations.MyRec.Fields("Comments").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=ЕСЛИ(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---Выгрузка справочной информации
        MyWRKBook.ActiveSheet.Range("N10") = "Единица измерения"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(шт.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (м)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (кг)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (км)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (литр)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (Упак.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (Компл.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (пара)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("D1").Locked = False
        MyWRKBook.ActiveSheet.Range("D1").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportToSpecToLO(ByVal MyCol As Integer, ByVal MySort As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запрошенных товаров в спецификацию в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---ширина колонок
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        '---защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""Скандика"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "Адрес:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "Спецификация поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        oSheet.getCellRangeByName("F1").String = "Запрос N"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "F1", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "F1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "F1", 9)

        oSheet.getCellRangeByName("H1").String = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "H1", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "H1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "H1", 9)

        '--заголовок таблицы
        oSheet.getCellRangeByName("A10").String = "N п/п"
        oSheet.getCellRangeByName("B10").String = "Код товара Scala"
        oSheet.getCellRangeByName("C10").String = "Код товара поставщика"
        oSheet.getCellRangeByName("D10").String = "Наименование товара"
        oSheet.getCellRangeByName("E10").String = "Ед измерения"
        oSheet.getCellRangeByName("F10").String = "Кол-во"
        oSheet.getCellRangeByName("G10").String = "Цена без НДС"
        oSheet.getCellRangeByName("H10").String = "Сумма без НДС"
        oSheet.getCellRangeByName("I10").String = "Срок поставки (нед)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--Вывод строк спецификации
        MySQLStr = "SELECT ISNULL(tbl_SupplSearchItems.ItemID, N'') AS ItemCode, ISNULL(tbl_SupplSearchItems.ItemSuppID, N'') AS SuppItemCode, "
        MySQLStr = MySQLStr & "CASE WHEN LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) + LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) "
        MySQLStr = MySQLStr & "= '' THEN ISNULL(tbl_SupplSearchItems.ItemName, '') ELSE LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(SC010300.SC01002, ''))) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(ISNULL(SC010300.SC01003, ''))))) END AS ItemName, CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'шт' THEN 'pcs(шт.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'кг' THEN 'kg (кг)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'л' THEN 'litre (литр)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') "
        MySQLStr = MySQLStr & "= 'компл' THEN 'set (Компл.)' ELSE CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)' END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_SupplSearchItems.QTY, '' AS Price, ISNULL(tbl_SupplSearchItems.LeadTimeWeek, 1) AS WeekQTY, "
        MySQLStr = MySQLStr & "ISNULL(tbl_SupplSearchItems.Comments, '') AS Comments "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_SupplSearchItems.ItemID = SC010300.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_SupplSearchItems.UOM = View_1.Expr1 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearchItems.SupplSearchID = " & DataGridView1.SelectedRows.Item(0).Cells("ID").Value & ") "
        Select Case MyCol
            Case 0
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
            Case 2
                MySQLStr = MySQLStr & "ORDER BY SuppItemCode "
            Case 3
                MySQLStr = MySQLStr & "ORDER BY ItemName "
            Case 4
                MySQLStr = MySQLStr & "ORDER BY UOM "
            Case 5
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.QTY "
            Case Else
                MySQLStr = MySQLStr & "ORDER BY tbl_SupplSearchItems.ID "
        End Select
        If MySort = 2 Then
            MySQLStr = MySQLStr & "Desc "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("QTY").Value
                If Not Declarations.MyRec.Fields("Price").Value.Equals("") Then
                    oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("Price").Value
                End If
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value
                oSheet.getCellRangeByName("J" & CStr(i + 11)).String = Declarations.MyRec.Fields("Comments").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----Защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---Выгрузка справочной информации
        oSheet.getCellRangeByName("N10").String = "Единица измерения"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(шт.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (м)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (кг)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (км)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (литр)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (Упак.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (Компл.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (пара)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----в начало файла
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----закрытие паролем
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----видимость
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление поставщика в условия поиска
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer

        MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value

        myValue = ""
        myValue = InputBox("Введите название поставщика", "Поставщик", "")

        If Trim(myValue) <> "" Then
            MySQLStr = "INSERT INTO tbl_SupplSearch_Suppliers "
            MySQLStr = MySQLStr & "(SupplSearchID, SupplierCode, SupplierName) "
            MySQLStr = MySQLStr & "VALUES (" & CStr(MyID) & ", "
            MySQLStr = MySQLStr & "N'', "
            MySQLStr = MySQLStr & "N'" & Trim(myValue) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "exec spp_SupplSearch_CheckProposal " & CStr(MyID)
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            LoadSuppliers()
            LoadProposal()
            CheckSupplierButtons()
            CheckProposalButtons()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление поставщика из условий поиска
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyID As Integer

        MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value

        MySQLStr = "DELETE FROM tbl_SupplSearch_Suppliers "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView3.SelectedRows.Item(0).Cells(0).Value & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "exec spp_SupplSearch_CheckProposal " & CStr(MyID)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        LoadSuppliers()
        LoadProposal()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление поставщика из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyID As Integer

        MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "ADD"
        MySupplierSelect.ShowDialog()

        MySQLStr = "exec spp_SupplSearch_CheckProposal " & CStr(MyID)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        LoadSuppliers()
        LoadProposal()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление поставщика из Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyID As Integer

        MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "UPDATE"
        MySupplierSelect.ShowDialog()

        MySQLStr = "exec spp_SupplSearch_CheckProposal " & CStr(MyID)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        LoadSuppliers()
        LoadProposal()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Назначение поисковика
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySearcherList = New SearcherList
        MySearcherList.ShowDialog()

        Application.DoEvents()
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// отказ в выполнении поиска
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim MyRez As Object
        Dim EmailStr As String
        Dim RequestStatus As String

        MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        myValue = ""
        myValue = InputBox("Введите комментарий (если необходимо)", "Комментарий", "")
        MyRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = vbYes Then
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr + "SET SearchStatus = 4, "
            MySQLStr = MySQLStr + "SearcherComments = ISNULL(SearcherComments,'') + '" + Chr(10) + Chr(13) + " --" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
            MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '--------------отправка почты
            EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString, 4)))
            If EmailStr = "" Then
                MsgBox("Для пользователя " & DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString & " в БД не занесена почта. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            Else
                RequestStatus = "Отказано в поиске"
                SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString(), _
                   EmailStr, DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString(), _
                   RequestStatus, DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString())
            End If
            '---------------------------
            LoadRequests()
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Item(0, i).Value = MyID Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Exit For
                End If
            Next
            CheckRequestButtons()
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// предложение варианта поиска
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim MyRez As Object
        Dim EmailStr As String
        Dim RequestStatus As String

        If CheckReadinesForProposal() = True Then
            If CheckSalesmanConfirmation() = False Then
                MySQLStr = "В предложении есть запасы, которые продавец пометил как необходимые, " + Chr(10) + Chr(13)
                MySQLStr = MySQLStr + "но которые не помечены галочкой с вашей стороны и, возможно, не обработаны. " + Chr(10) + Chr(13)
                MySQLStr = MySQLStr + "Вы уверены, что необходимо отправить предложение именно в этом варианте?"
                MyRez = MsgBox(MySQLStr, MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
            myValue = ""
            myValue = InputBox("Введите комментарий (если необходимо)", "Комментарий", "")
            MyRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                MySQLStr = "UPDATE tbl_SupplSearch "
                MySQLStr = MySQLStr + "SET SearchStatus = 2, "
                MySQLStr = MySQLStr + "SalesStatus = CASE WHEN SalesStatus = 1 THEN 2 ELSE SalesStatus END, "
                MySQLStr = MySQLStr + "SearcherComments = ISNULL(SearcherComments,'') + '" + Chr(10) + Chr(13) + "--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
                MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '--------------отправка почты
                EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString, 4)))
                If EmailStr = "" Then
                    MsgBox("Для пользователя " & DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString & " в БД не занесена почта. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    RequestStatus = "Предложен вариант"
                    SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString(), _
                       EmailStr, DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString(), _
                       RequestStatus, DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString())
                End If
                '---------------------------
                LoadRequests()
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Item(0, i).Value = MyID Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit For
                    End If
                Next
                CheckRequestButtons()
            End If
        End If
    End Sub

    Private Function CheckSalesmanConfirmation() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка выполнения пожелания продавца по поиску
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim NonConfCount As Integer

        NonConfCount = 0
        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            '---Отметки продавца
            If DataGridView4.Item("IsSelected", i).Value = False And DataGridView4.Item("SelectedBySalesman", i).Value = True Then
                NonConfCount = NonConfCount + 1
            End If
        Next

        If NonConfCount > 0 Then
            CheckSalesmanConfirmation = False
            Exit Function
        End If
        CheckSalesmanConfirmation = True
    End Function

    Private Function CheckReadinesForProposal() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка готовности к предложению варианта поиска
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim PropCount As Integer
        Dim MyRez As MsgBoxResult

        PropCount = 0
        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Item("IsSelected", i).Value = True Then
                '---------выбран
                '---Цена
                If DataGridView4.Item("Price", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("Price", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесена закупочная цена ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForProposal = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf DataGridView4.Item("Price", i).Value = 0 Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесена закупочная цена ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForProposal = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If
                '---Срок поставки
                If DataGridView4.Item("LeadTimeWeek", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("LeadTimeWeek", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен срок поставки ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForProposal = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf DataGridView4.Item("LeadTimeWeek", i).Value = 0 Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен срок поставки ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForProposal = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If

                PropCount = PropCount + 1
            End If
        Next
        If PropCount = 0 Then
            '---не выбран ни один товар
            MyRez = MsgBox("Вы не выбрали ни одного товара в предложении поставки. Продолжить? ", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = MsgBoxResult.No Then
                CheckReadinesForProposal = False
                Exit Function
            End If
        End If
        CheckReadinesForProposal = True
    End Function

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование предложенного варианта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Declarations.MyItemSrchID = DataGridView4.SelectedRows.Item(0).Cells("ID").Value
        MyAddItem = New AddItem
        MyAddItem.StartParam = "Edit"
        MyAddItem.ShowDialog()
        LoadProposal()

        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Item("ID", i).Value = Declarations.MyItemSrchID Then
                DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                Exit For
            End If
        Next
        CheckProposalButtons()
        ReCalculateSumm(DataGridView1.SelectedRows.Item(0).Cells("ID").Value)
    End Sub

    Private Sub DataGridView4_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellDoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование предложенного варианта по двойному щелчку
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyItemSrchID = DataGridView4.SelectedRows.Item(0).Cells("ID").Value
        MyAddItem = New AddItem
        MyAddItem.StartParam = "Edit"
        MyAddItem.ShowDialog()
        LoadProposal()

        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Item("ID", i).Value = Declarations.MyItemSrchID Then
                DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                Exit For
            End If
        Next
        CheckProposalButtons()
        ReCalculateSumm(DataGridView1.SelectedRows.Item(0).Cells("ID").Value)
    End Sub


    Private Sub DataGridView4_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView4.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния выбора поставляемого элемента
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyID As Integer

        If e.Button = Windows.Forms.MouseButtons.Left Then
            '------Колонка с комбобоксами
            If e.ColumnIndex = 1 Then
                If e.RowIndex <> -1 Then
                    If Button9.Enabled = True Then
                        MyID = DataGridView4.SelectedRows.Item(0).Cells("ID").Value
                        ChangeReadyState(MyID, DataGridView4.SelectedRows.Item(0).Cells("IsSelected").Value, DataGridView4.SelectedRows.Item(0).Cells("ItemID").Value)
                        '---загрузка данных
                        LoadProposal()
                        '---текущей строкой сделать редактируемую
                        For i As Integer = 0 To DataGridView4.Rows.Count - 1
                            If Trim(DataGridView4.Item("ID", i).Value.ToString) = MyID Then
                                DataGridView4.CurrentCell = DataGridView4.Item("IsSelected", i)
                            End If
                        Next
                        CheckProposalButtons()
                        ReCalculateSumm(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
                    End If
                Else
                    If IsNothing(SortColumnNum) = True Then
                        SortColumnNum = e.ColumnIndex
                        SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                    Else
                        If DataGridView4.Columns(SortColumnNum).Equals(DataGridView4.Columns(e.ColumnIndex)) Then
                            '------колонка та же - меняем сортировку
                            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                                SortColOrder = System.ComponentModel.ListSortDirection.Descending
                            Else
                                SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                            End If
                        Else
                            '------колонка новая - сортировка по возрастающей
                            SortColumnNum = e.ColumnIndex
                            SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                        End If
                    End If
                    SetSorting()
                    FormatDataGridView4()
                End If
            Else '------остальные колонки
                If e.RowIndex = -1 Then
                    If IsNothing(SortColumnNum) = True Then
                        SortColumnNum = e.ColumnIndex
                        SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                    Else
                        If DataGridView4.Columns(SortColumnNum).Equals(DataGridView4.Columns(e.ColumnIndex)) Then
                            '------колонка та же - меняем сортировку
                            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                                SortColOrder = System.ComponentModel.ListSortDirection.Descending
                            Else
                                SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                            End If
                        Else
                            '------колонка новая - сортировка по возрастающей
                            SortColumnNum = e.ColumnIndex
                            SortColOrder = System.ComponentModel.ListSortDirection.Ascending
                        End If
                    End If
                    SetSorting()
                    FormatDataGridView4()
                End If
            End If
        End If
    End Sub

    Private Sub SetSorting()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена сортировки таблицы
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If IsNothing(SortColumnNum) = False Then
            DataGridView4.Sort(DataGridView4.Columns(SortColumnNum), SortColOrder)
            If SortColOrder = System.ComponentModel.ListSortDirection.Ascending Then
                DataGridView4.Columns(SortColumnNum).HeaderCell.SortGlyphDirection = SortOrder.Ascending
            Else
                DataGridView4.Columns(SortColumnNum).HeaderCell.SortGlyphDirection = SortOrder.Descending
            End If
        End If
    End Sub

    Private Sub ChangeReadyState(ByVal MyID As Integer, ByVal CurrState As Boolean, ByVal ItemID As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния выбранности элемента в предложении
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----выбор только 1 поставщика
        'MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
        'MySQLStr = MySQLStr & "SET IsSelected = 0 "
        'MySQLStr = MySQLStr & "WHERE (ItemID = " & ItemID.ToString & ") "
        'InitMyConn(False)
        'Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
        If CurrState = False Then
            MySQLStr = MySQLStr & "SET IsSelected = 1 "
        Else
            MySQLStr = MySQLStr & "SET IsSelected = 0 "
        End If
        MySQLStr = MySQLStr & "WHERE (ID = " & MyID.ToString & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна с аттачментами продавца
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Sales"
        MyAttachmentsList.WhoStart = "Search"
        MyAttachmentsList.MyPlace = "List"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна с аттачментами поисковика
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells(0).Value
        MyAttachmentsList = New AttachmentsList
        MyAttachmentsList.AttType = "Search"
        MyAttachmentsList.WhoStart = "Search"
        MyAttachmentsList.MyPlace = "List"
        MyAttachmentsList.ShowDialog()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Информация о завершении всех действий и предложение по закрытию заявки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim MyID As Integer
        Dim MyRez As Object
        Dim EmailStr As String
        Dim RequestStatus As String

        If CheckReadinesForClose() = True Then
            If CheckSalesmanConfirmation() = False Then
                MySQLStr = "В предложении есть запасы, которые продавец пометил как необходимые, " + Chr(10) + Chr(13)
                MySQLStr = MySQLStr + "но которые не помечены галочкой с вашей стороны и, возможно, не обработаны. " + Chr(10) + Chr(13)
                MySQLStr = MySQLStr + "Вы уверены, что необходимо отправить предложение именно в этом варианте?"
                MyRez = MsgBox(MySQLStr, MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value
            myValue = ""
            myValue = InputBox("Введите комментарий (если необходимо)", "Комментарий", "")
            MyRez = MsgBox("Вы уверены?", MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = vbYes Then
                MySQLStr = "UPDATE tbl_SupplSearch "
                MySQLStr = MySQLStr + "SET SearchStatus = 3, "
                MySQLStr = MySQLStr + "SearcherComments = ISNULL(SearcherComments,'') + '" + Chr(10) + Chr(13) + "--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
                MySQLStr = MySQLStr + "WHERE (ID = " & MyID & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                '--------------отправка почты
                EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString, 4)))
                If EmailStr = "" Then
                    MsgBox("Для пользователя " & DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString & " в БД не занесена почта. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                Else
                    RequestStatus = "Поиск выполнен"
                    SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString(), _
                       EmailStr, DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString(), _
                       RequestStatus, DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString())
                End If
                '---------------------------
                LoadRequests()
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Item(0, i).Value = MyID Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Exit For
                    End If
                Next
                CheckRequestButtons()
            End If
        End If
    End Sub

    Private Function CheckReadinesForClose() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка готовности к завершению всех действий и предложению по закрытию заявки
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim PropCount As Integer

        PropCount = 0
        For i As Integer = 0 To DataGridView4.Rows.Count - 1
            If DataGridView4.Item("IsSelected", i).Value = True Then
                '---------выбран
                '---Код Скала
                If DataGridView4.Item("ItemCode", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("ItemCode", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен код товара в Scala ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf Trim(DataGridView4.Item("ItemCode", i).Value.ToString).Equals("") Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен код товара в Scala ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If
                '---Цена
                If DataGridView4.Item("Price", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("Price", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесена закупочная цена ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf DataGridView4.Item("Price", i).Value = 0 Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесена закупочная цена ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If
                '---Срок поставки
                If DataGridView4.Item("LeadTimeWeek", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("LeadTimeWeek", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен срок поставки ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf DataGridView4.Item("LeadTimeWeek", i).Value = 0 Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Не занесен срок поставки ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If
                '---Код поставщика в Scala
                If DataGridView4.Item("SupplierCode", i).Value Is Nothing Or IsDBNull(DataGridView4.Item("SupplierCode", i).Value) Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Нет кода поставщика в Scala ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                ElseIf Trim(DataGridView4.Item("SupplierCode", i).Value.ToString).Equals("") Then
                    MsgBox("Строка с кодом " + DataGridView4.Item("ID", i).Value.ToString + ". Нет кода поставщика в Scala ", MsgBoxStyle.Critical, "Внимание!")
                    CheckReadinesForClose = False
                    DataGridView4.CurrentCell = DataGridView4.Item("ID", i)
                    Exit Function
                End If
                PropCount = PropCount + 1
            End If
        Next
        If PropCount = 0 Then
            '---не выбран ни один товар
            MsgBox("Вы не выбрали ни одного товара в предложении поставки ", MsgBoxStyle.Critical, "Внимание!")
            CheckReadinesForClose = False
            Exit Function
        End If
        CheckReadinesForClose = True
    End Function

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации из спецификации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MyTxt As String

        MyTxt = "Для обновления данных вам необходимо использовать файл Excel, в котором начиная со строки 11 указать: " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "в колонке A указать номер по порядку (не обязательный параметр), " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке B указать код запаса Scala. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке C указать код запаса поставщика. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке D указать название запаса . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке E выбрать единицу измерения запаса . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке F проставить количество запаса . " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке G проставить цену запаса без НДС. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "В колонке I проставить срок поставки в неделях. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Если товар в наличии - срок поставки = 0. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Строки должны быть заполнены без пропусков. " & Chr(13) & Chr(10)
        MyTxt = MyTxt & "Все колонки должны быть заполнены, кроме B: и C:" & Chr(13) & Chr(10)
        MyTxt = MyTxt & "При импорте валюта будет проставлена РУБЛИ!!" & Chr(13) & Chr(10)
        MyTxt = MyTxt & "У Вас есть подготовленный файл Excel и вы готовы начать импорт?" & Chr(13) & Chr(10)
        MyRez = MsgBox(MyTxt, MsgBoxStyle.OkCancel, "Внимание!")

        If (MyRez = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                OpenFileDialog2.ShowDialog()
                If (OpenFileDialog2.FileName = "") Then
                Else
                    Declarations.ImportFileName = OpenFileDialog2.FileName
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells(0).Value
                    Declarations.MySupplierID = DataGridView3.SelectedRows.Item(0).Cells(0).Value
                    Declarations.MySupplierCode = DataGridView3.SelectedRows.Item(0).Cells(1).Value
                    Declarations.MySupplierName = DataGridView3.SelectedRows.Item(0).Cells(2).Value
                    UpdateRequestDataFromLO()
                End If
            Else
                OpenFileDialog1.ShowDialog()
                If (OpenFileDialog1.FileName = "") Then
                Else
                    Declarations.ImportFileName = OpenFileDialog1.FileName
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    Declarations.MyRequestNum = DataGridView1.SelectedRows.Item(0).Cells(0).Value
                    Declarations.MySupplierID = DataGridView3.SelectedRows.Item(0).Cells(0).Value
                    Declarations.MySupplierCode = DataGridView3.SelectedRows.Item(0).Cells(1).Value
                    Declarations.MySupplierName = DataGridView3.SelectedRows.Item(0).Cells(2).Value
                    UpdateRequestDataFromExcel()
                End If
            End If
            Me.Cursor = Cursors.Default
            LoadProposal()
            CheckProposalButtons()
            ReCalculateSumm(DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна проверки заказов на закупку 0 типа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyOrdersList Is Nothing Then
            MyOrdersList = New OrdersList
            MyOrdersList.Show()
        Else
            MyOrdersList.Close()
            MyOrdersList = New OrdersList
            MyOrdersList.Show()
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// нажатие кнопки полной / сокращенной информации
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If FullInfoFlag = 0 Then
            FullInfoFlag = 1
            Button10.Text = "Сокращенная информация"
        Else
            FullInfoFlag = 0
            Button10.Text = "Полная информация"
        End If
        ChangeColumnsVisibility()
    End Sub


    Private Sub ReCalculateSumm(ByVal SearchRequestNum As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление суммы найденного
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MySumm As Double

        MySQLStr = "SELECT ISNULL(SUM(ISNULL(QTY, 0) * ISNULL(Price, 0)), 0) AS SearchSumm "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_PropItems "
        MySQLStr = MySQLStr & "WHERE (IsSelected = 1) "
        MySQLStr = MySQLStr & "AND (SupplSearchID = " & CStr(SearchRequestNum) & ") "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Невозможно пересчитать сумму найденного. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
        Else
            MySumm = Declarations.MyRec.Fields("SearchSumm").Value
            trycloseMyRec()
            DataGridView1.SelectedRows.Item(0).Cells(17).Value = MySumm
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запросов на поиск
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            ExportReqToLO()
        Else
            ExportReqToExcel()
        End If
        Me.Cursor = Cursors.Default
        
    End Sub

    Private Sub ExportReqToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запросов на поиск в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyObj As Object
        Dim MyWRKBook As Object

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---------------------заголовок-----------------------------
        MyWRKBook.ActiveSheet.Range("B1") = "Список запросов на поиск"
        MyWRKBook.ActiveSheet.Range("A3") = "N запроса"
        MyWRKBook.ActiveSheet.Range("B3") = "Дата запроса"
        MyWRKBook.ActiveSheet.Range("C3") = "Название клиента"
        MyWRKBook.ActiveSheet.Range("D3") = "Контактное лицо"
        MyWRKBook.ActiveSheet.Range("E3") = "Телефон"
        MyWRKBook.ActiveSheet.Range("F3") = "Email клиента"
        MyWRKBook.ActiveSheet.Range("G3") = "Срок представления КП"
        MyWRKBook.ActiveSheet.Range("H3") = "Продавец"
        MyWRKBook.ActiveSheet.Range("I3") = "ID Статус продавца"
        MyWRKBook.ActiveSheet.Range("J3") = "Статус продавца"
        MyWRKBook.ActiveSheet.Range("K3") = "          Комментарий продавца          "
        MyWRKBook.ActiveSheet.Range("L3") = "Поисковик"
        MyWRKBook.ActiveSheet.Range("M3") = "ID Статус поисковика"
        MyWRKBook.ActiveSheet.Range("N3") = "Статус поисковика"
        MyWRKBook.ActiveSheet.Range("O3") = "          Комментарий поисковика          "
        MyWRKBook.ActiveSheet.Range("P3") = "Документы продавца"
        MyWRKBook.ActiveSheet.Range("Q3") = "Документы поисковика"
        MyWRKBook.ActiveSheet.Range("R3") = "Сумма найденного"

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 15

        '---Форматирование заголовка
        MyWRKBook.ActiveSheet.Range("A3:R3").Select()
        MyWRKBook.ActiveSheet.Range("A3:R3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:R3").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A3:R3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:R3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:R3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:R3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:R3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:R3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("B1").Select()
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:R3").Select()
        MyWRKBook.ActiveSheet.Range("A3:R3").Font.Bold = True

        MyWRKBook.ActiveSheet.Cells.Select()
        MyWRKBook.ActiveSheet.Cells.EntireColumn.AutoFit()


        '---------------------таблица-------------------------------
        If ComboBoxAct.Text = "Только активные для поисковика" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-1, N'' "
        ElseIf ComboBoxAct.Text = "Только активные" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "0, N'' "
        ElseIf ComboBoxAct.Text = "Все запросы" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "1, N'' "
        ElseIf ComboBoxAct.Text = "Нераспределенные" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "2, N'' "
        ElseIf ComboBoxAct.Text = "Только активные всех поисковиков для поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-3, N'' "
        ElseIf ComboBoxAct.Text = "Только активные всех поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "3, N'' "
        ElseIf ComboBoxAct.Text = "Все запросы всех поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "4, N'' "
        Else
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo 0, 0, N'' "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A4").CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Cells.Select()
        MyWRKBook.ActiveSheet.Cells.WrapText = True

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportReqToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка запросов на поиск в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---ширина колонок
        oSheet.getColumns().getByName("A").Width = 2490
        oSheet.getColumns().getByName("B").Width = 4060
        oSheet.getColumns().getByName("C").Width = 8960
        oSheet.getColumns().getByName("D").Width = 6020
        oSheet.getColumns().getByName("E").Width = 4060
        oSheet.getColumns().getByName("F").Width = 6020
        oSheet.getColumns().getByName("G").Width = 4060
        oSheet.getColumns().getByName("H").Width = 5040
        oSheet.getColumns().getByName("I").Width = 5040
        oSheet.getColumns().getByName("J").Width = 9940
        oSheet.getColumns().getByName("K").Width = 5040
        oSheet.getColumns().getByName("L").Width = 5040
        oSheet.getColumns().getByName("M").Width = 9940
        oSheet.getColumns().getByName("N").Width = 4060
        oSheet.getColumns().getByName("O").Width = 4060
        oSheet.getColumns().getByName("P").Width = 6020

        '---колонки
        oSheet.getCellRangeByName("B1").String = "Список запросов на поиск поставщика"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B1", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B1", 11)

        oSheet.getCellRangeByName("A3").String = "N запроса"
        oSheet.getCellRangeByName("B3").String = "Дата запроса"
        oSheet.getCellRangeByName("C3").String = "Клиент"
        oSheet.getCellRangeByName("D3").String = "Контакт клиента"
        oSheet.getCellRangeByName("E3").String = "Телефон контакта"
        oSheet.getCellRangeByName("F3").String = "EMail контакта"
        oSheet.getCellRangeByName("G3").String = "Запрошенная дата"
        oSheet.getCellRangeByName("H3").String = "Продавец"
        oSheet.getCellRangeByName("I3").String = "Статус продавца"
        oSheet.getCellRangeByName("J3").String = "Комментарий продавца"
        oSheet.getCellRangeByName("K3").String = "Поисковик"
        oSheet.getCellRangeByName("L3").String = "Статус поисковика"
        oSheet.getCellRangeByName("M3").String = "Комментарий поисковика"
        oSheet.getCellRangeByName("N3").String = "Документы продавца"
        oSheet.getCellRangeByName("O3").String = "Документы поисковика"
        oSheet.getCellRangeByName("P3").String = "Сумма найденного"

        Dim i As Integer
        i = 3
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).CellBackColor = RGB(174, 249, 255)  '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":P" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":P" & CStr(i))
        
        '---таблица
        If ComboBoxAct.Text = "Только активные для поисковика" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-1, N'' "
        ElseIf ComboBoxAct.Text = "Только активные" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "0, N'' "
        ElseIf ComboBoxAct.Text = "Все запросы" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "1, N'' "
        ElseIf ComboBoxAct.Text = "Нераспределенные" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "2, N'' "
        ElseIf ComboBoxAct.Text = "Только активные всех поисковиков для поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "-3, N'' "
        ElseIf ComboBoxAct.Text = "Только активные всех поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "3, N'' "
        ElseIf ComboBoxAct.Text = "Все запросы всех поисковиков" Then
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo " + Declarations.PurchCode + ", " + "4, N'' "
        Else
            MySQLStr = "exec spp_SupplSearch_SearchPRequestInfo 0, 0, N'' "
        End If

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            'i = 0
            'Declarations.MyRec.MoveFirst()
            'While Not Declarations.MyRec.EOF
            '    oSheet.getCellRangeByName("A" & CStr(i + 4)).String = Declarations.MyRec.Fields("ID").Value
            '    oSheet.getCellRangeByName("B" & CStr(i + 4)).Value = Declarations.MyRec.Fields("StartDate").Value
            '    LOFormatCells(oServiceManager, oDispatcher, oFrame, "B" & CStr(i + 4) & ":B" & CStr(i + 4), 51)
            '    oSheet.getCellRangeByName("C" & CStr(i + 4)).String = Declarations.MyRec.Fields("Customer").Value
            '    oSheet.getCellRangeByName("D" & CStr(i + 4)).String = Declarations.MyRec.Fields("CustomerContactName").Value
            '    oSheet.getCellRangeByName("E" & CStr(i + 4)).String = Declarations.MyRec.Fields("CustomerPhone").Value
            '    oSheet.getCellRangeByName("F" & CStr(i + 4)).String = Declarations.MyRec.Fields("CustomerEmail").Value
            '    oSheet.getCellRangeByName("G" & CStr(i + 4)).Value = Declarations.MyRec.Fields("RequestDate").Value
            '    LOFormatCells(oServiceManager, oDispatcher, oFrame, "G" & CStr(i + 4) & ":G" & CStr(i + 4), 51)
            '    oSheet.getCellRangeByName("H" & CStr(i + 4)).String = Declarations.MyRec.Fields("Salesman").Value
            '    oSheet.getCellRangeByName("I" & CStr(i + 4)).String = Declarations.MyRec.Fields("SalesStatus").Value
            '    oSheet.getCellRangeByName("J" & CStr(i + 4)).String = Declarations.MyRec.Fields("Comments").Value
            '    oSheet.getCellRangeByName("K" & CStr(i + 4)).String = Declarations.MyRec.Fields("Searcher").Value
            '    If Not IsDBNull(Declarations.MyRec.Fields("SearchStatus").Value) Then
            '        oSheet.getCellRangeByName("L" & CStr(i + 4)).String = Declarations.MyRec.Fields("SearchStatus").Value
            '    End If
            '    If Not IsDBNull(Declarations.MyRec.Fields("SearcherComments").Value) Then
            '        oSheet.getCellRangeByName("M" & CStr(i + 4)).String = Declarations.MyRec.Fields("SearcherComments").Value
            '    End If
            '    If Not IsDBNull(Declarations.MyRec.Fields("CCSal").Value) Then
            '        oSheet.getCellRangeByName("N" & CStr(i + 4)).Value = Declarations.MyRec.Fields("CCSal").Value
            '    End If
            '    If Not IsDBNull(Declarations.MyRec.Fields("CCSearch").Value) Then
            '        oSheet.getCellRangeByName("O" & CStr(i + 4)).Value = Declarations.MyRec.Fields("CCSearch").Value
            '    End If
            '    If Not IsDBNull(Declarations.MyRec.Fields("SearchSumm").Value) Then
            '        oSheet.getCellRangeByName("P" & CStr(i + 4)).Value = Declarations.MyRec.Fields("SearchSumm").Value
            '    End If
            '    LOFormatCells(oServiceManager, oDispatcher, oFrame, "P" & CStr(i + 4) & ":P" & CStr(i + 4), 4)

            '    Declarations.MyRec.MoveNext()
            '    i = i + 1
            'End While
            i = 4
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(15)
                MyArr(0) = CInt(Declarations.MyRec.Fields(0).Value)
                MyArr(1) = Declarations.MyRec.Fields(1).Value.ToOADate
                MyArr(2) = Declarations.MyRec.Fields(2).Value.ToString
                MyArr(3) = Declarations.MyRec.Fields(3).Value.ToString
                MyArr(4) = Declarations.MyRec.Fields(4).Value.ToString
                MyArr(5) = Declarations.MyRec.Fields(5).Value.ToString
                MyArr(6) = Declarations.MyRec.Fields(6).Value.ToOADate
                MyArr(7) = Declarations.MyRec.Fields(7).Value.ToString
                MyArr(8) = Declarations.MyRec.Fields(9).Value.ToString
                MyArr(9) = Declarations.MyRec.Fields(10).Value.ToString
                MyArr(10) = Declarations.MyRec.Fields(11).Value.ToString
                MyArr(11) = Declarations.MyRec.Fields(13).Value.ToString
                If Not IsDBNull(Declarations.MyRec.Fields(14).Value) Then
                    MyArr(12) = Declarations.MyRec.Fields(14).Value
                Else
                    MyArr(12) = ""
                End If
                If Not IsDBNull(Declarations.MyRec.Fields(15).Value) Then
                    MyArr(13) = Declarations.MyRec.Fields(15).Value.ToString
                Else
                    MyArr(13) = ""
                End If
                If Not IsDBNull(Declarations.MyRec.Fields(16).Value) Then
                    MyArr(14) = Declarations.MyRec.Fields(16).Value.ToString
                Else
                    MyArr(14) = ""
                End If
                If Not IsDBNull(Declarations.MyRec.Fields(17).Value) Then
                    MyArr(15) = CDbl(Declarations.MyRec.Fields(17).Value)
                Else
                    MyArr(15) = 0
                End If
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":P" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
        End If

        '---Формат таблицы
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + j - 1), "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + j - 1), 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A4:P" & CStr(i + j - 1))
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A4:P" & CStr(i + j - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + j - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + j - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A4:P" & CStr(i + j - 1)).BottomBorder = LineFormat

        '---ячейки с датами 
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "B4:B" & CStr(i + j - 1), 36)
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "G4:G" & CStr(i + j - 1), 36)

        '---
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Приемка заявки в работу поисковиком
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim EmailStr As String
        Dim RequestStatus As String

        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr & "SET SearchStatus = 1 "
        MySQLStr = MySQLStr & "WHERE (ID = " & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString() & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        '--------------отправка почты
        EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString, 4)))
        If EmailStr = "" Then
            MsgBox("Для пользователя " & DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString & " в БД не занесена почта. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
        Else
            RequestStatus = "Принят в работу"
            SendInfoByEmail(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString(), _
               EmailStr, DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString(), _
               RequestStatus, DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString())
        End If
        '---------------------------
        DataGridView1.SelectedRows.Item(0).Cells(12).Value = 1
        DataGridView1.SelectedRows.Item(0).Cells(13).Value = "Поиск начат"
        FormatDataGridView1()
        CheckRequestButtons()
        CheckSupplierButtons()
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Снятие фильтра
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        LoadRequests()
        LoadItems()
        LoadSuppliers()
        LoadProposal()
        CheckRequestButtons()
        CheckItemButtons()
        CheckSupplierButtons()
        CheckProposalButtons()
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор контекстного меню ввод комментария
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim myValue As String
        Dim EmailStr As String

        myValue = ""
        myValue = InputBox("Введите комментарий", "Комментарий", "")
        If myValue <> "" Then
            MySQLStr = "UPDATE tbl_SupplSearch "
            MySQLStr = MySQLStr + "SET SearcherComments = ISNULL(SearcherComments, '') + '" + CStr(Chr(10) + Chr(13)) + "' + '--" + Format(Now, "dd/MM/yyyy HH:mm") + "-->' + N'" & myValue & "' "
            MySQLStr = MySQLStr + "WHERE (ID = " & Declarations.MyRequestNum & ") "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            DataGridView1.Rows(Declarations.MyRowIndex).Cells(14).Value = DataGridView1.Rows(Declarations.MyRowIndex).Cells(14).Value _
                & Chr(10) & Chr(13) & "--" & Format(Now, "dd/MM/yyyy HH:mm") & "-->" & myValue
            '--------------отправка почты
            EmailStr = GetEmailFromDB(Trim(Strings.Left(DataGridView1.Rows(Declarations.MyRowIndex).Cells(7).Value.ToString, 4)))
            If EmailStr = "" Then
                MsgBox("Для пользователя " & DataGridView1.Rows(Declarations.MyRowIndex).Cells(7).Value.ToString & " в БД не занесена почта. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            Else
                SendCommentByEmail(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString(), _
                    EmailStr, DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString(), DataGridView1.SelectedRows.Item(0).Cells(7).Value.ToString(), _
                    Trim(myValue), DataGridView1.SelectedRows.Item(0).Cells(11).Value.ToString())
            End If
        End If
    End Sub
End Class
