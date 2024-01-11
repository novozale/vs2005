Public Class Form1
    Structure FinRez
        Public MyRezStr As String
        Public MyRelocOrderNum As String
    End Structure

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из программы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyDs As New DataSet                       '

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            Declarations.ScalaDate = CDate(Scala.ActiveFrame.Parent.ScalaDate)


            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
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
            End If
        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---Заполнение формы (ComboBox)
        BuildWHListFrom()
        BuildWHListTo()

        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today

        CheckBox1.Checked = False
        CheckBox2.Checked = False
    End Sub

    Private Sub BuildWHListFrom()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в Combobox список складов, с которых возможно перемещение, и выбор первого
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "SC23002" 'Это то что будет отображаться
            ComboBox1.ValueMember = "SC23001"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub BuildWHListTo()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод в Combobox список складов, на которые возможно перемещение, и выбор первого
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE(SC23001 <> N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "SC23002" 'Это то что будет отображаться
            ComboBox2.ValueMember = "SC23001"   'это то что будет храниться
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// после выбора склада, с которого будет перемещение,
        '/// Вывод в Combobox список складов, на которые возможно перемещение, и выбор первого
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        BuildWHListTo()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки создание заказа на премещение
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckCorrect() = True Then
            If My.Settings.UseOffice = "LibreOffice" Then
                CreateDisplacementOrderLO()
            Else
                CreateDisplacementOrder()
            End If

        End If
    End Sub

    Private Function CheckCorrect() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности выбранных значений
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        '---даты заказа
        If DateTimePicker1.Value < Today Then
            MsgBox("Введите корректную даты предполагаемой отгрузки заказа (не должна быть раньше сегодняшнего дня)", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckCorrect = False
            Exit Function
        End If

        If DateTimePicker2.Value < Today Then
            MsgBox("Введите корректную даты предполагаемого получения заказа (не должна быть раньше сегодняшнего дня)", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker2.Select()
            CheckCorrect = False
            Exit Function
        End If

        If DateTimePicker1.Value > DateTimePicker2.Value Then
            MsgBox("Дата предполагаемого получения заказа не должна быть раньше даты предполагаемой отгрузки заказа)", MsgBoxStyle.Critical, "Внимание!")
            DateTimePicker1.Select()
            CheckCorrect = False
            Exit Function
        End If

        '---мы не создаем заказ на перемещение со склада давальческого сырья на склад давальческого сырья
        If IsRawMaterialsWH(ComboBox1.SelectedValue) = True And IsRawMaterialsWH(ComboBox2.SelectedValue) = True Then
            MsgBox("Склад источник " & ComboBox1.SelectedValue & " и склад назначения " & ComboBox2.SelectedValue & " являются складами давальческого сырья. Делать заказ на перемещение с одного склада давальческого сырья на другой нельзя.", MsgBoxStyle.Critical, "Внимание!")
            CheckCorrect = False
            Exit Function
        End If

        CheckCorrect = True

    End Function

    Private Sub CreateDisplacementOrder()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание заказа на перемещение на основе данных из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim i As Double                             'счетчик строк
        Dim MySQLStr As String
        Dim MyProductCode As String                 'код запаса
        Dim MyQTY As Double                         'перемещаемое количество

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Try
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---Удаление старой временной таблицы
                    MySQLStr = "IF exists(select * from tempdb..sysobjects where "
                    MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
                    MySQLStr = MySQLStr & "and xtype = N'U') "
                    MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---Создание новой временной таблицы
                    MySQLStr = "CREATE TABLE #_MyOrder( "
                    MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                '--код товара в Scala
                    MySQLStr = MySQLStr & "[QTY] float, "                              '--количество
                    MySQLStr = MySQLStr & "[RestQTY] float  "                          '--Остаток - неперемещенное количество
                    MySQLStr = MySQLStr & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    appXLSRC = CreateObject("Excel.Application")
                    appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                    i = 2 '---начинаем со 2 строки
                    While Not appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value = Nothing
                        MyProductCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value.ToString)
                        If appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value = Nothing Then
                            Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & ". Не указано количество для перемещения. Проверьте корректность информации, занесенной в Excel.")
                        Else
                            Try
                                MyQTY = CDbl(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString)
                            Catch ex As Exception
                                appXLSRC.DisplayAlerts = 0
                                appXLSRC.Workbooks.Close()
                                appXLSRC.DisplayAlerts = 1
                                appXLSRC.Quit()
                                appXLSRC = Nothing
                                Me.Cursor = Cursors.Default
                                Me.Refresh()
                                System.Windows.Forms.Application.DoEvents()
                                MsgBox("Строка " & CStr(i) & ". Код товара " & MyProductCode & ". Некорректно указано количество для перемещения. " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                Exit Sub
                            End Try
                        End If
                        '---Есть ли в Scala товары с таким кодом
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            '--запаса в Scala нет
                            Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " в Scala отсутствует. Проверьте корректность информации, занесенной в Excel.")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            If (CheckBox2.Checked = True) Or (CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) <> "02" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "03" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "04" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "05" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "06")) Then
                                '---Занесение во временную таблицу
                                '---Сначала проверяем - может, такой код уже есть во временной таблице
                                MySQLStr = "SELECT COUNT(*) AS CC "
                                MySQLStr = MySQLStr & "FROM #_MyOrder "
                                MySQLStr = MySQLStr & "WHERE (ItemCode = N'" & MyProductCode & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                    trycloseMyRec()
                                    '---и еще проверка - не перемещать комплексные продукты
                                    MySQLStr = "SELECT SC01066 "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Trim(Declarations.MyRec.Fields("SC01066").Value.ToString) = "8" Then
                                        trycloseMyRec()
                                        Throw (New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " является комплексным. При помощи данной программы комплексные товары перемещать нельзя."))
                                        trycloseMyRec()
                                    Else
                                        trycloseMyRec()
                                        MySQLStr = "INSERT INTO #_MyOrder "
                                        MySQLStr = MySQLStr & "(ItemCode, QTY, RestQTY) "
                                        MySQLStr = MySQLStr & "VALUES (N'" & MyProductCode & "', "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
                                        InitMyConn(False)
                                        Declarations.MyConn.Execute(MySQLStr)
                                    End If
                                Else
                                    trycloseMyRec()
                                    Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " присутствует в Excel более чем в одной строке. Код запаса на перемещение можно вставлять в Excel только 1 раз, без дублирования.")
                                End If
                            ElseIf CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) = "02" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "03" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "04" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "05" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "06") Then
                                '---сообщаем, что кабель не заносим
                                MsgBox("Строка " & CStr(i) & ". Код товара " & MyProductCode & " является кабельной продукцией и не будет включен в заказ на перемещение.", MsgBoxStyle.Critical, "Внимание!")
                            End If
                        End If
                        i = i + 1
                    End While

                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---запуск хранимой процедуры формирования заказа на перемещение
                    '---блокировка
                    '--SetBlock("0000000009") --блокировки в хранимую процедуру

                    '---и вывод результата в окно
                    ResultOutput(ExecSppCreateRelocOrder())

                    'MsgBox("Процедура создания заказа на перемещение завершена.", MsgBoxStyle.OkOnly, "Внимание!")
                Catch ex As Exception
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
        End If
    End Sub

    Private Sub CreateDisplacementOrderLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание заказа на перемещение на основе данных из LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim i As Double                             'счетчик строк
        Dim MySQLStr As String
        Dim MyProductCode As String                 'код запаса
        Dim MyQTY As Double                         'перемещаемое количество
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Try
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---Удаление старой временной таблицы
                    MySQLStr = "IF exists(select * from tempdb..sysobjects where "
                    MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
                    MySQLStr = MySQLStr & "and xtype = N'U') "
                    MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---Создание новой временной таблицы
                    MySQLStr = "CREATE TABLE #_MyOrder( "
                    MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                '--код товара в Scala
                    MySQLStr = MySQLStr & "[QTY] float, "                              '--количество
                    MySQLStr = MySQLStr & "[RestQTY] float  "                          '--Остаток - неперемещенное количество
                    MySQLStr = MySQLStr & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    LOSetNotation(0)
                    oServiceManager = CreateObject("com.sun.star.ServiceManager")
                    oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                    oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                    oFileName = Replace(OpenFileDialog2.FileName, "\", "/")
                    oFileName = "file:///" + oFileName
                    Dim arg(1)
                    arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                    arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                    oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                    oSheet = oWorkBook.getSheets().getByIndex(0)

                    i = 2 '---начинаем со 2 строки
                    While oSheet.getCellRangeByName("A" & i).String.Equals("") = False
                        MyProductCode = Trim(oSheet.getCellRangeByName("A" & i).String)
                        If oSheet.getCellRangeByName("B" & i).Value = 0 Then
                            Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & ". Не указано количество для перемещения. Проверьте корректность информации, занесенной в Excel.")
                        Else
                            Try
                                MyQTY = CDbl(oSheet.getCellRangeByName("B" & i).Value)
                            Catch ex As Exception
                                oWorkBook.Close(True)
                                Me.Cursor = Cursors.Default
                                Me.Refresh()
                                System.Windows.Forms.Application.DoEvents()
                                MsgBox("Строка " & CStr(i) & ". Код товара " & MyProductCode & ". Некорректно указано количество для перемещения. " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                                Exit Sub
                            End Try
                        End If

                        '---Есть ли в Scala товары с таким кодом
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            '--запаса в Scala нет
                            Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " в Scala отсутствует. Проверьте корректность информации, занесенной в Excel.")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            If (CheckBox2.Checked = True) Or (CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) <> "02" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "03" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "04" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "05" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "06")) Then
                                '---Занесение во временную таблицу
                                '---Сначала проверяем - может, такой код уже есть во временной таблице
                                MySQLStr = "SELECT COUNT(*) AS CC "
                                MySQLStr = MySQLStr & "FROM #_MyOrder "
                                MySQLStr = MySQLStr & "WHERE (ItemCode = N'" & MyProductCode & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                    trycloseMyRec()
                                    '---и еще проверка - не перемещать комплексные продукты
                                    MySQLStr = "SELECT SC01066 "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Trim(Declarations.MyRec.Fields("SC01066").Value.ToString) = "8" Then
                                        trycloseMyRec()
                                        Throw (New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " является комплексным. При помощи данной программы комплексные товары перемещать нельзя."))
                                        trycloseMyRec()
                                    Else
                                        trycloseMyRec()
                                        MySQLStr = "INSERT INTO #_MyOrder "
                                        MySQLStr = MySQLStr & "(ItemCode, QTY, RestQTY) "
                                        MySQLStr = MySQLStr & "VALUES (N'" & MyProductCode & "', "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
                                        InitMyConn(False)
                                        Declarations.MyConn.Execute(MySQLStr)
                                    End If
                                Else
                                    trycloseMyRec()
                                    Throw New ArgumentException("Строка " & CStr(i) & ". Код товара " & MyProductCode & " присутствует в Excel более чем в одной строке. Код запаса на перемещение можно вставлять в Excel только 1 раз, без дублирования.")
                                End If
                            ElseIf CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) = "02" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "03" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "04" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "05" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "06") Then
                                '---сообщаем, что кабель не заносим
                                MsgBox("Строка " & CStr(i) & ". Код товара " & MyProductCode & " является кабельной продукцией и не будет включен в заказ на перемещение.", MsgBoxStyle.Critical, "Внимание!")
                            End If
                        End If
                        i = i + 1
                    End While
                    oWorkBook.Close(True)
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---запуск хранимой процедуры формирования заказа на перемещение
                    ResultOutput(ExecSppCreateRelocOrder())
                Catch ex As Exception
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
            End If
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Function ExecSppCreateRelocOrder() As FinRez
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// вызов хранимой процедуры создания заказа на перемещение на основе данных из Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String
        Dim MyRelocOrderNum As String                   'номер создаваемого заказа на перемещение
        Dim MyFinRez As FinRez                          'возвращаемая структура
        Dim cmd As New ADODB.Command                    'команда (spp процедура)
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1 //с какого склада перемещать
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2 //на какой склад перемещать
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3 //Включать в заказ на перемещение запасы для заказов на продажу на других складах
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4 //предполагаемая дата отгрузки
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5 //предполагаемая дата приемки
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6 //возвращаемая строка - результат работы
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7 //возвращаемая строка - номер заказа на перемещение

        MyRezStr = ""
        MyRelocOrderNum = ""
        InitMyConn(False)
        Try
            cmd.ActiveConnection = Declarations.MyConn
            cmd.CommandText = "spp_DisplacementOrderCreationFromExcel"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.CommandTimeout = 300

            '----Создание параметров---------------------------------------------------
            '---Исходный склад
            MyParam = cmd.CreateParameter("@SrcWarNo", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam)
            '---Склад назначения
            MyParam1 = cmd.CreateParameter("@DestWarNo", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam1)
            '--Флаг - Включать в заказ на перемещение запасы для заказов на продажу на других складах
            MyParam2 = cmd.CreateParameter("@MyOtherWHFlag", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam2)
            '---Дата отправки
            MyParam3 = cmd.CreateParameter("@MyOrderDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam3)
            '---Дата получения
            MyParam4 = cmd.CreateParameter("@MyShipDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam4)
            '---Возвращаемый параметр (строка) - результат работы
            MyParam5 = cmd.CreateParameter("@MyRezStr", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
            cmd.Parameters.Append(MyParam5)
            '---Возвращаемый параметр (строка) - номер заказа на перемещение
            MyParam6 = cmd.CreateParameter("@MyRelocOrderNum", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 30)
            cmd.Parameters.Append(MyParam6)

            '----значения параметров---------------------------------------------------
            '---Исходный склад
            MyParam.Value = Trim(ComboBox1.SelectedValue.ToString)
            '---Склад назначения
            MyParam1.Value = Trim(ComboBox2.SelectedValue.ToString)
            '--Флаг - Включать в заказ на перемещение запасы для заказов на продажу на других складах
            If CheckBox1.Checked = True Then
                MyParam2.Value = 1 'включать
            Else
                MyParam2.Value = 0 'не включать
            End If
            '---Дата отправки
            MyParam3.Value = DateTimePicker1.Value
            '---Дата получения
            MyParam4.Value = DateTimePicker2.Value
            '---запуск хранимой процедуры------------------------------------------------
            cmd.Execute()
            MyRezStr = MyRezStr + LTrim(RTrim(MyParam5.Value))
            MyRelocOrderNum = Trim(MyParam6.Value)

        Catch ex As Exception
            MyRezStr = MyRezStr + ex.Message
        End Try
        MyFinRez.MyRezStr = MyRezStr
        MyFinRez.MyRelocOrderNum = MyRelocOrderNum
        ExecSppCreateRelocOrder = MyFinRez
    End Function

    Private Sub ResultOutput(ByVal MyFinRez As FinRez)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// вывод результата работы хранимой процедуры создания заказа на перемещение на основе данных из Excel в окно
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---снятие блокировки
        '----RemoveBlock()

        MyErrorForm = New ErrorForm
        If MyFinRez.MyRezStr = "" Then
        Else
            MyErrorForm.MyHdr = "Во время импорта данных из Excel были ошибки " & Chr(13)
        End If

        '------------Вывод информации о номере заказа на перемещение
        If MyFinRez.MyRezStr <> "" Then
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & Chr(13) & Chr(13)
        End If

        If Trim(MyFinRez.MyRelocOrderNum) = "" Then
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "В результате выполнения операции заказ на перемещение не был создан. " & Chr(13)
        Else
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "В результате выполнения операции был создан заказ на перемещение: " & Trim(MyFinRez.MyRelocOrderNum) & Chr(13)
        End If

        '------------Вывод информации о неперемещенных заказах
        MySQLStr = "SELECT  ItemCode, RestQTY "
        MySQLStr = MySQLStr & "FROM #_MyOrder "
        MySQLStr = MySQLStr & "WHERE (RestQTY <> 0) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & Chr(13)
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "В результате импорта из Excel заказ на перемещение не сформировался для следующих запасов: " & Chr(13)
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "Код товара в Scala  Неперемещенное количество" & Chr(13)
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                MyFinRez.MyRezStr = MyFinRez.MyRezStr & Microsoft.VisualBasic.Left(Declarations.MyRec.Fields("ItemCode").Value & "                    ", 20) & MyRec.Fields("RestQTY").Value.ToString & Chr(13)
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        MyErrorForm.MyMsg = MyFinRez.MyRezStr & Chr(13)
        MyErrorForm.ShowDialog()
    End Sub
End Class
