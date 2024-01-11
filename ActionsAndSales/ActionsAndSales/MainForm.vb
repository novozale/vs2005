Public Class MainForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запуск процедуры создания акции / распродажи  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button1.Enabled = False
        Button2.Enabled = False
        System.Windows.Forms.Application.DoEvents()
        If My.Settings.UseOffice = "LibreOffice" Then
            ImportDataFromExcel_LO()
        Else
            ImportDataFromExcel()
        End If
        Button1.Enabled = True
        Button2.Enabled = True
        System.Windows.Forms.Application.DoEvents()
        MsgBox("Процедура загрузки информации по акции / распродаже завершена.", vbOKOnly, "Внимание!")
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Excel информации по акции / распродаже 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyTableName As String                   'Имя временной таблицы
        Dim MyGuid As String                          '
        Dim connStr As String                       'строка соединения с Excel
        Dim MySQLStr As String                      'SQL запрос
        Dim cn As OleDbConnection                   'объект соединение с OLE
        Dim FirstExcelSheetName As String           'название первого листа Excel
        Dim myds As DataSet                         'Excel dataset
        Dim MyVersion As String                     'Версия документа
        Dim mycount As Integer
        Dim MyDBL As Double                         'для проверки
        Dim MyInt As Integer                        'для проверки
        Dim MyStr As String                         'для проверки
        Dim MyDatetimeStart As Date                 'для проверки
        Dim MyDatetimeFin As Date                   'для проверки
        Dim MySQLAdapter As SqlClient.SqlDataAdapter 'для временной таблицы
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyErrStr As String
        Dim MyContinueFlag As Integer
        Dim MyRez As MsgBoxResult                   'результат выбора
        Dim MyActOrSalesFlag As String

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_ActionsAndSales_Tmp_" + MyGuid

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                '----------------------------Подпись операции
                Label3.Text = "Выполнение проверок Excel файла"
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                connStr = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" & OpenFileDialog1.FileName & ";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;"""
                Try
                    cn = New OleDbConnection(connStr)
                    FirstExcelSheetName = GetFirstExcelSheetName(cn)

                    '============================проверки============================================================================
                    '---Проверяем версию листа Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A1:A1]"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds Is Nothing = False Then
                        If IsDBNull(myds.Tables(0).Rows(0).Item(0)) Then
                            MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            MyVersion = Trim(myds.Tables(0).Rows(0).Item(0))
                            MySQLStr = "SELECT Version "
                            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                            MySQLStr = MySQLStr & "WHERE (Name = N'Создание акции или распродажи') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            Else
                                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                    trycloseMyRec()
                                Else
                                    MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "Внимание!")
                                    trycloseMyRec()
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        MsgBox("Невозможно прочитать версию листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '---Проверяем - занесена ли информация что это - акция или распродажа
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "C2:C2]"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds Is Nothing = False Then
                        If IsDBNull(myds.Tables(0).Rows(0).Item(0)) Then
                            MsgBox("В импортируемом листе Excel в ячейке ""C2"" не проставлено что это - акция или неликвид. ", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            MyActOrSalesFlag = Trim(myds.Tables(0).Rows(0).Item(0))
                            If (MyActOrSalesFlag.Equals("акция") = False And MyActOrSalesFlag.Equals("неликвид") = False) Then
                                MsgBox("В импортируемом листе Excel в ячейке ""C2"" должно быть проставлено - ""акция"" или ""неликвид"". ", MsgBoxStyle.Critical, "Внимание!")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        End If
                    Else
                            MsgBox("Невозможно прочитать что это - акция или распродажа. Обратитесь к администратору.", vbCritical, "Внимание!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                    End If


                    '---Проверяем корректность данных в Excel
                    '-----Дублированные коды
                    MySQLStr = "SELECT F1 FROM [" & FirstExcelSheetName & "A5:A] group by F1 having(count(F1) > 1)"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds.Tables(0).Rows.Count > 0 Then
                        MsgBox("В файле находятся " & myds.Tables(0).Rows.Count & " дублированных записей кодов товаров в Scala. Воспользуйтесь кнопкой ""Подсветить дублированные"" в Excel, проверьте и удалите лишние коды ")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----правильность занесения данных в Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A5:J] where(F1 <> """")"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    '-----правильность занесения
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        '-----заполнение кода товара Scala
                        If Trim(myds.Tables(0).Rows(mycount).Item(0).ToString) = "" Then
                            MsgBox("Строка " & CStr(mycount + 5) & " не занесен код товара в Scala")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        '-----закупочная цена Электроскандии
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(1)
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена закупочная цена Электроскандии")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----Валюта закупки
                        Try
                            MyInt = myds.Tables(0).Rows(mycount).Item(2)
                            If (MyInt <> 0 And MyInt <> 1 And MyInt <> 4 And MyInt <> 12) Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена валюта закупки")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена валюта закупки")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----коэффициент маржинальности
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(3)
                            If (MyDBL < 1 Or MyDBL > 2) Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен коэффициент маржинальности - должен быть в промежутке от 1 до 2")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен коэффициент маржинальности")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----Признак - акция по количеству
                        Try
                            MyStr = myds.Tables(0).Rows(mycount).Item(4).ToString
                            If MyStr.ToUpper.Equals("ДА") = False And MyStr.ToUpper.Equals("НЕТ") = False Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен признак - акция по количеству: должно быть да или нет")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен признак - акция по количеству")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----Количество, продаваемое по акции
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(5))) Then
                        Else
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(5)
                                If (MyDBL < 0) Then
                                    MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесено количество, продаваемое по акции - должно быть пусто или больше или равно 0")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесено количество, продаваемое по акции - должно быть пусто или больше или равно 0")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If
                        '-----Признак - акция по сроку
                        Try
                            MyStr = myds.Tables(0).Rows(mycount).Item(6).ToString
                            If MyStr.ToUpper.Equals("ДА") = False And MyStr.ToUpper.Equals("НЕТ") = False Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен признак - акция по сроку: должно быть да или нет")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесен признак - акция по сроку")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----Дата начала акции
                        Try
                            MyDatetimeStart = myds.Tables(0).Rows(mycount).Item(7)
                            If MyDatetimeStart < Today Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата начала акции: дата начала акции не может быть меньше текущей даты")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата начала акции")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----Дата окончания акции
                        MyStr = myds.Tables(0).Rows(mycount).Item(8).ToString
                        If MyStr.Equals("") Then
                            If (myds.Tables(0).Rows(mycount).Item(6).ToString.ToUpper.Equals("ДА")) Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата окончания акции: так как акция по сроку - дата ококнчания акции должна быть указана обязательно.")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Else
                            If (myds.Tables(0).Rows(mycount).Item(6).ToString.ToUpper.Equals("НЕТ")) Then
                                MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата окончания акции: так как акция не по сроку - дата ококнчания акции должна быть пустой.")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            Else
                                Try
                                    MyDatetimeFin = myds.Tables(0).Rows(mycount).Item(8)
                                    If MyDatetimeFin < MyDatetimeStart Then
                                        MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата окончания акции: дата окончания акции не может быть меньше даты начала акции")
                                        Me.Cursor = Cursors.Default
                                        Exit Sub
                                    End If
                                Catch ex As Exception
                                    MsgBox("Строка " & CStr(mycount + 5) & " некорректно занесена дата окончания акции")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End Try
                            End If
                        End If
                        '-----проверка - обязательно должна быть акция или по количеству или по сроку
                        If myds.Tables(0).Rows(mycount).Item(4).ToString().ToUpper.Equals("НЕТ") And myds.Tables(0).Rows(mycount).Item(6).ToString().ToUpper.Equals("НЕТ") Then
                            MsgBox("Строка " & CStr(mycount + 5) & " акция должна быть обязательно или по сроку или по количеству или по сроку и количеству.")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If

                        '-----Название и условия акции
                        MyStr = myds.Tables(0).Rows(mycount).Item(9).ToString
                        If (MyStr.Equals("")) Then
                            MsgBox("Строка " & CStr(mycount + 5) & " не указано название и условие акции")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        mycount = mycount + 1
                    End While

                    '========================================загрузка датасета во временную таблицу=================================
                    '----------------------------Подпись операции
                    Label3.Text = "Загрузка данных на сервер"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----Создание временной таблицы
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[ScalaCode] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePrice] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePriceCurr] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MarginCoeff] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[QTYAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionStopQTY] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[TimeAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[DateStart] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[DateFinish] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionName] [nvarchar](4000) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionOrSales] [nvarchar](50) NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----данные из временной таблицы
                    InitMyConn(False)
                    MySQLStr = "SELECT  ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    Try
                        MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                        MySQLAdapter.SelectCommand.CommandTimeout = 1200
                        Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                        MySQLAdapter.Fill(MySQLDs)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try
                    '-----Перенос данных из Excel dataset в SQL dataset
                    Dim dt As DataTable
                    Dim dr As DataRow

                    dt = MySQLDs.Tables(0)
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        dr = dt.NewRow
                        dr.Item(0) = myds.Tables(0).Rows(mycount).Item(0)
                        dr.Item(1) = myds.Tables(0).Rows(mycount).Item(1)
                        dr.Item(2) = myds.Tables(0).Rows(mycount).Item(2)
                        dr.Item(3) = myds.Tables(0).Rows(mycount).Item(3)
                        dr.Item(4) = myds.Tables(0).Rows(mycount).Item(4)
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(5))) Then
                            dr.Item(5) = 999999999
                        Else
                            If (myds.Tables(0).Rows(mycount).Item(5) = 0) Then
                                dr.Item(5) = 999999999
                            Else
                                dr.Item(5) = myds.Tables(0).Rows(mycount).Item(5)
                            End If
                        End If
                        dr.Item(6) = myds.Tables(0).Rows(mycount).Item(6)
                        dr.Item(7) = myds.Tables(0).Rows(mycount).Item(7)
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(8))) Then
                            dr.Item(8) = New DateTime(9999, 12, 31, 0, 0, 0)
                        Else
                            dr.Item(8) = myds.Tables(0).Rows(mycount).Item(8)
                        End If
                        dr.Item(9) = myds.Tables(0).Rows(mycount).Item(9)
                        dr.Item(10) = MyActOrSalesFlag
                        dt.Rows.Add(dr)
                        mycount = mycount + 1
                    End While
                    Try
                        MySQLAdapter.Update(MySQLDs, "Table")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '========================================Выполнение проверок на сервере=================================
                    '==============================Проверка наличия в БД скальских кодов====================================
                    Label3.Text = "Проверка наличия в БД скальских кодов"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "Select " & MyTableName & ".ScalaCode "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "SC010300 ON " & MyTableName & ".ScalaCode = SC010300.SC01001 "
                    MySQLStr = MySQLStr & "WHERE(SC010300.SC01001 Is NULL) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Следующие коды товаров отсутствуют в Scala:" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Declarations.MyRec.Fields("ScalaCode").Value & Chr(13) & Chr(10)
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверные коды товаров в Excel файле.")
                        End If
                    End If

                    '==============================Проверка наличия в БД акции с таким же названием=============================
                    Label3.Text = "Проверка наличия в БД акции с таким же названием"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MyContinueFlag = 1
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, CONVERT(nvarchar(30), MIN(tbl_ActionsAndSales.DateStart), 103) AS DateStart, "
                    MySQLStr = MySQLStr & "CONVERT(nvarchar(30), MAX(tbl_ActionsAndSales.DateFinish), 103) AS DateFinish, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN 'Уже стартовала' ELSE 'Еще не стартовала' END AS MyState "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN tbl_ActionsAndSales ON "
                    MySQLStr = MySQLStr & " " & MyTableName & ".ActionName = tbl_ActionsAndSales.ActionName "
                    MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales.ActionName, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN 'Уже стартовала' ELSE 'Еще не стартовала' END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Данные акции / распродажи уже присутствуют в Scala:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) & Declarations.MyRec.Fields("MyState").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) + "С" + Chr(9) & Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value & Chr(13) & Chr(10)
                            If (Declarations.MyRec.Fields("MyState").Value.ToString.Equals("Уже стартовала")) Then
                                MyContinueFlag = 0
                            End If
                            Declarations.MyRec.MoveNext()
                        End While
                        If (MyContinueFlag = 0) Then
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "Поменяйте название акции / распродажи в Excel файле на уникальное " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        Else
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то указанные " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "выше акции / распродажи будут удалены и заменены акциями из Excel. " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "Если вы не хотите удалять уже существующие акции / распродажи, то " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "выберите ""Выход"", " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "поменяйте название акции / распродажи в Excel файле на уникальное." & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        End If
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        If (MyContinueFlag = 0) Then
                            MyErrForm.Button2.Visible = False
                            MyErrForm.Button2.Enabled = False
                        Else
                            MyErrForm.Button2.Visible = True
                            MyErrForm.Button2.Enabled = True
                        End If
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Акция / распродажа уже существует.")
                        Else    '----удаление ранее занесенных акций
                            MySQLStr = "DELETE FROM tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & ".ActionName = tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales_1.ActionName) AS View_2 ON tbl_ActionsAndSales.ActionName = View_2.ActionName "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '==============================Закрытие (смена дат) акций в случае перекрытия=============================
                    Label3.Text = "Закрытие (смена дат) акций в случае перекрытия"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '----------------------------Диапазон дат акции / распродажи внутри уже существующего---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart >= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish <= "
                    MySQLStr = MySQLStr & "CASE WHEN tbl_ActionsAndSales.DateFinish = CONVERT(datetime, '31/12/9999', 103) THEN "
                    MySQLStr = MySQLStr & "dateadd(dd, - 1, tbl_ActionsAndSales.DateFinish) ELSE tbl_ActionsAndSales.DateFinish END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи находится внутри уже существующего диапазона:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Поменяйте Диапазон акции / распродажи в Excel файле " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        End If
                    End If

                    '---------------Диапазон дат акции / распродажи охватывает с обоих сторон существующий--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart <= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish >= tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи охватывает с обоих сторон существующий:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Поменяйте Диапазон акции / распродажи в Excel файле " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        End If
                    End If

                    '------------------Диапазон дат акции / распродажи Слева перекрывает уже существующие---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи по дате начала перекрывается с уже существующим:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то даты " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "окончания предыдущей акции / распродажи будут изменены на дату начала текущей акции - 1 день. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "Если вы не хотите менять таким образом дату окончания предыдущей акции / распродажи, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "то выберите ""Выход"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "поменяйте даты начала акции / распродажи в Excel файле на необходимую" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        Else    '-----меняем даты окончания акции
                            'MySQLStr = "UPDATE " & MyTableName & " "
                            'MySQLStr = MySQLStr & "SET DateStart = DateAdd(dd, 1, CASE WHEN View_2.DateFinish = CONVERT(datetime, '31/12/9999', 103) "
                            'MySQLStr = MySQLStr & "THEN dateadd(dd, - 1, View_2.DateFinish) ELSE View_2.DateFinish END)) "
                            'MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            'MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            'MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales.DateStart AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            'MySQLStr = MySQLStr & " " & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            MySQLStr = "Update tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart), "
                            MySQLStr = MySQLStr & "TimeAction = N'да' "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName, tbl_ActionsAndSales_1.ScalaCode, " & MyTableName & "_1.DateStart, "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales_1.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales_1.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales_1.DateFinish) AS View_2 ON "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales.ScalaCode = View_2.ScalaCode And tbl_ActionsAndSales.DateStart < View_2.DateStart And tbl_ActionsAndSales.DateFinish > View_2.DateStart "

                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '------------------Диапазон дат акции / распродажи Справа перекрывает уже существующие--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи по дате окончания перекрывается с уже существующим:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то даты " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "окончания акции / распродажи будут изменены на дату начала последующей  акции минус 1 день. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "Если вы не хотите менять таким образом дату окончания акции / распродажи, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "то выберите ""Выход"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "поменяйте даты окончания акции / распродажи в Excel файле на необходимыю" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        Else    '-----меняем даты начала акции
                            MySQLStr = "UPDATE " & MyTableName & " "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart) "
                            MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish > tbl_ActionsAndSales.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            MySQLStr = MySQLStr & "" & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If



                    '==============================Занесение акции / распродажи в БД=============================
                    Label3.Text = "Занесение акции / распродажи в БД"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "INSERT INTO tbl_ActionsAndSales "
                    MySQLStr = MySQLStr & "(ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, ActionFinished, ActionFinishedDate) "
                    MySQLStr = MySQLStr & "SELECT ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, 0 AS ActionFinished, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, '01/01/1900', 103) AS ActionFinishedDate "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '==============================================Запуск пересчета прайс листа на продажу============================================
                    MyRez = MsgBox("Произвести расчет прайс - листа на продажу сейчас? Иначе расчет будет выполнен ночью.", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = MsgBoxResult.Yes Then
                        '----------------------------Подпись операции
                        Label3.Text = "Формирование прайс листа на продажу. "
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()

                        MySQLStr = "Exec spp_PrepareCommonPriceList_PriCost "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If


                Catch ex As Exception
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                Finally
                    cn.Close()
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                    '----------------------------Подпись операции
                    Label3.Text = ""
                End Try
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub ImportDataFromExcel_LO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Excel информации по акции / распродаже при помощи LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL запрос
        Dim MyTableName As String                   'Имя временной таблицы
        Dim MyGuid As String                          '
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                     'Версия документа
        Dim MyActOrSalesFlag As String              'что это - акция или распродажа
        Dim MyDBL As Double                         'для проверки
        Dim MyInt As Integer                        'для проверки
        Dim MyStr As String                         'для проверки
        Dim MyDatetimeStart As Date                 'для проверки
        Dim MyDatetimeFin As Date                   'для проверки
        Dim MySQLAdapter As SqlClient.SqlDataAdapter 'для временной таблицы
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyErrStr As String
        Dim MyContinueFlag As Integer
        Dim MyRez As MsgBoxResult                   'результат выбора


        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_ActionsAndSales_Tmp_" + MyGuid

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                '----------------------------Подпись операции
                Label3.Text = "Выполнение проверок Excel файла"
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                Try
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
                    '============================проверки============================================================================
                    '---Проверяем версию листа Excel
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    If MyVersion = "" Then
                        MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        MySQLStr = "SELECT Version "
                        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                        MySQLStr = MySQLStr & "WHERE (Name = N'Создание акции или распродажи') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                trycloseMyRec()
                            Else
                                MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "Внимание!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    End If

                    '---Проверяем - занесена ли информация что это - акция или распродажа
                    MyActOrSalesFlag = oSheet.getCellRangeByName("C2").String
                    If MyActOrSalesFlag.Equals("") Then
                        MsgBox("В импортируемом листе Excel в ячейке ""C2"" не проставлено что это - акция или неликвид. ", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If (MyActOrSalesFlag.Equals("акция") = False And MyActOrSalesFlag.Equals("неликвид") = False) Then
                            MsgBox("В импортируемом листе Excel в ячейке ""C2"" должно быть проставлено - ""акция"" или ""неликвид"". ", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If

                    '---Проверяем корректность данных в Excel
                    '-----Дублированные коды
                    oSheet.unprotect("!pass2022")

                    Dim args() As Object
                    ReDim args(0)
                    args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args(0).Name = "ToPoint"
                    args(0).Value = "$A$5:$K$100000"
                    Dim oFrame As Object
                    oFrame = oWorkBook.getCurrentController.getFrame
                    oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
                    Dim args1() As Object
                    ReDim args1(6)
                    args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(0).Name = "ByRows"
                    args1(0).Value = True
                    args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(1).Name = "HasHeader"
                    args1(1).Value = False
                    args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(2).Name = "CaseSensitive"
                    args1(2).Value = False
                    args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(3).Name = "NaturalSort"
                    args1(3).Value = False
                    args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(4).Name = "IncludeAttribs"
                    args1(4).Value = True
                    args1(5) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(5).Name = "Col1"
                    args1(5).Value = 1
                    args1(6) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(6).Name = "Ascending1"
                    args1(6).Value = True
                    oDispatcher.executeDispatch(oFrame, ".uno:DataSort", "", 0, args1)

                    'Dim oSortFields(0) As Object
                    'Dim oSortDesc(0) As Object
                    'Dim srange = oSheet.getCellRangeByName("A5:K100000")
                    ' ''oSortFields(0) = oServiceManager.Bridge_GetStruct("com.sun.star.table.TableSortField")
                    ' ''oSortFields(0).Field = 0
                    ' ''oSortFields(0).IsAscending = False
                    ''oSortFields(0) = oServiceManager.Bridge_GetStruct("com.sun.star.util.SortField")
                    ''oSortFields(0).Field = 0
                    ''oSortFields(0).SortAscending = False
                    ' ''oSortDesc = srange.createSortDescriptor
                    ' ''oSortDesc(1).Value = False
                    ' ''oSortDesc(3).Value = oSortFields
                    ''oSortDesc(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    ''oSortDesc(0).Name = "SortFields"
                    ''oSortDesc(0).Value = oSortFields

                    'Dim oReflection As Object
                    ''Dim unoWrap As Object
                    'oReflection = oServiceManager.createInstance("com.sun.star.reflection.CoreReflection")
                    ''Dim sortFields(0) As Object
                    ''Dim sortProperties(0) As Object
                    'oReflection.forName("com.sun.star.table.TableSortField").CreateObject(oSortFields(0))
                    'oSortFields(0).Field = 0
                    'oSortFields(0).IsAscending = False
                    ''unoWrap = oServiceManager.Bridge_GetValueObject
                    ''unoWrap.set("[]com.sun.star.table.TableSortField", oSortFields)
                    'oReflection.forName("com.sun.star.beans.PropertyValue").createObject(oSortDesc(0))
                    'oSortDesc(0).Name = "SortFields"
                    'oSortDesc(0).Value = oSortFields

                    'srange.Sort(oSortDesc)


                    Dim srange = oSheet.getCellRangeByName("A5:K100000")
                    Dim myarr = srange.DataArray
                    For i As Integer = 1 To 99995
                        If myarr(i)(0) = myarr(i - 1)(0) And myarr(i)(0) <> "" Then
                            MsgBox("В файле находятся дублированные записи кодов товаров в Scala. Воспользуйтесь кнопкой ""Показать дублированные данные"" в LibreOffice, проверьте и удалите лишние коды ")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    Next i
                    Dim args2() As Object
                    ReDim args2(0)
                    oDispatcher.executeDispatch(oFrame, ".uno:Save", "", 0, args2)

                    '-----правильность занесения данных в Excel
                    For i As Integer = 0 To 99995
                        If myarr(i)(0).ToString = "" Then
                            Exit For
                        Else
                            '-----закупочная цена Электроскандии
                            Try
                                MyDBL = myarr(i)(1)
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесена закупочная цена Электроскандии")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Валюта закупки
                            Try
                                MyInt = myarr(i)(2)
                                If (MyInt <> 0 And MyInt <> 1 And MyInt <> 4 And MyInt <> 12) Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесена валюта закупки")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесена валюта закупки")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----коэффициент маржинальности
                            Try
                                MyDBL = myarr(i)(3)
                                If (MyDBL < 1 Or MyDBL > 2) Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесен коэффициент маржинальности - должен быть в промежутке от 1 до 2")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесен коэффициент маржинальности")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Признак - акция по количеству
                            Try
                                MyStr = myarr(i)(4)
                                If MyStr.ToUpper.Equals("ДА") = False And MyStr.ToUpper.Equals("НЕТ") = False Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесен признак - акция по количеству: должно быть да или нет")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесен признак - акция по количеству")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Количество, продаваемое по акции
                            Try
                                MyDBL = myarr(i)(5)
                                If (MyDBL < 0) Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесено количество, продаваемое по акции - должно быть пусто или больше или равно 0")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесено количество, продаваемое по акции - должно быть пусто или больше или равно 0")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Признак - акция по сроку
                            Try
                                MyStr = myarr(i)(6)
                                If MyStr.ToUpper.Equals("ДА") = False And MyStr.ToUpper.Equals("НЕТ") = False Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесен признак - акция по сроку: должно быть да или нет")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесен признак - акция по сроку")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Дата начала акции
                            Try
                                MyDatetimeStart = DateTime.FromOADate(myarr(i)(7))
                                If MyDatetimeStart < Today Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата начала акции: дата начала акции не может быть меньше текущей даты")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата начала акции")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----Дата окончания акции
                            MyStr = myarr(i)(8)
                            If MyStr.Equals("") Then
                                If (myarr(i)(6).ToString.ToUpper.Equals("ДА")) Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата окончания акции: так как акция по сроку - дата ококнчания акции должна быть указана обязательно.")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Else
                                If (myarr(i)(6).ToString.ToUpper.Equals("НЕТ")) Then
                                    MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата окончания акции: так как акция не по сроку - дата ококнчания акции должна быть пустой.")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                Else
                                    Try
                                        MyDatetimeFin = DateTime.FromOADate(myarr(i)(8))
                                        If MyDatetimeFin < MyDatetimeStart Then
                                            MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата окончания акции: дата окончания акции не может быть меньше даты начала акции")
                                            Me.Cursor = Cursors.Default
                                            oWorkBook.Close(True)
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        MsgBox("Строка " & CStr(i + 5) & " некорректно занесена дата окончания акции")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If
                            End If
                            '-----проверка - обязательно должна быть акция или по количеству или по сроку
                            If myarr(i)(4).ToString().ToUpper.Equals("НЕТ") And myarr(i)(6).ToString().ToUpper.Equals("НЕТ") Then
                                MsgBox("Строка " & CStr(i + 5) & " акция должна быть обязательно или по сроку или по количеству или по сроку и количеству.")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                            '-----Название и условия акции
                            MyStr = myarr(i)(9)
                            If (MyStr.Equals("")) Then
                                MsgBox("Строка " & CStr(i + 5) & " не указано название и условие акции")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    Next i

                    '========================================загрузка датасета во временную таблицу=================================
                    '----------------------------Подпись операции
                    Label3.Text = "Загрузка данных на сервер"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----Создание временной таблицы
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[ScalaCode] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePrice] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePriceCurr] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MarginCoeff] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[QTYAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionStopQTY] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[TimeAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[DateStart] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[DateFinish] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionName] [nvarchar](4000) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionOrSales] [nvarchar](50) NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----данные из временной таблицы
                    InitMyConn(False)
                    MySQLStr = "SELECT  ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    Try
                        MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                        MySQLAdapter.SelectCommand.CommandTimeout = 1200
                        Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                        MySQLAdapter.Fill(MySQLDs)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '-----Перенос данных из Excel dataset в SQL dataset
                    Dim dt As DataTable
                    Dim dr As DataRow

                    dt = MySQLDs.Tables(0)
                    For i As Integer = 0 To 99995
                        If myarr(i)(0).ToString = "" Then
                            Exit For
                        Else
                            dr = dt.NewRow
                            dr.Item(0) = myarr(i)(0).ToString
                            dr.Item(1) = myarr(i)(1)
                            dr.Item(2) = myarr(i)(2)
                            dr.Item(3) = myarr(i)(3)
                            dr.Item(4) = myarr(i)(4)
                            If (myarr(i)(5) = 0) Then
                                dr.Item(5) = 999999999
                            Else
                                dr.Item(5) = myarr(i)(5)
                            End If
                            dr.Item(6) = myarr(i)(6)
                            dr.Item(7) = DateTime.FromOADate(myarr(i)(7))
                            If (myarr(i)(8).ToString().Equals("")) Then
                                dr.Item(8) = New DateTime(9999, 12, 31, 0, 0, 0)
                            Else
                                dr.Item(8) = DateTime.FromOADate(myarr(i)(8))
                            End If
                            dr.Item(9) = myarr(i)(9)
                            dr.Item(10) = MyActOrSalesFlag
                            dt.Rows.Add(dr)
                        End If
                    Next i
                    Try
                        MySQLAdapter.Update(MySQLDs, "Table")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '========================================Выполнение проверок на сервере=================================
                    '==============================Проверка наличия в БД скальских кодов====================================
                    Label3.Text = "Проверка наличия в БД скальских кодов"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "Select " & MyTableName & ".ScalaCode "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "SC010300 ON " & MyTableName & ".ScalaCode = SC010300.SC01001 "
                    MySQLStr = MySQLStr & "WHERE(SC010300.SC01001 Is NULL) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Следующие коды товаров отсутствуют в Scala:" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Declarations.MyRec.Fields("ScalaCode").Value & Chr(13) & Chr(10)
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверные коды товаров в Excel файле.")
                        End If
                    End If

                    '==============================Проверка наличия в БД акции с таким же названием=============================
                    Label3.Text = "Проверка наличия в БД акции с таким же названием"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MyContinueFlag = 1
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, CONVERT(nvarchar(30), MIN(tbl_ActionsAndSales.DateStart), 103) AS DateStart, "
                    MySQLStr = MySQLStr & "CONVERT(nvarchar(30), MAX(tbl_ActionsAndSales.DateFinish), 103) AS DateFinish, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN 'Уже стартовала' ELSE 'Еще не стартовала' END AS MyState "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN tbl_ActionsAndSales ON "
                    MySQLStr = MySQLStr & " " & MyTableName & ".ActionName = tbl_ActionsAndSales.ActionName "
                    MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales.ActionName, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN 'Уже стартовала' ELSE 'Еще не стартовала' END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Данные акции / распродажи уже присутствуют в Scala:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) & Declarations.MyRec.Fields("MyState").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) + "С" + Chr(9) & Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value & Chr(13) & Chr(10)
                            If (Declarations.MyRec.Fields("MyState").Value.ToString.Equals("Уже стартовала")) Then
                                MyContinueFlag = 0
                            End If
                            Declarations.MyRec.MoveNext()
                        End While
                        If (MyContinueFlag = 0) Then
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "Поменяйте название акции / распродажи в Excel файле на уникальное " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        Else
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то указанные " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "выше акции / распродажи будут удалены и заменены акциями из Excel. " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "Если вы не хотите удалять уже существующие акции / распродажи, то " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "выберите ""Выход"", " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "поменяйте название акции / распродажи в Excel файле на уникальное." & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        End If
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        If (MyContinueFlag = 0) Then
                            MyErrForm.Button2.Visible = False
                            MyErrForm.Button2.Enabled = False
                        Else
                            MyErrForm.Button2.Visible = True
                            MyErrForm.Button2.Enabled = True
                        End If
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Акция / распродажа уже существует.")
                        Else    '----удаление ранее занесенных акций
                            MySQLStr = "DELETE FROM tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & ".ActionName = tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales_1.ActionName) AS View_2 ON tbl_ActionsAndSales.ActionName = View_2.ActionName "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '==============================Закрытие (смена дат) акций в случае перекрытия=============================
                    Label3.Text = "Закрытие (смена дат) акций в случае перекрытия"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '----------------------------Диапазон дат акции / распродажи внутри уже существующего---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart >= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish <= "
                    MySQLStr = MySQLStr & "CASE WHEN tbl_ActionsAndSales.DateFinish = CONVERT(datetime, '31/12/9999', 103) THEN "
                    MySQLStr = MySQLStr & "dateadd(dd, - 1, tbl_ActionsAndSales.DateFinish) ELSE tbl_ActionsAndSales.DateFinish END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи находится внутри уже существующего диапазона:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Поменяйте Диапазон акции / распродажи в Excel файле " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        End If
                    End If

                    '---------------Диапазон дат акции / распродажи охватывает с обоих сторон существующий--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart <= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish >= tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи охватывает с обоих сторон существующий:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Поменяйте Диапазон акции / распродажи в Excel файле " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        End If
                    End If

                    '------------------Диапазон дат акции / распродажи Слева перекрывает уже существующие---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи по дате начала перекрывается с уже существующим:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то даты " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "окончания предыдущей акции / распродажи будут изменены на дату начала текущей акции - 1 день. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "Если вы не хотите менять таким образом дату окончания предыдущей акции / распродажи, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "то выберите ""Выход"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "поменяйте даты начала акции / распродажи в Excel файле на необходимую" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        Else    '-----меняем даты окончания акции
                            'MySQLStr = "UPDATE " & MyTableName & " "
                            'MySQLStr = MySQLStr & "SET DateStart = DateAdd(dd, 1, CASE WHEN View_2.DateFinish = CONVERT(datetime, '31/12/9999', 103) "
                            'MySQLStr = MySQLStr & "THEN dateadd(dd, - 1, View_2.DateFinish) ELSE View_2.DateFinish END)) "
                            'MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            'MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            'MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales.DateStart AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            'MySQLStr = MySQLStr & " " & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            MySQLStr = "Update tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart), "
                            MySQLStr = MySQLStr & "TimeAction = N'да' "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName, tbl_ActionsAndSales_1.ScalaCode, " & MyTableName & "_1.DateStart, "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales_1.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales_1.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales_1.DateFinish) AS View_2 ON "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales.ScalaCode = View_2.ScalaCode And tbl_ActionsAndSales.DateStart < View_2.DateStart And tbl_ActionsAndSales.DateFinish > View_2.DateStart "

                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '------------------Диапазон дат акции / распродажи Справа перекрывает уже существующие--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "Диапазон дат акции / распродажи по дате окончания перекрывается с уже существующим:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "Товар" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "С" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "По" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "Если вы выберите ""Продолжить процесс загрузки"", то даты " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "окончания акции / распродажи будут изменены на дату начала последующей  акции минус 1 день. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "Если вы не хотите менять таким образом дату окончания акции / распродажи, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "то выберите ""Выход"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "поменяйте даты окончания акции / распродажи в Excel файле на необходимыю" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "и загрузите файл по новой." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("Неверный диапазон дат в Excel файле.")
                        Else    '-----меняем даты начала акции
                            MySQLStr = "UPDATE " & MyTableName & " "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart) "
                            MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish > tbl_ActionsAndSales.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            MySQLStr = MySQLStr & "" & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If



                    '==============================Занесение акции / распродажи в БД=============================
                    Label3.Text = "Занесение акции / распродажи в БД"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "INSERT INTO tbl_ActionsAndSales "
                    MySQLStr = MySQLStr & "(ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, ActionFinished, ActionFinishedDate) "
                    MySQLStr = MySQLStr & "SELECT ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, 0 AS ActionFinished, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, '01/01/1900', 103) AS ActionFinishedDate "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    '==============================================Запуск пересчета прайс листа на продажу============================================
                    MyRez = MsgBox("Произвести расчет прайс - листа на продажу сейчас? Иначе расчет будет выполнен ночью.", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = MsgBoxResult.Yes Then
                        '----------------------------Подпись операции
                        Label3.Text = "Формирование прайс листа на продажу. "
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()

                        MySQLStr = "Exec spp_PrepareCommonPriceList_PriCost "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If


                Catch ex As Exception
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                Finally
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                    '----------------------------Подпись операции
                    Label3.Text = ""
                End Try

                Me.Cursor = Cursors.Default
                oWorkBook.Close(True)
            End If
        End If
    End Sub

    Private Function GetFirstExcelSheetName(ByRef cn As OleDbConnection) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение имени первого листа Excel  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyTable As String
        Dim dt As DataTable

        Try
            cn.Open()
            MyTable = cn.GetSchema("Tables").Rows(0)("TABLE_NAME")
            cn.Close()
            GetFirstExcelSheetName = MyTable
        Catch ex As Exception
            GetFirstExcelSheetName = ""
        End Try
    End Function

    Private Function GetExcelDataSet(ByRef cn As OleDbConnection, ByVal MySQLStr As String) As DataSet
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение dataset  Excel  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim cmd As OleDbDataAdapter
        Dim ds As New DataSet()

        Try
            cmd = New OleDbDataAdapter(MySQLStr, cn)
            cn.Open()
            cmd.Fill(ds, "Table1")
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        GetExcelDataSet = ds
    End Function

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка программы, загрузка исходных данных 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        LoadActionsList()
        DateTimePicker1.Value = CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))
    End Sub

    Private Sub LoadActionsList()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка акций, которые можно закрыть 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка акций
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ActionName, ActionName + ' From ' + CONVERT(nvarchar(30), DateStart, 103) + ' To ' + CONVERT(nvarchar(30), DateFinish, 103) AS ActionFullName "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionFinished = 0) "
        MySQLStr = MySQLStr & "GROUP BY ActionName, DateStart, DateFinish "
        MySQLStr = MySQLStr & "ORDER BY ActionFullName "
        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "ActionFullName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "ActionName"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие акции / распродажи 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If CheckDate() = True Then
            MySQLStr = "UPDATE tbl_ActionsAndSales "
            MySQLStr = MySQLStr & "SET ActionFinished = 1, "
            MySQLStr = MySQLStr & "ActionFinishedDate = CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103)"
            MySQLStr = MySQLStr & "WHERE (ActionName = N'" & ComboBox1.SelectedValue & "') "
            MySQLStr = MySQLStr & "AND (ActionFinished = 0) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MsgBox("Произведено закрытие " & ComboBox1.SelectedValue & ".", MsgBoxStyle.OkOnly, "Внимание!")
            LoadActionsList()
        Else
            DateTimePicker1.Select()
        End If
    End Sub

    Private Function CheckDate() As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка корректности ввода даты закрытия  акции / распродажи
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "SELECT DateStart, DateFinish "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionName = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (ActionFinished = 0) "
        MySQLStr = MySQLStr & "GROUP BY DateStart, DateFinish "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("Невозможно проверить корректность выставления даты закрытия акции / распродажи. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDate = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("DateStart").Value <= DateTimePicker1.Value And Declarations.MyRec.Fields("DateFinish").Value >= DateTimePicker1.Value _
                And CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())) <= DateTimePicker1.Value Then
                CheckDate = True
                Exit Function
            Else
                MsgBox("Дата закрытия акции / распродажи должна быть больше или равна текущей и быть в диапазоне от даты начала акции / распродажи до ококнчания акции / распродажи.", MsgBoxStyle.Critical, "Внимание!")
                CheckDate = False
                Exit Function
            End If
        End If
    End Function
End Class
