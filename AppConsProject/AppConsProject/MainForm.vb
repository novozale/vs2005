Imports System.Runtime.InteropServices

Public Class MainForm
    <Runtime.InteropServices.DllImport( _
  "user32.dll", _
  CharSet:=CharSet.Auto, _
  CallingConvention:=CallingConvention.StdCall _
)> _
Public Shared Function SetWindowPos( _
  ByVal hWnd As IntPtr, _
  ByVal hWndInsertAfter As IntPtr, _
  ByVal X As Int32, _
  ByVal Y As Int32, _
  ByVal cx As Int32, _
  ByVal cy As Int32, _
  ByVal uFlags As Int32) _
  As Boolean
    End Function
    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Список складов ComboBox1
        WHList()

        'Каталог продуктов
        BuildProductList()

        'Запасы на консигнационном складе
        BuildConsStock()


    End Sub

    Private Sub WHList()
        Dim SQLStr As String
        Dim SQLAdapter As SqlClient.SqlDataAdapter
        Dim DS As New DataSet

        SQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        SQLStr = SQLStr & "FROM SC230300 WITH(NOLOCK) "
        SQLStr = SQLStr & "WHERE (LEFT(SC23001,1) = N'K') "
        SQLStr = SQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            SQLAdapter = New SqlClient.SqlDataAdapter(SQLStr, Declarations.NETConnStr)
            SQLAdapter.SelectCommand.CommandTimeout = 600
            SQLAdapter.Fill(DS)
            ComboBox1.DisplayMember = "SC23002"
            ComboBox1.ValueMember = "SC23001"
            ComboBox1.DataSource = DS.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub BuildProductList()
        Dim SQLStr As String
        Dim SQLAdapter As SqlClient.SqlDataAdapter
        Dim DS As New DataSet

        SQLStr = "SELECT SC01001, SC01002 + SC01003 AS SC01002  FROM ScaDataDB.dbo.SC010300 WITH(NOLOCK)"
        InitMyConn(False)
        Try
            SQLAdapter = New SqlClient.SqlDataAdapter(SQLStr, Declarations.NETConnStr)
            SQLAdapter.SelectCommand.CommandTimeout = 600
            SQLAdapter.Fill(DS)

            DataGridView1.DataSource = DS.Tables(0)

            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).HeaderText = "Запас"
            DataGridView1.Columns(1).Width = 1000

            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Private Sub BuildConsStock()

        Dim SQLStr As String
        Dim SQLAdapter As SqlClient.SqlDataAdapter
        Dim DS As New DataSet

        SQLStr = "SELECT Code, SC01002 + SC01003 as Name, MinQty, MaxQty"
        SQLStr = SQLStr & " FROM  [dbo].[tbl_ConsStocks] INNER JOIN  [dbo].[SC010300] ON Code = SC01001"
        SQLStr = SQLStr & " WHERE WH ='" & ComboBox1.SelectedValue & "'"

        InitMyConn(False)
        Try
            SQLAdapter = New SqlClient.SqlDataAdapter(SQLStr, Declarations.NETConnStr)
            SQLAdapter.SelectCommand.CommandTimeout = 600
            SQLAdapter.Fill(DS)
            DataGridView2.DataSource = DS.Tables(0)

            DataGridView2.Columns(0).HeaderText = "ID"
            DataGridView2.Columns(0).Width = 80
            DataGridView2.Columns(1).HeaderText = "Запас"
            DataGridView2.Columns(1).Width = 500
            DataGridView2.Columns(2).HeaderText = "Минимальный уровень"
            DataGridView2.Columns(2).Width = 100
            DataGridView2.Columns(3).HeaderText = "Максимальный уровень"
            DataGridView2.Columns(3).Width = 100

            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            CheckButtons()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CheckButtons()
        '--------------------------------------------------------------------------------
        '--
        '-- проверка состояния кнопок
        '--
        '--------------------------------------------------------------------------------

        If DataGridView2.SelectedRows.Count = 0 Then
            Button5.Enabled = False
            Button4.Enabled = False
        Else
            Button5.Enabled = True
            Button4.Enabled = True
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged

        'При изменении склада в ComboBox1 - перезагружаем данные в DataGridView2

        '---список запасов на выбранном консигнационном складе
        BuildConsStock()

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Выход из приложения
        Application.Exit()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        AddItemProduct()
    End Sub

    Private Sub AddItemProduct()

        Dim SQLStr As String
        FrmAddProduct = New AddProduct
        FrmAddProduct.ShowDialog()
        If Declarations.IsSuccess = False Then
            Exit Sub
        Else
            'Заносим данные в рабочу таблицу
            SQLStr = "INSERT INTO tbl_ConsStocks "
            SQLStr = SQLStr & "(Code, WH, MinQty, MaxQty ) "
            SQLStr = SQLStr & "VALUES('" & DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString() & "', "
            SQLStr = SQLStr & "'" & ComboBox1.SelectedValue & "', "
            SQLStr = SQLStr & Replace(CStr(Declarations.MinQty), ",", ".") & ", "
            SQLStr = SQLStr & Replace(CStr(Declarations.MaxQty), ",", ".") & ") "
            InitMyConn(False)
            Declarations.Conn.Execute(SQLStr)

        End If

        BuildConsStock()

    End Sub

    Private Sub EditItemProduct()
        Dim SQLStr As String
        FrmEditProduct = New EditProduct
        FrmEditProduct.ShowDialog()
        If Declarations.IsSuccess = False Then
            Exit Sub
        Else
            SQLStr = "UPDATE tbl_ConsStocks "
            SQLStr = SQLStr & "SET MinQty=" & Replace(CStr(Declarations.MinQty), ",", ".") & ","
            SQLStr = SQLStr & "MaxQty=" & Replace(CStr(Declarations.MaxQty), ",", ".") & " "
            SQLStr = SQLStr & "WHERE Code=N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            SQLStr = SQLStr & "AND (WH=N'" & ComboBox1.SelectedValue & "')"
            InitMyConn(False)
            Declarations.Conn.Execute(SQLStr)

        End If

        BuildConsStock()

    End Sub
    Private Sub RemoveItemProduct()

        Dim SQLStr As String

        Dim Message As DialogResult = MessageBox.Show("Удалить строку с уровнем запаса?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

        If Message = DialogResult.No Then Exit Sub

        If Message = DialogResult.Yes Then
            SQLStr = "DELETE FROM tbl_ConsStocks "
            SQLStr = SQLStr & "WHERE Code=N'" & Trim(DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            SQLStr = SQLStr & "AND (WH=N'" & ComboBox1.SelectedValue & "')"
            InitMyConn(False)
            Declarations.Conn.Execute(SQLStr)
        End If
        'Обновляем список
        BuildConsStock()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        EditItemProduct()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        RemoveItemProduct()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Dim Msg As MsgBoxResult
        Dim Str As String

        Str = "Для импорта данных необходимо подготовить файл Excel следующего формата: " & Chr(13) & Chr(10)
        Str = Str + "Колонка A - Код запаса" & Chr(13) & Chr(10)
        Str = Str + "Колонка B - Код склада" & Chr(13) & Chr(10)
        Str = Str + "Колонка C - Минимальный уровень запаса" & Chr(13) & Chr(10)
        Str = Str + "Колонка D - Максимальный уровень запаса" & Chr(13) & Chr(10)
        Msg = MsgBox(Str, MsgBoxStyle.OkCancel, "Внимание!!!")
        If (Msg = MsgBoxResult.Ok) Then
            If My.Settings.UseOffice = "LibreOffice" Then
                ImportDataFromLO()
            Else
                ImportDataFromExcel()
            End If

        Else
        End If
        BuildConsStock()
    End Sub

    Private Sub ImportDataFromExcel()

        Dim appXLSRC As Object
        Dim WH As String
        Dim Code As String
        Dim MinQ As Double
        Dim MaxQ As Double
        Dim SQLStr As String
        Dim RowNum As Integer

        OpenFileDialog1.FileName = ""
        OpenFileDialog1.ShowDialog()

        If (OpenFileDialog1.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

            appXLSRC = CreateObject("Excel.Application")
            appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)
            WH = appXLSRC.Worksheets(1).Range("B2").Value

            '---проверяем что в Excel проставлен код склада
            If WH = Nothing Then
                MsgBox("В импортируемом листе Excel в ячейке 'B2' не проставлен код склада в Scala ", MsgBoxStyle.Critical, "Внимание!")
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                Exit Sub

            End If

            RowNum = 2
            While Not appXLSRC.Worksheets(1).Range("A" & RowNum).Value = Nothing
                Code = appXLSRC.Worksheets(1).Range("A" & RowNum).Value.ToString()
                WH = appXLSRC.Worksheets(1).Range("B" & RowNum).Value.ToString()
                MinQ = appXLSRC.Worksheets(1).Range("C" & RowNum).Value.ToString()
                MaxQ = appXLSRC.Worksheets(1).Range("D" & RowNum).Value.ToString()

                DeleteDBProduct(Code, WH)
                InsertDBProduct(Code, WH, MinQ, MaxQ)

                RowNum = RowNum + 1

            End While

            'SQLStr = "SELECT COUNT(SC23001) AS CC "
            'SQLStr = SQLStr & "FROM SC230300 WITH(NOLOCK) "
            'SQLStr = SQLStr & "WHERE (SC23001 = N'" & WH & "')"
            'InitMyConn(False)
            'InitMyRec(False, SQLStr)

            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
            MsgBox("Импорт данных произведен.", MsgBoxStyle.OkOnly, "Внимание!")

            Me.Cursor = Cursors.Default
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()

        End If

    End Sub

    Private Sub ImportDataFromLO()
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim counter As Integer
        Dim MyDBL As Double
        Dim WH As String
        Dim Code As String
        Dim MinQ As Double
        Dim MaxQ As Double

        OpenFileDialog2.FileName = ""
        OpenFileDialog2.ShowDialog()

        If (OpenFileDialog2.FileName = "") Then
        Else
            Me.Cursor = Cursors.WaitCursor
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

                '-----правильность занесения данных в Calc
                counter = 2
                Do
                    '-----Код склада
                    If oSheet.getCellRangeByName("B" & counter).String.Equals("") Then
                        MsgBox("Строка " & CStr(counter) & " не проставлен код склада в Scala", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If
                    '-----минимальное количество
                    Try
                        MyDBL = oSheet.getCellRangeByName("C" & counter).Value
                    Catch ex As Exception
                        MsgBox("Строка " & CStr(counter) & " некорректно указано минимальное количество")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    '-----максимальное количество
                    Try
                        MyDBL = oSheet.getCellRangeByName("D" & counter).Value
                    Catch ex As Exception
                        MsgBox("Строка " & CStr(counter) & " некорректно указано максимальное количество")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try
                    counter = counter + 1
                Loop Until oSheet.getCellRangeByName("A" & counter).String.Equals("")

                '-----Занесение данных в Calc
                counter = 2
                Do
                    Code = oSheet.getCellRangeByName("A" & counter).String
                    WH = oSheet.getCellRangeByName("B" & counter).String
                    MinQ = oSheet.getCellRangeByName("C" & counter).Value
                    MaxQ = oSheet.getCellRangeByName("D" & counter).Value

                    DeleteDBProduct(Code, WH)
                    InsertDBProduct(Code, WH, MinQ, MaxQ)
                    counter = counter + 1
                Loop Until oSheet.getCellRangeByName("A" & counter).String.Equals("")
            Catch ex As Exception
                MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
            End Try
            Me.Cursor = Cursors.Default
            oWorkBook.Close(True)
            MsgBox("Импорт данных произведен.", MsgBoxStyle.OkOnly, "Внимание!")

            Me.Cursor = Cursors.Default
            Me.Refresh()
            System.Windows.Forms.Application.DoEvents()
        End If
    End Sub

    Private Sub InsertDBProduct(ByVal Code As String, ByVal WH As String, ByVal MinQ As Double, ByVal MaxQ As Double)
        Dim SQLStr As String
        '---заносим в рабочую таблицу
        SQLStr = "INSERT INTO [ScaDataDB].[dbo].[tbl_ConsStocks] "
        SQLStr = SQLStr & "(Code, WH, MinQty, MaxQty) "
        SQLStr = SQLStr & " VALUES("
        SQLStr = SQLStr & "N'" & Code & "', "
        SQLStr = SQLStr & "N'" & WH & "', "
        SQLStr = SQLStr & Replace(CStr(MinQ), ",", ".") & ", "
        SQLStr = SQLStr & Replace(CStr(MaxQ), ",", ".") & ") "
        InitMyConn(False)
        Declarations.Conn.Execute(SQLStr)

    End Sub

    Private Sub DeleteDBProduct(ByVal Code As String, ByVal WH As String)
        Dim SQLStr As String

        SQLStr = "DELETE FROM [ScaDataDB].[dbo].[tbl_ConsStocks]"
        SQLStr = SQLStr & " WHERE Code=N'" & Code & "'"
        SQLStr = SQLStr & " AND WH=N'" & WH & "'"
        InitMyConn(False)
        Declarations.Conn.Execute(SQLStr)
    End Sub

    Private Sub ExportToXls()

        Dim WH As String
        Dim SQLStr As String
        Dim i As Integer
        Dim StrNum As Double      'номер строки
        Dim Obj As Object       'Excel
        Dim WRKBook As Object   'книга

        WH = ComboBox1.SelectedValue.ToString()
        SQLStr = "SELECT  a.Code, b.SC01002 + b.SC01003 as Name, a.WH, a.MinQty, a.MaxQty "
        SQLStr = SQLStr + " FROM [ScaDataDB].[dbo].[tbl_ConsStocks] a INNER JOIN [dbo].[SC010300] b"
        SQLStr = SQLStr + " ON Code  = SC01001"
        SQLStr = SQLStr + " WHERE a.WH = N'" & WH & "'"
        InitMyConn(False)
        InitMyRec(False, SQLStr)

        Obj = CreateObject("Excel.Application")
        Obj.SheetsInNewWorkbook = 1
        WRKBook = Obj.Workbooks.Add

        StrNum = 2

        WRKBook.ActiveSheet.Range("A1") = "Код продукта"
        WRKBook.ActiveSheet.Range("B1") = "Название"
        WRKBook.ActiveSheet.Range("C1") = "Код склада"
        WRKBook.ActiveSheet.Range("D1") = "Мин. кол-во"
        WRKBook.ActiveSheet.Range("E1") = "Макс. кол-во"

        'Declarations.Rec.MoveFirst()
        While Declarations.Rec.EOF <> True
            WRKBook.ActiveSheet.Range("A" & StrNum) = Declarations.Rec.Fields("Code").Value
            WRKBook.ActiveSheet.Range("B" & StrNum) = Declarations.Rec.Fields("Name").Value
            WRKBook.ActiveSheet.Range("C" & StrNum) = Declarations.Rec.Fields("WH").Value
            WRKBook.ActiveSheet.Range("D" & StrNum) = Declarations.Rec.Fields("MinQty").Value
            WRKBook.ActiveSheet.Range("E" & StrNum) = Declarations.Rec.Fields("MaxQty").Value

            StrNum = StrNum + 1
            Declarations.Rec.MoveNext()
        End While
        UploadCommonHeader(WRKBook)
        trycloseRec()

        WRKBook.ActiveSheet.Range("A1").Select()
        Obj.Application.Visible = True
        Obj = Nothing

    End Sub

    Private Sub ExportToOds()

        Dim WH As String
        Dim SQLStr As String
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim StrNum As Integer      'номер строки

        WH = ComboBox1.SelectedValue.ToString()
        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        UploadCommonHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher)

        WH = ComboBox1.SelectedValue.ToString()
        SQLStr = "SELECT  a.Code, b.SC01002 + b.SC01003 as Name, a.WH, a.MinQty, a.MaxQty "
        SQLStr = SQLStr + " FROM [ScaDataDB].[dbo].[tbl_ConsStocks] a INNER JOIN [dbo].[SC010300] b"
        SQLStr = SQLStr + " ON Code  = SC01001"
        SQLStr = SQLStr + " WHERE a.WH = N'" & WH & "'"
        InitMyConn(False)
        InitMyRec(False, SQLStr)
        StrNum = 1

        While Declarations.Rec.EOF <> True
            StrNum = StrNum + 1

            oSheet.getCellRangeByName("A" & StrNum).String = Declarations.Rec.Fields("Code").Value
            oSheet.getCellRangeByName("B" & StrNum).String = Declarations.Rec.Fields("Name").Value
            oSheet.getCellRangeByName("C" & StrNum).String = Declarations.Rec.Fields("WH").Value
            oSheet.getCellRangeByName("D" & StrNum).Value = Declarations.Rec.Fields("MinQty").Value
            oSheet.getCellRangeByName("E" & StrNum).Value = Declarations.Rec.Fields("MaxQty").Value

            Declarations.Rec.MoveNext()
        End While

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$2:$E$" & StrNum
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "Обычный"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub UploadCommonHeaderLO(ByRef oSheet, ByRef oServiceManager, ByRef oWorkBook, ByRef oDispatcher)

        oSheet.getCellRangeByName("A1").String = "Код продукта"
        oSheet.getCellRangeByName("B1").String = "Название"
        oSheet.getCellRangeByName("C1").String = "Код склада"
        oSheet.getCellRangeByName("D1").String = "Мин. кол-во"
        oSheet.getCellRangeByName("E1").String = "Макс. кол-во"

        oSheet.getColumns().getByName("A").Width = 4000
        oSheet.getColumns().getByName("B").Width = 11000
        oSheet.getColumns().getByName("C").Width = 4000
        oSheet.getColumns().getByName("D").Width = 3000
        oSheet.getColumns().getByName("E").Width = 3000
        oSheet.getCellRangeByName("A1").Rows.Height = 1000

        oSheet.getCellRangeByName("A1:E1").CellBackColor = 12510163
        oSheet.getCellRangeByName("A1:E1").VertJustify = 2
        oSheet.getCellRangeByName("A1:E1").HoriJustify = 2

        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A1:E1").TopBorder = LineFormat
        oSheet.getCellRangeByName("A1:E1").RightBorder = LineFormat
        oSheet.getCellRangeByName("A1:E1").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A1:E1").BottomBorder = LineFormat

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1:$E$1"
        Dim oFrame As Object
        oFrame = oWorkBook.getCurrentController.getFrame
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "Обычный"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = "Calibri"
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "Bold"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args2)
    End Sub

    Private Function UploadCommonHeader(ByVal WRKBook As Object)

        WRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        WRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 40
        WRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 30
        WRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 15
        WRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15

    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If My.Settings.UseOffice = "LibreOffice" Then
            ExportToOds()
        Else
            ExportToXls()
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        If Trim(TextBox2.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox2.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), Len(UCase(Trim(TextBox2.Text)))), 1) = 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim Rez As Object

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    Rez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If Rez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            SelectList = New ItemSelectList
            SelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If Trim(TextBox4.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox4.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), Len(UCase(Trim(TextBox4.Text)))), 1) = 0 Then
                    DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox4.Select()
        Else
            SelectList2 = New ItemSelectList2
            SelectList2.ShowDialog()
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 Then
                    DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim Rez As Object

        If Trim(TextBox4.Text) = "" And Trim(TextBox1.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox4.Select()
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView2.CurrentCellAddress.Y + 1 To DataGridView2.Rows.Count
                If i = DataGridView2.Rows.Count Then
                    Rez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If Rez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView2.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(0, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox4.Text))) <> 0 And InStr(UCase(Trim(DataGridView2.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 Then
                        DataGridView2.CurrentCell = DataGridView2.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub
End Class