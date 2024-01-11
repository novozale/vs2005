Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.VisualBasic.FileIO
Imports SalesComissionStarter.spbprd2
Imports System.Collections.ObjectModel


Public Class SalesComissionStarter

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MyMonth.SelectedIndex = CInt(Microsoft.VisualBasic.DateAndTime.Month(DateAdd(DateInterval.Month, -1, Now())) - 1)
        MyCatalogR.Text = My.Settings.DistributionCatalogPath

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FolderBrowserDialog1 = New FolderBrowserDialog
        'FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.SelectedPath = My.Settings.InitCatalog
        Dim dr As DialogResult = FolderBrowserDialog1.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            MyCatalogR.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1 = New OpenFileDialog()
        OpenFileDialog1.InitialDirectory = My.Settings.InitCatalog
        OpenFileDialog1.Filter = "*.xls|*.xls"
        Dim dr As DialogResult = OpenFileDialog1.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            MyFileR.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        AllocateData(MyMonth.Text, MyFileR.Text, MyCatalogR.Text)
        MsgBox("Закончено распределение отчетов", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub AllocateData(ByVal MyMonth As String, ByVal MyFile As String, ByVal MyCatalog As String)
        Dim appXLSRC As Object                             'Excel объект - общий отчет
        Dim appXLDST As Object                             'Excel объект - накопительный отчет по продавцу
        Dim Counter As Double
        Dim PBar As Double
        Dim Salesman As String
        Dim cc As Object
        Dim LineNumber As Integer

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(MyFile)
        appXLDST = CreateObject("Excel.Application")
        appXLDST.Application.displayalerts = False

        Counter = 100 / (My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls").Count)
        PBar = 0
        ProgressBar2.Value = 0
        Application.DoEvents()
        '---список файлов, по которым разносим данные-----------------------------------
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls")
            appXLDST.Workbooks.Open(foundFile)
            '---Создание нового листа за месяц из шаблона-------------------------------
            Try
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Delete()
            Catch ex As Exception
            End Try
            Try
                appXLDST.Worksheets("Шаблон").copy(appXLDST.Worksheets("Шаблон"))
                appXLDST.Worksheets("Шаблон (2)").name = GetRowNumber(MyMonth)
                appXLDST.ActiveWorkbook.Close(SaveChanges:=True)
                '---Поск продавцов----------------------------------------------------------
                appXLDST.Workbooks.Open(foundFile)
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("A2") = Microsoft.VisualBasic.Year(Today())
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D2") = MyMonth
                Salesman = "S"
                LineNumber = 9
                Do While ((Microsoft.VisualBasic.Left(Salesman, 1) = "S") Or (Microsoft.VisualBasic.Left(Salesman, 1) = "P"))
                    LineNumber = LineNumber + 1
                    Salesman = appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("C" + CStr(LineNumber)).Value
                    cc = appXLSRC.Worksheets(1).Range("C3:C1000").Find(Salesman)
                    '---перенос значений--------------------------------------------------------
                    If Not cc Is Nothing Then
                        '--- роль в группе
                        If (appXLSRC.Worksheets(1).Range("D" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D" + CStr(LineNumber)) = CStr(appXLSRC.Worksheets(1).Range("D" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D" + CStr(LineNumber)) = ""
                        End If

                        '--- % участия в группе
                        If (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("E" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("E" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("F" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("F" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("F" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("F" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("G" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("G" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("G" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("G" + CStr(LineNumber)) = 0
                        End If

                        '--- Исключено из расчета - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("H" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("H" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("H" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("H" + CStr(LineNumber)) = 0
                        End If

                        '--- Исключено из расчета - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("I" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("I" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("I" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("I" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой от 12 до 16% вкл. - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("J" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("J" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("J" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("J" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой от 12 до 16% вкл. - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("K" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("K" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("K" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("K" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой от 16 до 22% вкл. - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("L" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("L" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("L" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("L" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой от 16 до 22% вкл. - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("M" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("M" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("M" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("M" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой от 22 до 27% вкл. - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("N" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("N" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("N" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("N" + CStr(LineNumber)) = 0
                        End If
                        '--- Оплаченные отгрузки с маржой от 22 до 27% вкл. - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("O" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("O" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("O" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("o" + CStr(LineNumber)) = 0
                        End If

                        '--- Оплаченные отгрузки с маржой более 27% - Реализация без неликвидов, руб.
                        If (appXLSRC.Worksheets(1).Range("P" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("P" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("P" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("P" + CStr(LineNumber)) = 0
                        End If
                        '--- Оплаченные отгрузки с маржой более 27% - Маржа, руб.
                        If (appXLSRC.Worksheets(1).Range("Q" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("Q" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("Q" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("Q" + CStr(LineNumber)) = 0
                        End If

                        '--- Коэффициент выполнения плана продаж
                        If (appXLSRC.Worksheets(1).Range("V" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("W" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("V" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("W" + CStr(LineNumber)) = 0
                        End If

                        '--- Коэффициент активности
                        If (appXLSRC.Worksheets(1).Range("Z" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("X" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("Z" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("X" + CStr(LineNumber)) = 0
                        End If

                        '--- Комиссия единовременная за новых клиентов
                        If (appXLSRC.Worksheets(1).Range("AA" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AB" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AA" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AA" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AB" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AA" + CStr(LineNumber)) = 0
                        End If

                        '--- Сумма продаж новым клиентам с учетом коэффициента, руб.
                        If (appXLSRC.Worksheets(1).Range("AD" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AB" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AD" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AB" + CStr(LineNumber)) = 0
                        End If

                        '--- Комиссия за продажу неликвидов группы
                        If (appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AF" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AD" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AF" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) / 100
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AD" + CStr(LineNumber)) = 0
                        End If

                        '---Комиссия за продажу через Web-сайт
                        If (appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) / 100
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = 0
                        End If


                        '--- План группы, руб.
                        If (appXLSRC.Worksheets(1).Range("S" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AI" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("S" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AI" + CStr(LineNumber)) = 0
                        End If

                        '--- Факт группы, руб.
                        If (appXLSRC.Worksheets(1).Range("T" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AJ" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("T" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AJ" + CStr(LineNumber)) = 0
                        End If

                        '--- Фактически отработанное количество дней
                        If (appXLSRC.Worksheets(1).Range("W" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AL" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("W" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AL" + CStr(LineNumber)) = 0
                        End If

                        '--- Количество визитов
                        If (appXLSRC.Worksheets(1).Range("X" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AM" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("X" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AM" + CStr(LineNumber)) = 0
                        End If

                        '--- Сумма доставки
                        If (appXLSRC.Worksheets(1).Range("R" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AO" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("R" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AO" + CStr(LineNumber)) = 0
                        End If
                    End If
                    Salesman = appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("C" + CStr(LineNumber + 1)).Value
                Loop
                '---окончание переноса значений---------------------------------------------

            Catch ex As Exception
            End Try
            appXLDST.ActiveWorkbook.Close(SaveChanges:=True)
            PBar = PBar + Counter
            ProgressBar2.Value = PBar
            Application.DoEvents()
        Next
        appXLSRC.ActiveWorkbook.Close(SaveChanges:=True)
        appXLSRC.Quit()
        appXLSRC = Nothing
        appXLDST.Application.displayalerts = False
        appXLDST.Quit()
        appXLDST = Nothing

    End Sub

    Private Sub TotalData(ByVal MyMonth As String, ByVal MyFile As String, ByVal MyCatalog As String)
        Dim appXLSRC As Object                             'Excel объект - общий отчет
        Dim appXLDST As Object                             'Excel объект - накопительный отчет по продавцу
        Dim Counter As Double
        Dim PBar As Double
        Dim Salesman As String
        Dim cc As Object
        Dim LineNumber As Integer
        Dim Count As Integer

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(MyFile)
        appXLDST = CreateObject("Excel.Application")
        appXLDST.Application.displayalerts = False

        Counter = 100 / (My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls").Count)
        PBar = 0
        ProgressBar2.Value = 0
        Application.DoEvents()
        '---список файлов, по которым разносим данные-----------------------------------
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls")
            appXLDST.Workbooks.Open(foundFile)
            Try
                '---Расчет суммарных значений-----------------------------------------------
                '---создание нового листа "Total" из "шаблона"
                appXLDST.Application.displayalerts = False
                appXLDST.Worksheets("Total").Delete()
                appXLDST.Worksheets("Шаблон").copy(appXLDST.sheets(1))
                appXLDST.Application.displayalerts = False
                appXLDST.Worksheets("Шаблон (2)").name = "Total"
                appXLDST.Worksheets("Total").Range("A2") = Microsoft.VisualBasic.Year(Today())
                appXLDST.Worksheets("Total").Range("D2") = "С начала года"
                '---Расчет
                Salesman = "S"
                LineNumber = 9
                Do While (Microsoft.VisualBasic.Left(Salesman, 1) = "S" Or (Microsoft.VisualBasic.Left(Salesman, 1) = "P")) 'Перебор всех продавцов на листе "Total" начиная с 10-ой строки
                    LineNumber = LineNumber + 1                     'Переход на следующую строку
                    appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber) + ":Y" + CStr(LineNumber)).value = 0 '---Очистка ячеек от введенных формул
                    Salesman = appXLDST.Worksheets("Total").Range("C" + CStr(LineNumber)).Value
                    Count = 1
                    Do While Count <= appXLDST.Worksheets.count         'Перебор всех листов в рабочей книге для суммирования
                        If (appXLDST.Worksheets(Count).Name <> "Total") And (appXLDST.Worksheets(Count).Name <> "Шаблон") Then
                            cc = appXLDST.Worksheets(Count).Range("C3:C1000").Find(Salesman)
                            If Not cc Is Nothing Then
                                '--- роль в группе
                                If (appXLDST.Worksheets(Count).Range("D" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("D" + CStr(LineNumber)) = CStr(appXLDST.Worksheets(Count).Range("D" + CStr(cc.Row)).Value)
                                End If

                                '--- % участия в группе
                                If (appXLDST.Worksheets(Count).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("E" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("F" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("F" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("F" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("F" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("G" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("G" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("G" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("G" + CStr(cc.Row)).Value)
                                End If

                                '--- Исключено из расчета - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("H" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("H" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("H" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("H" + CStr(cc.Row)).Value)
                                End If

                                '--- Исключено из расчета - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("I" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("I" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("I" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("I" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой от 12 до 16% вкл. - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("J" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("J" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("J" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("J" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой от 12 до 16% вкл. - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("K" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("K" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("K" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("K" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой от 16 до 22% вкл. - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("L" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("L" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("L" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("L" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой от 16 до 22% вкл. - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("M" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("M" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("M" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("M" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой от 22 до 27% вкл. - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("N" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("N" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("N" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("N" + CStr(cc.Row)).Value)
                                End If
                                '--- Оплаченные отгрузки с маржой от 22 до 27% вкл. - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("O" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("O" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("O" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("O" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой более 27% - Реализация без неликвидов, руб.
                                If (appXLDST.Worksheets(Count).Range("P" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("P" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("P" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("P" + CStr(cc.Row)).Value)
                                End If

                                '--- Оплаченные отгрузки с маржой более 27% - Маржа, руб.
                                If (appXLDST.Worksheets(Count).Range("Q" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("Q" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("Q" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("Q" + CStr(cc.Row)).Value)
                                End If

                                '--- Коэффициент выполнения плана продаж
                                If (appXLDST.Worksheets(Count).Range("W" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("W" + CStr(cc.Row)).Value)
                                End If

                                '--- Коэффициент активности
                                If (appXLDST.Worksheets(Count).Range("X" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("X" + CStr(cc.Row)).Value)
                                End If

                                '--- Комиссия с учетом коэффициентов
                                If (appXLDST.Worksheets(Count).Range("Y" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("Y" + CStr(cc.Row)).Value)
                                End If

                                '--- Комиссия единовременная за новых клиентов
                                If (appXLDST.Worksheets(Count).Range("AA" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AA" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AA" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AA" + CStr(cc.Row)).Value)
                                End If

                                '--- Комиссия от продаж новым клиентам
                                If (appXLDST.Worksheets(Count).Range("AB" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AB" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AB" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AB" + CStr(cc.Row)).Value)
                                End If

                                '--- Комиссия за продажу неликвидов группы
                                If (appXLDST.Worksheets(Count).Range("AD" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AD" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AD" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AD" + CStr(cc.Row)).Value)
                                End If

                                '---Комиссия за продажу через Web-сайт
                                If (appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) / 100
                                Else
                                    appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = 0
                                End If

                                '--- План группы, руб.
                                If (appXLDST.Worksheets(Count).Range("AI" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AI" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AH" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AH" + CStr(cc.Row)).Value)
                                End If

                                '--- Факт группы, руб.
                                If (appXLDST.Worksheets(Count).Range("AJ" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AJ" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AI" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AI" + CStr(cc.Row)).Value)
                                End If

                                '--- Фактически отработанное количество дней
                                If (appXLDST.Worksheets(Count).Range("AL" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AL" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AK" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AK" + CStr(cc.Row)).Value)
                                End If

                                '--- Количество визитов
                                If (appXLDST.Worksheets(Count).Range("AM" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AM" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AL" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AL" + CStr(cc.Row)).Value)
                                End If

                                '--- Сумма доставки
                                If (appXLDST.Worksheets(Count).Range("AO" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AO" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AN" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AN" + CStr(cc.Row)).Value)
                                End If

                            End If
                        End If
                        Count = Count + 1       'Переход на следующий лист
                    Loop
                    '--- Нахождение средних значений
                    'appXLDST.Worksheets("Total").Range("AO" + CStr(LineNumber)) = appXLDST.Worksheets.count()

                    '--- % участия в группе
                    'appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)

                    '--- Коэффициент выполнения плана продаж
                    'appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)

                    '--- Коэффициент активности продавца
                    'appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)


                    '--- Переход к следующему продавцу
                    Salesman = appXLDST.Worksheets("Total").Range("C" + CStr(LineNumber + 1)).Value
                Loop
            Catch ex As Exception
            End Try
            appXLDST.ActiveWorkbook.Close(SaveChanges:=True)

            PBar = PBar + Counter
            ProgressBar2.Value = PBar
            Application.DoEvents()
        Next
        appXLSRC.ActiveWorkbook.Close(SaveChanges:=True)
        appXLSRC.Quit()
        appXLSRC = Nothing
        appXLDST.Quit()
        appXLDST = Nothing

    End Sub

    Private Function GetRowNumber(ByVal MyStr) As String
        Select Case LTrim(RTrim(MyStr))
            Case "Январь"
                Return "01"
            Case "Февраль"
                Return "02"
            Case "Март"
                Return "03"
            Case "Апрель"
                Return "04"
            Case "Май"
                Return "05"
            Case "Июнь"
                Return "06"
            Case "Июль"
                Return "07"
            Case "Август"
                Return "08"
            Case "Сентябрь"
                Return "09"
            Case "Октябрь"
                Return "10"
            Case "Ноябрь"
                Return "11"
            Case "Декабрь"
                Return "12"
            Case Else
                Return "1012"
        End Select
    End Function

    Public Sub New()
        InitializeComponent()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TotalData(MyMonth.Text, MyFileR.Text, MyCatalogR.Text)
        MsgBox("Закончен пересчет отчетов", MsgBoxStyle.Information, "Внимание!")
    End Sub

End Class
