Imports System.IO
Public Class CASH_CustomUpload

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор каталога для выгрузки файлов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--отмена выбора
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CASH_CustomUpload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = My.Settings.CASHCatalog
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка наличия каталога выгрузки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) <> "" Then
            If Directory.Exists(Trim(TextBox1.Text)) = False Then
                MsgBox("Введенный каталог не существует. Введите корректный или выберите.", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка "Условной" номенклатуры
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            Me.Cursor = Cursors.WaitCursor
            DataUpload()
            Me.Cursor = Cursors.Default
            MsgBox("Выгрузка данных завершена.", MsgBoxStyle.Information, "Внимание!")
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Каталог для выгрузки должен быть заполнен. Введите корректный или выберите.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckData = False
            Exit Function
        End If

        If Directory.Exists(Trim(TextBox1.Text)) = False Then
            MsgBox("Введенный каталог не существует. Введите корректный или выберите.", MsgBoxStyle.Critical, "Внимание!")
            TextBox1.Select()
            CheckData = False
            Exit Function
        End If
        CheckData = True
    End Function

    Private Sub DataUpload()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция выгрузки "условной" номенклатуры
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyFile As String
        Dim MyFlagFile As String
        Dim MyWrkStr As String

        MyFile = Trim(TextBox1.Text) + "\" + "goods.txt"
        MyFlagFile = Trim(TextBox1.Text) + "\" + "goods_flag.txt"

        '-----------Очистка каталога
        If File.Exists(MyFile) = True Then
            Try
                File.Delete(MyFile)
            Catch ex As Exception
                MsgBox("Невозможно очистить выбранный каталог. Попробуйте позже или обратитесь к администратору. " + ex.Message, MsgBoxStyle.Critical, "Внимание!")
                Exit Sub
            End Try
        End If
        If File.Exists(MyFlagFile) = True Then
            Try
                File.Delete(MyFlagFile)
            Catch ex As Exception
                MsgBox("Невозможно очистить выбранный каталог. Попробуйте позже или обратитесь к администратору. " + ex.Message, MsgBoxStyle.Critical, "Внимание!")
                Exit Sub
            End Try
        End If

        '----------Создание файла и заполнение
        Dim f As New StreamWriter(MyFile, False, System.Text.Encoding.GetEncoding(1251))
        '---налоговые ставки
        MyWrkStr = "$$$ADDTAXRATES" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;НДС 0%;;;0" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "2;НДС 10%;;;10" + vbCrLf
        f.Write(MyWrkStr)
        If Now < CDate("01/01/2019") Then
            MyWrkStr = "3;НДС 18%;;;18" + vbCrLf
        Else
            MyWrkStr = "3;НДС 20%;;;20" + vbCrLf
        End If
        f.Write(MyWrkStr)
        MyWrkStr = "4;Без НДС;;;0" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "5;Расчетная ставка 10/110;;;10" + vbCrLf
        f.Write(MyWrkStr)
        If Now < CDate("01/01/2019") Then
            MyWrkStr = "6;Расчетная ставка 18/118;;;18" + vbCrLf
        Else
            MyWrkStr = "6;Расчетная ставка 20/120;;;20" + vbCrLf
        End If
        f.Write(MyWrkStr)
        '---налоговые группы
        MyWrkStr = "$$$ADDTAXGROUPRATES" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;2;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;3;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;4;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;5;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;10;1" + vbCrLf
        f.Write(MyWrkStr)
        '---Удаление всех товаров
        MyWrkStr = "$$$DELETEALLWARES" + vbCrLf
        f.Write(MyWrkStr)
        '---выгрузка всех товаров
        MyWrkStr = "$$$ADDQUANTITY" + vbCrLf
        f.Write(MyWrkStr)

        MyWrkStr = "00000000;;;Электрооборудование;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000000;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000001;;;Услуги;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000001;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000002;;;Основные средства б.у;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000002;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000003;;;МПЗ (ПК, мониторы);0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000003;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000004;;;Прочие товары;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000004;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)

        '---закрытие файла
        f.Close()
        '---выгрузка файла - флага
        Dim f1 As New StreamWriter(MyFlagFile, False, System.Text.Encoding.GetEncoding(1251))
        f1.Close()
    End Sub
End Class