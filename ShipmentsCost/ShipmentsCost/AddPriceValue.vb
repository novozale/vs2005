Public Class AddPriceValue

    Private Sub AddPriceValue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по открытию окна
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        TextBox6.Text = MainForm.ComboBox1.Text
        TextBox1.Text = MainForm.ComboBox2.Text
        TextBox2.Text = MainForm.ComboBox3.Text
        If MainForm.ComboBox3.SelectedValue = 1 Then    '---килограммы
            Label5.Text = "Начиная с веса "
            Label6.Text = "По вес "
            Label7.Text = "Цена за кг (РУБ) "
        Else                                            '---кубометры
            Label5.Text = "Начиная с объема "
            Label6.Text = "По объем "
            Label7.Text = "Цена за куб м (РУБ) "
        End If

        MySQLStr = "SELECT ID, CONVERT(nvarchar, ID) + ' ' + Name AS Name "
        MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_PriceType WITH (NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ID"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox1.ValueMember = "ID"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение данных 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then

            Declarations.Destination = TextBox3.Text
            Declarations.PriceType = ComboBox1.SelectedValue
            Declarations.PriceFrom = TextBox4.Text
            Declarations.PriceTo = TextBox5.Text
            Declarations.PriceVal = TextBox7.Text
            Declarations.MinCost = TextBox8.Text
            Declarations.MySuccess = True                               'Успешность выполнения операции
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без добавления значения прайс - листа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле "Начиная с" вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Начиная с..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""Начиная с..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле "По" вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""По ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""По ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox7_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox7.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле "Цена" вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox7.Text) <> "" Then
            If InStr(TextBox7.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Цена ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox7.Text
                Catch ex As Exception
                    MsgBox("В поле ""Цена ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""Пункт назначения"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox3.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("Поле ""Начиная С ..."" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox4.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""По ..."" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox5.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox7.Text) = "" Then
            MsgBox("Поле ""Цена ..."" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox7.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox8.Text) = "" Then
            MsgBox("Поле ""Мин. цена ..."" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            TextBox8.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = 1 And Trim(TextBox3.Text) <> "Средняя по региону" Then
            MsgBox("Если вы выбрали тип прайс - листа ""За 100 километров по региону"", то в поле пункт назначения должно быть значение ""Средняя по региону"".", MsgBoxStyle.Critical, "Внимание")
            TextBox3.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        If ComboBox1.SelectedValue = 0 And Trim(TextBox3.Text) = "Средняя по региону" Then
            MsgBox("Если вы выбрали пункт назначения ""Средняя по региону"", то в поле тип прайс - листа должно быть значение ""За 100 километров по региону"".", MsgBoxStyle.Critical, "Внимание")
            ComboBox1.Focus()
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function

    Private Sub TextBox8_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox8.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле "Мин. цена" вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox8.Text) <> "" Then
            If InStr(TextBox8.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Мин. цена ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox8.Text
                Catch ex As Exception
                    MsgBox("В поле ""Мин. цена ..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub
End Class