Public Class EditCustom

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна без добавления ручных значений МЖЗ, ROP и страхового запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MySuccess = False
        Me.Close()
    End Sub

    Private Sub EditCustom_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка информации по открытию окна
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Label3.Text = "Редактирование ручных значений МЖЗ, ROP и страхового уровня для запаса " & Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & " " & Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString())
        TextBox6.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(8).Value.ToString())
        TextBox4.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(12).Value.ToString())
        TextBox3.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(16).Value.ToString())
        TextBox2.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(6).Value.ToString())
        TextBox1.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(10).Value.ToString())
        TextBox5.Text = Trim(MainForm.DataGridView2.SelectedRows.Item(0).Cells(14).Value.ToString())
    End Sub

    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле МЖЗ вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox2.Text) <> "" Then
            If InStr(TextBox2.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""МЖЗ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox2.Text
                Catch ex As Exception
                    MsgBox("В поле ""МЖЗ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка, что в поле ROP вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox1.Text) <> "" Then
            If InStr(TextBox1.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""ROP"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox1.Text
                Catch ex As Exception
                    MsgBox("В поле ""ROP"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
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
        '// Проверка, что в поле уровень страхового запаса вводится число
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Страховой уровень"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""Страховой уровень"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// пересчет пропорции ROP и страхового запаса при условии что МЖЗ введено
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) <> "" Then
            If CDbl(TextBox6.Text) <> 0 Then
                TextBox1.Text = Math.Round(CDbl(TextBox2.Text) * CDbl(TextBox4.Text) / CDbl(TextBox6.Text), 3)
                TextBox5.Text = Math.Round(CDbl(TextBox2.Text) * CDbl(TextBox3.Text) / CDbl(TextBox6.Text), 3)
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение данных 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then

            Declarations.MyMGZ = TextBox2.Text
            Declarations.MyROP = TextBox1.Text
            Declarations.MyInsuranceLVL = TextBox5.Text
            Declarations.MySuccess = True                               'Успешность выполнения операции
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""МЖЗ"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""ROP"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Страховой уровень"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function
End Class