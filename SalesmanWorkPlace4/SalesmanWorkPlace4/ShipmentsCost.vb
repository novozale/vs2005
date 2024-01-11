Public Class ShipmentsCost

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ¬ыход из окна без сохранени€ результатов
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ¬ыход из окна с сохранением результатов
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        'Dim MyPercent As Double        'на какой % увеличиваем стоимость заказа
        Dim MyOrder As String           'Ќомер заказа
        Dim DelCost As Double           'стоимость доставки
        'Dim i As Integer
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        'Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        'Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        'Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        'Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        'Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        'Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        'Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        'Dim MyParam9 As ADODB.Parameter                 'передаваемый параметр номер 10
        'Dim MyParam10 As ADODB.Parameter                'передаваемый параметр номер 11

        If CheckFilling = True Then
            MySQLStr = "DELETE FROM tbl_SW4SalesHdr_AddInfo "
            'MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(MyOrderLines.Label6.Text) & "')"
            MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(MyEditHeader.Label3.Text) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            If CDbl(NumericUpDown1.Text) <> 0 Then
                MySQLStr = "INSERT INTO tbl_SW4SalesHdr_AddInfo "
                MySQLStr = MySQLStr & "(ID, OrderID, DeliverySum) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                'MySQLStr = MySQLStr & "N'" & Trim(MyOrderLines.Label6.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(MyEditHeader.Label3.Text) & "', "
                MySQLStr = MySQLStr & Replace(NumericUpDown1.Text, ",", ".") & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---увеличение стоимости заказа на стоимость доставки (построчно)
                If CheckBox1.Checked = True Then
                    'MyPercent = CDbl(NumericUpDown1.Text) / (CDbl(MyOrderLines.Label24.Text) * Declarations.CurrencyValue)
                    'MyOrder = Trim(MyOrderLines.Label6.Text)

                    'cmd.ActiveConnection = Declarations.MyConn
                    'cmd.CommandText = "spp_SalesWorkplace4_EditOrder"
                    'cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    'cmd.CommandTimeout = 300

                    'MyParam = cmd.CreateParameter("@Order", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
                    'cmd.Parameters.Append(MyParam)
                    'MyParam1 = cmd.CreateParameter("@Str", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
                    'cmd.Parameters.Append(MyParam1)
                    'MyParam2 = cmd.CreateParameter("@MyItemID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
                    'cmd.Parameters.Append(MyParam2)
                    'MyParam3 = cmd.CreateParameter("@MyName1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
                    'cmd.Parameters.Append(MyParam3)
                    'MyParam4 = cmd.CreateParameter("@MyName2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
                    'cmd.Parameters.Append(MyParam4)
                    'MyParam5 = cmd.CreateParameter("@Cost", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                    'cmd.Parameters.Append(MyParam5)
                    'MyParam6 = cmd.CreateParameter("@CostIntr", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                    'cmd.Parameters.Append(MyParam6)
                    'MyParam7 = cmd.CreateParameter("@Qty", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                    'cmd.Parameters.Append(MyParam7)
                    'MyParam8 = cmd.CreateParameter("@Unit", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                    'cmd.Parameters.Append(MyParam8)
                    'MyParam9 = cmd.CreateParameter("@Discount", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
                    'cmd.Parameters.Append(MyParam9)
                    'MyParam10 = cmd.CreateParameter("@EditOrRecalc", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                    'cmd.Parameters.Append(MyParam10)
                    'For i = 0 To MyOrderLines.DataGridView2.Rows.Count - 1
                    '    MyParam.Value = MyOrder
                    '    MyParam1.Value = MyOrderLines.DataGridView2.Rows(i).Cells(1).Value
                    '    MyParam2.Value = ""
                    '    MyParam3.Value = ""
                    '    MyParam4.Value = ""
                    '    MyParam5.Value = MyOrderLines.DataGridView2.Rows(i).Cells(6).Value * (1 + MyPercent)
                    '    MyParam6.Value = 0
                    '    MyParam7.Value = 0
                    '    MyParam8.Value = 0
                    '    MyParam9.Value = MyOrderLines.DataGridView2.Rows(i).Cells(10).Value
                    '    MyParam10.Value = 1
                    '    Try
                    '        cmd.Execute()
                    '    Catch ex As Exception
                    '        MsgBox(ex.ToString)
                    '    End Try
                    'Next
                    MyOrder = Trim(MyEditHeader.Label3.Text)
                    DelCost = CDbl(NumericUpDown1.Text)

                    cmd.ActiveConnection = Declarations.MyConn
                    cmd.CommandText = "spp_SalesWorkplace4_AddDelCostToPrice"
                    cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    cmd.CommandTimeout = 300

                    MyParam = cmd.CreateParameter("@Order", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
                    cmd.Parameters.Append(MyParam)
                    MyParam1 = cmd.CreateParameter("@DelCost", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                    cmd.Parameters.Append(MyParam1)

                    MyParam.Value = MyOrder
                    MyParam1.Value = DelCost

                    Try
                        cmd.Execute()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
                '----—умма доставки не 0 - выставл€ем доставку
                MyEditHeader.ComboBox4.SelectedItem = "ƒоставка до двери"
                MySQLStr = "UPDATE tbl_OR010300 "
                MySQLStr = MySQLStr & "SET PriceCond = N 'ƒоставка до двери' "
                MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(MyEditHeader.Label3.Text) & "')"
                Try
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex As Exception
                End Try
            Else
                '----—умма доставки 0 - выставл€ем самовывоз
                MyEditHeader.ComboBox4.SelectedItem = "—амовывоз со склада"
                MySQLStr = "UPDATE tbl_OR010300 "
                MySQLStr = MySQLStr & "SET PriceCond = N '—амовывоз со склада' "
                MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Trim(MyEditHeader.Label3.Text) & "')"
                Try
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                Catch ex As Exception
                End Try
            End If
            Me.Close()
        End If
    End Sub

    Private Function CheckFilling() As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ѕроверка внесени€ суммы
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRes As Double

        Try
            MyRes = CDbl(NumericUpDown1.Text)
            CheckFilling = True
        Catch ex As Exception
            MsgBox("Ќеобходимо внести число в поле стоимости доставки", MsgBoxStyle.Critical, "¬нимание!")
            NumericUpDown1.Focus()
            CheckFilling = False
        End Try
    End Function

    Private Sub ShipmentsCost_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// «агрузка окна, загрузка данных
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT DeliverySum "
        MySQLStr = MySQLStr & "FROM tbl_SW4SalesHdr_AddInfo WITH (NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(Trim(MyOrderLines.Label6.Text)) & "')"
        MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(Trim(MyEditHeader.Label3.Text)) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            NumericUpDown1.Text = 0
        Else
            Declarations.MyRec.MoveFirst()
            NumericUpDown1.Text = CStr(Declarations.MyRec.Fields("DeliverySum").Value)
        End If
        trycloseMyRec()

        '---доставка - процент от суммы доставленного
        MySQLStr = "SELECT CONVERT(float, tbl_CustomerCard0300.TransportKoeff) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON tbl_OR010300.OR01003 = tbl_CustomerCard0300.SL01001 "
        'MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01001 = N'" & Trim(MyOrderLines.Label6.Text) & "')"
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01001 = N'" & Trim(MyEditHeader.Label3.Text) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            NumericUpDown2.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                NumericUpDown2.Text = ""
                Button3.Enabled = False
            Else
                NumericUpDown2.Text = CStr(Math.Round(Declarations.MyRec.Fields("CC").Value, 2))
                Button3.Enabled = True
            End If
        End If
        trycloseMyRec()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// –асчет стоимости доставки как процент
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////

        If NumericUpDown2.Text <> "" Then
            NumericUpDown1.Text = Math.Round(CDbl(NumericUpDown2.Text) * CDbl(MyOrderLines.Label24.Text) * CDbl(MyOrderLines.Label9.Text) / 100, 2)
        End If
    End Sub
End Class