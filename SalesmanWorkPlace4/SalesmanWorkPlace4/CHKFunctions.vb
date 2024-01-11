Module CHKFunctions

    Public Function CheckDataInProposal(ByVal MyOrder As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Для известного номера предложения - проверка корректности его заполнения
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim cmd As New ADODB.Command
        Dim MyCustomerCode

        '----Покупатель есть в Scala
        MySQLStr = "SELECT ISNULL(SL010300.SL01001, N'') AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON tbl_OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "WHERE     (tbl_OR010300.OR01001 = N'" & MyOrder & "') "
        MySQLStr = MySQLStr & "GROUP BY SL010300.SL01001 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            MyCustomerCode = Declarations.MyRec.Fields("CC").Value
            If Trim(MyCustomerCode) <> "" Then
                CheckDataInProposal = True
            Else
                MsgBox("Покупатель данного заказа отсутсвует в Scala. Для перевода предложения в счет 0 типа необходимо завести данного покупателя в Scala или в предложении о покупке поменять покупателя.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '-----продавец должен иметь тот же кост центр, что и покупатель
        Dim MyRezStr As String
        MyRezStr = CheckSalesman(Declarations.SalesmanCode, MyCustomerCode)
        If MyRezStr <> "" Then
            MsgBox(MyRezStr, MsgBoxStyle.OkOnly, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        End If


        '----В заказе есь строки
        MySQLStr = "SELECT COUNT(OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("В данном предложении нет ни одной строки. Для перевода предложения в счет 0 типа необходимо завести хотя бы 1 строку.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----Все запасы есть в Scala
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(SC010300.SC01001 IS NULL)"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("В данном заказе есть запасы, которые отсутсвуют в Scala. Для перевода предложения в счет 0 типа необходимо завести данные запасы в Scala или в предложении о покупке поменять запасы на существующие в Scala.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----Запасы в заказе не являются заблокированными
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard0300 ON SC010300.SC01001 = tbl_ItemCard0300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(tbl_ItemCard0300.IsBlocked = N'1') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("В данном заказе есть заблокированные запасы. Для перевода предложения в счет 0 типа необходимо удалить эти запасы из предложения, или в предложении о покупке поменять запасы на незаблокированные в Scala.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----В заказ не включены составные запасы
        MySQLStr = "SELECT COUNT(tbl_OR030300.OR03005) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(SC010300.SC01066 = 8) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("В данном заказе есть составные запасы. Использование их в предложении о покупке запрещено. Для перевода предложения в счет 0 типа необходимо удалить эти запасы из предложения.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '----Проверка, что выставлены корректные даты (больше текущей)
        MySQLStr = "SELECT PriceCond, ReadyDate, DeliveryDate, ExpirationDate "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & MyOrder & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            trycloseMyRec()
            Exit Function
        Else
            'If Declarations.MyRec.Fields("PriceCond").Value = "Доставка до двери" Then
            'If Declarations.MyRec.Fields("DeliveryDate").Value < Now() Then
            'MsgBox("В данном заказе срок доставки меньше текущей даты. Откорректируйте это значение.", MsgBoxStyle.Critical, "Внимание!")
            'CheckDataInProposal = False
            'trycloseMyRec()
            'Exit Function
            'Else
            'CheckDataInProposal = True
            'End If
            'Else
            'If Declarations.MyRec.Fields("ReadyDate").Value < Now() Then
            '    MsgBox("В данном заказе срок готовности к отгрузке меньше текущей даты. Откорректируйте это значение.", MsgBoxStyle.Critical, "Внимание!")
            '    CheckDataInProposal = False
            '    trycloseMyRec()
            '    Exit Function
            'Else
            '    CheckDataInProposal = True
            'End If
            'End If
            If Declarations.MyRec.Fields("ExpirationDate").Value < Now() Then
                MsgBox("В данном заказе срок действия предложения меньше текущей даты. Откорректируйте это значение.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                trycloseMyRec()
                Exit Function
            Else
                CheckDataInProposal = True
            End If
        End If
        trycloseMyRec()

        '----Проверка, что везде выставлены сроки поставки
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM  tbl_OR030300 "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') AND "
        MySQLStr = MySQLStr & "(WeekQTY IS NULL) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Ошибка перевода предложения в заказ 0 типа. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                CheckDataInProposal = True
            Else
                MsgBox("В данном заказе есть строки с непроставленными сроками поставки. Для перевода предложения в счет 0 типа необходимо во всех строках проставить сроки поставки.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataInProposal = False
                Exit Function
            End If
        End If
        trycloseMyRec()

        '-----Проверка, что единицы измерения в КП и в Scala равны
        cmd.ActiveConnection = Declarations.MyConn
        cmd.CommandText = "spp_Agents_CheckCPData"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.CommandTimeout = 300

        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2

        Dim MyRSTR As String
        MyRSTR = ""

        MyParam = cmd.CreateParameter("@MyCPID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
        cmd.Parameters.Append(MyParam)
        MyParam.Value = MyOrder

        MyParam1 = cmd.CreateParameter("@MyRetValue", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
        cmd.Parameters.Append(MyParam1)
        MyParam1.Value = ""

        Try
            cmd.Execute()
            MyRSTR = MyParam1.Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        End Try

        If Trim(MyRSTR) <> "" Then
            MsgBox("В коммерческом предложении есть запасы, единицы измерения которых отличаются от значений в карточке товара: " & Chr(13) & Chr(10) & Trim(MyRSTR), MsgBoxStyle.Critical, "Внимание!")
            CheckDataInProposal = False
            Exit Function
        End If


        CheckDataInProposal = True
    End Function
End Module
