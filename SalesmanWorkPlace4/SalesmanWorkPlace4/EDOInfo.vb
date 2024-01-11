Public Class EDOInfo

    Private Sub EDOInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна ввода дополнительной информации по ЭДО
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT CustomerPONum, CustomerAgreementNum, CustomerManagerName, DeliveryAddress, GovermentID, InternalComment, "
        MySQLStr = MySQLStr & "CustomerAgreementDateStart, CustomerAgreementDateFin "
        MySQLStr = MySQLStr & "FROM tbl_SalesHdrCP_EDOInfo "
        MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(MyEditHeader.Label3.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            trycloseMyRec()
            '-----получение информации по договору
            MySQLStr = "SELECT AgreementN, DataFrom AS DateFrom, DataTo AS DateTo "
            MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 "
            MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(MyEditHeader.TextBox1.Text) & "') AND (DataFrom <= GETDATE()) AND (DataTo >= GETDATE()) OR "
            MySQLStr = MySQLStr & "(SL01001 = N'" & Trim(MyEditHeader.TextBox1.Text) & "') AND (DataTo <> DataFrom AND DataTo = CONVERT(datetime, '01/01/1900', 103)) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                TextBox2.Text = ""
                DateTimePicker1.Value = CDate(Format("dd/MM/yyyy", "01/01/1900"))
                DateTimePicker2.Value = CDate(Format("dd/MM/yyyy", "01/01/1900"))
                trycloseMyRec()
            Else
                Declarations.MyRec.MoveFirst()
                If Trim(Declarations.MyRec.Fields("AgreementN").Value).Equals("") Then
                    TextBox2.Text = ""
                    DateTimePicker1.Value = CDate(Format("dd/MM/yyyy", "01/01/1900"))
                    DateTimePicker2.Value = CDate(Format("dd/MM/yyyy", "01/01/1900"))
                Else
                    TextBox2.Text = Declarations.MyRec.Fields("AgreementN").Value
                    DateTimePicker1.Value = Declarations.MyRec.Fields("DateFrom").Value
                    DateTimePicker2.Value = Declarations.MyRec.Fields("DateTo").Value
                End If
                trycloseMyRec()
            End If
        Else
            Declarations.MyRec.MoveFirst()
            TextBox1.Text = Declarations.MyRec.Fields("CustomerPONum").Value
            TextBox2.Text = Declarations.MyRec.Fields("CustomerAgreementNum").Value
            TextBox3.Text = Declarations.MyRec.Fields("CustomerManagerName").Value
            TextBox4.Text = Declarations.MyRec.Fields("DeliveryAddress").Value
            TextBox5.Text = Declarations.MyRec.Fields("GovermentID").Value
            TextBox6.Text = Declarations.MyRec.Fields("InternalComment").Value
            DateTimePicker1.Value = Declarations.MyRec.Fields("CustomerAgreementDateStart").Value
            DateTimePicker1.Value = Declarations.MyRec.Fields("CustomerAgreementDateFin").Value
            trycloseMyRec()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранение значений
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением значений
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckValues = True Then
            MySQLStr = "DELETE FROM tbl_SalesHdrCP_EDOInfo "
            MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(MyEditHeader.Label3.Text) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO tbl_SalesHdrCP_EDOInfo "
            MySQLStr = MySQLStr & "(OrderID, CustomerPONum, CustomerAgreementNum, CustomerManagerName, DeliveryAddress, GovermentID, InternalComment, "
            MySQLStr = MySQLStr & "CustomerAgreementDateStart, CustomerAgreementDateFin) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyEditHeader.Label3.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox1.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox2.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox3.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox4.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox5.Text) & "', "
            MySQLStr = MySQLStr & "N'" & Trim(TextBox6.Text) & "', "
            MySQLStr = MySQLStr & "Convert(datetime, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103), "
            MySQLStr = MySQLStr & "Convert(datetime, '" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103)) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            Me.Close()
        End If
    End Sub

    Private Function CheckValues() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка правильности заполнения строк
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyStr As String

        MyStr = Trim(TextBox5.Text)
        If Len(MyStr) <> 0 And Len(MyStr) <> 20 And Len(MyStr) <> 25 Then
            MsgBox("Длина идентификатора госконтракта должна быть 20 или 25 знаков.", MsgBoxStyle.Critical, "Внимание!")
            TextBox5.Select()
            CheckValues = False
        Else
            CheckValues = True
        End If
    End Function
End Class