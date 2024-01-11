Public Class AddRelOrder

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Закрытие окна ввода заказа на перемещение
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// добавление заказа на перемещение
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String

        If Trim(ComboBox1.SelectedItem) = "" Then
            MsgBox("Необходимо выбрать, на какой склад произведено перемещение.", MsgBoxStyle.Critical, "Внимание")
        Else
            '---проверка что в этот заказ не включены другие заказы на перемещение (должен быть только 1!!!)
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices "
            MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value <> 0 Then
                trycloseMyRec()
                MsgBox("В данную доставку уже включены документы. Информация о перемещении может вноситься в документ только 1 раз, информация о перемещении всегда должна заноситься отдельно от остальных документов (СФ на продажу и инвойсов на закупку)", MsgBoxStyle.Critical, "Внимание")
            Else
                trycloseMyRec()
                '---Добавление заказа в таблицу
                MySQLStr = "INSERT INTO tbl_ShipmentsCost_FactByInvoices "
                MySQLStr = MySQLStr & "SELECT NEWID() AS ID, "
                MySQLStr = MySQLStr & "'" & Declarations.MyRecordID & "' AS DocID, "
                If Trim(ComboBox1.SelectedItem) = "01 Санкт Петербург" Then
                    MySQLStr = MySQLStr & "'Перемещение на WH01', "
                Else
                    MySQLStr = MySQLStr & "'Перемещение на WH03', "
                End If
                MySQLStr = MySQLStr & "1 AS InvoiceSumm, "
                MySQLStr = MySQLStr & "NULL AS ShipmentCost, "
                MySQLStr = MySQLStr & "3 AS DocType, "
                MySQLStr = MySQLStr & "NULL AS SupplierCode, "
                MySQLStr = MySQLStr & "NULL AS DocYear "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                Me.Close()
            End If
        End If
    End Sub
End Class