Public Class AddRelOrder

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ����� ������ �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ �� �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String

        If Trim(ComboBox1.SelectedItem) = "" Then
            MsgBox("���������� �������, �� ����� ����� ����������� �����������.", MsgBoxStyle.Critical, "��������")
        Else
            '---�������� ��� � ���� ����� �� �������� ������ ������ �� ����������� (������ ���� ������ 1!!!)
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices "
            MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value <> 0 Then
                trycloseMyRec()
                MsgBox("� ������ �������� ��� �������� ���������. ���������� � ����������� ����� ��������� � �������� ������ 1 ���, ���������� � ����������� ������ ������ ���������� �������� �� ��������� ���������� (�� �� ������� � �������� �� �������)", MsgBoxStyle.Critical, "��������")
            Else
                trycloseMyRec()
                '---���������� ������ � �������
                MySQLStr = "INSERT INTO tbl_ShipmentsCost_FactByInvoices "
                MySQLStr = MySQLStr & "SELECT NEWID() AS ID, "
                MySQLStr = MySQLStr & "'" & Declarations.MyRecordID & "' AS DocID, "
                If Trim(ComboBox1.SelectedItem) = "01 ����� ���������" Then
                    MySQLStr = MySQLStr & "'����������� �� WH01', "
                Else
                    MySQLStr = MySQLStr & "'����������� �� WH03', "
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