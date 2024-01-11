Public Class AddSInvoice

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ����� ������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������� �� �������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ��������� ����� ���� ������� �� ������� (� ���������).", MsgBoxStyle.Critical, "��������")
        Else
            '---��������, ��� � ���� �������� �� �������� ������ �� �����������
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices "
            MySQLStr = MySQLStr & "WHERE (DocID = '" & Declarations.MyRecordID & "') "
            MySQLStr = MySQLStr & "AND (DocType = 3) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value <> 0 Then
                trycloseMyRec()
                MsgBox("� ������ �������� ��� �������� ��������� �� �����������. ���������� � ����������� ������ ������ ���������� �������� �� ��������� ���������� (�� �� ������� � �������� �� �������)", MsgBoxStyle.Critical, "��������")
            Else
                trycloseMyRec()
                '---�������� ��� ����� ������ ����
                MySQLStr = "SELECT COUNT(dbo.ST030300.ST03001) AS CC "
                MySQLStr = MySQLStr & "FROM dbo.ST030300 WITH (NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "dbo.OR200300 ON dbo.ST030300.ST03009 = dbo.OR200300.OR20001 "
                MySQLStr = MySQLStr & "WHERE (dbo.OR200300.OR20109 + dbo.ST030300.ST03014 = N'" & Trim(TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    MsgBox("����� ���� ������� �� ������� (� ���������) � ���� ���. ������� ���������� ��������.", MsgBoxStyle.Critical, "��������")
                Else
                    trycloseMyRec()
                    '---�������� ��� ���� ������ �� ������� � ������ ��������� �� ���������
                    MySQLStr = "SELECT COUNT(ID) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices WITH (NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SL03002 = N'" & Trim(TextBox1.Text) & "') "
                    MySQLStr = MySQLStr & "AND (DocType = 1) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.Fields("CC").Value <> 0 Then
                        trycloseMyRec()
                        MsgBox("����� ���� ������� �� ������� ��� ������� � ��������� ��������. ������� ������ ��������.", MsgBoxStyle.Critical, "��������")
                    Else
                        trycloseMyRec()
                        '---���������� ������� � �������
                        MySQLStr = "INSERT INTO tbl_ShipmentsCost_FactByInvoices "
                        MySQLStr = MySQLStr & "SELECT NEWID() AS ID, "
                        MySQLStr = MySQLStr & "'" & Declarations.MyRecordID & "' AS DocID, "
                        MySQLStr = MySQLStr & "View_1.OR20109 + ST030300.ST03014 AS SL03002, "
                        MySQLStr = MySQLStr & "SUM(ROUND(ST030300.ST03021 * ST030300.ST03020, 2) - ROUND(ST030300.ST03021 * ST030300.ST03020 * ST030300.ST03022 / 100, 2)) AS InvoiceSumm, "
                        MySQLStr = MySQLStr & "NULL AS Expr1, 1 AS Expr2, NULL AS Expr3, NULL AS Expr4  "
                        MySQLStr = MySQLStr & "FROM ST030300 INNER JOIN "
                        MySQLStr = MySQLStr & "(SELECT OR20001, OR20109 "
                        MySQLStr = MySQLStr & "FROM OR200300 "
                        MySQLStr = MySQLStr & "GROUP BY OR20001, OR20109) AS View_1 ON ST030300.ST03009 = View_1.OR20001 "
                        MySQLStr = MySQLStr & "WHERE (View_1.OR20109 + ST030300.ST03014 = N'" & Trim(TextBox1.Text) & "') "
                        MySQLStr = MySQLStr & "GROUP BY View_1.OR20109 + ST030300.ST03014 "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        Me.Close()
                    End If
                End If
            End If
        End If
    End Sub
End Class