Public Class AddPInvoice

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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

        If CheckDataFiling() = True Then

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
                '---�������� ��� ����� ����� ����
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM PL030300 WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (PL03001 = N'" & Trim(TextBox2.Text) & "') "
                MySQLStr = MySQLStr & "AND (PL03002 = N'" & Trim(TextBox1.Text) & "') "
                MySQLStr = MySQLStr & "AND (YEAR(PL03004) = " & Trim(TextBox3.Text) & ") "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                    MsgBox("������ ������� �� ������� � ���� ���. ������� ���������� ��������.", MsgBoxStyle.Critical, "��������")
                Else
                    trycloseMyRec()
                    '---�������� ��� ���� ������ �� ������� � ������ ��������� �� ���������
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM tbl_ShipmentsCost_FactByInvoices "
                    MySQLStr = MySQLStr & "WHERE (SL03002 = N'" & Trim(TextBox1.Text) & "') "
                    MySQLStr = MySQLStr & "AND (SupplierCode = N'" & Trim(TextBox2.Text) & "') "
                    MySQLStr = MySQLStr & "AND (DocYear = N'" & Trim(TextBox3.Text) & "') "
                    MySQLStr = MySQLStr & "AND (DocType = 2) "
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
                        MySQLStr = MySQLStr & "PL03002 AS SL03002, "
                        MySQLStr = MySQLStr & "PL03013 AS InvoiceSumm, "
                        MySQLStr = MySQLStr & "NULL AS ShipmentCost, "
                        MySQLStr = MySQLStr & "2 AS DocType, "
                        MySQLStr = MySQLStr & "PL03001 AS SupplierCode, "
                        MySQLStr = MySQLStr & "YEAR(PL03004) "
                        MySQLStr = MySQLStr & "FROM PL030300 "
                        MySQLStr = MySQLStr & "WHERE (PL03002 = N'" & Trim(TextBox1.Text) & "') "
                        MySQLStr = MySQLStr & "AND (PL03001 = N'" & Trim(TextBox2.Text) & "') "
                        MySQLStr = MySQLStr & "AND (YEAR(PL03004) = " & Trim(TextBox3.Text) & ") "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        Me.Close()
                    End If
                End If
            End If
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ���������� ����� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyInt As Integer            '---��� �������� ��������������
        Dim MyRet As Object             '---��� ������

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���������� ��������� ����� ������� �� �������.", MsgBoxStyle.Critical, "��������")
            TextBox1.Select()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("���������� ��������� ��� ����������.", MsgBoxStyle.Critical, "��������")
            TextBox2.Select()
            CheckDataFiling = False
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("���������� ��������� ��� ��������� (������� ����������).", MsgBoxStyle.Critical, "��������")
            TextBox3.Select()
            CheckDataFiling = False
            Exit Function
        End If

        CheckDataFiling = True
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ������ ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "AddPInvoice"
        MySupplierSelect.ShowDialog()
    End Sub

    Private Sub TextBox3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ����� ���������� � ���� ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyInt As Integer            '---��� �������� ��������������
        Dim MyRet As Object             '---��� ������

        If TextBox3.Modified = True Then
            If Trim(TextBox3.Text) <> "" Then
                Try
                    MyInt = TextBox3.Text
                Catch ex As Exception
                    MsgBox("��� ��������� (������� ����������) ������ ���� 4 - � ������� ������ ", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End Try

                If (CInt(TextBox3.Text) < 0) Then
                    MsgBox("��� ��������� (������� ����������) ������ ���� ������ 0 ", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End If

                If Len(Trim(TextBox3.Text)) <> 4 Then
                    MsgBox("��� ��������� (������� ����������) ������ ���� ������ 4 ������� ��� �������� ", MsgBoxStyle.Critical, "��������!")
                    e.Cancel = True
                    Exit Sub
                End If

                If (CInt(TextBox3.Text) <> Now().Year) Then
                    MyRet = MsgBox("��������� ���� ��� ��������� (������� ����������) �� ��������� � �������. �� �������, ��� ��� ������ ���������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRet = vbYes Then

                    Else
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ����� ���������� � ���� ��� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If TextBox2.Modified = True Then
            If Trim(TextBox2.Text) <> "" Then
                MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(TextBox2.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MsgBox("�� ����� �������� ��� ����������. ������� ���������� ��� �������������� �������.", vbCritical, "��������!")
                    e.Cancel = True
                    Exit Sub
                Else
                    trycloseMyRec()
                End If
            End If
        End If
    End Sub

    Private Sub AddPInvoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox3.Text = Now().Year
    End Sub
End Class