Public Class CreditDialog

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ����� �� ����� � ����������� �� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox6.Text = "" Then
            MsgBox("��������� ������� �������� ������ ��� ���������� ���������� ������.", vbCritical, "��������!")
            TextBox6.Select()
            Exit Sub
        Else
            Declarations.CmdToShip = True
            Declarations.MyReason = TextBox6.Text
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���������� � �������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyCreditInfo = New CreditInfo
        MyCreditInfo.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ����� �� �����, ���������� �� �������� ���
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CreditDialog_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim CardPayedSum As Double          '---���������� �� �������� � WEB �����

        TextBox1.Text = Declarations.OrderSum
        TextBox2.Text = Declarations.Avance1Type + Declarations.Avance2Type
        TextBox3.Text = Declarations.InvoiceDebt
        TextBox4.Text = Declarations.OrderDebt
        TextBox5.Text = "RUR"
        TextBox7.Text = Declarations.CreditAmount
        TextBox8.Text = Declarations.OverduePaymentQTY
        TextBox9.Text = Declarations.CreditInDays
        TextBox10.Text = Declarations.Overdue
        If Declarations.MyPermission = True Then
            Button3.Visible = True
            Button3.Enabled = True
            TextBox6.Visible = True
            Label10.Visible = True
        Else
            Button3.Visible = False
            Button3.Enabled = False
            TextBox6.Visible = False
            Label10.Visible = False
        End If

        '---------���������� �� WEB
        Declarations.IsWEBOrder = CheckWEBOrNot(Declarations.OrderID)
        If Declarations.IsWEBOrder = 0 Then '---�� �������� ������� � WEB �����
            Label11.Visible = False
            Label12.Visible = False
        Else
            CardPayedSum = GetCardPayment(Declarations.OrderID)
            Label12.Text = "�� WEB ����� �������� " & CStr(CardPayedSum) & " ���"
            Label11.Visible = True
            Label12.Visible = True
        End If
    End Sub
End Class