Public Class CorrectRequestDate
    Public MyID As Integer
    Public MyDate As DateTime

    Private Sub CorrectRequestDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub CorrectRequestDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DateTimePicker1.Value = MyDate
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� � �����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckFormFilling() = True Then
            '----���������� �����������
            If SaveRequest() = True Then
                Me.Close()
            End If
        End If
    End Sub

    Private Function CheckFormFilling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ����� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----���� ������ ��� ����� �������
        If DateTimePicker1.Value < Now().AddDays(-1) Then
            MsgBox("���� ""���� ������������� ��"" ������ ���� ������ ��� ����� �������", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        '-----����� 15:00 ���� ������ ������ �������
        If Hour(Now()) >= 15 Then
            If DateTimePicker1.Value < Now() Then
                MsgBox("����� 15:00 ���� ""���� ������������� ��"" ������ ���� ������ �������", MsgBoxStyle.Critical, "��������!")
                DateTimePicker1.Select()
                CheckFormFilling = False
                Exit Function
            End If
        End If

        '-----������ �������� ������� � �����������
        If Weekday(DateTimePicker1.Value, 2) = 6 Or Weekday(DateTimePicker1.Value, 2) = 7 Then
            MsgBox("���� ""���� ������������� ��"" �� ������ ���� �������� ��� ������������.", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckFormFilling = False
            Exit Function
        End If

        CheckFormFilling = True
    End Function

    Private Function SaveRequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ ��������� � �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        MySQLStr = "UPDATE tbl_SupplSearch "
        MySQLStr = MySQLStr & "SET RequestDate = CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103) "
        MySQLStr = MySQLStr & "WHERE (ID = " & CStr(MyID) & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        SaveRequest = True
    End Function
End Class