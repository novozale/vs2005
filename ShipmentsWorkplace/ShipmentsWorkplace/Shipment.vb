Public Class Shipment
    Public MyAction As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Shipment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����, ����������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustDelAddress As String

        CustDelAddress = ""
        '------���������� �� ����������------------------
        'MySQLStr = "SELECT     CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) <> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) "
        'MySQLStr = MySQLStr & "ELSE SL010300.SL01002 END AS CustomerName, CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) "
        'MySQLStr = MySQLStr & "<> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) ELSE Ltrim(Rtrim(Ltrim(Rtrim(SL010300.SL01003)) + ' ' + Ltrim(Rtrim(SL010300.SL01004)) "
        'MySQLStr = MySQLStr & "+ ' ' + Ltrim(Rtrim(SL010300.SL01005)))) END AS CustomerAddress, SL010300.SL01021 AS CustomerINN, "
        'MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(SL140300.SL14004)) + ' ' + LTRIM(RTRIM(SL140300.SL14005)) + ' ' + LTRIM(RTRIM(SL140300.SL14006)))) "
        'MySQLStr = MySQLStr & "AS CustomerDelAddress "
        'MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
        'MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON SL010300.SL01001 = tbl_CustomerCard0300.SL01001 INNER JOIN "
        'MySQLStr = MySQLStr & "SL140300 ON SL010300.SL01001 = SL140300.SL14001 "
        'MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(Declarations.MyCustomerCode) & "') "
        'MySQLStr = MySQLStr & "AND (SL140300.SL14002 = N'00') "

        MySQLStr = "SELECT CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) <> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongName)) "
        MySQLStr = MySQLStr & "ELSE SL010300.SL01002 END AS CustomerName, CASE WHEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) "
        MySQLStr = MySQLStr & "<> '' THEN Ltrim(Rtrim(tbl_CustomerCard0300.LongAddress)) ELSE Ltrim(Rtrim(Ltrim(Rtrim(SL010300.SL01003)) + ' ' + Ltrim(Rtrim(SL010300.SL01004)) "
        MySQLStr = MySQLStr & "+ ' ' + Ltrim(Rtrim(SL010300.SL01005)))) END AS CustomerAddress, SL010300.SL01021 AS CustomerINN, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_10.SL14004, N''))) + ' ' + LTRIM(RTRIM(ISNULL(View_10.SL14005, N''))) "
        MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(View_10.SL14006, N''))))) AS CustomerDelAddress "
        MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON SL010300.SL01001 = tbl_CustomerCard0300.SL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SL14001, SL14004, SL14005, SL14006 "
        MySQLStr = MySQLStr & "FROM SL140300 "
        MySQLStr = MySQLStr & "WHERE (SL14002 = N'00')) AS View_10 ON SL010300.SL01001 = View_10.SL14001 "
        MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(Declarations.MyCustomerCode) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("�� ������� ���������� �� ������� " & Trim(Declarations.MyCustomerCode) & ". ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            trycloseMyRec()
            Me.Close()
        Else
            LblCustomerCode.Text = Declarations.MyCustomerCode
            LblCustomerName.Text = Declarations.MyRec.Fields("CustomerName").Value
            LblCustomerLegalAddress.Text = Declarations.MyRec.Fields("CustomerAddress").Value
            LblCustomerINN.Text = Declarations.MyRec.Fields("CustomerINN").Value
            CustDelAddress = Declarations.MyRec.Fields("CustomerDelAddress").Value
            trycloseMyRec()
        End If

        '------���������� �� ��������--------------------
        LblSalesmanCode.Text = Declarations.SalesmanCode
        LblSalesmanName.Text = Declarations.UserName
        LblWHCode.Text = Declarations.MyWH

        '------���������� �� ��������--------------------
        '------��������
        If MyAction = "Create" Then
            ComboBox1.Text = "��������"
            TextBox1.Text = ""
            TextBox3.Text = ""
            '---����� ��������
            MySQLStr = "SELECT COUNT(SL14001) AS CC "
            MySQLStr = MySQLStr & "FROM SL140300 "
            MySQLStr = MySQLStr & "WHERE (SL14002 <> N'00') AND (SL14001 = N'" & Trim(Declarations.MyCustomerCode) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                CustDelAddress = ""
                trycloseMyRec()
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                Else
                    CustDelAddress = ""
                End If
                trycloseMyRec()
            End If
            TextBox2.Text = CustDelAddress
            '-----������� (� ����������� �� ������������)
            If MyContact = 1 Then
                MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(SL01006)) + ' ' + LTRIM(RTRIM(SL01007)) + ' ' + LTRIM(RTRIM(SL01008)) + ' ' + LTRIM(RTRIM(SL01009)) "
                MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(SL01010)) + ' ' + LTRIM(RTRIM(SL01011)) + ' ' + LTRIM(RTRIM(SL01012)) + ' ' + LTRIM(RTRIM(SL01013)))) AS Contact "
                MySQLStr = MySQLStr & "FROM SL010300 "
                MySQLStr = MySQLStr & "WHERE (SL01001 = N'" & Trim(Declarations.MyCustomerCode) & "')"
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    TextBox1.Text = ""
                    trycloseMyRec()
                Else
                    TextBox1.Text = Declarations.MyRec.Fields("Contact").Value
                    trycloseMyRec()
                End If
            End If
        ElseIf MyAction = "Copy" Then
            '------�����������
            If MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "��������" Then
                ComboBox1.Text = "��������"
            ElseIf MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "�������� � ������� ��������" Then
                ComboBox1.Text = "�������� � ������� ��������"
            Else
                ComboBox1.Text = "���������"
            End If
            TextBox1.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(6).Value
            TextBox2.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(7).Value
            TextBox3.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(8).Value
            TextBox5.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(18).Value
            TextBox6.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(19).Value
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(9).Value) = "" Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(10).Value) = "" Then
                CheckBox2.Checked = False
            Else
                CheckBox2.Checked = True
            End If
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(11).Value) = "" Then
                CheckBox3.Checked = False
            Else
                CheckBox3.Checked = True
            End If

        Else            '---Edit
            '------��������������
            If MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "��������" Then
                ComboBox1.Text = "��������"
            ElseIf MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(3).Value = "�������� � ������� ��������" Then
                ComboBox1.Text = "�������� � ������� ��������"
            Else
                ComboBox1.Text = "���������"
            End If
            TextBox1.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(6).Value
            TextBox2.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(7).Value
            TextBox3.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(8).Value
            TextBox5.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(18).Value
            TextBox6.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(19).Value
            DateTimePicker1.Value = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(12).Value
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(9).Value) = "" Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(10).Value) = "" Then
                CheckBox2.Checked = False
            Else
                CheckBox2.Checked = True
            End If
            If Trim(MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(11).Value) = "" Then
                CheckBox3.Checked = False
            Else
                CheckBox3.Checked = True
            End If
            TextBox4.Text = MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(17).Value
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� � ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If CheckFiields = True Then
            If MyAction = "Create" Or MyAction = "Copy" Then
                MySQLStr = "INSERT INTO tbl_Shipments_SalesmanWP_Info "
                MySQLStr = MySQLStr & "(CustomerCode, CustomerName, CustomerINN, CustomerLegalAddress, SalesmanCode, WHCode, DeliveryOrNot, DeliverySumm, DeliveredSumm, "
                MySQLStr = MySQLStr & "ContactInfo, DeliveryAddress, Comment, PrintBillOrNot, PrintBillOrNot1, PrintFullBillOrNot, RequestedShipmentDate, IsRequestSend, "
                MySQLStr = MySQLStr & "IsReminderSend, AttFile, CommentTr, CommentDoc) "
                MySQLStr = MySQLStr & "VALUES (N'" & Trim(LblCustomerCode.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(LblCustomerName.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(LblCustomerINN.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(LblCustomerLegalAddress.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(LblSalesmanCode.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(LblWHCode.Text) & "', "
                If ComboBox1.Text = "��������" Then
                    MySQLStr = MySQLStr & "1, "
                ElseIf ComboBox1.Text = "�������� � ������� ��������" Then
                    MySQLStr = MySQLStr & "2, "
                Else
                    MySQLStr = MySQLStr & "0, "
                End If
                MySQLStr = MySQLStr & "0, "
                MySQLStr = MySQLStr & "0, "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox1.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox2.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox3.Text) & "', "
                If CheckBox1.Checked = True Then
                    MySQLStr = MySQLStr & "1, "
                Else
                    MySQLStr = MySQLStr & "0, "
                End If
                If CheckBox2.Checked = True Then
                    MySQLStr = MySQLStr & "1, "
                Else
                    MySQLStr = MySQLStr & "0, "
                End If
                If CheckBox3.Checked = True Then
                    MySQLStr = MySQLStr & "1, "
                Else
                    MySQLStr = MySQLStr & "0, "
                End If
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Month), 2) & "/" & DateTimePicker1.Value.Year & "', 103), "
                MySQLStr = MySQLStr & "0, "
                MySQLStr = MySQLStr & "0, "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox4.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox5.Text) & "', "
                MySQLStr = MySQLStr & "N'" & Trim(TextBox6.Text) & "' "
                MySQLStr = MySQLStr & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else    '------��������������
                MySQLStr = "UPDATE tbl_Shipments_SalesmanWP_Info "
                If ComboBox1.Text = "��������" Then
                    MySQLStr = MySQLStr & "SET DeliveryOrNot = 1, "
                ElseIf ComboBox1.Text = "�������� � ������� ��������" Then
                    MySQLStr = MySQLStr & "SET DeliveryOrNot = 2, "
                Else
                    MySQLStr = MySQLStr & "SET DeliveryOrNot = 0, "
                End If
                MySQLStr = MySQLStr & "ContactInfo = N'" & Trim(TextBox1.Text) & "', "
                MySQLStr = MySQLStr & "DeliveryAddress = N'" & Trim(TextBox2.Text) & "', "
                MySQLStr = MySQLStr & "Comment = N'" & Trim(TextBox3.Text) & "', "
                MySQLStr = MySQLStr & "CommentTr = N'" & Trim(TextBox5.Text) & "', "
                MySQLStr = MySQLStr & "CommentDoc = N'" & Trim(TextBox6.Text) & "', "
                If CheckBox1.Checked = True Then
                    MySQLStr = MySQLStr & "PrintBillOrNot = 1, "
                Else
                    MySQLStr = MySQLStr & "PrintBillOrNot = 0, "
                End If
                If CheckBox2.Checked = True Then
                    MySQLStr = MySQLStr & "PrintBillOrNot1 = 1, "
                Else
                    MySQLStr = MySQLStr & "PrintBillOrNot1 = 0, "
                End If
                If CheckBox3.Checked = True Then
                    MySQLStr = MySQLStr & "PrintFullBillOrNot = 1, "
                Else
                    MySQLStr = MySQLStr & "PrintFullBillOrNot = 0, "
                End If
                MySQLStr = MySQLStr & "RequestedShipmentDate = CONVERT(DATETIME, '" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Day), 2) & "/" & Microsoft.VisualBasic.Right("00" & CStr(DateTimePicker1.Value.Month), 2) & "/" & DateTimePicker1.Value.Year & "', 103), "
                MySQLStr = MySQLStr & "AttFile = N'" & Trim(TextBox4.Text) & "' "
                MySQLStr = MySQLStr & "WHERE (ID = " & MyShipmentsList.DataGridView1.SelectedRows.Item(0).Cells(0).Value & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If
            '------
            MyOperationFlag = 1
            Me.Close()
        End If
    End Sub

    Private Function CheckFiields() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim DelDate As Date
        Dim CurrDate As Date

        If Trim(TextBox1.Text) = "" Then
            MsgBox("���� ""���������� ����������"" ������ ���� ���������. ", MsgBoxStyle.Critical, "��������!")
            TextBox1.Select()
            CheckFiields = False
            Exit Function
        End If

        If ComboBox1.Text = "��������" Or ComboBox1.Text = "�������� � ������� ��������" Then
            If Trim(TextBox2.Text) = "" Then
                MsgBox("���� ""����� ��������"" ������ ���� ���������. ", MsgBoxStyle.Critical, "��������!")
                TextBox2.Select()
                CheckFiields = False
                Exit Function
            End If
        End If

        DelDate = DateTimePicker1.Value.Date
        CurrDate = Now().Date
        If DateTime.Compare(DelDate, CurrDate) < 0 Then
            MsgBox("���� �������� �� ������ �� ������ ���� ������ ������� ����. ", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckFiields = False
            Exit Function
        End If

        CheckFiields = True
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ��������� �� CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MyContactInfo = New ContactInfo
        MyContactInfo.StartParam = "Contact"
        MyContactInfo.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���� ������� �� Scala
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim MyDelAddresses = New DelAddresses
        MyDelAddresses.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���� � �����, ������� ����� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDlg As OpenFileDialog

        MyDlg = New OpenFileDialog
        MyDlg.Filter = "��� ����� (*.*)|*.*"
        MyDlg.Multiselect = False
        MyDlg.SupportMultiDottedExtensions = True
        MyDlg.Title = "�������� ����"
        MyDlg.InitialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        MyDlg.RestoreDirectory = True
        If MyDlg.ShowDialog() <> DialogResult.OK Then
        Else
            TextBox4.Text = MyDlg.FileName
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox4.Text = ""
    End Sub
End Class