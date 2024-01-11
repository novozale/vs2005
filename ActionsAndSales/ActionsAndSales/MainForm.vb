Public Class MainForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ��������� �������� ����� / ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button1.Enabled = False
        Button2.Enabled = False
        System.Windows.Forms.Application.DoEvents()
        If My.Settings.UseOffice = "LibreOffice" Then
            ImportDataFromExcel_LO()
        Else
            ImportDataFromExcel()
        End If
        Button1.Enabled = True
        Button2.Enabled = True
        System.Windows.Forms.Application.DoEvents()
        MsgBox("��������� �������� ���������� �� ����� / ���������� ���������.", vbOKOnly, "��������!")
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ���������� �� ����� / ���������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyTableName As String                   '��� ��������� �������
        Dim MyGuid As String                          '
        Dim connStr As String                       '������ ���������� � Excel
        Dim MySQLStr As String                      'SQL ������
        Dim cn As OleDbConnection                   '������ ���������� � OLE
        Dim FirstExcelSheetName As String           '�������� ������� ����� Excel
        Dim myds As DataSet                         'Excel dataset
        Dim MyVersion As String                     '������ ���������
        Dim mycount As Integer
        Dim MyDBL As Double                         '��� ��������
        Dim MyInt As Integer                        '��� ��������
        Dim MyStr As String                         '��� ��������
        Dim MyDatetimeStart As Date                 '��� ��������
        Dim MyDatetimeFin As Date                   '��� ��������
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyErrStr As String
        Dim MyContinueFlag As Integer
        Dim MyRez As MsgBoxResult                   '��������� ������
        Dim MyActOrSalesFlag As String

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_ActionsAndSales_Tmp_" + MyGuid

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                '----------------------------������� ��������
                Label3.Text = "���������� �������� Excel �����"
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                connStr = "provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" & OpenFileDialog1.FileName & ";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;"""
                Try
                    cn = New OleDbConnection(connStr)
                    FirstExcelSheetName = GetFirstExcelSheetName(cn)

                    '============================��������============================================================================
                    '---��������� ������ ����� Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A1:A1]"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds Is Nothing = False Then
                        If IsDBNull(myds.Tables(0).Rows(0).Item(0)) Then
                            MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Excel ", MsgBoxStyle.Critical, "��������!")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            MyVersion = Trim(myds.Tables(0).Rows(0).Item(0))
                            MySQLStr = "SELECT Version "
                            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                            MySQLStr = MySQLStr & "WHERE (Name = N'�������� ����� ��� ����������') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                                MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            Else
                                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                    trycloseMyRec()
                                Else
                                    MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                                    trycloseMyRec()
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                            End If
                        End If
                    Else
                        MsgBox("���������� ��������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '---��������� - �������� �� ���������� ��� ��� - ����� ��� ����������
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "C2:C2]"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds Is Nothing = False Then
                        If IsDBNull(myds.Tables(0).Rows(0).Item(0)) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C2"" �� ����������� ��� ��� - ����� ��� ��������. ", MsgBoxStyle.Critical, "��������!")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            MyActOrSalesFlag = Trim(myds.Tables(0).Rows(0).Item(0))
                            If (MyActOrSalesFlag.Equals("�����") = False And MyActOrSalesFlag.Equals("��������") = False) Then
                                MsgBox("� ������������� ����� Excel � ������ ""C2"" ������ ���� ����������� - ""�����"" ��� ""��������"". ", MsgBoxStyle.Critical, "��������!")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        End If
                    Else
                            MsgBox("���������� ��������� ��� ��� - ����� ��� ����������. ���������� � ��������������.", vbCritical, "��������!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                    End If


                    '---��������� ������������ ������ � Excel
                    '-----������������� ����
                    MySQLStr = "SELECT F1 FROM [" & FirstExcelSheetName & "A5:A] group by F1 having(count(F1) > 1)"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds.Tables(0).Rows.Count > 0 Then
                        MsgBox("� ����� ��������� " & myds.Tables(0).Rows.Count & " ������������� ������� ����� ������� � Scala. �������������� ������� ""���������� �������������"" � Excel, ��������� � ������� ������ ���� ")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----������������ ��������� ������ � Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A5:J] where(F1 <> """")"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    '-----������������ ���������
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        '-----���������� ���� ������ Scala
                        If Trim(myds.Tables(0).Rows(mycount).Item(0).ToString) = "" Then
                            MsgBox("������ " & CStr(mycount + 5) & " �� ������� ��� ������ � Scala")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        '-----���������� ���� ��������������
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(1)
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���������� ���� ��������������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----������ �������
                        Try
                            MyInt = myds.Tables(0).Rows(mycount).Item(2)
                            If (MyInt <> 0 And MyInt <> 1 And MyInt <> 4 And MyInt <> 12) Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ������ �������")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ������ �������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----����������� ��������������
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(3)
                            If (MyDBL < 1 Or MyDBL > 2) Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ����������� �������������� - ������ ���� � ���������� �� 1 �� 2")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ����������� ��������������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----������� - ����� �� ����������
                        Try
                            MyStr = myds.Tables(0).Rows(mycount).Item(4).ToString
                            If MyStr.ToUpper.Equals("��") = False And MyStr.ToUpper.Equals("���") = False Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ������� - ����� �� ����������: ������ ���� �� ��� ���")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ������� - ����� �� ����������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----����������, ����������� �� �����
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(5))) Then
                        Else
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(5)
                                If (MyDBL < 0) Then
                                    MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ����������, ����������� �� ����� - ������ ���� ����� ��� ������ ��� ����� 0")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ����������, ����������� �� ����� - ������ ���� ����� ��� ������ ��� ����� 0")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If
                        '-----������� - ����� �� �����
                        Try
                            MyStr = myds.Tables(0).Rows(mycount).Item(6).ToString
                            If MyStr.ToUpper.Equals("��") = False And MyStr.ToUpper.Equals("���") = False Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ������� - ����� �� �����: ������ ���� �� ��� ���")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ������� - ����� �� �����")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----���� ������ �����
                        Try
                            MyDatetimeStart = myds.Tables(0).Rows(mycount).Item(7)
                            If MyDatetimeStart < Today Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ������ �����: ���� ������ ����� �� ����� ���� ������ ������� ����")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ������ �����")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----���� ��������� �����
                        MyStr = myds.Tables(0).Rows(mycount).Item(8).ToString
                        If MyStr.Equals("") Then
                            If (myds.Tables(0).Rows(mycount).Item(6).ToString.ToUpper.Equals("��")) Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ��������� �����: ��� ��� ����� �� ����� - ���� ���������� ����� ������ ���� ������� �����������.")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End If
                        Else
                            If (myds.Tables(0).Rows(mycount).Item(6).ToString.ToUpper.Equals("���")) Then
                                MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ��������� �����: ��� ��� ����� �� �� ����� - ���� ���������� ����� ������ ���� ������.")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            Else
                                Try
                                    MyDatetimeFin = myds.Tables(0).Rows(mycount).Item(8)
                                    If MyDatetimeFin < MyDatetimeStart Then
                                        MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ��������� �����: ���� ��������� ����� �� ����� ���� ������ ���� ������ �����")
                                        Me.Cursor = Cursors.Default
                                        Exit Sub
                                    End If
                                Catch ex As Exception
                                    MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ���� ��������� �����")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End Try
                            End If
                        End If
                        '-----�������� - ����������� ������ ���� ����� ��� �� ���������� ��� �� �����
                        If myds.Tables(0).Rows(mycount).Item(4).ToString().ToUpper.Equals("���") And myds.Tables(0).Rows(mycount).Item(6).ToString().ToUpper.Equals("���") Then
                            MsgBox("������ " & CStr(mycount + 5) & " ����� ������ ���� ����������� ��� �� ����� ��� �� ���������� ��� �� ����� � ����������.")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If

                        '-----�������� � ������� �����
                        MyStr = myds.Tables(0).Rows(mycount).Item(9).ToString
                        If (MyStr.Equals("")) Then
                            MsgBox("������ " & CStr(mycount + 5) & " �� ������� �������� � ������� �����")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        mycount = mycount + 1
                    End While

                    '========================================�������� �������� �� ��������� �������=================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������ �� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----�������� ��������� �������
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[ScalaCode] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePrice] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePriceCurr] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MarginCoeff] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[QTYAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionStopQTY] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[TimeAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[DateStart] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[DateFinish] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionName] [nvarchar](4000) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionOrSales] [nvarchar](50) NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ �� ��������� �������
                    InitMyConn(False)
                    MySQLStr = "SELECT  ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    Try
                        MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                        MySQLAdapter.SelectCommand.CommandTimeout = 1200
                        Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                        MySQLAdapter.Fill(MySQLDs)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try
                    '-----������� ������ �� Excel dataset � SQL dataset
                    Dim dt As DataTable
                    Dim dr As DataRow

                    dt = MySQLDs.Tables(0)
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        dr = dt.NewRow
                        dr.Item(0) = myds.Tables(0).Rows(mycount).Item(0)
                        dr.Item(1) = myds.Tables(0).Rows(mycount).Item(1)
                        dr.Item(2) = myds.Tables(0).Rows(mycount).Item(2)
                        dr.Item(3) = myds.Tables(0).Rows(mycount).Item(3)
                        dr.Item(4) = myds.Tables(0).Rows(mycount).Item(4)
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(5))) Then
                            dr.Item(5) = 999999999
                        Else
                            If (myds.Tables(0).Rows(mycount).Item(5) = 0) Then
                                dr.Item(5) = 999999999
                            Else
                                dr.Item(5) = myds.Tables(0).Rows(mycount).Item(5)
                            End If
                        End If
                        dr.Item(6) = myds.Tables(0).Rows(mycount).Item(6)
                        dr.Item(7) = myds.Tables(0).Rows(mycount).Item(7)
                        If (IsDBNull(myds.Tables(0).Rows(mycount).Item(8))) Then
                            dr.Item(8) = New DateTime(9999, 12, 31, 0, 0, 0)
                        Else
                            dr.Item(8) = myds.Tables(0).Rows(mycount).Item(8)
                        End If
                        dr.Item(9) = myds.Tables(0).Rows(mycount).Item(9)
                        dr.Item(10) = MyActOrSalesFlag
                        dt.Rows.Add(dr)
                        mycount = mycount + 1
                    End While
                    Try
                        MySQLAdapter.Update(MySQLDs, "Table")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '========================================���������� �������� �� �������=================================
                    '==============================�������� ������� � �� ��������� �����====================================
                    Label3.Text = "�������� ������� � �� ��������� �����"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "Select " & MyTableName & ".ScalaCode "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "SC010300 ON " & MyTableName & ".ScalaCode = SC010300.SC01001 "
                    MySQLStr = MySQLStr & "WHERE(SC010300.SC01001 Is NULL) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "��������� ���� ������� ����������� � Scala:" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Declarations.MyRec.Fields("ScalaCode").Value & Chr(13) & Chr(10)
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� ���� ������� � Excel �����.")
                        End If
                    End If

                    '==============================�������� ������� � �� ����� � ����� �� ���������=============================
                    Label3.Text = "�������� ������� � �� ����� � ����� �� ���������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MyContinueFlag = 1
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, CONVERT(nvarchar(30), MIN(tbl_ActionsAndSales.DateStart), 103) AS DateStart, "
                    MySQLStr = MySQLStr & "CONVERT(nvarchar(30), MAX(tbl_ActionsAndSales.DateFinish), 103) AS DateFinish, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN '��� ����������' ELSE '��� �� ����������' END AS MyState "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN tbl_ActionsAndSales ON "
                    MySQLStr = MySQLStr & " " & MyTableName & ".ActionName = tbl_ActionsAndSales.ActionName "
                    MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales.ActionName, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN '��� ����������' ELSE '��� �� ����������' END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "������ ����� / ���������� ��� ������������ � Scala:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) & Declarations.MyRec.Fields("MyState").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) + "�" + Chr(9) & Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value & Chr(13) & Chr(10)
                            If (Declarations.MyRec.Fields("MyState").Value.ToString.Equals("��� ����������")) Then
                                MyContinueFlag = 0
                            End If
                            Declarations.MyRec.MoveNext()
                        End While
                        If (MyContinueFlag = 0) Then
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� �� ���������� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        Else
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ��������� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "���� ����� / ���������� ����� ������� � �������� ������� �� Excel. " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "���� �� �� ������ ������� ��� ������������ ����� / ����������, �� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "�������� ""�����"", " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "��������� �������� ����� / ���������� � Excel ����� �� ����������." & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        End If
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        If (MyContinueFlag = 0) Then
                            MyErrForm.Button2.Visible = False
                            MyErrForm.Button2.Enabled = False
                        Else
                            MyErrForm.Button2.Visible = True
                            MyErrForm.Button2.Enabled = True
                        End If
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("����� / ���������� ��� ����������.")
                        Else    '----�������� ����� ���������� �����
                            MySQLStr = "DELETE FROM tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & ".ActionName = tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales_1.ActionName) AS View_2 ON tbl_ActionsAndSales.ActionName = View_2.ActionName "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '==============================�������� (����� ���) ����� � ������ ����������=============================
                    Label3.Text = "�������� (����� ���) ����� � ������ ����������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '----------------------------�������� ��� ����� / ���������� ������ ��� �������������---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart >= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish <= "
                    MySQLStr = MySQLStr & "CASE WHEN tbl_ActionsAndSales.DateFinish = CONVERT(datetime, '31/12/9999', 103) THEN "
                    MySQLStr = MySQLStr & "dateadd(dd, - 1, tbl_ActionsAndSales.DateFinish) ELSE tbl_ActionsAndSales.DateFinish END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� ��������� ������ ��� ������������� ���������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        End If
                    End If

                    '---------------�������� ��� ����� / ���������� ���������� � ����� ������ ������������--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart <= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish >= tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� ���������� � ����� ������ ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        End If
                    End If

                    '------------------�������� ��� ����� / ���������� ����� ����������� ��� ������������---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� �� ���� ������ ������������� � ��� ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ���� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���������� ����� / ���������� ����� �������� �� ���� ������ ������� ����� - 1 ����. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "���� �� �� ������ ������ ����� ������� ���� ��������� ���������� ����� / ����������, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "�� �������� ""�����"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���� ������ ����� / ���������� � Excel ����� �� �����������" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        Else    '-----������ ���� ��������� �����
                            'MySQLStr = "UPDATE " & MyTableName & " "
                            'MySQLStr = MySQLStr & "SET DateStart = DateAdd(dd, 1, CASE WHEN View_2.DateFinish = CONVERT(datetime, '31/12/9999', 103) "
                            'MySQLStr = MySQLStr & "THEN dateadd(dd, - 1, View_2.DateFinish) ELSE View_2.DateFinish END)) "
                            'MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            'MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            'MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales.DateStart AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            'MySQLStr = MySQLStr & " " & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            MySQLStr = "Update tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart), "
                            MySQLStr = MySQLStr & "TimeAction = N'��' "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName, tbl_ActionsAndSales_1.ScalaCode, " & MyTableName & "_1.DateStart, "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales_1.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales_1.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales_1.DateFinish) AS View_2 ON "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales.ScalaCode = View_2.ScalaCode And tbl_ActionsAndSales.DateStart < View_2.DateStart And tbl_ActionsAndSales.DateFinish > View_2.DateStart "

                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '------------------�������� ��� ����� / ���������� ������ ����������� ��� ������������--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� �� ���� ��������� ������������� � ��� ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ���� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ����� / ���������� ����� �������� �� ���� ������ �����������  ����� ����� 1 ����. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "���� �� �� ������ ������ ����� ������� ���� ��������� ����� / ����������, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "�� �������� ""�����"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���� ��������� ����� / ���������� � Excel ����� �� �����������" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        Else    '-----������ ���� ������ �����
                            MySQLStr = "UPDATE " & MyTableName & " "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart) "
                            MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish > tbl_ActionsAndSales.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            MySQLStr = MySQLStr & "" & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If



                    '==============================��������� ����� / ���������� � ��=============================
                    Label3.Text = "��������� ����� / ���������� � ��"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "INSERT INTO tbl_ActionsAndSales "
                    MySQLStr = MySQLStr & "(ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, ActionFinished, ActionFinishedDate) "
                    MySQLStr = MySQLStr & "SELECT ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, 0 AS ActionFinished, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, '01/01/1900', 103) AS ActionFinishedDate "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '==============================================������ ��������� ����� ����� �� �������============================================
                    MyRez = MsgBox("���������� ������ ����� - ����� �� ������� ������? ����� ������ ����� �������� �����.", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = MsgBoxResult.Yes Then
                        '----------------------------������� ��������
                        Label3.Text = "������������ ����� ����� �� �������. "
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()

                        MySQLStr = "Exec spp_PrepareCommonPriceList_PriCost "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If


                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                Finally
                    cn.Close()
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                    '----------------------------������� ��������
                    Label3.Text = ""
                End Try
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub ImportDataFromExcel_LO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ���������� �� ����� / ���������� ��� ������ LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyTableName As String                   '��� ��������� �������
        Dim MyGuid As String                          '
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                     '������ ���������
        Dim MyActOrSalesFlag As String              '��� ��� - ����� ��� ����������
        Dim MyDBL As Double                         '��� ��������
        Dim MyInt As Integer                        '��� ��������
        Dim MyStr As String                         '��� ��������
        Dim MyDatetimeStart As Date                 '��� ��������
        Dim MyDatetimeFin As Date                   '��� ��������
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyErrStr As String
        Dim MyContinueFlag As Integer
        Dim MyRez As MsgBoxResult                   '��������� ������


        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_ActionsAndSales_Tmp_" + MyGuid

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                '----------------------------������� ��������
                Label3.Text = "���������� �������� Excel �����"
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                Try
                    LOSetNotation(0)

                    oServiceManager = CreateObject("com.sun.star.ServiceManager")
                    oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                    oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                    oFileName = Replace(OpenFileDialog2.FileName, "\", "/")
                    oFileName = "file:///" + oFileName
                    Dim arg(1)
                    arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                    arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                    oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                    oSheet = oWorkBook.getSheets().getByIndex(0)
                    '============================��������============================================================================
                    '---��������� ������ ����� Excel
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    If MyVersion = "" Then
                        MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Excel ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        MySQLStr = "SELECT Version "
                        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                        MySQLStr = MySQLStr & "WHERE (Name = N'�������� ����� ��� ����������') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                trycloseMyRec()
                            Else
                                MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    End If

                    '---��������� - �������� �� ���������� ��� ��� - ����� ��� ����������
                    MyActOrSalesFlag = oSheet.getCellRangeByName("C2").String
                    If MyActOrSalesFlag.Equals("") Then
                        MsgBox("� ������������� ����� Excel � ������ ""C2"" �� ����������� ��� ��� - ����� ��� ��������. ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        If (MyActOrSalesFlag.Equals("�����") = False And MyActOrSalesFlag.Equals("��������") = False) Then
                            MsgBox("� ������������� ����� Excel � ������ ""C2"" ������ ���� ����������� - ""�����"" ��� ""��������"". ", MsgBoxStyle.Critical, "��������!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    End If

                    '---��������� ������������ ������ � Excel
                    '-----������������� ����
                    oSheet.unprotect("!pass2022")

                    Dim args() As Object
                    ReDim args(0)
                    args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args(0).Name = "ToPoint"
                    args(0).Value = "$A$5:$K$100000"
                    Dim oFrame As Object
                    oFrame = oWorkBook.getCurrentController.getFrame
                    oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
                    Dim args1() As Object
                    ReDim args1(6)
                    args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(0).Name = "ByRows"
                    args1(0).Value = True
                    args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(1).Name = "HasHeader"
                    args1(1).Value = False
                    args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(2).Name = "CaseSensitive"
                    args1(2).Value = False
                    args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(3).Name = "NaturalSort"
                    args1(3).Value = False
                    args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(4).Name = "IncludeAttribs"
                    args1(4).Value = True
                    args1(5) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(5).Name = "Col1"
                    args1(5).Value = 1
                    args1(6) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args1(6).Name = "Ascending1"
                    args1(6).Value = True
                    oDispatcher.executeDispatch(oFrame, ".uno:DataSort", "", 0, args1)

                    'Dim oSortFields(0) As Object
                    'Dim oSortDesc(0) As Object
                    'Dim srange = oSheet.getCellRangeByName("A5:K100000")
                    ' ''oSortFields(0) = oServiceManager.Bridge_GetStruct("com.sun.star.table.TableSortField")
                    ' ''oSortFields(0).Field = 0
                    ' ''oSortFields(0).IsAscending = False
                    ''oSortFields(0) = oServiceManager.Bridge_GetStruct("com.sun.star.util.SortField")
                    ''oSortFields(0).Field = 0
                    ''oSortFields(0).SortAscending = False
                    ' ''oSortDesc = srange.createSortDescriptor
                    ' ''oSortDesc(1).Value = False
                    ' ''oSortDesc(3).Value = oSortFields
                    ''oSortDesc(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    ''oSortDesc(0).Name = "SortFields"
                    ''oSortDesc(0).Value = oSortFields

                    'Dim oReflection As Object
                    ''Dim unoWrap As Object
                    'oReflection = oServiceManager.createInstance("com.sun.star.reflection.CoreReflection")
                    ''Dim sortFields(0) As Object
                    ''Dim sortProperties(0) As Object
                    'oReflection.forName("com.sun.star.table.TableSortField").CreateObject(oSortFields(0))
                    'oSortFields(0).Field = 0
                    'oSortFields(0).IsAscending = False
                    ''unoWrap = oServiceManager.Bridge_GetValueObject
                    ''unoWrap.set("[]com.sun.star.table.TableSortField", oSortFields)
                    'oReflection.forName("com.sun.star.beans.PropertyValue").createObject(oSortDesc(0))
                    'oSortDesc(0).Name = "SortFields"
                    'oSortDesc(0).Value = oSortFields

                    'srange.Sort(oSortDesc)


                    Dim srange = oSheet.getCellRangeByName("A5:K100000")
                    Dim myarr = srange.DataArray
                    For i As Integer = 1 To 99995
                        If myarr(i)(0) = myarr(i - 1)(0) And myarr(i)(0) <> "" Then
                            MsgBox("� ����� ��������� ������������� ������ ����� ������� � Scala. �������������� ������� ""�������� ������������� ������"" � LibreOffice, ��������� � ������� ������ ���� ")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        End If
                    Next i
                    Dim args2() As Object
                    ReDim args2(0)
                    oDispatcher.executeDispatch(oFrame, ".uno:Save", "", 0, args2)

                    '-----������������ ��������� ������ � Excel
                    For i As Integer = 0 To 99995
                        If myarr(i)(0).ToString = "" Then
                            Exit For
                        Else
                            '-----���������� ���� ��������������
                            Try
                                MyDBL = myarr(i)(1)
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���������� ���� ��������������")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----������ �������
                            Try
                                MyInt = myarr(i)(2)
                                If (MyInt <> 0 And MyInt <> 1 And MyInt <> 4 And MyInt <> 12) Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ������ �������")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� �������� ������ �������")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----����������� ��������������
                            Try
                                MyDBL = myarr(i)(3)
                                If (MyDBL < 1 Or MyDBL > 2) Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ����������� �������������� - ������ ���� � ���������� �� 1 �� 2")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� ������� ����������� ��������������")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----������� - ����� �� ����������
                            Try
                                MyStr = myarr(i)(4)
                                If MyStr.ToUpper.Equals("��") = False And MyStr.ToUpper.Equals("���") = False Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ������� - ����� �� ����������: ������ ���� �� ��� ���")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� ������� ������� - ����� �� ����������")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----����������, ����������� �� �����
                            Try
                                MyDBL = myarr(i)(5)
                                If (MyDBL < 0) Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ����������, ����������� �� ����� - ������ ���� ����� ��� ������ ��� ����� 0")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� �������� ����������, ����������� �� ����� - ������ ���� ����� ��� ������ ��� ����� 0")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----������� - ����� �� �����
                            Try
                                MyStr = myarr(i)(6)
                                If MyStr.ToUpper.Equals("��") = False And MyStr.ToUpper.Equals("���") = False Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ������� - ����� �� �����: ������ ���� �� ��� ���")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� ������� ������� - ����� �� �����")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----���� ������ �����
                            Try
                                MyDatetimeStart = DateTime.FromOADate(myarr(i)(7))
                                If MyDatetimeStart < Today Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ������ �����: ���� ������ ����� �� ����� ���� ������ ������� ����")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Catch ex As Exception
                                MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ������ �����")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End Try
                            '-----���� ��������� �����
                            MyStr = myarr(i)(8)
                            If MyStr.Equals("") Then
                                If (myarr(i)(6).ToString.ToUpper.Equals("��")) Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ��������� �����: ��� ��� ����� �� ����� - ���� ���������� ����� ������ ���� ������� �����������.")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            Else
                                If (myarr(i)(6).ToString.ToUpper.Equals("���")) Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ��������� �����: ��� ��� ����� �� �� ����� - ���� ���������� ����� ������ ���� ������.")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                Else
                                    Try
                                        MyDatetimeFin = DateTime.FromOADate(myarr(i)(8))
                                        If MyDatetimeFin < MyDatetimeStart Then
                                            MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ��������� �����: ���� ��������� ����� �� ����� ���� ������ ���� ������ �����")
                                            Me.Cursor = Cursors.Default
                                            oWorkBook.Close(True)
                                            Exit Sub
                                        End If
                                    Catch ex As Exception
                                        MsgBox("������ " & CStr(i + 5) & " ����������� �������� ���� ��������� �����")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If
                            End If
                            '-----�������� - ����������� ������ ���� ����� ��� �� ���������� ��� �� �����
                            If myarr(i)(4).ToString().ToUpper.Equals("���") And myarr(i)(6).ToString().ToUpper.Equals("���") Then
                                MsgBox("������ " & CStr(i + 5) & " ����� ������ ���� ����������� ��� �� ����� ��� �� ���������� ��� �� ����� � ����������.")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                            '-----�������� � ������� �����
                            MyStr = myarr(i)(9)
                            If (MyStr.Equals("")) Then
                                MsgBox("������ " & CStr(i + 5) & " �� ������� �������� � ������� �����")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    Next i

                    '========================================�������� �������� �� ��������� �������=================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������ �� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----�������� ��������� �������
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[ScalaCode] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePrice] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[PurchasePriceCurr] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MarginCoeff] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[QTYAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionStopQTY] [numeric](28, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[TimeAction] [nvarchar](10) NOT NULL, "
                    MySQLStr = MySQLStr & "[DateStart] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[DateFinish] [datetime] NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionName] [nvarchar](4000) NOT NULL, "
                    MySQLStr = MySQLStr & "[ActionOrSales] [nvarchar](50) NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ �� ��������� �������
                    InitMyConn(False)
                    MySQLStr = "SELECT  ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    Try
                        MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                        MySQLAdapter.SelectCommand.CommandTimeout = 1200
                        Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
                        MySQLAdapter.Fill(MySQLDs)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '-----������� ������ �� Excel dataset � SQL dataset
                    Dim dt As DataTable
                    Dim dr As DataRow

                    dt = MySQLDs.Tables(0)
                    For i As Integer = 0 To 99995
                        If myarr(i)(0).ToString = "" Then
                            Exit For
                        Else
                            dr = dt.NewRow
                            dr.Item(0) = myarr(i)(0).ToString
                            dr.Item(1) = myarr(i)(1)
                            dr.Item(2) = myarr(i)(2)
                            dr.Item(3) = myarr(i)(3)
                            dr.Item(4) = myarr(i)(4)
                            If (myarr(i)(5) = 0) Then
                                dr.Item(5) = 999999999
                            Else
                                dr.Item(5) = myarr(i)(5)
                            End If
                            dr.Item(6) = myarr(i)(6)
                            dr.Item(7) = DateTime.FromOADate(myarr(i)(7))
                            If (myarr(i)(8).ToString().Equals("")) Then
                                dr.Item(8) = New DateTime(9999, 12, 31, 0, 0, 0)
                            Else
                                dr.Item(8) = DateTime.FromOADate(myarr(i)(8))
                            End If
                            dr.Item(9) = myarr(i)(9)
                            dr.Item(10) = MyActOrSalesFlag
                            dt.Rows.Add(dr)
                        End If
                    Next i
                    Try
                        MySQLAdapter.Update(MySQLDs, "Table")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '========================================���������� �������� �� �������=================================
                    '==============================�������� ������� � �� ��������� �����====================================
                    Label3.Text = "�������� ������� � �� ��������� �����"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "Select " & MyTableName & ".ScalaCode "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "SC010300 ON " & MyTableName & ".ScalaCode = SC010300.SC01001 "
                    MySQLStr = MySQLStr & "WHERE(SC010300.SC01001 Is NULL) "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "��������� ���� ������� ����������� � Scala:" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Declarations.MyRec.Fields("ScalaCode").Value & Chr(13) & Chr(10)
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� ���� ������� � Excel �����.")
                        End If
                    End If

                    '==============================�������� ������� � �� ����� � ����� �� ���������=============================
                    Label3.Text = "�������� ������� � �� ����� � ����� �� ���������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MyContinueFlag = 1
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, CONVERT(nvarchar(30), MIN(tbl_ActionsAndSales.DateStart), 103) AS DateStart, "
                    MySQLStr = MySQLStr & "CONVERT(nvarchar(30), MAX(tbl_ActionsAndSales.DateFinish), 103) AS DateFinish, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN '��� ����������' ELSE '��� �� ����������' END AS MyState "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN tbl_ActionsAndSales ON "
                    MySQLStr = MySQLStr & " " & MyTableName & ".ActionName = tbl_ActionsAndSales.ActionName "
                    MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales.ActionName, CASE WHEN tbl_ActionsAndSales.DateStart <= dateadd(day, "
                    MySQLStr = MySQLStr & "datediff(day, 0, GETDATE()), 0) THEN '��� ����������' ELSE '��� �� ����������' END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "������ ����� / ���������� ��� ������������ � Scala:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) & Declarations.MyRec.Fields("MyState").Value & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + Chr(9) + "�" + Chr(9) & Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value & Chr(13) & Chr(10)
                            If (Declarations.MyRec.Fields("MyState").Value.ToString.Equals("��� ����������")) Then
                                MyContinueFlag = 0
                            End If
                            Declarations.MyRec.MoveNext()
                        End While
                        If (MyContinueFlag = 0) Then
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� �� ���������� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        Else
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ��������� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "���� ����� / ���������� ����� ������� � �������� ������� �� Excel. " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "���� �� �� ������ ������� ��� ������������ ����� / ����������, �� " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "�������� ""�����"", " & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "��������� �������� ����� / ���������� � Excel ����� �� ����������." & Chr(13) & Chr(10)
                            MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        End If
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        If (MyContinueFlag = 0) Then
                            MyErrForm.Button2.Visible = False
                            MyErrForm.Button2.Enabled = False
                        Else
                            MyErrForm.Button2.Visible = True
                            MyErrForm.Button2.Enabled = True
                        End If
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("����� / ���������� ��� ����������.")
                        Else    '----�������� ����� ���������� �����
                            MySQLStr = "DELETE FROM tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & ".ActionName = tbl_ActionsAndSales_1.ActionName "
                            MySQLStr = MySQLStr & "GROUP BY tbl_ActionsAndSales_1.ActionName) AS View_2 ON tbl_ActionsAndSales.ActionName = View_2.ActionName "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '==============================�������� (����� ���) ����� � ������ ����������=============================
                    Label3.Text = "�������� (����� ���) ����� � ������ ����������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '----------------------------�������� ��� ����� / ���������� ������ ��� �������������---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart >= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish <= "
                    MySQLStr = MySQLStr & "CASE WHEN tbl_ActionsAndSales.DateFinish = CONVERT(datetime, '31/12/9999', 103) THEN "
                    MySQLStr = MySQLStr & "dateadd(dd, - 1, tbl_ActionsAndSales.DateFinish) ELSE tbl_ActionsAndSales.DateFinish END "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� ��������� ������ ��� ������������� ���������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        End If
                    End If

                    '---------------�������� ��� ����� / ���������� ���������� � ����� ������ ������������--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart <= tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish >= tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� ���������� � ����� ������ ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "��������� �������� ����� / ���������� � Excel ����� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.Button2.Visible = False
                        MyErrForm.Button2.Enabled = False
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        End If
                    End If

                    '------------------�������� ��� ����� / ���������� ����� ����������� ��� ������������---------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateStart < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� �� ���� ������ ������������� � ��� ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ���� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���������� ����� / ���������� ����� �������� �� ���� ������ ������� ����� - 1 ����. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "���� �� �� ������ ������ ����� ������� ���� ��������� ���������� ����� / ����������, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "�� �������� ""�����"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���� ������ ����� / ���������� � Excel ����� �� �����������" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        Else    '-----������ ���� ��������� �����
                            'MySQLStr = "UPDATE " & MyTableName & " "
                            'MySQLStr = MySQLStr & "SET DateStart = DateAdd(dd, 1, CASE WHEN View_2.DateFinish = CONVERT(datetime, '31/12/9999', 103) "
                            'MySQLStr = MySQLStr & "THEN dateadd(dd, - 1, View_2.DateFinish) ELSE View_2.DateFinish END)) "
                            'MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            'MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            'MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales.DateStart AND "
                            'MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            'MySQLStr = MySQLStr & " " & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            MySQLStr = "Update tbl_ActionsAndSales "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart), "
                            MySQLStr = MySQLStr & "TimeAction = N'��' "
                            MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales INNER JOIN "
                            MySQLStr = MySQLStr & "(SELECT tbl_ActionsAndSales_1.ActionName, tbl_ActionsAndSales_1.ScalaCode, " & MyTableName & "_1.DateStart, "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales AS tbl_ActionsAndSales_1 ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales_1.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart > tbl_ActionsAndSales_1.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateStart < tbl_ActionsAndSales_1.DateFinish) AS View_2 ON "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales.ScalaCode = View_2.ScalaCode And tbl_ActionsAndSales.DateStart < View_2.DateStart And tbl_ActionsAndSales.DateFinish > View_2.DateStart "

                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If

                    '------------------�������� ��� ����� / ���������� ������ ����������� ��� ������������--------
                    MySQLStr = "SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & ".ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish > tbl_ActionsAndSales.DateStart AND "
                    MySQLStr = MySQLStr & " " & MyTableName & ".DateFinish < tbl_ActionsAndSales.DateFinish "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyErrStr = "�������� ��� ����� / ���������� �� ���� ��������� ������������� � ��� ������������:" & Chr(13) & Chr(10)
                        While Declarations.MyRec.EOF = False
                            MyErrStr = MyErrStr + Chr(13) & Chr(10) & Declarations.MyRec.Fields("ActionName").Value + Chr(9) + "�����" + Chr(9) + Declarations.MyRec.Fields("ScalaCode").Value + Chr(9) + "�" + Chr(9) + Declarations.MyRec.Fields("DateStart").Value & Chr(9) + "��" + Chr(9) & Declarations.MyRec.Fields("DateFinish").Value
                            Declarations.MyRec.MoveNext()
                        End While
                        MyErrStr = MyErrStr + Chr(13) & Chr(10) + Chr(13) & Chr(10) & "���� �� �������� ""���������� ������� ��������"", �� ���� " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ����� / ���������� ����� �������� �� ���� ������ �����������  ����� ����� 1 ����. " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "���� �� �� ������ ������ ����� ������� ���� ��������� ����� / ����������, " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "�� �������� ""�����"", " & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "��������� ���� ��������� ����� / ���������� � Excel ����� �� �����������" & Chr(13) & Chr(10)
                        MyErrStr = MyErrStr + "� ��������� ���� �� �����." & Chr(13) & Chr(10)
                        MyErrForm = New ErrForm
                        MyErrForm.MyErrStr = MyErrStr
                        MyErrForm.ShowDialog()
                        If MyErrRezult = 0 Then
                            Throw New System.Exception("�������� �������� ��� � Excel �����.")
                        Else    '-----������ ���� ������ �����
                            MySQLStr = "UPDATE " & MyTableName & " "
                            MySQLStr = MySQLStr & "SET DateFinish = DateAdd(dd, -1, View_2.DateStart) "
                            MySQLStr = MySQLStr & "FROM (SELECT tbl_ActionsAndSales.ActionName, tbl_ActionsAndSales.ScalaCode, tbl_ActionsAndSales.DateStart, tbl_ActionsAndSales.DateFinish "
                            MySQLStr = MySQLStr & "FROM " & MyTableName & " AS " & MyTableName & "_1 INNER JOIN "
                            MySQLStr = MySQLStr & "tbl_ActionsAndSales ON " & MyTableName & "_1.ScalaCode = tbl_ActionsAndSales.ScalaCode AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish > tbl_ActionsAndSales.DateStart AND "
                            MySQLStr = MySQLStr & " " & MyTableName & "_1.DateFinish < tbl_ActionsAndSales.DateFinish) AS View_2 INNER JOIN "
                            MySQLStr = MySQLStr & "" & MyTableName & " ON View_2.ScalaCode = " & MyTableName & ".ScalaCode "
                            InitMyConn(False)
                            Declarations.MyConn.Execute(MySQLStr)
                        End If
                    End If



                    '==============================��������� ����� / ���������� � ��=============================
                    Label3.Text = "��������� ����� / ���������� � ��"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "INSERT INTO tbl_ActionsAndSales "
                    MySQLStr = MySQLStr & "(ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, ActionFinished, ActionFinishedDate) "
                    MySQLStr = MySQLStr & "SELECT ScalaCode, PurchasePrice, PurchasePriceCurr, MarginCoeff, QTYAction, ActionStopQTY, TimeAction, DateStart, DateFinish, ActionName, ActionOrSales, 0 AS ActionFinished, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, '01/01/1900', 103) AS ActionFinishedDate "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    '==============================================������ ��������� ����� ����� �� �������============================================
                    MyRez = MsgBox("���������� ������ ����� - ����� �� ������� ������? ����� ������ ����� �������� �����.", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = MsgBoxResult.Yes Then
                        '----------------------------������� ��������
                        Label3.Text = "������������ ����� ����� �� �������. "
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()

                        MySQLStr = "Exec spp_PrepareCommonPriceList_PriCost "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If


                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                Finally
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Declarations.MyConn.Close()
                    Declarations.MyConn = Nothing
                    '----------------------------������� ��������
                    Label3.Text = ""
                End Try

                Me.Cursor = Cursors.Default
                oWorkBook.Close(True)
            End If
        End If
    End Sub

    Private Function GetFirstExcelSheetName(ByRef cn As OleDbConnection) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������� ����� Excel  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyTable As String
        Dim dt As DataTable

        Try
            cn.Open()
            MyTable = cn.GetSchema("Tables").Rows(0)("TABLE_NAME")
            cn.Close()
            GetFirstExcelSheetName = MyTable
        Catch ex As Exception
            GetFirstExcelSheetName = ""
        End Try
    End Function

    Private Function GetExcelDataSet(ByRef cn As OleDbConnection, ByVal MySQLStr As String) As DataSet
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� dataset  Excel  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim cmd As OleDbDataAdapter
        Dim ds As New DataSet()

        Try
            cmd = New OleDbDataAdapter(MySQLStr, cn)
            cn.Open()
            cmd.Fill(ds, "Table1")
            cn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        GetExcelDataSet = ds
    End Function

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������, �������� �������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        LoadActionsList()
        DateTimePicker1.Value = CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))
    End Sub

    Private Sub LoadActionsList()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �����, ������� ����� ������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '��� ������ �����
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ActionName, ActionName + ' From ' + CONVERT(nvarchar(30), DateStart, 103) + ' To ' + CONVERT(nvarchar(30), DateFinish, 103) AS ActionFullName "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionFinished = 0) "
        MySQLStr = MySQLStr & "GROUP BY ActionName, DateStart, DateFinish "
        MySQLStr = MySQLStr & "ORDER BY ActionFullName "
        InitMyConn(False)

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "ActionFullName" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "ActionName"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� / ���������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        If CheckDate() = True Then
            MySQLStr = "UPDATE tbl_ActionsAndSales "
            MySQLStr = MySQLStr & "SET ActionFinished = 1, "
            MySQLStr = MySQLStr & "ActionFinishedDate = CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103)"
            MySQLStr = MySQLStr & "WHERE (ActionName = N'" & ComboBox1.SelectedValue & "') "
            MySQLStr = MySQLStr & "AND (ActionFinished = 0) "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MsgBox("����������� �������� " & ComboBox1.SelectedValue & ".", MsgBoxStyle.OkOnly, "��������!")
            LoadActionsList()
        Else
            DateTimePicker1.Select()
        End If
    End Sub

    Private Function CheckDate() As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ����� ���� ��������  ����� / ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������

        MySQLStr = "SELECT DateStart, DateFinish "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionName = N'" & ComboBox1.SelectedValue & "') "
        MySQLStr = MySQLStr & "AND (ActionFinished = 0) "
        MySQLStr = MySQLStr & "GROUP BY DateStart, DateFinish "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("���������� ��������� ������������ ����������� ���� �������� ����� / ����������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
            CheckDate = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            If Declarations.MyRec.Fields("DateStart").Value <= DateTimePicker1.Value And Declarations.MyRec.Fields("DateFinish").Value >= DateTimePicker1.Value _
                And CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())) <= DateTimePicker1.Value Then
                CheckDate = True
                Exit Function
            Else
                MsgBox("���� �������� ����� / ���������� ������ ���� ������ ��� ����� ������� � ���� � ��������� �� ���� ������ ����� / ���������� �� ���������� ����� / ����������.", MsgBoxStyle.Critical, "��������!")
                CheckDate = False
                Exit Function
            End If
        End If
    End Function
End Class
