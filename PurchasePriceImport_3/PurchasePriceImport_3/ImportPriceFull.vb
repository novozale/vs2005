Public Class ImportPriceFull

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ImportPriceFull_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        '// ������ ��������� �������� �� Excel �������� ����� ����� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Button1.Enabled = False
        Button2.Enabled = False
        System.Windows.Forms.Application.DoEvents()
        If My.Settings.UseOffice = "LibreOffice" Then
            ImportDataFromLO()
        Else
            ImportDataFromExcel()
        End If
        Button1.Enabled = True
        Button2.Enabled = True
        System.Windows.Forms.Application.DoEvents()
        MsgBox("��������� �������� �������������� ���������� �� ������� ���������.", vbOKOnly, "��������!")
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel �������������� ���������� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyGuid As String                          '
        Dim MyTableName As String                   '��� ��������� �������
        Dim MySysTableName As String                '��� ��������� ��������� �������
        Dim connStr As String                       '������ ���������� � Excel
        Dim cn As OleDbConnection                   '������ ���������� � OLE
        Dim MySQLStr As String                      'SQL ������
        Dim FirstExcelSheetName As String           '�������� ������� ����� Excel
        Dim myds As DataSet                         'Excel dataset
        Dim MyVersion As String                     '������ ���������
        Dim MySuppCode As String                    '��� ����������
        Dim mycount As Integer
        Dim MyDBL As Double                         '��� ��������
        Dim MyParams(10) As Integer                  '��������� �������� �������
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_PurchasePriceItems_AddInfo_Tmp_" + MyGuid
        MySysTableName = "tbl_PurchasePriceItems_AddInfoSys_Tmp_" + MyGuid

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
                            MySQLStr = MySQLStr & "WHERE (Name = N'������ �������������� ���������� �� ������� �� ����� ������') "
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

                    '---���������, ��� ���������� ��� ���������� � Scala 
                    MySuppCode = ""
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "C2:C2]"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds Is Nothing = False Then
                        If IsDBNull(myds.Tables(0).Rows(0).Item(0)) Then
                            MsgBox("� ������������� ����� Excel � ������ 'C2' �� ���������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            MySuppCode = Trim(myds.Tables(0).Rows(0).Item(0))
                        End If
                    Else
                        MsgBox("���������� ��������� ��� ���������. ���������� � ��������������.", vbCritical, "��������!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '---��������� ��� ���� ��������� ���� � Scala
                    MySQLStr = "SELECT COUNT(PL01001) AS CC "
                    MySQLStr = MySQLStr & "FROM PL010300 "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ 'C2' ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If

                    '========================================��������� ���������� �������� �������=================================
                    Label3.Text = "�������� ���������� �������� ������� Excel �����"
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A5:L5] "
                    myds = GetExcelDataSet(cn, MySQLStr)
                    '-----�������� ��������� �������
                    If Trim(myds.Tables(0).Rows(0).Item(1).ToString) = "" Then
                        MsgBox("��� '�������� ��������� �������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(1).ToString) = "�� ���������" Then
                        MyParams(0) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(1).ToString) = "��������� ��������" Then
                        MyParams(0) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(1).ToString) = "��������� ���" Then
                        MyParams(0) = 2
                    Else
                        MsgBox("��� '�������� ��������� �������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----�������� ������
                    If Trim(myds.Tables(0).Rows(0).Item(2).ToString) = "" Then
                        MsgBox("��� '�������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(2).ToString) = "�� ���������" Then
                        MyParams(1) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(2).ToString) = "��������� ��������" Then
                        MyParams(1) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(2).ToString) = "��������� ���" Then
                        MyParams(1) = 2
                    Else
                        MsgBox("��� '�������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----�������� ������
                    If Trim(myds.Tables(0).Rows(0).Item(3).ToString) = "" Then
                        MsgBox("��� '�������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(3).ToString) = "�� ���������" Then
                        MyParams(2) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(3).ToString) = "��������� ��������" Then
                        MyParams(2) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(3).ToString) = "��������� ���" Then
                        MyParams(2) = 2
                    Else
                        MsgBox("��� '�������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----��� ������ ������
                    If Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "" Then
                        MsgBox("��� '��� ������ ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "�� ���������" Then
                        MyParams(3) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "��������� ��������" Then
                        MyParams(3) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "��������� ���" Then
                        MyParams(3) = 2
                    Else
                        MsgBox("��� '��� ������ ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----��� ��������� ������
                    If Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "" Then
                        MsgBox("��� '��� ��������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "�� ���������" Then
                        MyParams(3) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "��������� ��������" Then
                        MyParams(3) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(4).ToString) = "��������� ���" Then
                        MyParams(3) = 2
                    Else
                        MsgBox("��� '��� ��������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----��� �������������
                    If Trim(myds.Tables(0).Rows(0).Item(5).ToString) = "" Then
                        MsgBox("��� '��� �������������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(5).ToString) = "�� ���������" Then
                        MyParams(4) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(5).ToString) = "��������� ��������" Then
                        MyParams(4) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(5).ToString) = "��������� ���" Then
                        MyParams(4) = 2
                    Else
                        MsgBox("��� '��� �������������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----��� ������ �������������
                    If Trim(myds.Tables(0).Rows(0).Item(6).ToString) = "" Then
                        MsgBox("��� '��� ������ �������������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(6).ToString) = "�� ���������" Then
                        MyParams(5) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(6).ToString) = "��������� ��������" Then
                        MyParams(5) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(6).ToString) = "��������� ���" Then
                        MyParams(5) = 2
                    Else
                        MsgBox("��� '��� ������ �������������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----�����
                    If Trim(myds.Tables(0).Rows(0).Item(7).ToString) = "" Then
                        MsgBox("��� '�����' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(7).ToString) = "�� ���������" Then
                        MyParams(6) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(7).ToString) = "��������� ��������" Then
                        MyParams(6) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(7).ToString) = "��������� ���" Then
                        MyParams(6) = 2
                    Else
                        MsgBox("��� '�����' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----������
                    If Trim(myds.Tables(0).Rows(0).Item(8).ToString) = "" Then
                        MsgBox("��� '������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(8).ToString) = "�� ���������" Then
                        MyParams(7) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(8).ToString) = "��������� ��������" Then
                        MyParams(7) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(8).ToString) = "��������� ���" Then
                        MyParams(7) = 2
                    Else
                        MsgBox("��� '������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----������
                    If Trim(myds.Tables(0).Rows(0).Item(9).ToString) = "" Then
                        MsgBox("��� '������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(9).ToString) = "�� ���������" Then
                        MyParams(8) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(9).ToString) = "��������� ��������" Then
                        MyParams(8) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(9).ToString) = "��������� ���" Then
                        MyParams(8) = 2
                    Else
                        MsgBox("��� '������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----���
                    If Trim(myds.Tables(0).Rows(0).Item(10).ToString) = "" Then
                        MsgBox("��� '���' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(10).ToString) = "�� ���������" Then
                        MyParams(9) = 0
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(10).ToString) = "��������� ��������" Then
                        MyParams(9) = 1
                    ElseIf Trim(myds.Tables(0).Rows(0).Item(10).ToString) = "��������� ���" Then
                        MyParams(9) = 2
                    Else
                        MsgBox("��� '���' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----���� �� 1 ������ ���� � ���������, �������� �� "�� ���������"
                    mycount = 0
                    For i As Integer = 0 To 10
                        If MyParams(i) <> 0 Then
                            mycount = mycount + 1
                        End If
                    Next
                    If mycount = 0 Then
                        MsgBox("��� ���� ������� ������� �������� '�� ���������'. ��� �������� ����������, ����� ���� �� ��� ����� ������� ���� ������� ��������, �������� �� '�� ���������'.")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '---��������� ������������ ������ � Excel
                    Label3.Text = "�������� ������������ ������ Excel �����"
                    '-----������������� ����
                    MySQLStr = "SELECT F1 FROM [" & FirstExcelSheetName & "A6:A] group by F1 having(count(F1) > 1)"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds.Tables(0).Rows.Count > 0 Then
                        MsgBox("� ����� ��������� " & myds.Tables(0).Rows.Count & " ������������� ������� ����� ������� ����������. �������������� ������� ""���������� �������������"" � Excel, ��������� � ������� ������ ���� ")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----������������ ��������� ������ � Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A6:L] where(F1 <> """")"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    '-----������������ ���������
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        '-----���������� ���� ������ ����������
                        If Trim(myds.Tables(0).Rows(mycount).Item(0).ToString) = "" Then
                            MsgBox("������ " & CStr(mycount + 6) & " �� ������� ��� ������ ����������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If

                        '-----���������� �����
                        If Trim(myds.Tables(0).Rows(mycount).Item(8).ToString) <> "" Then
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(8)
                            Catch ex As Exception
                                MsgBox("������ " & CStr(mycount + 6) & " ����������� �������� �����")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If

                        '-----���������� ������
                        If Trim(myds.Tables(0).Rows(mycount).Item(9).ToString) <> "" Then
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(9)
                            Catch ex As Exception
                                MsgBox("������ " & CStr(mycount + 6) & " ����������� �������� ������")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If

                        '-----���������� ������
                        If Trim(myds.Tables(0).Rows(mycount).Item(10).ToString) <> "" Then
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(10)
                            Catch ex As Exception
                                MsgBox("������ " & CStr(mycount + 6) & " ����������� �������� ������")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If

                        '-----���������� ����
                        If Trim(myds.Tables(0).Rows(mycount).Item(11).ToString) <> "" Then
                            Try
                                MyDBL = myds.Tables(0).Rows(mycount).Item(11)
                            Catch ex As Exception
                                MsgBox("������ " & CStr(mycount + 6) & " ����������� ������� ���")
                                Me.Cursor = Cursors.Default
                                Exit Sub
                            End Try
                        End If
                        mycount = mycount + 1
                    End While


                    '========================================�������� �������� �� ��������� �������=================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������ �� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----�������� ��������� ������
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) NOT NULL, "
                    MySQLStr = MySQLStr & "[CollectionName] [nvarchar](255) NULL, "
                    MySQLStr = MySQLStr & "[WEBName] [nvarchar](250) NULL, "
                    MySQLStr = MySQLStr & "[Description] [nvarchar](max) NULL, "
                    MySQLStr = MySQLStr & "[ItemGroupCode] [nvarchar](50) NULL, "
                    MySQLStr = MySQLStr & "[ItemSubGroupCode] [nvarchar](50) NULL, "
                    MySQLStr = MySQLStr & "[ManufacturerCode] [bigint] NULL, "
                    MySQLStr = MySQLStr & "[ManufacturerItemCode] [nvarchar](100) NULL, "
                    MySQLStr = MySQLStr & "[ItemLength] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemWidth] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemHeight] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemWeight] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "CONSTRAINT [PK_" & MyTableName & "] PRIMARY KEY CLUSTERED "
                    MySQLStr = MySQLStr & "( "
                    MySQLStr = MySQLStr & "[SC01060] Asc "
                    MySQLStr = MySQLStr & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
                    MySQLStr = MySQLStr & "ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    Try
                        MySQLStr = "DROP TABLE " & MySysTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MySysTableName & "]( "
                    MySQLStr = MySQLStr & "[MyParam0] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam1] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam2] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam3] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam4] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam5] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam6] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam7] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam8] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam9] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam10] [int] NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ 
                    InitMyConn(False)
                    MySQLStr = "SELECT SC01060, CollectionName, WEBName, Description, ItemGroupCode, ItemSubGroupCode, ManufacturerCode, "
                    MySQLStr = MySQLStr & "ManufacturerItemCode, ItemLength, ItemWidth, ItemHeight, ItemWeight "
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
                        dr.Item(5) = myds.Tables(0).Rows(mycount).Item(5)
                        dr.Item(6) = myds.Tables(0).Rows(mycount).Item(6)
                        dr.Item(7) = myds.Tables(0).Rows(mycount).Item(7)
                        dr.Item(8) = myds.Tables(0).Rows(mycount).Item(8)
                        dr.Item(9) = myds.Tables(0).Rows(mycount).Item(9)
                        dr.Item(10) = myds.Tables(0).Rows(mycount).Item(10)
                        dr.Item(11) = myds.Tables(0).Rows(mycount).Item(11)
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

                    '-----���������� ������� � �����������
                    MySQLStr = "INSERT INTO " & MySysTableName & " "
                    MySQLStr = MySQLStr & "(MyParam0, MyParam1, MyParam2, MyParam3, MyParam4, MyParam5, MyParam6, MyParam7, MyParam8, MyParam9, MyParam10) "
                    MySQLStr = MySQLStr & "VALUES (" & CStr(MyParams(0)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(1)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(2)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(3)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(4)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(5)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(6)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(7)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(8)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(9)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(10)) & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)


                    '==============================================���������� �������� �� �������============================================
                    '----------------------------������� ��������
                    Label3.Text = "���������� �������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    If ServerChecks(MyTableName, MySysTableName) = False Then
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If


                    '======================================��������� �������� ���������� � �������===========================================
                    '-----���������
                    '----------------------------������� ��������
                    Label3.Text = "���������� �������������� ���������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MySQLStr = "UPDATE tbl_PurchasePriceItems_AddInfo "
                    MySQLStr = MySQLStr + "SET CollectionName = CASE WHEN MyParam0 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam0 = 1 AND " + Trim(MyTableName) + ".CollectionName IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".CollectionName ELSE tbl_PurchasePriceItems_AddInfo.CollectionName "
                    MySQLStr = MySQLStr + "END, WEBName = CASE WHEN MyParam1 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam1 = 1 AND " + Trim(MyTableName) + ".WEBName IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".WEBName ELSE tbl_PurchasePriceItems_AddInfo.WEBName END, "
                    MySQLStr = MySQLStr + "Description = CASE WHEN MyParam2 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam2 = 1 AND " + Trim(MyTableName) + ".Description IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".Description ELSE tbl_PurchasePriceItems_AddInfo.Description END, "
                    MySQLStr = MySQLStr + "ItemGroupCode = CASE WHEN MyParam3 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam3 = 1 AND " + Trim(MyTableName) + ".ItemGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemGroupCode ELSE tbl_PurchasePriceItems_AddInfo.ItemGroupCode "
                    MySQLStr = MySQLStr + "END, ItemSubGroupCode = CASE WHEN MyParam4 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam4 = 1 AND " + Trim(MyTableName) + ".ItemSubGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemSubGroupCode ELSE tbl_PurchasePriceItems_AddInfo.ItemSubGroupCode "
                    MySQLStr = MySQLStr + "END, ManufacturerCode = CASE WHEN MyParam5 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam5 = 1 AND " + Trim(MyTableName) + ".ManufacturerCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ManufacturerCode ELSE tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "END, ManufacturerItemCode = CASE WHEN MyParam6 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam6 = 1 AND " + Trim(MyTableName) + ".ManufacturerItemCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ManufacturerItemCode ELSE tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "END, ItemLength = CASE WHEN MyParam7 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam7 = 1 AND " + Trim(MyTableName) + ".ItemLength IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemLength ELSE tbl_PurchasePriceItems_AddInfo.ItemLength END, "
                    MySQLStr = MySQLStr + "ItemWidth = CASE WHEN MyParam8 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam8 = 1 AND " + Trim(MyTableName) + ".ItemWidth IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemWidth ELSE tbl_PurchasePriceItems_AddInfo.ItemWidth END, "
                    MySQLStr = MySQLStr + "ItemHeight = CASE WHEN MyParam9 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam9 = 1 AND " + Trim(MyTableName) + ".ItemHeight IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemHeight ELSE tbl_PurchasePriceItems_AddInfo.ItemHeight END, "
                    MySQLStr = MySQLStr + "ItemWeight = CASE WHEN MyParam10 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam10 = 1 AND " + Trim(MyTableName) + ".ItemWeight IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemWeight ELSE tbl_PurchasePriceItems_AddInfo.ItemWeight END "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo INNER JOIN "
                    MySQLStr = MySQLStr + Trim(MyTableName) + " ON "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo.SC01060 = " + Trim(MyTableName) + ".SC01060 CROSS JOIN "
                    MySQLStr = MySQLStr + Trim(MySysTableName) + " "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.PL01001 = N'" + Trim(MySuppCode) + "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----��������
                    Label3.Text = "�������� ����� ���������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MySQLStr = "INSERT INTO tbl_PurchasePriceItems_AddInfo "
                    MySQLStr = MySQLStr + "SELECT NEWID() AS ID, N'" + Trim(MySuppCode) + "' AS PL01001, " + Trim(MyTableName) + ".SC01060, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".CollectionName, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".WEBName, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".Description, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemGroupCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemSubGroupCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ManufacturerCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ManufacturerItemCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemLength, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemWidth, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemHeight, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemWeight "
                    MySQLStr = MySQLStr + "FROM " + Trim(MyTableName) + " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr + "(SELECT ID, PL01001, SC01060, CollectionName, WEBName, Description, ItemGroupCode, ItemSubGroupCode, ManufacturerCode, "
                    MySQLStr = MySQLStr + "ManufacturerItemCode, ItemLength, ItemWidth, ItemHeight, ItemWeight "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo AS tbl_PurchasePriceItems_AddInfo_1 "
                    MySQLStr = MySQLStr + "WHERE (PL01001 = N'" + Trim(MySuppCode) + "')) AS View_8 ON "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".SC01060 = View_8.SC01060 "
                    MySQLStr = MySQLStr + "WHERE (View_8.ID IS NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� ������ � �������� ������
                    Label3.Text = "����� ���� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01037 = tbl_PurchasePriceItems_AddInfo.ItemGroupCode "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_PurchasePriceItems_AddInfo.PL01001 = SC010300.SC01058 AND "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo.SC01060 = SC010300.SC01060 "
                    MySQLStr = MySQLStr + "And tbl_PurchasePriceItems_AddInfo.ItemGroupCode <> SC010300.SC01037 "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "AND (tbl_PurchasePriceItems_AddInfo.ItemGroupCode <> N'') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ����� � �������� ������
                    Label3.Text = "����� ����� � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01007 = tbl_PurchasePriceItems_AddInfo.ItemLength "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01007 <> tbl_PurchasePriceItems_AddInfo.ItemLength "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemLength IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ������ � �������� ������
                    Label3.Text = "����� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01008 = tbl_PurchasePriceItems_AddInfo.ItemWidth "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01008 <> tbl_PurchasePriceItems_AddInfo.ItemWidth "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemWidth IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ������ � �������� ������
                    Label3.Text = "����� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01009 = tbl_PurchasePriceItems_AddInfo.ItemHeight "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01009 <> tbl_PurchasePriceItems_AddInfo.ItemHeight "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemHeight IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� � �������� ������
                    Label3.Text = "����� ���� � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01069 = tbl_PurchasePriceItems_AddInfo.ItemWeight "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01069 <> tbl_PurchasePriceItems_AddInfo.ItemWeight "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemWeight IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� �������������
                    Label3.Text = "����� ���� �������������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "UPDATE tbl_ItemCard0300 "
                    MySQLStr = MySQLStr + "SET Manufacturer = tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "FROM tbl_ItemCard0300 INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_ItemCard0300.SC01001 = SC010300.SC01001 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 AND "
                    MySQLStr = MySQLStr + "tbl_ItemCard0300.Manufacturer <> tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ManufacturerCode IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� ������ �������������
                    Label3.Text = "����� ���� ������ �������������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "UPDATE tbl_ItemCard0300 "
                    MySQLStr = MySQLStr + "SET ManufacturerItemCode = tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "FROM tbl_ItemCard0300 INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_ItemCard0300.SC01001 = SC010300.SC01001 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 AND "
                    MySQLStr = MySQLStr + "tbl_ItemCard0300.ManufacturerItemCode <> tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)





                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                Finally
                    cn.Close()
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                        MySQLStr = "DROP TABLE " & MySysTableName & " "
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

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Libre Office �������������� ���������� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyGuid As String                          '
        Dim MyTableName As String                   '��� ��������� �������
        Dim MySysTableName As String                '��� ��������� ��������� �������
        Dim MySQLStr As String                      'SQL ������
        Dim MyVersion As String                     '������ ���������
        Dim MySuppCode As String                    '��� ����������
        Dim MyParams(10) As Integer                  '��������� �������� �������
        Dim mycount As Integer
        Dim MyStr As String
        Dim MyDBL As Double
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_PurchasePriceItems_AddInfo_Tmp_" + MyGuid
        MySysTableName = "tbl_PurchasePriceItems_AddInfoSys_Tmp_" + MyGuid
        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                '----------------------------������� ��������
                Label3.Text = "���������� �������� Libre Office �����"
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
                    '---��������� ������ ����� Libre Office
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    If MyVersion = "" Then
                        MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Libre Office ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        MySQLStr = "SELECT Version "
                        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
                        MySQLStr = MySQLStr & "WHERE (Name = N'������ �������������� ���������� �� ������� �� ����� ������') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            MsgBox("� Scala �� ����������� ������� ������ ����� Libre Office. ���������� � ��������������", vbCritical, "��������!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                trycloseMyRec()
                            Else
                                MsgBox("�� ��������� �������� � ������������ ������� ����� Libre Office. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    End If

                    '---���������, ��� ���������� ��� ���������� � Scala � ��� �� ����������
                    MySuppCode = Trim(oSheet.getCellRangeByName("C2").String)
                    If MySuppCode.Equals("") Then
                        MsgBox("� ������������� ����� Excel � ������ ""C2"" �� ���������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '---��������� ��� ���� ��������� ���� � Scala
                    MySQLStr = "SELECT COUNT(PL01001) AS CC "
                    MySQLStr = MySQLStr & "FROM PL010300 "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If (Declarations.MyRec.Fields("CC").Value = 0) Then
                        MsgBox("� ������������� ����� Excel � ������ 'C2' ���������� �������� ��� ���������� � Scala ", MsgBoxStyle.Critical, "��������!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        trycloseMyRec()
                    End If

                    '========================================��������� ���������� �������� �������=================================
                    Label3.Text = "�������� ���������� �������� ������� LibreOffice �����"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----�������� ��������� ������� (0)
                    MyStr = Trim(oSheet.getCellRangeByName("B5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '�������� ��������� �������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(0) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(0) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(0) = 2
                    Else
                        MsgBox("��� '�������� ��������� �������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----�������� ������ (1)
                    MyStr = Trim(oSheet.getCellRangeByName("C5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '�������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(1) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(1) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(1) = 2
                    Else
                        MsgBox("��� '�������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----�������� ������ (2)
                    MyStr = Trim(oSheet.getCellRangeByName("D5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '�������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(2) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(2) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(2) = 2
                    Else
                        MsgBox("��� '�������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----��� ������ ������ (3)
                    MyStr = Trim(oSheet.getCellRangeByName("E5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '��� ������ ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(3) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(3) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(3) = 2
                    Else
                        MsgBox("��� '��� ������ ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----��� ��������� ������ (4)
                    MyStr = Trim(oSheet.getCellRangeByName("F5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '��� ��������� ������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(4) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(4) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(4) = 2
                    Else
                        MsgBox("��� '��� ��������� ������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----��� ������������� (5)
                    MyStr = Trim(oSheet.getCellRangeByName("G5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '��� �������������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(5) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(5) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(5) = 2
                    Else
                        MsgBox("��� '��� �������������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----��� ������ ������������� (6)
                    MyStr = Trim(oSheet.getCellRangeByName("H5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '��� ������ �������������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(6) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(6) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(6) = 2
                    Else
                        MsgBox("��� '��� ������ �������������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----����� (7)
                    MyStr = Trim(oSheet.getCellRangeByName("I5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '�����' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(7) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(7) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(7) = 2
                    Else
                        MsgBox("��� '�����' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----������ (8)
                    MyStr = Trim(oSheet.getCellRangeByName("J5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(8) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(8) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(8) = 2
                    Else
                        MsgBox("��� '������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----������ (9)
                    MyStr = Trim(oSheet.getCellRangeByName("K5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '������' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(9) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(9) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(9) = 2
                    Else
                        MsgBox("��� '������' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----��� (10)
                    MyStr = Trim(oSheet.getCellRangeByName("L5").String)
                    If MyStr.Equals("") Then
                        MsgBox("��� '���' �� ������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    ElseIf MyStr.Equals("�� ���������") Then
                        MyParams(10) = 0
                    ElseIf MyStr.Equals("��������� ��������") Then
                        MyParams(10) = 1
                    ElseIf MyStr.Equals("��������� ���") Then
                        MyParams(10) = 2
                    Else
                        MsgBox("��� '���' ������ ������������ ��� ��������. �������� �� ����������� ����.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '-----���� �� 1 ������ ���� � ���������, �������� �� "�� ���������"
                    mycount = 0
                    For i As Integer = 0 To 10
                        If MyParams(i) <> 0 Then
                            mycount = mycount + 1
                        End If
                    Next
                    If mycount = 0 Then
                        MsgBox("��� ���� ������� ������� �������� '�� ���������'. ��� �������� ����������, ����� ���� �� ��� ����� ������� ���� ������� ��������, �������� �� '�� ���������'.")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '---��������� ������������ ������ � Excel
                    Label3.Text = "�������� ������������ ������ Libre Office �����"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----������������� ����
                    oSheet.unprotect("!pass2022")

                    Dim args() As Object
                    ReDim args(0)
                    args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args(0).Name = "ToPoint"
                    args(0).Value = "$A$6:$M$300000"
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

                    Dim StopFlag = False
                    Dim MyStep As Integer = 50000
                    Dim srange As Object
                    Dim myarr As Object
                    For j As Integer = 0 To 300000 Step MyStep
                        srange = oSheet.getCellRangeByName("A" & 6 + j & ":M" & 6 + j + MyStep + 1)
                        myarr = srange.DataArray
                        For i As Integer = 1 To MyStep
                            If myarr(i)(0).ToString.Equals("") Then
                                StopFlag = True
                                Exit For
                            Else
                                If myarr(i)(0).ToString.Equals(myarr(i - 1)(0).ToString) And Not myarr(i)(0).ToString.Equals("") Then
                                    MsgBox("� ����� ��������� ������������� ������ ����� ������� � Scala. �������������� ������� ""�������� ������������� ������"" � LibreOffice, ��������� � ������� ������ ���� ")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If
                            End If
                        Next i
                        If StopFlag Then Exit For
                    Next j
                    Dim args2() As Object
                    ReDim args2(0)
                    oDispatcher.executeDispatch(oFrame, ".uno:Save", "", 0, args2)

                    '-----������������ ��������� ������ � LibreOffice
                    StopFlag = False
                    For j As Integer = 0 To 300000 Step MyStep
                        srange = oSheet.getCellRangeByName("A" & 6 + j & ":M" & 6 + j + MyStep)
                        myarr = srange.DataArray
                        For i As Integer = 0 To MyStep
                            If myarr(i)(0).ToString.Equals("") Then
                                StopFlag = True
                                Exit For
                            Else
                                '-----���������� �����
                                If Not myarr(i)(8).ToString.Equals("") Then
                                    Try
                                        MyDBL = myarr(i)(8)
                                    Catch ex As Exception
                                        MsgBox("������ " & CStr(i + 6) & " ����������� �������� �����")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If

                                '-----���������� ������
                                If Not myarr(i)(9).ToString.Equals("") Then
                                    Try
                                        MyDBL = myarr(i)(9)
                                    Catch ex As Exception
                                        MsgBox("������ " & CStr(i + 6) & " ����������� �������� ������")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If

                                '-----���������� ������
                                If Not myarr(i)(10).ToString.Equals("") Then
                                    Try
                                        MyDBL = myarr(i)(10)
                                    Catch ex As Exception
                                        MsgBox("������ " & CStr(i + 6) & " ����������� �������� ������")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If

                                '-----���������� ����
                                If Not myarr(i)(11).ToString.Equals("") Then
                                    Try
                                        MyDBL = myarr(i)(11)
                                    Catch ex As Exception
                                        MsgBox("������ " & CStr(i + 6) & " ����������� ������� ���")
                                        Me.Cursor = Cursors.Default
                                        oWorkBook.Close(True)
                                        Exit Sub
                                    End Try
                                End If
                            End If
                        Next i
                        If StopFlag Then Exit For
                    Next j

                    '========================================�������� �������� �� ��������� �������=================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������ �� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----�������� ��������� ������
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) NOT NULL, "
                    MySQLStr = MySQLStr & "[CollectionName] [nvarchar](255) NULL, "
                    MySQLStr = MySQLStr & "[WEBName] [nvarchar](250) NULL, "
                    MySQLStr = MySQLStr & "[Description] [nvarchar](max) NULL, "
                    MySQLStr = MySQLStr & "[ItemGroupCode] [nvarchar](50) NULL, "
                    MySQLStr = MySQLStr & "[ItemSubGroupCode] [nvarchar](50) NULL, "
                    MySQLStr = MySQLStr & "[ManufacturerCode] [bigint] NULL, "
                    MySQLStr = MySQLStr & "[ManufacturerItemCode] [nvarchar](100) NULL, "
                    MySQLStr = MySQLStr & "[ItemLength] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemWidth] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemHeight] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "[ItemWeight] [numeric](28, 8) NULL, "
                    MySQLStr = MySQLStr & "CONSTRAINT [PK_" & MyTableName & "] PRIMARY KEY CLUSTERED "
                    MySQLStr = MySQLStr & "( "
                    MySQLStr = MySQLStr & "[SC01060] Asc "
                    MySQLStr = MySQLStr & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, "
                    MySQLStr = MySQLStr & "ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    Try
                        MySQLStr = "DROP TABLE " & MySysTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MySysTableName & "]( "
                    MySQLStr = MySQLStr & "[MyParam0] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam1] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam2] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam3] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam4] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam5] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam6] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam7] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam8] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam9] [int] NOT NULL, "
                    MySQLStr = MySQLStr & "[MyParam10] [int] NOT NULL "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ 
                    InitMyConn(False)
                    MySQLStr = "SELECT SC01060, CollectionName, WEBName, Description, ItemGroupCode, ItemSubGroupCode, ManufacturerCode, "
                    MySQLStr = MySQLStr & "ManufacturerItemCode, ItemLength, ItemWidth, ItemHeight, ItemWeight "
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
                    StopFlag = False
                    For j As Integer = 0 To 300000 Step MyStep
                        srange = oSheet.getCellRangeByName("A" & 6 + j & ":L" & 5 + j + MyStep)
                        myarr = srange.DataArray
                        For i As Integer = 0 To MyStep - 1
                            If myarr(i)(0).ToString.Equals("") Then
                                StopFlag = True
                                Exit For
                            Else
                                dr = dt.NewRow
                                '-----
                                If myarr(i)(0) Is Nothing Then
                                    MsgBox("������ ������ � ������ A:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(0).ToString.Equals("") Then
                                    dr.Item(0) = DBNull.Value
                                Else
                                    dr.Item(0) = myarr(i)(0)
                                End If
                                '-----
                                If myarr(i)(1) Is Nothing Then
                                    MsgBox("������ ������ � ������ B:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(1).ToString.Equals("") Then
                                    dr.Item(1) = DBNull.Value
                                Else
                                    dr.Item(1) = myarr(i)(1)
                                End If
                                '-----
                                If myarr(i)(2) Is Nothing Then
                                    MsgBox("������ ������ � ������ C:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(2).ToString.Equals("") Then
                                    dr.Item(2) = DBNull.Value
                                Else
                                    dr.Item(2) = myarr(i)(2)
                                End If
                                '-----
                                If myarr(i)(3) Is Nothing Then
                                    MsgBox("������ ������ � ������ D:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(3).ToString.Equals("") Then
                                    dr.Item(3) = DBNull.Value
                                Else
                                    dr.Item(3) = myarr(i)(3)
                                End If
                                '-----
                                If myarr(i)(4) Is Nothing Then
                                    MsgBox("������ ������ � ������ E:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(4).ToString.Equals("") Then
                                    dr.Item(4) = DBNull.Value
                                Else
                                    dr.Item(4) = myarr(i)(4)
                                End If
                                '-----
                                If myarr(i)(5) Is Nothing Then
                                    MsgBox("������ ������ � ������ F:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(5).ToString.Equals("") Then
                                    dr.Item(5) = DBNull.Value
                                Else
                                    dr.Item(5) = myarr(i)(5)
                                End If
                                '-----
                                If myarr(i)(6) Is Nothing Then
                                    MsgBox("������ ������ � ������ G:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(6).ToString.Equals("") Then
                                    dr.Item(6) = DBNull.Value
                                Else
                                    dr.Item(6) = myarr(i)(6)
                                End If
                                '-----
                                If myarr(i)(7) Is Nothing Then
                                    MsgBox("������ ������ � ������ H:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(7).ToString.Equals("") Then
                                    dr.Item(7) = DBNull.Value
                                Else
                                    dr.Item(7) = myarr(i)(7)
                                End If
                                '-----
                                If myarr(i)(8) Is Nothing Then
                                    MsgBox("������ ������ � ������ I:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(8).ToString.Equals("") Then
                                    dr.Item(8) = DBNull.Value
                                Else
                                    dr.Item(8) = myarr(i)(8)
                                End If
                                '-----
                                If myarr(i)(9) Is Nothing Then
                                    MsgBox("������ ������ � ������ J:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(9).ToString.Equals("") Then
                                    dr.Item(9) = DBNull.Value
                                Else
                                    dr.Item(9) = myarr(i)(9)
                                End If
                                '-----
                                If myarr(i)(10) Is Nothing Then
                                    MsgBox("������ ������ � ������ K:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(10).ToString.Equals("") Then
                                    dr.Item(10) = DBNull.Value
                                Else
                                    dr.Item(10) = myarr(i)(10)
                                End If
                                '-----
                                If myarr(i)(11) Is Nothing Then
                                    MsgBox("������ ������ � ������ L:" & 6 + j + i, MsgBoxStyle.OkOnly, "��������!")
                                    Me.Cursor = Cursors.Default
                                    Exit Sub
                                End If
                                If myarr(i)(11).ToString.Equals("") Then
                                    dr.Item(11) = DBNull.Value
                                Else
                                    dr.Item(11) = myarr(i)(11)
                                End If

                                dt.Rows.Add(dr)
                            End If
                        Next i
                        If StopFlag Then Exit For
                    Next j
                    Try
                        MySQLAdapter.Update(MySQLDs, "Table")
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End Try

                    '-----���������� ������� � �����������
                    MySQLStr = "INSERT INTO " & MySysTableName & " "
                    MySQLStr = MySQLStr & "(MyParam0, MyParam1, MyParam2, MyParam3, MyParam4, MyParam5, MyParam6, MyParam7, MyParam8, MyParam9, MyParam10) "
                    MySQLStr = MySQLStr & "VALUES (" & CStr(MyParams(0)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(1)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(2)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(3)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(4)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(5)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(6)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(7)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(8)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(9)) & " "
                    MySQLStr = MySQLStr & ", " & CStr(MyParams(10)) & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '==============================================���������� �������� �� �������============================================
                    '----------------------------������� ��������
                    Label3.Text = "���������� �������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    If ServerChecks(MyTableName, MySysTableName) = False Then
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '======================================��������� �������� ���������� � �������===========================================
                    '-----���������
                    '----------------------------������� ��������
                    Label3.Text = "���������� �������������� ���������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "UPDATE tbl_PurchasePriceItems_AddInfo "
                    MySQLStr = MySQLStr + "SET CollectionName = CASE WHEN MyParam0 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam0 = 1 AND " + Trim(MyTableName) + ".CollectionName IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".CollectionName ELSE tbl_PurchasePriceItems_AddInfo.CollectionName "
                    MySQLStr = MySQLStr + "END, WEBName = CASE WHEN MyParam1 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam1 = 1 AND " + Trim(MyTableName) + ".WEBName IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".WEBName ELSE tbl_PurchasePriceItems_AddInfo.WEBName END, "
                    MySQLStr = MySQLStr + "Description = CASE WHEN MyParam2 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam2 = 1 AND " + Trim(MyTableName) + ".Description IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".Description ELSE tbl_PurchasePriceItems_AddInfo.Description END, "
                    MySQLStr = MySQLStr + "ItemGroupCode = CASE WHEN MyParam3 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam3 = 1 AND " + Trim(MyTableName) + ".ItemGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemGroupCode ELSE tbl_PurchasePriceItems_AddInfo.ItemGroupCode "
                    MySQLStr = MySQLStr + "END, ItemSubGroupCode = CASE WHEN MyParam4 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam4 = 1 AND " + Trim(MyTableName) + ".ItemSubGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemSubGroupCode ELSE tbl_PurchasePriceItems_AddInfo.ItemSubGroupCode "
                    MySQLStr = MySQLStr + "END, ManufacturerCode = CASE WHEN MyParam5 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam5 = 1 AND " + Trim(MyTableName) + ".ManufacturerCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ManufacturerCode ELSE tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "END, ManufacturerItemCode = CASE WHEN MyParam6 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam6 = 1 AND " + Trim(MyTableName) + ".ManufacturerItemCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ManufacturerItemCode ELSE tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "END, ItemLength = CASE WHEN MyParam7 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam7 = 1 AND " + Trim(MyTableName) + ".ItemLength IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemLength ELSE tbl_PurchasePriceItems_AddInfo.ItemLength END, "
                    MySQLStr = MySQLStr + "ItemWidth = CASE WHEN MyParam8 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam8 = 1 AND " + Trim(MyTableName) + ".ItemWidth IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemWidth ELSE tbl_PurchasePriceItems_AddInfo.ItemWidth END, "
                    MySQLStr = MySQLStr + "ItemHeight = CASE WHEN MyParam9 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam9 = 1 AND " + Trim(MyTableName) + ".ItemHeight IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemHeight ELSE tbl_PurchasePriceItems_AddInfo.ItemHeight END, "
                    MySQLStr = MySQLStr + "ItemWeight = CASE WHEN MyParam10 = 2 OR "
                    MySQLStr = MySQLStr + "(MyParam10 = 1 AND " + Trim(MyTableName) + ".ItemWeight IS NOT NULL) "
                    MySQLStr = MySQLStr + "THEN " + Trim(MyTableName) + ".ItemWeight ELSE tbl_PurchasePriceItems_AddInfo.ItemWeight END "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo INNER JOIN "
                    MySQLStr = MySQLStr + Trim(MyTableName) + " ON "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo.SC01060 = " + Trim(MyTableName) + ".SC01060 CROSS JOIN "
                    MySQLStr = MySQLStr + Trim(MySysTableName) + " "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.PL01001 = N'" + Trim(MySuppCode) + "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----��������
                    Label3.Text = "�������� ����� ���������� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "INSERT INTO tbl_PurchasePriceItems_AddInfo "
                    MySQLStr = MySQLStr + "SELECT NEWID() AS ID, N'" + Trim(MySuppCode) + "' AS PL01001, " + Trim(MyTableName) + ".SC01060, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".CollectionName, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".WEBName, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".Description, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemGroupCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemSubGroupCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ManufacturerCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ManufacturerItemCode, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemLength, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemWidth, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemHeight, "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".ItemWeight "
                    MySQLStr = MySQLStr + "FROM " + Trim(MyTableName) + " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr + "(SELECT ID, PL01001, SC01060, CollectionName, WEBName, Description, ItemGroupCode, ItemSubGroupCode, ManufacturerCode, "
                    MySQLStr = MySQLStr + "ManufacturerItemCode, ItemLength, ItemWidth, ItemHeight, ItemWeight "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo AS tbl_PurchasePriceItems_AddInfo_1 "
                    MySQLStr = MySQLStr + "WHERE (PL01001 = N'" + Trim(MySuppCode) + "')) AS View_8 ON "
                    MySQLStr = MySQLStr + Trim(MyTableName) + ".SC01060 = View_8.SC01060 "
                    MySQLStr = MySQLStr + "WHERE (View_8.ID IS NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� ������ � �������� ������
                    Label3.Text = "����� ���� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01037 = tbl_PurchasePriceItems_AddInfo.ItemGroupCode "
                    MySQLStr = MySQLStr + "FROM tbl_PurchasePriceItems_AddInfo INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_PurchasePriceItems_AddInfo.PL01001 = SC010300.SC01058 AND "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo.SC01060 = SC010300.SC01060 "
                    MySQLStr = MySQLStr + "And tbl_PurchasePriceItems_AddInfo.ItemGroupCode <> SC010300.SC01037 "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemGroupCode IS NOT NULL) "
                    MySQLStr = MySQLStr + "AND (tbl_PurchasePriceItems_AddInfo.ItemGroupCode <> N'') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ����� � �������� ������
                    Label3.Text = "����� ����� � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01007 = tbl_PurchasePriceItems_AddInfo.ItemLength "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01007 <> tbl_PurchasePriceItems_AddInfo.ItemLength "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemLength IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ������ � �������� ������
                    Label3.Text = "����� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01008 = tbl_PurchasePriceItems_AddInfo.ItemWidth "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01008 <> tbl_PurchasePriceItems_AddInfo.ItemWidth "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemWidth IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ������ � �������� ������
                    Label3.Text = "����� ������ � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01009 = tbl_PurchasePriceItems_AddInfo.ItemHeight "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01009 <> tbl_PurchasePriceItems_AddInfo.ItemHeight "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemHeight IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� � �������� ������
                    Label3.Text = "����� ���� � �������� ������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "exec spp_System_SetBlock N'0000000009', 1 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "UPDATE SC010300 "
                    MySQLStr = MySQLStr + "SET SC01069 = tbl_PurchasePriceItems_AddInfo.ItemWeight "
                    MySQLStr = MySQLStr + "FROM SC010300 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 "
                    MySQLStr = MySQLStr + "And SC010300.SC01069 <> tbl_PurchasePriceItems_AddInfo.ItemWeight "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ItemWeight IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    MySQLStr = "exec spp_System_RemoveBlock "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� �������������
                    Label3.Text = "����� ���� �������������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "UPDATE tbl_ItemCard0300 "
                    MySQLStr = MySQLStr + "SET Manufacturer = tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "FROM tbl_ItemCard0300 INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_ItemCard0300.SC01001 = SC010300.SC01001 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 AND "
                    MySQLStr = MySQLStr + "tbl_ItemCard0300.Manufacturer <> tbl_PurchasePriceItems_AddInfo.ManufacturerCode "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ManufacturerCode IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----����� ���� ������ �������������
                    Label3.Text = "����� ���� ������ �������������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "UPDATE tbl_ItemCard0300 "
                    MySQLStr = MySQLStr + "SET ManufacturerItemCode = tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "FROM tbl_ItemCard0300 INNER JOIN "
                    MySQLStr = MySQLStr + "SC010300 ON tbl_ItemCard0300.SC01001 = SC010300.SC01001 INNER JOIN "
                    MySQLStr = MySQLStr + "tbl_PurchasePriceItems_AddInfo ON SC010300.SC01058 = tbl_PurchasePriceItems_AddInfo.PL01001 AND "
                    MySQLStr = MySQLStr + "SC010300.SC01060 = tbl_PurchasePriceItems_AddInfo.SC01060 AND "
                    MySQLStr = MySQLStr + "tbl_ItemCard0300.ManufacturerItemCode <> tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode "
                    MySQLStr = MySQLStr + "WHERE (tbl_PurchasePriceItems_AddInfo.ManufacturerItemCode IS NOT NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                Finally
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Try
                        MySQLStr = "DROP TABLE " & MySysTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try
                    Try
                        oWorkBook.Close(True)
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

    Private Function ServerChecks(ByVal MyTableName As String, ByVal MySysTableName As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� �������� ������ �� �������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� �����������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim WrkStr As String = ""
        Dim MySQLStr As String = ""
        Dim i As Integer

        MySQLStr = "exec spp_PurchasePriceItems_AddInfo_Check N'" + Trim(MyTableName) + "', N'" + Trim(MySysTableName) + "'"
        InitMyConn(False)
        Try
            MySQLAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MySQLAdapter.SelectCommand.CommandTimeout = 1200
            Dim builder As SqlClient.SqlCommandBuilder = New SqlClient.SqlCommandBuilder(MySQLAdapter)
            MySQLAdapter.Fill(MySQLDs)
        Catch ex As Exception
            MsgBox(ex.ToString)
            Me.Cursor = Cursors.Default
            Exit Function
        End Try

        '-----���������
        If MySQLDs.Tables(0).Rows.Count > 0 Or MySQLDs.Tables(1).Rows.Count > 0 _
            Or MySQLDs.Tables(2).Rows.Count > 0 Or MySQLDs.Tables(3).Rows.Count > 0 Then
            WrkStr = "����������� ���������� � Excel ����: " + Chr(13) + Chr(10)
        End If

        '-----���� ������ ������
        If MySQLDs.Tables(0).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "������������ ���� ����� �������: " + Chr(13) + Chr(10)
            WrkStr = WrkStr + "��� ������                    ������������ ��� ������ ������" + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(0).Rows.Count
                WrkStr = WrkStr + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(0).Rows(i).Item(0) + "                              ", 30) _
                    + MySQLDs.Tables(0).Rows(i).Item(1) + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If

        '-----���� ��������� ������
        If MySQLDs.Tables(1).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "������������ ���� �������� �������: " + Chr(13) + Chr(10)
            WrkStr = WrkStr + "��� ������                    ������������ ��� ��������� ������" + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(1).Rows.Count
                WrkStr = WrkStr + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(1).Rows(i).Item(0) + "                              ", 30) _
                    + MySQLDs.Tables(1).Rows(i).Item(1) + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If

        '-----���� �������������
        If MySQLDs.Tables(2).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "������������ ���� �������������: " + Chr(13) + Chr(10)
            WrkStr = WrkStr + "��� ������                    ������������ ��� �������������" + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(2).Rows.Count
                WrkStr = WrkStr + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(2).Rows(i).Item(0) + "                              ", 30) _
                    + MySQLDs.Tables(2).Rows(i).Item(1).ToString + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If

        '-----���� ������ � ���������
        If MySQLDs.Tables(3).Rows.Count > 0 Then
            WrkStr = WrkStr + Chr(13) + Chr(10)
            WrkStr = WrkStr + "������������� ���� ������ � ��������� ������: " + Chr(13) + Chr(10)
            WrkStr = WrkStr + "��� ������                    ��� ������ ������ ��� ��������� ������" + Chr(13) + Chr(10)
            i = 0
            While i < MySQLDs.Tables(3).Rows.Count
                WrkStr = WrkStr + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(3).Rows(i).Item(0) + "                              ", 30) _
                    + Microsoft.VisualBasic.Strings.Left(MySQLDs.Tables(3).Rows(i).Item(1) + "                  ", 18) _
                    + MySQLDs.Tables(3).Rows(i).Item(2) + Chr(13) + Chr(10)
                i = i + 1
            End While
        End If

        If WrkStr.Length > 0 Then
            MyErrorMessage = New ErrorMessage
            MyErrorMessage.TextBox1.Text = WrkStr
            MyErrorMessage.ShowDialog()
            ServerChecks = False
        Else
            ServerChecks = True
        End If
    End Function

End Class