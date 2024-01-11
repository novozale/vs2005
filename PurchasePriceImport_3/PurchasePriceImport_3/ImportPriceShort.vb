Imports System.Guid

Public Class ImportPriceShort

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ImportPriceShort_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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
        MsgBox("��������� �������� ����� - ����� �� ������� ���������.", vbOKOnly, "��������!")
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ����� ����� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyVersion As String                     '������ ���������
        Dim MySuppCode As String                    '��� ����������
        Dim cn As OleDbConnection                   '������ ���������� � OLE
        Dim connStr As String                       '������ ����������� � Excel
        Dim FirstExcelSheetName As String           '�������� ������� ����� Excel
        Dim myds As DataSet                         'Excel dataset
        Dim MyDBL As Double                         '��� ��������
        Dim MyStr As String                         '��� ��������
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim mycount As Integer
        Dim MyTableName As String                   '��� ��������� �������
        Dim MyGuid As String                          '
        Dim MyRez As MsgBoxResult                   '��������� ������

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_PurchasePriceHistory_Tmp_" + MyGuid

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
                            MySQLStr = MySQLStr & "WHERE (Name = N'������ ����� ����� �� �������') "
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

                    '---���������, ��� ���������� ��� ���������� � Scala � ��� �� ����������
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

                    '---��������� ������������ ������ � Excel
                    '-----������������� ����
                    MySQLStr = "SELECT F1 FROM [" & FirstExcelSheetName & "A5:A] group by F1 having(count(F1) > 1)"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    If myds.Tables(0).Rows.Count > 0 Then
                        MsgBox("� ����� ��������� " & myds.Tables(0).Rows.Count & " ������������� ������� ����� ������� ����������. �������������� ������� ""���������� �������������"" � Excel, ��������� � ������� ������ ���� ")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If

                    '-----������������ ��������� ������ � Excel
                    MySQLStr = "SELECT * FROM [" & FirstExcelSheetName & "A5:G] where(F1 <> """")"
                    myds = GetExcelDataSet(cn, MySQLStr)
                    '-----������������ ���������
                    mycount = 0
                    While mycount < myds.Tables(0).Rows.Count
                        '-----���������� ���� ������ ����������
                        If Trim(myds.Tables(0).Rows(mycount).Item(0).ToString) = "" Then
                            MsgBox("������ " & CStr(mycount + 5) & " �� ������� ��� ������ ����������")
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
                        '-----���� �������� � ����
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(2)
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ���� �������� � ����")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----����������� ������ ������
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(3)
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ����������� ������ ������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----������� ����� ����������
                        Try
                            MyDBL = myds.Tables(0).Rows(mycount).Item(4)
                        Catch ex As Exception
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ������� ����� ����������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End Try
                        '-----��� ������
                        MyStr = Trim(myds.Tables(0).Rows(mycount).Item(5).ToString)
                        If MyStr = "" Or (MyStr <> "0" And MyStr <> "00" And MyStr <> "1" And MyStr <> "12" And MyStr <> "4" _
                            And MyStr <> "6" And MyStr <> "11" And MyStr <> "13") Then
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� ������� ��� ������")
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        '-----������ ���������� �����
                        MyStr = Trim(myds.Tables(0).Rows(mycount).Item(6).ToString)
                        If (MyStr <> "A" And MyStr <> "F") Then
                            MsgBox("������ " & CStr(mycount + 5) & " ����������� �������� ������ ���������� ����� �������� - ������ ���� ��������� ���������� ����� A ��� F.")
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
                    MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) NOT NULL, "
                    MySQLStr = MySQLStr & "[Price] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[LT] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[MinQTY] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[BasePrice] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[Curr] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[LifeCycleStage] [nvarchar](3) NOT NULL, "
                    MySQLStr = MySQLStr & "CONSTRAINT [PK_" & MyTableName & "] PRIMARY KEY CLUSTERED "
                    MySQLStr = MySQLStr & "( "
                    MySQLStr = MySQLStr & "[SC01060] Asc "
                    MySQLStr = MySQLStr & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ �� ��������� �������
                    InitMyConn(False)
                    MySQLStr = "SELECT SC01060, Price, LT, MinQTY, BasePrice, Curr, LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    'MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory_Tmp "
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

                    '=====��������� ���������� ����� �� ����� ���������� ������� ����� ��� ������� ����������, ���� �� ���===============
                    '---��������� ����������� ����
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������� ����� ����� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    
                    MySQLStr = "SELECT MAX(DateFrom) AS DateFrom, "
                    MySQLStr = MySQLStr & "MAX(DateTo) AS DateTo, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar, DATEPART(mm, GETDATE())) + '/' + CONVERT(nvarchar, DATEPART(yyyy, GETDATE())), 103) AS DateCurr "
                    MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        '---����� ����� �� ����� ���������� �� �����������
                    Else
                        If Not IsDBNull(Declarations.MyRec.Fields("DateFrom").Value) Then
                            '---����� ����� �� ����� ���������� �����������
                            If Declarations.MyRec.Fields("DateFrom").Value = Declarations.MyRec.Fields("DateCurr").Value Then
                                '---� ���� ���� ����� ��� ����������� - ������� ���������� ����� �� ���� ����
                                MySQLStr = "DELETE FROM tbl_PurchasePriceHistory "
                                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "') AND "
                                MySQLStr = MySQLStr & "(DateFrom = CONVERT(DATETIME,'" & Declarations.MyRec.Fields("DateFrom").Value & "',103))"
                                trycloseMyRec()
                                InitMyConn(False)
                                Declarations.MyConn.Execute(MySQLStr)
                            Else
                                '---��������� ������ ����� ������� �����
                                If Declarations.MyRec.Fields("DateTo").Value = CDate("31/12/9999") Then
                                    MySQLStr = "Update tbl_PurchasePriceHistory "
                                    MySQLStr = MySQLStr & "SET DateTo = CONVERT(DATETIME,'" & Declarations.MyRec.Fields("DateCurr").Value & "', 103) "
                                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "') AND "
                                    MySQLStr = MySQLStr & "(DateTo = CONVERT(DATETIME,'31/12/9999',103))"
                                    trycloseMyRec()
                                    InitMyConn(False)
                                    Declarations.MyConn.Execute(MySQLStr)
                                End If
                            End If
                        End If
                    End If

                    '==============================================������������ ������ ����� �����============================================
                    '----------------------------������� ��������
                    Label3.Text = "������������ ������ ����� ����� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----���������� �� - ��������� �����
                    MySQLStr = "INSERT INTO tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "SELECT NEWID() AS Expr1, N'" & MySuppCode & "' AS Expr2, '' AS Expr3, " & MyTableName & ".Price, " & MyTableName & ".Curr, DATEADD(DAY, "
                    MySQLStr = MySQLStr & "DATEDIFF(DAY, 0, CURRENT_TIMESTAMP), 0) AS Expr4, CONVERT(DATETIME, '9999-12-31 00:00:00', 102) AS Expr5, "
                    MySQLStr = MySQLStr & MyTableName & ".LT, " & MyTableName & ".MinQTY, " & MyTableName & ".SC01060, "
                    MySQLStr = MySQLStr & MyTableName & ".BasePrice, " & MyTableName & ".LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "(SELECT SC01060 "
                    MySQLStr = MySQLStr & "FROM dbo.SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01058 = N'" & MySuppCode & "')) AS  View_2 ON " & MyTableName & ".SC01060 = View_2.SC01060 "
                    MySQLStr = MySQLStr & "WHERE (View_2.SC01060 Is NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)


                    '-----���������� ��������� �����
                    MySQLStr = "INSERT INTO tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "SELECT NEWID() AS Expr1, N'" & MySuppCode & "' AS Expr2, View_2.SC01001, " & MyTableName & ".Price, " & MyTableName & ".Curr, DATEADD(DAY, "
                    MySQLStr = MySQLStr & "DATEDIFF(DAY, 0, CURRENT_TIMESTAMP), 0) AS Expr4, CONVERT(DATETIME, '9999-12-31 00:00:00', 102) AS Expr5, "
                    MySQLStr = MySQLStr & MyTableName & ".LT, " & MyTableName & ".MinQTY, " & MyTableName & ".SC01060, "
                    MySQLStr = MySQLStr & MyTableName & ".BasePrice, " & MyTableName & ".LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "(SELECT SC01001, SC01060 "
                    MySQLStr = MySQLStr & "FROM SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01058 = N'" & MySuppCode & "')) AS View_2 ON " & MyTableName & ".SC01060 = View_2.SC01060 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '==============================================�������� ��������� ����� � 2 � ����� ���==========================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ��������� ������ � 2 � ����� ���. "
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    If CheckPriceChanging(MySuppCode) = False Then
                        MyRez = MsgBox("� ����� - ����� ������� ���������� ���� ����� ��������, ������� � 2 � ����� ��� ���������� �� ������. ���������� ������ ����� - ����� �� �������? ", MsgBoxStyle.YesNo, "��������!")
                    Else
                        MyRez = MsgBox("���������� ������ ����� - ����� �� ������� ������? ����� ������ ����� �������� �����.", MsgBoxStyle.YesNo, "��������!")
                    End If


                    '==============================================������ ��������� ����� ����� �� �������============================================
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

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Libre Office ����� ����� �� ������� �� 1 ����������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                      'SQL ������
        Dim MyVersion As String                     '������ ���������
        Dim MySuppCode As String                    '��� ����������
        Dim MyTableName As String                   '��� ��������� �������
        Dim MyDBL As Double                         '��� ��������
        Dim MyStr As String                         '��� ��������
        Dim MyGuid As String                        '
        Dim MySQLAdapter As SqlClient.SqlDataAdapter '��� ��������� �������
        Dim MySQLDs As New DataSet                  'SQL dataset
        Dim MyRez As MsgBoxResult                   '��������� ������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        MyGuid = Replace(Guid.NewGuid.ToString, "-", "")
        MyTableName = "tbl_PurchasePriceHistory_Tmp_" + MyGuid

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
                        MySQLStr = MySQLStr & "WHERE (Name = N'������ ����� ����� �� �������') "
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

                    '---��������� ������������ ������ � Excel
                    '-----������������� ����
                    oSheet.unprotect("!pass2022")

                    Dim args() As Object
                    ReDim args(0)
                    args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
                    args(0).Name = "ToPoint"
                    args(0).Value = "$A$5:$H$300000"
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
                            If myarr(i)(0).ToString = "" Then
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

                    '-----������������ ��������� ������ � Libre Ofice
                    StopFlag = False
                    For j As Integer = 0 To 300000 Step MyStep
                        srange = oSheet.getCellRangeByName("A" & 6 + j & ":M" & 6 + j + MyStep)
                        myarr = srange.DataArray
                        For i As Integer = 0 To MyStep
                            If myarr(i)(0).ToString.Equals("") Then
                                StopFlag = True
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

                                '-----���� �������� � ����
                                Try
                                    MyDBL = myarr(i)(2)
                                Catch ex As Exception
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ���� �������� � ����")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End Try

                                '-----����������� ������ ������
                                Try
                                    MyDBL = myarr(i)(3)
                                Catch ex As Exception
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ����������� ������ ������")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End Try

                                '-----������� ����� ����������
                                Try
                                    MyDBL = myarr(i)(4)
                                Catch ex As Exception
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ������� ����� ����������")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End Try

                                '-----��� ������
                                MyStr = Trim(myarr(i)(5).ToString)
                                If MyStr = "" Or (MyStr <> "0" And MyStr <> "00" And MyStr <> "1" And MyStr <> "12" And MyStr <> "4" _
                                    And MyStr <> "6" And MyStr <> "11" And MyStr <> "13") Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� ������� ��� ������")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
                                End If

                                '-----������ ���������� �����
                                MyStr = Trim(myarr(i)(6).ToString.ToString)
                                If (MyStr <> "A" And MyStr <> "F") Then
                                    MsgBox("������ " & CStr(i + 5) & " ����������� �������� ������ ���������� ����� �������� - ������ ���� ��������� ���������� ����� A ��� F.")
                                    Me.Cursor = Cursors.Default
                                    oWorkBook.Close(True)
                                    Exit Sub
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

                    '-----�������� ��������� �������
                    Try
                        MySQLStr = "DROP TABLE " & MyTableName & " "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    Catch ex As Exception
                    End Try

                    MySQLStr = "CREATE TABLE [dbo].[" & MyTableName & "]( "
                    MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) NOT NULL, "
                    MySQLStr = MySQLStr & "[Price] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[LT] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[MinQTY] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[BasePrice] [numeric](20, 8) NOT NULL, "
                    MySQLStr = MySQLStr & "[Curr] [nvarchar](50) NOT NULL, "
                    MySQLStr = MySQLStr & "[LifeCycleStage] [nvarchar](3) NOT NULL, "
                    MySQLStr = MySQLStr & "CONSTRAINT [PK_" & MyTableName & "] PRIMARY KEY CLUSTERED "
                    MySQLStr = MySQLStr & "( "
                    MySQLStr = MySQLStr & "[SC01060] Asc "
                    MySQLStr = MySQLStr & ")WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY] "
                    MySQLStr = MySQLStr & ") ON [PRIMARY] "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----������ �� ��������� �������
                    InitMyConn(False)
                    MySQLStr = "SELECT SC01060, Price, LT, MinQTY, BasePrice, Curr, LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " "
                    'MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory_Tmp "
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

                    '-----������� ������ �� Libre Office � SQL dataset
                    Dim dt As DataTable
                    Dim dr As DataRow

                    dt = MySQLDs.Tables(0)
                    StopFlag = False
                    For j As Integer = 0 To 300000 Step MyStep
                        srange = oSheet.getCellRangeByName("A" & 5 + j & ":M" & 4 + j + MyStep)
                        myarr = srange.DataArray
                        For i As Integer = 0 To MyStep - 1
                            If myarr(i)(0).ToString.Equals("") Then
                                StopFlag = True
                                Exit For
                            Else
                                dr = dt.NewRow
                                dr.Item(0) = myarr(i)(0).ToString
                                dr.Item(1) = myarr(i)(1)
                                dr.Item(2) = myarr(i)(2)
                                dr.Item(3) = myarr(i)(3)
                                dr.Item(4) = myarr(i)(4)
                                dr.Item(5) = myarr(i)(5)
                                dr.Item(6) = myarr(i)(6)
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

                    '=====��������� ���������� ����� �� ����� ���������� ������� ����� ��� ������� ����������, ���� �� ���===============
                    '---��������� ����������� ����
                    '----------------------------������� ��������
                    Label3.Text = "�������� ������� ����� ����� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    MySQLStr = "SELECT MAX(DateFrom) AS DateFrom, "
                    MySQLStr = MySQLStr & "MAX(DateTo) AS DateTo, "
                    MySQLStr = MySQLStr & "CONVERT(datetime, CONVERT(nvarchar, DATEPART(dd, GETDATE())) + '/' + CONVERT(nvarchar, DATEPART(mm, GETDATE())) + '/' + CONVERT(nvarchar, DATEPART(yyyy, GETDATE())), 103) AS DateCurr "
                    MySQLStr = MySQLStr & "FROM tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "')"
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        '---����� ����� �� ����� ���������� �� �����������
                    Else
                        If Not IsDBNull(Declarations.MyRec.Fields("DateFrom").Value) Then
                            '---����� ����� �� ����� ���������� �����������
                            If Declarations.MyRec.Fields("DateFrom").Value = Declarations.MyRec.Fields("DateCurr").Value Then
                                '---� ���� ���� ����� ��� ����������� - ������� ���������� ����� �� ���� ����
                                MySQLStr = "DELETE FROM tbl_PurchasePriceHistory "
                                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "') AND "
                                MySQLStr = MySQLStr & "(DateFrom = CONVERT(DATETIME,'" & Declarations.MyRec.Fields("DateFrom").Value & "',103))"
                                trycloseMyRec()
                                InitMyConn(False)
                                Declarations.MyConn.Execute(MySQLStr)
                            Else
                                '---��������� ������ ����� ������� �����
                                If Declarations.MyRec.Fields("DateTo").Value = CDate("31/12/9999") Then
                                    MySQLStr = "Update tbl_PurchasePriceHistory "
                                    MySQLStr = MySQLStr & "SET DateTo = CONVERT(DATETIME,'" & Declarations.MyRec.Fields("DateCurr").Value & "', 103) "
                                    MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & MySuppCode & "') AND "
                                    MySQLStr = MySQLStr & "(DateTo = CONVERT(DATETIME,'31/12/9999',103))"
                                    trycloseMyRec()
                                    InitMyConn(False)
                                    Declarations.MyConn.Execute(MySQLStr)
                                End If
                            End If
                        End If
                    End If

                    '==============================================������������ ������ ����� �����============================================
                    '----------------------------������� ��������
                    Label3.Text = "������������ ������ ����� ����� �� �������"
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '-----���������� �� - ��������� �����
                    MySQLStr = "INSERT INTO tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "SELECT NEWID() AS Expr1, N'" & MySuppCode & "' AS Expr2, '' AS Expr3, " & MyTableName & ".Price, " & MyTableName & ".Curr, DATEADD(DAY, "
                    MySQLStr = MySQLStr & "DATEDIFF(DAY, 0, CURRENT_TIMESTAMP), 0) AS Expr4, CONVERT(DATETIME, '9999-12-31 00:00:00', 102) AS Expr5, "
                    MySQLStr = MySQLStr & MyTableName & ".LT, " & MyTableName & ".MinQTY, " & MyTableName & ".SC01060, "
                    MySQLStr = MySQLStr & MyTableName & ".BasePrice, " & MyTableName & ".LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " LEFT OUTER JOIN "
                    MySQLStr = MySQLStr & "(SELECT SC01060 "
                    MySQLStr = MySQLStr & "FROM dbo.SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01058 = N'" & MySuppCode & "')) AS  View_2 ON " & MyTableName & ".SC01060 = View_2.SC01060 "
                    MySQLStr = MySQLStr & "WHERE (View_2.SC01060 Is NULL) "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '-----���������� ��������� �����
                    MySQLStr = "INSERT INTO tbl_PurchasePriceHistory "
                    MySQLStr = MySQLStr & "SELECT NEWID() AS Expr1, N'" & MySuppCode & "' AS Expr2, View_2.SC01001, " & MyTableName & ".Price, " & MyTableName & ".Curr, DATEADD(DAY, "
                    MySQLStr = MySQLStr & "DATEDIFF(DAY, 0, CURRENT_TIMESTAMP), 0) AS Expr4, CONVERT(DATETIME, '9999-12-31 00:00:00', 102) AS Expr5, "
                    MySQLStr = MySQLStr & MyTableName & ".LT, " & MyTableName & ".MinQTY, " & MyTableName & ".SC01060, "
                    MySQLStr = MySQLStr & MyTableName & ".BasePrice, " & MyTableName & ".LifeCycleStage "
                    MySQLStr = MySQLStr & "FROM " & MyTableName & " INNER JOIN "
                    MySQLStr = MySQLStr & "(SELECT SC01001, SC01060 "
                    MySQLStr = MySQLStr & "FROM SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01058 = N'" & MySuppCode & "')) AS View_2 ON " & MyTableName & ".SC01060 = View_2.SC01060 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '==============================================�������� ��������� ����� � 2 � ����� ���==========================================
                    '----------------------------������� ��������
                    Label3.Text = "�������� ��������� ������ � 2 � ����� ���. "
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    If CheckPriceChangingLO(MySuppCode) = False Then
                        MyRez = MsgBox("� ����� - ����� ������� ���������� ���� ����� ��������, ������� � 2 � ����� ��� ���������� �� ������. ���������� ������ ����� - ����� �� �������? ", MsgBoxStyle.YesNo, "��������!")
                    Else
                        MyRez = MsgBox("���������� ������ ����� - ����� �� ������� ������? ����� ������ ����� �������� �����.", MsgBoxStyle.YesNo, "��������!")
                    End If

                    '==============================================������ ��������� ����� ����� �� �������============================================
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

    Private Function CheckPriceChanging(ByVal MySuppCode As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ���������� �� �������, ���������� ����� �� ������� ���������� � 2 � ����� ���  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyObj As Object       'Excel
        Dim MyWRKBook As Object   '�����
        Dim StrNum As Double      '����� ������

        MySQLStr = "exec spp_PurchasePrices_History_CheckPrice N'" & MySuppCode & "' "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---���������� ��� - ��� OK
            CheckPriceChanging = True
            trycloseMyRec()
        Else
            '---���� ���������� - ������� � Excel
            CheckPriceChanging = False

            MyObj = CreateObject("Excel.Application")
            MyObj.SheetsInNewWorkbook = 1
            MyWRKBook = MyObj.Workbooks.Add

            UploadHeader(MyWRKBook, MySuppCode)
            StrNum = 5
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & StrNum) = "'" & Declarations.MyRec.Fields("Code").Value
                MyWRKBook.ActiveSheet.Range("B" & StrNum) = "'" & Declarations.MyRec.Fields("SuppCode").Value
                MyWRKBook.ActiveSheet.Range("C" & StrNum) = Declarations.MyRec.Fields("NewPrice").Value
                MyWRKBook.ActiveSheet.Range("D" & StrNum) = Declarations.MyRec.Fields("NewCurr").Value
                MyWRKBook.ActiveSheet.Range("E" & StrNum) = Declarations.MyRec.Fields("OldPrice").Value
                MyWRKBook.ActiveSheet.Range("F" & StrNum) = Declarations.MyRec.Fields("OldCurr").Value
                Declarations.MyRec.MoveNext()
                StrNum = StrNum + 1
            End While
            trycloseMyRec()

            MyWRKBook.ActiveSheet.Range("A1").Select()
            MyObj.Application.Visible = True
            MyObj = Nothing
        End If
        CheckPriceChanging = True
    End Function

    Private Function CheckPriceChangingLO(ByVal MySuppCode As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� � Libre Office ���������� �� �������, ���������� ����� �� ������� ���������� � 2 � ����� ���  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim StrNum As Double      '����� ������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object

        MySQLStr = "exec spp_PurchasePrices_History_CheckPrice N'" & MySuppCode & "' "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---���������� ��� - ��� OK
            CheckPriceChangingLO = True
            trycloseMyRec()
        Else
            '---���� ���������� - ������� � Libre Office
            CheckPriceChangingLO = False

            LOSetNotation(0)
            oServiceManager = CreateObject("com.sun.star.ServiceManager")
            oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
            oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
            Dim arg(1)
            arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
            arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
            oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
            oSheet = oWorkBook.getSheets().getByIndex(0)
            oFrame = oWorkBook.getCurrentController.getFrame

            UploadHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, MySuppCode)

            StrNum = 5
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & StrNum).String = Declarations.MyRec.Fields("Code").Value
                oSheet.getCellRangeByName("B" & StrNum).String = Declarations.MyRec.Fields("SuppCode").Value
                oSheet.getCellRangeByName("C" & StrNum).Value = Declarations.MyRec.Fields("NewPrice").Value
                oSheet.getCellRangeByName("D" & StrNum).Value = Declarations.MyRec.Fields("NewCurr").Value
                oSheet.getCellRangeByName("E" & StrNum).Value = Declarations.MyRec.Fields("OldPrice").Value
                oSheet.getCellRangeByName("F" & StrNum).Value = Declarations.MyRec.Fields("OldCurr").Value

                Declarations.MyRec.MoveNext()
                StrNum = StrNum + 1
            End While
            trycloseMyRec()

            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = "$A$1"
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
            oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
        End If
    End Function


    Private Function UploadHeader(ByVal MyWRKBook As Object, ByVal MySuppCode As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Excel ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("B1") = "��������� ���������� ���� ����� ��� � 2 ���� "
        MyWRKBook.ActiveSheet.Range("B2") = "��������� " & MySuppCode & " ���� ������: " & Now
        MyWRKBook.ActiveSheet.Range("B1:B2").Select()
        MyWRKBook.ActiveSheet.Range("B1:B2").Font.Bold = True

        '--- � ������� �����
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20

        MyWRKBook.ActiveSheet.Range("A4") = "��� ������"
        MyWRKBook.ActiveSheet.Range("B4") = "��� ������ ����������"
        MyWRKBook.ActiveSheet.Range("C4") = "����� ����"
        MyWRKBook.ActiveSheet.Range("D4") = "����� ������"
        MyWRKBook.ActiveSheet.Range("E4") = "������ ����"
        MyWRKBook.ActiveSheet.Range("F4") = "������ ������"

        MyWRKBook.ActiveSheet.Range("A4:F4").Select()
        MyWRKBook.ActiveSheet.Range("A4:F4").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A4:F4").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A4:F4").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:F4").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:F4").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:F4").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:F4").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A4:F4").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

    End Function

    Private Function UploadHeaderLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal MySuppCode As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� � Libre Office ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 4000
        oSheet.getColumns().getByName("B").Width = 4000
        oSheet.getColumns().getByName("C").Width = 4000
        oSheet.getColumns().getByName("D").Width = 4000
        oSheet.getColumns().getByName("E").Width = 4000
        oSheet.getColumns().getByName("F").Width = 4000

        oSheet.getCellRangeByName("B1").String = "��������� ���������� ���� ����� ��� � 2 ����"
        oSheet.getCellRangeByName("B2").String = "��������� " & MySuppCode & " ���� ������: " & Now
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B1:B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B1:B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B1:B2", 11)
        oSheet.getCellRangeByName("B2").VertJustify = 2

        oSheet.getCellRangeByName("A4").String = "��� ������"
        oSheet.getCellRangeByName("B4").String = "��� ������ ����������"
        oSheet.getCellRangeByName("C4").String = "����� ����"
        oSheet.getCellRangeByName("D4").String = "����� ������"
        oSheet.getCellRangeByName("E4").String = "������ ����"
        oSheet.getCellRangeByName("F4").String = "������ ������"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4:F4", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4:F4", 11)
        oSheet.getCellRangeByName("A4:F4").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A4:F4").TopBorder = LineFormat
        oSheet.getCellRangeByName("A4:F4").RightBorder = LineFormat
        oSheet.getCellRangeByName("A4:F4").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A4:F4").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A4:F4").VertJustify = 2
        oSheet.getCellRangeByName("A4:F4").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A4:F4")
    End Function

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
End Class