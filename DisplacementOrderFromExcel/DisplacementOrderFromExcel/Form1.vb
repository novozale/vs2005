Public Class Form1
    Structure FinRez
        Public MyRezStr As String
        Public MyRelocOrderNum As String
    End Structure

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �� ���������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyDs As New DataSet                       '

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            Declarations.ScalaDate = CDate(Scala.ActiveFrame.Parent.ScalaDate)


            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                MsgBox("�� ������ ��� ��������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Application.Exit()
            Else
                Declarations.SalesmanCode = Declarations.MyRec.Fields("SC").Value
                Declarations.SalesmanName = Declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
            End If
        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try

        '---���������� ����� (ComboBox)
        BuildWHListFrom()
        BuildWHListTo()

        DateTimePicker1.Value = Today
        DateTimePicker2.Value = Today

        CheckBox1.Checked = False
        CheckBox2.Checked = False
    End Sub

    Private Sub BuildWHListFrom()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � Combobox ������ �������, � ������� �������� �����������, � ����� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        'MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "SC23002" '��� �� ��� ����� ������������
            ComboBox1.ValueMember = "SC23001"   '��� �� ��� ����� ���������
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub BuildWHListTo()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� � Combobox ������ �������, �� ������� �������� �����������, � ����� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '


        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE(SC23001 <> N'" & ComboBox1.SelectedValue & "') "
        'MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001"
        InitMyConn(False)
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "SC23002" '��� �� ��� ����� ������������
            ComboBox2.ValueMember = "SC23001"   '��� �� ��� ����� ���������
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ������, � �������� ����� �����������,
        '/// ����� � Combobox ������ �������, �� ������� �������� �����������, � ����� �������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        BuildWHListTo()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ �������� ������ �� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckCorrect() = True Then
            If My.Settings.UseOffice = "LibreOffice" Then
                CreateDisplacementOrderLO()
            Else
                CreateDisplacementOrder()
            End If

        End If
    End Sub

    Private Function CheckCorrect() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������ ��������� ��������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        '---���� ������
        If DateTimePicker1.Value < Today Then
            MsgBox("������� ���������� ���� �������������� �������� ������ (�� ������ ���� ������ ������������ ���)", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckCorrect = False
            Exit Function
        End If

        If DateTimePicker2.Value < Today Then
            MsgBox("������� ���������� ���� ��������������� ��������� ������ (�� ������ ���� ������ ������������ ���)", MsgBoxStyle.Critical, "��������!")
            DateTimePicker2.Select()
            CheckCorrect = False
            Exit Function
        End If

        If DateTimePicker1.Value > DateTimePicker2.Value Then
            MsgBox("���� ��������������� ��������� ������ �� ������ ���� ������ ���� �������������� �������� ������)", MsgBoxStyle.Critical, "��������!")
            DateTimePicker1.Select()
            CheckCorrect = False
            Exit Function
        End If

        '---�� �� ������� ����� �� ����������� �� ������ ������������� ����� �� ����� ������������� �����
        If IsRawMaterialsWH(ComboBox1.SelectedValue) = True And IsRawMaterialsWH(ComboBox2.SelectedValue) = True Then
            MsgBox("����� �������� " & ComboBox1.SelectedValue & " � ����� ���������� " & ComboBox2.SelectedValue & " �������� �������� ������������� �����. ������ ����� �� ����������� � ������ ������ ������������� ����� �� ������ ������.", MsgBoxStyle.Critical, "��������!")
            CheckCorrect = False
            Exit Function
        End If

        CheckCorrect = True

    End Function

    Private Sub CreateDisplacementOrder()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ����������� �� ������ ������ �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim i As Double                             '������� �����
        Dim MySQLStr As String
        Dim MyProductCode As String                 '��� ������
        Dim MyQTY As Double                         '������������ ����������

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Try
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---�������� ������ ��������� �������
                    MySQLStr = "IF exists(select * from tempdb..sysobjects where "
                    MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
                    MySQLStr = MySQLStr & "and xtype = N'U') "
                    MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---�������� ����� ��������� �������
                    MySQLStr = "CREATE TABLE #_MyOrder( "
                    MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                '--��� ������ � Scala
                    MySQLStr = MySQLStr & "[QTY] float, "                              '--����������
                    MySQLStr = MySQLStr & "[RestQTY] float  "                          '--������� - �������������� ����������
                    MySQLStr = MySQLStr & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    appXLSRC = CreateObject("Excel.Application")
                    appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                    i = 2 '---�������� �� 2 ������
                    While Not appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value = Nothing
                        MyProductCode = Trim(appXLSRC.Worksheets(1).Range("A" & CStr(i)).Value.ToString)
                        If appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value = Nothing Then
                            Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & ". �� ������� ���������� ��� �����������. ��������� ������������ ����������, ���������� � Excel.")
                        Else
                            Try
                                MyQTY = CDbl(appXLSRC.Worksheets(1).Range("B" & CStr(i)).Value.ToString)
                            Catch ex As Exception
                                appXLSRC.DisplayAlerts = 0
                                appXLSRC.Workbooks.Close()
                                appXLSRC.DisplayAlerts = 1
                                appXLSRC.Quit()
                                appXLSRC = Nothing
                                Me.Cursor = Cursors.Default
                                Me.Refresh()
                                System.Windows.Forms.Application.DoEvents()
                                MsgBox("������ " & CStr(i) & ". ��� ������ " & MyProductCode & ". ����������� ������� ���������� ��� �����������. " & ex.Message, MsgBoxStyle.Critical, "��������!")
                                Exit Sub
                            End Try
                        End If
                        '---���� �� � Scala ������ � ����� �����
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            '--������ � Scala ���
                            Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " � Scala �����������. ��������� ������������ ����������, ���������� � Excel.")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            If (CheckBox2.Checked = True) Or (CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) <> "02" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "03" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "04" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "05" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "06")) Then
                                '---��������� �� ��������� �������
                                '---������� ��������� - �����, ����� ��� ��� ���� �� ��������� �������
                                MySQLStr = "SELECT COUNT(*) AS CC "
                                MySQLStr = MySQLStr & "FROM #_MyOrder "
                                MySQLStr = MySQLStr & "WHERE (ItemCode = N'" & MyProductCode & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                    trycloseMyRec()
                                    '---� ��� �������� - �� ���������� ����������� ��������
                                    MySQLStr = "SELECT SC01066 "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Trim(Declarations.MyRec.Fields("SC01066").Value.ToString) = "8" Then
                                        trycloseMyRec()
                                        Throw (New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " �������� �����������. ��� ������ ������ ��������� ����������� ������ ���������� ������."))
                                        trycloseMyRec()
                                    Else
                                        trycloseMyRec()
                                        MySQLStr = "INSERT INTO #_MyOrder "
                                        MySQLStr = MySQLStr & "(ItemCode, QTY, RestQTY) "
                                        MySQLStr = MySQLStr & "VALUES (N'" & MyProductCode & "', "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
                                        InitMyConn(False)
                                        Declarations.MyConn.Execute(MySQLStr)
                                    End If
                                Else
                                    trycloseMyRec()
                                    Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " ������������ � Excel ����� ��� � ����� ������. ��� ������ �� ����������� ����� ��������� � Excel ������ 1 ���, ��� ������������.")
                                End If
                            ElseIf CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) = "02" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "03" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "04" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "05" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "06") Then
                                '---��������, ��� ������ �� �������
                                MsgBox("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " �������� ��������� ���������� � �� ����� ������� � ����� �� �����������.", MsgBoxStyle.Critical, "��������!")
                            End If
                        End If
                        i = i + 1
                    End While

                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---������ �������� ��������� ������������ ������ �� �����������
                    '---����������
                    '--SetBlock("0000000009") --���������� � �������� ���������

                    '---� ����� ���������� � ����
                    ResultOutput(ExecSppCreateRelocOrder())

                    'MsgBox("��������� �������� ������ �� ����������� ���������.", MsgBoxStyle.OkOnly, "��������!")
                Catch ex As Exception
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    MsgBox(ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
            End If
        End If
    End Sub

    Private Sub CreateDisplacementOrderLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ����������� �� ������ ������ �� LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim i As Double                             '������� �����
        Dim MySQLStr As String
        Dim MyProductCode As String                 '��� ������
        Dim MyQTY As Double                         '������������ ����������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Try
                    Me.Cursor = Cursors.WaitCursor
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---�������� ������ ��������� �������
                    MySQLStr = "IF exists(select * from tempdb..sysobjects where "
                    MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyOrder') "
                    MySQLStr = MySQLStr & "and xtype = N'U') "
                    MySQLStr = MySQLStr & "DROP TABLE #_MyOrder "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

                    '---�������� ����� ��������� �������
                    MySQLStr = "CREATE TABLE #_MyOrder( "
                    MySQLStr = MySQLStr & "[ItemCode] [nvarchar](35), "                '--��� ������ � Scala
                    MySQLStr = MySQLStr & "[QTY] float, "                              '--����������
                    MySQLStr = MySQLStr & "[RestQTY] float  "                          '--������� - �������������� ����������
                    MySQLStr = MySQLStr & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)

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

                    i = 2 '---�������� �� 2 ������
                    While oSheet.getCellRangeByName("A" & i).String.Equals("") = False
                        MyProductCode = Trim(oSheet.getCellRangeByName("A" & i).String)
                        If oSheet.getCellRangeByName("B" & i).Value = 0 Then
                            Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & ". �� ������� ���������� ��� �����������. ��������� ������������ ����������, ���������� � Excel.")
                        Else
                            Try
                                MyQTY = CDbl(oSheet.getCellRangeByName("B" & i).Value)
                            Catch ex As Exception
                                oWorkBook.Close(True)
                                Me.Cursor = Cursors.Default
                                Me.Refresh()
                                System.Windows.Forms.Application.DoEvents()
                                MsgBox("������ " & CStr(i) & ". ��� ������ " & MyProductCode & ". ����������� ������� ���������� ��� �����������. " & ex.Message, MsgBoxStyle.Critical, "��������!")
                                Exit Sub
                            End Try
                        End If

                        '---���� �� � Scala ������ � ����� �����
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.Fields("CC").Value = 0 Then
                            '--������ � Scala ���
                            Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " � Scala �����������. ��������� ������������ ����������, ���������� � Excel.")
                            trycloseMyRec()
                        Else
                            trycloseMyRec()
                            If (CheckBox2.Checked = True) Or (CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) <> "02" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "03" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "04" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "05" And _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) <> "06")) Then
                                '---��������� �� ��������� �������
                                '---������� ��������� - �����, ����� ��� ��� ���� �� ��������� �������
                                MySQLStr = "SELECT COUNT(*) AS CC "
                                MySQLStr = MySQLStr & "FROM #_MyOrder "
                                MySQLStr = MySQLStr & "WHERE (ItemCode = N'" & MyProductCode & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.Fields("CC").Value = 0 Then
                                    trycloseMyRec()
                                    '---� ��� �������� - �� ���������� ����������� ��������
                                    MySQLStr = "SELECT SC01066 "
                                    MySQLStr = MySQLStr & "FROM SC010300 "
                                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & MyProductCode & "') "
                                    InitMyConn(False)
                                    InitMyRec(False, MySQLStr)
                                    If Trim(Declarations.MyRec.Fields("SC01066").Value.ToString) = "8" Then
                                        trycloseMyRec()
                                        Throw (New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " �������� �����������. ��� ������ ������ ��������� ����������� ������ ���������� ������."))
                                        trycloseMyRec()
                                    Else
                                        trycloseMyRec()
                                        MySQLStr = "INSERT INTO #_MyOrder "
                                        MySQLStr = MySQLStr & "(ItemCode, QTY, RestQTY) "
                                        MySQLStr = MySQLStr & "VALUES (N'" & MyProductCode & "', "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ") "
                                        InitMyConn(False)
                                        Declarations.MyConn.Execute(MySQLStr)
                                    End If
                                Else
                                    trycloseMyRec()
                                    Throw New ArgumentException("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " ������������ � Excel ����� ��� � ����� ������. ��� ������ �� ����������� ����� ��������� � Excel ������ 1 ���, ��� ������������.")
                                End If
                            ElseIf CheckBox2.Checked = False And _
                                (Microsoft.VisualBasic.Left(MyProductCode, 2) = "02" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "03" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "04" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "05" Or _
                                Microsoft.VisualBasic.Left(MyProductCode, 2) = "06") Then
                                '---��������, ��� ������ �� �������
                                MsgBox("������ " & CStr(i) & ". ��� ������ " & MyProductCode & " �������� ��������� ���������� � �� ����� ������� � ����� �� �����������.", MsgBoxStyle.Critical, "��������!")
                            End If
                        End If
                        i = i + 1
                    End While
                    oWorkBook.Close(True)
                    Me.Cursor = Cursors.Default
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()

                    '---������ �������� ��������� ������������ ������ �� �����������
                    ResultOutput(ExecSppCreateRelocOrder())
                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
            End If
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Function ExecSppCreateRelocOrder() As FinRez
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �������� ��������� �������� ������ �� ����������� �� ������ ������ �� Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyRezStr As String
        Dim MyRelocOrderNum As String                   '����� ������������ ������ �� �����������
        Dim MyFinRez As FinRez                          '������������ ���������
        Dim cmd As New ADODB.Command                    '������� (spp ���������)
        Dim MyParam As ADODB.Parameter                  '������������ �������� ����� 1 //� ������ ������ ����������
        Dim MyParam1 As ADODB.Parameter                 '������������ �������� ����� 2 //�� ����� ����� ����������
        Dim MyParam2 As ADODB.Parameter                 '������������ �������� ����� 3 //�������� � ����� �� ����������� ������ ��� ������� �� ������� �� ������ �������
        Dim MyParam3 As ADODB.Parameter                 '������������ �������� ����� 4 //�������������� ���� ��������
        Dim MyParam4 As ADODB.Parameter                 '������������ �������� ����� 5 //�������������� ���� �������
        Dim MyParam5 As ADODB.Parameter                 '������������ �������� ����� 6 //������������ ������ - ��������� ������
        Dim MyParam6 As ADODB.Parameter                 '������������ �������� ����� 7 //������������ ������ - ����� ������ �� �����������

        MyRezStr = ""
        MyRelocOrderNum = ""
        InitMyConn(False)
        Try
            cmd.ActiveConnection = Declarations.MyConn
            cmd.CommandText = "spp_DisplacementOrderCreationFromExcel"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.CommandTimeout = 300

            '----�������� ����������---------------------------------------------------
            '---�������� �����
            MyParam = cmd.CreateParameter("@SrcWarNo", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam)
            '---����� ����������
            MyParam1 = cmd.CreateParameter("@DestWarNo", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam1)
            '--���� - �������� � ����� �� ����������� ������ ��� ������� �� ������� �� ������ �������
            MyParam2 = cmd.CreateParameter("@MyOtherWHFlag", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam2)
            '---���� ��������
            MyParam3 = cmd.CreateParameter("@MyOrderDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam3)
            '---���� ���������
            MyParam4 = cmd.CreateParameter("@MyShipDate", ADODB.DataTypeEnum.adDBDate, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam4)
            '---������������ �������� (������) - ��������� ������
            MyParam5 = cmd.CreateParameter("@MyRezStr", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
            cmd.Parameters.Append(MyParam5)
            '---������������ �������� (������) - ����� ������ �� �����������
            MyParam6 = cmd.CreateParameter("@MyRelocOrderNum", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamOutput, 30)
            cmd.Parameters.Append(MyParam6)

            '----�������� ����������---------------------------------------------------
            '---�������� �����
            MyParam.Value = Trim(ComboBox1.SelectedValue.ToString)
            '---����� ����������
            MyParam1.Value = Trim(ComboBox2.SelectedValue.ToString)
            '--���� - �������� � ����� �� ����������� ������ ��� ������� �� ������� �� ������ �������
            If CheckBox1.Checked = True Then
                MyParam2.Value = 1 '��������
            Else
                MyParam2.Value = 0 '�� ��������
            End If
            '---���� ��������
            MyParam3.Value = DateTimePicker1.Value
            '---���� ���������
            MyParam4.Value = DateTimePicker2.Value
            '---������ �������� ���������------------------------------------------------
            cmd.Execute()
            MyRezStr = MyRezStr + LTrim(RTrim(MyParam5.Value))
            MyRelocOrderNum = Trim(MyParam6.Value)

        Catch ex As Exception
            MyRezStr = MyRezStr + ex.Message
        End Try
        MyFinRez.MyRezStr = MyRezStr
        MyFinRez.MyRelocOrderNum = MyRelocOrderNum
        ExecSppCreateRelocOrder = MyFinRez
    End Function

    Private Sub ResultOutput(ByVal MyFinRez As FinRez)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ���������� ������ �������� ��������� �������� ������ �� ����������� �� ������ ������ �� Excel � ����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---������ ����������
        '----RemoveBlock()

        MyErrorForm = New ErrorForm
        If MyFinRez.MyRezStr = "" Then
        Else
            MyErrorForm.MyHdr = "�� ����� ������� ������ �� Excel ���� ������ " & Chr(13)
        End If

        '------------����� ���������� � ������ ������ �� �����������
        If MyFinRez.MyRezStr <> "" Then
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & Chr(13) & Chr(13)
        End If

        If Trim(MyFinRez.MyRelocOrderNum) = "" Then
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "� ���������� ���������� �������� ����� �� ����������� �� ��� ������. " & Chr(13)
        Else
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "� ���������� ���������� �������� ��� ������ ����� �� �����������: " & Trim(MyFinRez.MyRelocOrderNum) & Chr(13)
        End If

        '------------����� ���������� � �������������� �������
        MySQLStr = "SELECT  ItemCode, RestQTY "
        MySQLStr = MySQLStr & "FROM #_MyOrder "
        MySQLStr = MySQLStr & "WHERE (RestQTY <> 0) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & Chr(13)
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "� ���������� ������� �� Excel ����� �� ����������� �� ������������� ��� ��������� �������: " & Chr(13)
            MyFinRez.MyRezStr = MyFinRez.MyRezStr & "��� ������ � Scala  �������������� ����������" & Chr(13)
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                MyFinRez.MyRezStr = MyFinRez.MyRezStr & Microsoft.VisualBasic.Left(Declarations.MyRec.Fields("ItemCode").Value & "                    ", 20) & MyRec.Fields("RestQTY").Value.ToString & Chr(13)
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        MyErrorForm.MyMsg = MyFinRez.MyRezStr & Chr(13)
        MyErrorForm.ShowDialog()
    End Sub
End Class
