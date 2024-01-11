Module Functions
    Public Sub InitMyConn(ByVal IsSystem As Boolean)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ������������� ���������� � ��, ������ ���������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim Scala As New SfwIII.Application

        On Error GoTo MyCatch
        If MyConn Is Nothing Then
            MyConn = New ADODB.Connection
            MyConn.CursorLocation = 3
            MyConn.CommandTimeout = 600
            MyConn.ConnectionTimeout = 300
            If Declarations.MyConnStr = "" Then
                Declarations.MyConnStr = Scala.ActiveProcess.UserContext.GetConnectionString(1)
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SPBDVL3"
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=sqlcls"
                Declarations.MyNETConnStr = Replace(Declarations.MyConnStr, "Provider=SQLOLEDB;", "")
                Declarations.MyNETConnStr = Declarations.MyNETConnStr & ";Timeout=0;"
            End If
            If IsSystem = True Then
                MyConn.Open(Replace(Declarations.MyConnStr, "ScaDataDB", "ScalaSystemDB"))
            Else
                MyConn.Open(Declarations.MyConnStr)
            End If
            If Declarations.CompanyID = "" Then
                Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            End If
            If Declarations.Year = "" Then
                Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            End If
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 1")
    End Sub

    Public Sub trycloseMyRec()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//������� �������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        On Error Resume Next
        MyRec.Close()
    End Sub

    Public Sub InitMyRec(ByVal IsSystem As Boolean, ByVal sql As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//�������� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyErr

        On Error GoTo MyCatch
        InitMyConn(IsSystem)
        If MyRec Is Nothing Then
            MyRec = New ADODB.Recordset
        End If
        trycloseMyRec()
        MyRec.LockType = LockTypeEnum.adLockOptimistic
        MyRec.Open(sql, MyConn)
        If MyConn.Errors.Count > 0 Then
            For Each MyErr In MyConn.Errors
                Err.Raise(MyErr.Number, MyErr.Source, MyErr.Description)
            Next MyErr
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 2")
    End Sub

    Public Function GetNewID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ���������� ID ����������� �� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyID As Double                            'ID
        Dim MyTID As Double                           '

        Do
            MyID = GetNextID()
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            MyTID = Declarations.MyRec.Fields("CC").Value
        Loop While MyTID <> 0
        GetNewID = MyID
    End Function

    Public Function GetNextID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ID ����������� �� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyID As Double

        MySQLStr = "Select SY68002 "
        MySQLStr = MySQLStr & "FROM tbl_SY6803XX "
        MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyID = 1
            trycloseMyRec()
        Else
            MyID = Declarations.MyRec.Fields("SY68002").Value
            trycloseMyRec()
            MySQLStr = "UPDATE tbl_SY6803XX "
            MySQLStr = MySQLStr & "SET SY68002 = " & CStr(MyID + 1) & " "
            MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
            Declarations.MyConn.Execute(MySQLStr)
        End If
        GetNextID = MyID
    End Function

    Public Function GetNewPRDID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ���������� ID ������ �� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyID As Double                            'ID
        Dim MyTID As Double                           '

        Do
            MyID = GetNextPRDID()
            'MySQLStr = "SELECT COUNT(*) AS CC "
            'MySQLStr = MySQLStr & "FROM OR010300 WITH (NOLOCK) "
            'MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')"

            MySQLStr = "SELECT COUNT(OR01001) AS CC "
            MySQLStr = MySQLStr & "FROM (SELECT OR01001 "
            MySQLStr = MySQLStr & "FROM OR010300 "
            MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "') "
            MySQLStr = MySQLStr & "UNION ALL "
            MySQLStr = MySQLStr & "Select OR20001 "
            MySQLStr = MySQLStr & "FROM OR200300 "
            MySQLStr = MySQLStr & "WHERE (OR20001 = N'" & Microsoft.VisualBasic.Right("0000000000" & CStr(MyID), 10) & "')) AS View_5 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            MyTID = Declarations.MyRec.Fields("CC").Value
        Loop While MyTID <> 0
        GetNewPRDID = MyID
    End Function

    Public Function GetNextPRDID() As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ���������� ID ������ �� ������� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyID As Double

        MySQLStr = "Select SY68002 "
        MySQLStr = MySQLStr & "FROM SY6803XX "
        MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MyID = 1
            trycloseMyRec()
        Else
            MyID = Declarations.MyRec.Fields("SY68002").Value
            trycloseMyRec()
            MySQLStr = "UPDATE SY6803XX "
            MySQLStr = MySQLStr & "SET SY68002 = " & CStr(MyID + 1) & " "
            MySQLStr = MySQLStr & "WHERE (SY68001 = N'OR01') "
            Declarations.MyConn.Execute(MySQLStr)
        End If
        GetNextPRDID = MyID
    End Function

    Public Function GetExchangeRate(ByVal MyCurr As Integer, ByVal Mydate As DateTime) As Double
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ������ ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT SYCH006 "
        MySQLStr = MySQLStr & "FROM SYCH0100 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = " & CStr(MyCurr) & ") AND "
        MySQLStr = MySQLStr & "(SYCH004 <= CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(MM, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(yyyy, CONVERT(datetime,'" & Mydate & "',103))), 103)) AND "
        MySQLStr = MySQLStr & "(SYCH005 > CONVERT(DATETIME, CONVERT(nvarchar, DATEPART(dd, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(MM, CONVERT(datetime,'" & Mydate & "',103))) + '/' + CONVERT(nvarchar,DATEPART(yyyy, CONVERT(datetime,'" & Mydate & "',103))), 103)) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            GetExchangeRate = 1
        Else
            GetExchangeRate = Declarations.MyRec.Fields("SYCH006").Value
        End If
        trycloseMyRec()

    End Function

    Public Function ImportDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ����� ������ �� Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySQLStr As String
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim MyOrderCounter As Double                  '������� ����� � ������
        Dim MyItemCounter As Double                   '������� ��� - �� ��������� ����� ������ �����. ���� ������ ����������

        Dim MySuppItemCode As String                  '��� ������ ����������
        Dim MyItemName As String                      '�������� ������
        Dim MyUOM As Integer                          '��� ������� ���������
        Dim MyQTY As Double                           '��� - ��
        Dim MyPrice As Double                         '���� �� 1
        Dim MyPriCost As Double                       '��������� �������������
        Dim c As Object
        Dim MyWeekQTY As Double                       '���� ��������
        Dim MyRez As Object

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            '---�������� ������ �������� �� ������� (��� ������� ������)
            MyRez = MsgBox("������� ������ ������ �� ������?", MsgBoxStyle.YesNo, "��������!")
            If MyRez = vbYes Then
                MySQLStr = "DELETE FROM  tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "')  "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                MyOrderCounter = 1
            Else
                MySQLStr = "SELECT MAX(OR03002) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MyOrderCounter = 1
                Else
                    If IsDBNull(Declarations.MyRec.Fields("CC").Value) = True Then
                        MyOrderCounter = 1
                    Else
                        MyOrderCounter = CInt(Declarations.MyRec.Fields("CC").Value) / 10 + 1
                    End If
                End If
            End If

            MyExcelCounter = 11

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing
                '------��� ������
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    Declarations.MyItemCode = ""
                Else
                    Declarations.MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                End If
                '------��� ������ ����������
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                End If
                If Len(MySuppItemCode) > 32 Then
                    MyRez = MsgBox("������ C" & MyExcelCounter & " ��� ������ ���������� �� ������ ��������� 32 ����a. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" Then
                        MsgBox("������ " & MyExcelCounter & " ����������� ������ ���� ������� ��� ��� ������ Scala, ��� ��� ������ ����������!", MsgBoxStyle.Critical, "��������!")
                    Else
                        If Trim(Declarations.MyItemCode) = "" Then
                            MySQLStr = "SELECT COUNT(*) AS CC "
                            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                trycloseMyRec()
                                MsgBox("���������� �������� ���������� � ������� �� Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                                Exit While
                            Else
                                Declarations.MyRec.MoveFirst()
                                MyItemCounter = Declarations.MyRec.Fields("CC").Value
                                trycloseMyRec()
                            End If
                            If MyItemCounter = 0 Then
                                '---������ � ����� ����� ������ ���������� � Scala ��� 
                                Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                            ElseIf MyItemCounter = 1 Then
                                '---����� � ����� ����� ������ ���������� � Scala ������ ����
                                MySQLStr = "SELECT SC01001 AS CC "
                                MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                                MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                                InitMyConn(False)
                                InitMyRec(False, MySQLStr)
                                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                    trycloseMyRec()
                                    Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                                Else
                                    Declarations.MyRec.MoveFirst()
                                    Declarations.MyItemCode = Declarations.MyRec.Fields("CC").Value
                                    trycloseMyRec()
                                End If
                            Else
                                '---� Scala ��������� ������� � ����� ����� ������ ���������� 
                                MySelectItemBySuppCode = New SelectItemBySuppCode
                                MySelectItemBySuppCode.MyItemSuppCode = Trim(MySuppItemCode)
                                MySelectItemBySuppCode.MyWindowFrom = "Import"
                                MySelectItemBySuppCode.ShowDialog()
                            End If
                        End If
                    End If



                    If (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) Is Double) Then
                        MyRez = MsgBox("������ D" & MyExcelCounter & " ������������ ������ ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    Else
                        If Len(appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) > 50 Then
                            MyRez = MsgBox("������ D" & MyExcelCounter & " ������������ ������ �� ������ ��������� 50 ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                            If MyRez = vbYes Then
                                Exit While
                            End If
                        Else
                            '------�������� ������
                            MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value
                            If (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value) Is Double) Then
                                MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                                If c.Text = "" Then
                                    MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' �� ������� �� ������ �����. ���������� ������� �������� �� ����������� ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                    If MyRez = vbYes Then
                                        Exit While
                                    End If
                                Else
                                    '------��� ������� ���������
                                    MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                                    If (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                        MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                        If MyRez = vbYes Then
                                            Exit While
                                        End If
                                    Else
                                        If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                            MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                            If MyRez = vbYes Then
                                                Exit While
                                            End If
                                        Else
                                            '----------��� - �� � ������
                                            MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                                            If (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                                MyRez = MsgBox("������ G" & MyExcelCounter & " �������� '���� ��� ���' ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                                If MyRez = vbYes Then
                                                    Exit While
                                                End If
                                            Else
                                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                                    MyRez = MsgBox("������ G" & MyExcelCounter & " �������� '���� ��� ���' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                                    If MyRez = vbYes Then
                                                        Exit While
                                                    End If
                                                Else
                                                    '----------����
                                                    MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                                                    '----------���� ��������
                                                    Try
                                                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                                                        If MyWeekQTY >= 0 Then
                                                            '----------��������� �������������
                                                            Try
                                                                MyPriCost = Math.Round(appXLSRC.Worksheets(1).Range("J" & MyExcelCounter).Value, 2)
                                                            Catch ex As Exception
                                                                MyPriCost = 0
                                                            End Try
                                                            AddRow(Declarations.MyOrderNum, Right("000000" & CStr(MyOrderCounter * 10), 6), Declarations.MyItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY, Trim(MySuppItemCode), MyPriCost)

                                                            MyOrderCounter = MyOrderCounter + 1.0
                                                        Else
                                                            MyRez = MsgBox("������ I" & MyExcelCounter & " ���� �������� ������ �����������. ������ ���� ����� ������� ��� ������ ����. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                                            If MyRez = vbYes Then
                                                                Exit While
                                                            End If
                                                        End If
                                                    Catch ex As Exception
                                                        MyRez = MsgBox("������ I" & MyExcelCounter & " ���� �������� ������ �����������. ������ ���� ����� ������� ��� ������ ����. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                                        If MyRez = vbYes Then
                                                            Exit While
                                                        End If
                                                    End Try
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------�������� ��������
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("��������� ������� ����� ������ �� ������� ���������.", vbOKOnly, "��������!")
    End Function

    Public Function ImportDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ����� ������ �� Libre Office
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim MyOrderCounter As Double                  '������� ����� � ������
        Dim MyItemCounter As Double                   '������� ��� - �� ��������� ����� ������ �����. ���� ������ ����������
        Dim MySuppItemCode As String                  '��� ������ ����������
        Dim MyItemName As String                      '�������� ������
        Dim MyUOMStr As String                        '������� ���������
        Dim MyUOM As Integer                          '��� ������� ���������
        Dim MyQTY As Double                           '��� - ��
        Dim MyPrice As Double                         '���� �� 1
        Dim MyPriCost As Double                       '��������� �������������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       '������ ���������
        Dim MySQLStr As String                        'SQL ������
        Dim MyRez As Object
        Dim MyWeekQTY As Double                       '���� ��������

        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        oFileName = Replace(Declarations.ImportFileName, "\", "/")
        oFileName = "file:///" + oFileName
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        '---��������� ������ ����� Excel
        MyVersion = oSheet.getCellRangeByName("A1").String
        If MyVersion = "" Then
            MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Excel ", MsgBoxStyle.Critical, "��������!")
            oWorkBook.Close(True)
            Exit Function
        Else
            MySQLStr = "SELECT Version "
            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
            MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                trycloseMyRec()
                oWorkBook.Close(True)
                Exit Function
            Else
                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                    trycloseMyRec()
                Else
                    MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                    trycloseMyRec()
                    oWorkBook.Close(True)
                    Exit Function
                End If
            End If
        End If

        '---�������� ������ �������� �� ������� (��� ������� ������)
        MyRez = MsgBox("������� ������ ������ �� ������?", MsgBoxStyle.YesNo, "��������!")
        If MyRez = vbYes Then
            MySQLStr = "DELETE FROM  tbl_OR030300 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "')  "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            MyOrderCounter = 1
        Else
            MySQLStr = "SELECT MAX(OR03002) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Declarations.MyOrderNum & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MyOrderCounter = 1
            Else
                If IsDBNull(Declarations.MyRec.Fields("CC").Value) = True Then
                    MyOrderCounter = 1
                Else
                    MyOrderCounter = CInt(Declarations.MyRec.Fields("CC").Value) / 10 + 1
                End If
            End If
        End If

        MyExcelCounter = 11

        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------��� ������
            Declarations.MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            '------��� ������ ����������
            MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
            If Len(MySuppItemCode) > 32 Then
                MyRez = MsgBox("������ C" & MyExcelCounter & " ��� ������ ���������� �� ������ ��������� 32 ����a. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                If MyRez = vbYes Then
                    Exit While
                End If
            Else
                If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" Then
                    MsgBox("������ " & MyExcelCounter & " ����������� ������ ���� ������� ��� ��� ������ Scala, ��� ��� ������ ����������!", MsgBoxStyle.Critical, "��������!")
                Else
                    If Trim(Declarations.MyItemCode) = "" Then
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("���������� �������� ���������� � ������� �� Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyItemCounter = Declarations.MyRec.Fields("CC").Value
                            trycloseMyRec()
                        End If
                        If MyItemCounter = 0 Then
                            '---������ � ����� ����� ������ ���������� � Scala ��� 
                            Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                        ElseIf MyItemCounter = 1 Then
                            '---����� � ����� ����� ������ ���������� � Scala ������ ����
                            MySQLStr = "SELECT SC01001 AS CC "
                            MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                            MySQLStr = MySQLStr & "WHERE (SC01060 = N'" & Trim(MySuppItemCode) & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                trycloseMyRec()
                                Declarations.MyItemCode = "NN_" & Trim(MySuppItemCode)
                            Else
                                Declarations.MyRec.MoveFirst()
                                Declarations.MyItemCode = Declarations.MyRec.Fields("CC").Value
                                trycloseMyRec()
                            End If
                        Else
                            '---� Scala ��������� ������� � ����� ����� ������ ���������� 
                            MySelectItemBySuppCode = New SelectItemBySuppCode
                            MySelectItemBySuppCode.MyItemSuppCode = Trim(MySuppItemCode)
                            MySelectItemBySuppCode.MyWindowFrom = "Import"
                            MySelectItemBySuppCode.ShowDialog()
                        End If
                    End If
                End If
            End If

            If oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") Then
                MyRez = MsgBox("������ D" & MyExcelCounter & " ������������ ������ ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                If MyRez = vbYes Then
                    Exit While
                End If
            Else
                If Len(Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)) > 50 Then
                    MyRez = MsgBox("������ D" & MyExcelCounter & " ������������ ������ �� ������ ��������� 50 ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    '------�������� ������
                    MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
                    '------������� ���������
                    MyUOMStr = oSheet.getCellRangeByName("E" & MyExcelCounter).String
                    Dim MyRange As Object = oSheet.getCellrangeByName("O11:O18")
                    Dim Search_Desc As Object = MyRange.createSearchDescriptor()
                    Search_Desc.SearchString = MyUOMStr
                    Dim Search_Result As Object = MyRange.findAll(Search_Desc)
                    If Search_Result.Count < 1 Then
                        MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' �� ������� �� ������ �����. ���������� ������� �������� �� ����������� ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    Else
                        Dim Last_Occur As Object = Search_Result.getByIndex(Search_Result.Count - 1)
                        MyUOM = oSheet.getCellRangeByName("N" & Last_Occur.CellAddress.Row + 1).String
                    End If
                    '-----����������
                    Try
                        MyQTY = oSheet.getCellRangeByName("F" & MyExcelCounter).Value
                    Catch ex As Exception
                        MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    End Try
                    '----------����
                    Try
                        MyPrice = oSheet.getCellRangeByName("G" & MyExcelCounter).Value
                    Catch ex As Exception
                        MyRez = MsgBox("������ G" & MyExcelCounter & " �������� '���� ��� ���' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    End Try
                    '----------���� ��������
                    Try
                        MyWeekQTY = oSheet.getCellRangeByName("I" & MyExcelCounter).Value
                    Catch ex As Exception
                        MyRez = MsgBox("������ I" & MyExcelCounter & " �������� '���� ��������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        If MyRez = vbYes Then
                            Exit While
                        End If
                    End Try
                    '----------��������� �������������
                    Try
                        MyPriCost = oSheet.getCellRangeByName("J" & MyExcelCounter).Value
                    Catch ex As Exception
                        MyPriCost = 0
                    End Try

                    AddRow(Declarations.MyOrderNum, Right("000000" & CStr(MyOrderCounter * 10), 6), Declarations.MyItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY, Trim(MySuppItemCode), MyPriCost)
                    MyOrderCounter = MyOrderCounter + 1.0
                End If
            End If
            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("��������� ������� ����� ������ �� ������� ���������.", vbOKOnly, "��������!")
    End Function

    Public Function ImportRequestDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ����� ������� �� ����� �� Excel (������������)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MySQLStr As String
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim MyItemCounter As Double                   '������� ��� - �� ��������� ����� ������ �����. ���� ������ ����������

        Dim MySuppItemCode As String                  '��� ������ ����������
        Dim MyScalaSuppItemCode As String             '��� ������ ���������� � Scala
        Dim MyItemName As String                      '�������� ������
        Dim MyScalaItemName As String                 '�������� ������ � Scala
        Dim MyUOM As Integer                          '��� ������� ���������
        Dim MyScalaUOM As Integer                     '��� ������� ��������� � Scala
        Dim MyQTY As Double                           '��� - ��
        Dim MyPrice As Double                         '���� �� 1
        Dim c As Object
        Dim MyWeekQTY As Double                       '���� ��������
        Dim MyRez As Object

        MyScalaSuppItemCode = ""
        MyScalaItemName = ""
        MyScalaUOM = -1

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            '---�������� ������ �������� �� ������� (��� ������� ������)
            MyRez = MsgBox("������� ������ ������ �� �������?", MsgBoxStyle.YesNo, "��������!")
            If MyRez = vbYes Then
                MySQLStr = "DELETE FROM  tbl_SupplSearchItems "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = N'" & Declarations.MyRequestNum & "')  "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            MyExcelCounter = 11

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Or _
                Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Or _
                Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing

                '------��� ������ Scala
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    Declarations.MyItemCode = ""
                Else
                    Declarations.MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                    '---�������� ��� ����� ��� ���� � Scala
                    MySQLStr = "SELECT COUNT(*) AS CC "
                    MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)
                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                        trycloseMyRec()
                        MyItemCounter = 0
                        'MsgBox("������ B" & MyExcelCounter & ". ���������� �������� ���������� � ������� �� Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                        'Exit While
                    Else
                        Declarations.MyRec.MoveFirst()
                        MyItemCounter = Declarations.MyRec.Fields("CC").Value
                        trycloseMyRec()
                    End If
                    If MyItemCounter = 0 Then
                        'MsgBox("������ B" & MyExcelCounter & ". ������ ���� ������ � Scala ���.", MsgBoxStyle.Critical, "��������!")
                        'Exit While
                        Declarations.MyItemCode = ""
                    Else
                        '-----�������� ���������� �� ���������� ������
                        MySQLStr = "SELECT SC01135 AS UOM, LTRIM(RTRIM(LTRIM(RTRIM(SC01002)) + ' ' + LTRIM(RTRIM(SC01003)))) AS Name, SC01060 AS SuppItemID "
                        MySQLStr = MySQLStr & "FROM SC010300 "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "

                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
 
                        MyScalaUOM = CInt(Declarations.MyRec.Fields("UOM").Value)
                        MyScalaItemName = Declarations.MyRec.Fields("Name").Value
                        MyScalaSuppItemCode = Declarations.MyRec.Fields("SuppItemID").Value
                        trycloseMyRec()
                    End If
                End If

                '------��� ������ ����������
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    If Trim(MyItemCode) = "" Then
                        MySuppItemCode = ""
                    Else
                        MySuppItemCode = MyScalaSuppItemCode
                    End If
                Else
                    If Trim(MyItemCode) = "" Then
                        MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                    Else
                        MySuppItemCode = MyScalaSuppItemCode
                    End If
                    ''------�������� ������������ ���� ������ ���������� � �������
                    'MySQLStr = "SELECT COUNT(ID) AS CC "
                    'MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                    'MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    'MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
                    'InitMyConn(False)
                    'InitMyRec(False, MySQLStr)
                    'If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    '    trycloseMyRec()
                    '    MsgBox("������ C" & MyExcelCounter & " ������ �������� ������������ ���� ������ ������������� � �������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                    '    Exit While
                    'Else
                    '    If Declarations.MyRec.Fields("CC").Value = 0 Then
                    '        trycloseMyRec()
                    '    Else
                    '        trycloseMyRec()
                    '        MsgBox("������ C" & MyExcelCounter & " ������ - ����� � ����� ������ ������������� " & Trim(MySuppItemCode) & " ��� ������������ � ������� �� ����� ����������.", MsgBoxStyle.Critical, "��������!")
                    '        Exit While
                    '    End If
                    'End If
                End If


                '------�������� ������
                If (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) Is Double) Then
                    If Trim(MyItemCode) = "" Then
                        MyItemName = ""
                    Else
                        MyItemName = MyScalaItemName
                    End If
                Else
                    If Len(appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value) > 254 Then
                        MsgBox("������ D" & MyExcelCounter & " ������������ ������ �� ������ ��������� 254 �����. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                        Exit While
                    Else
                        If Trim(MyItemCode) = "" Then
                            MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value
                        Else
                            MyItemName = MyScalaItemName
                        End If
                        ''------�������� ������������ �������� ������ � �������
                        'MySQLStr = "SELECT COUNT(ID) AS CC "
                        'MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                        'MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                        'MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
                        'InitMyConn(False)
                        'InitMyRec(False, MySQLStr)
                        'If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                        '    trycloseMyRec()
                        '    MsgBox("������ D" & MyExcelCounter & " ������ �������� ������������ �������� ������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                        '    Exit While
                        'Else
                        '    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        '        trycloseMyRec()
                        '    Else
                        '        trycloseMyRec()
                        '        MsgBox("������ C" & MyExcelCounter & " ������ - ����� � ��������� " & Trim(MyItemName) & " ��� ������������ � ������� �� ����� ����������.", MsgBoxStyle.Critical, "��������!")
                        '        Exit While
                        '    End If
                        'End If
                    End If
                End If

                '------�������� ������������ ���� ������ ���������� + �������� ������ � �������
                MySQLStr = "SELECT COUNT(ID) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
                MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    trycloseMyRec()
                    MsgBox("������ D" & MyExcelCounter & " ������ �������� ������������ ���� ������ ������������� + �������� ������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                    Exit While
                Else
                    If Declarations.MyRec.Fields("CC").Value = 0 Then
                        trycloseMyRec()
                    Else
                        trycloseMyRec()
                        MsgBox("������ C" & MyExcelCounter & " ������ - ����� � ����� ������������� " & Trim(MySuppItemCode) & " � ��������� " & Trim(MyItemName) & " ��� ������������ � ������� �� ����� ����������.", MsgBoxStyle.Critical, "��������!")
                        Exit While
                    End If
                End If


                '-------�������� ���������� ������ 3 - � �����
                'If Trim(Declarations.MyItemCode) = "" And Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                '-----��� Scala �� �������
                If Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                    MsgBox("������ " & MyExcelCounter & " ����������� ������ ���� ������� ��� ������ ���������� ��� �������� ������!", MsgBoxStyle.Critical, "��������!")
                    Exit While
                Else
                    '------��� ������� ���������
                    If (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value) Is Double) Then
                        If Trim(MyItemCode) = "" Then
                            MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                            If MyRez = vbYes Then
                                Exit While
                            End If
                        Else
                            MyUOM = MyScalaUOM
                        End If
                    Else
                        c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                        If c.Text = "" Then
                            If Trim(MyItemCode) = "" Then
                                MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' �� ������� �� ������ �����. ���������� ������� �������� �� ����������� ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                MyUOM = MyScalaUOM
                            End If
                        Else
                            If Trim(MyItemCode) = "" Then
                                MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                            Else
                                MyUOM = MyScalaUOM
                            End If
                            '----------��� - �� � �������
                            If (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                If MyRez = vbYes Then
                                    Exit While
                                End If
                            Else
                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                                    MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                                    If MyRez = vbYes Then
                                        Exit While
                                    End If
                                Else
                                    MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                                    '----------����
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                                        MyPrice = 0
                                    Else
                                        MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                                    End If
                                    '----------���� ��������
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value) Is Double) Then
                                        MyWeekQTY = 0
                                    Else
                                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                                    End If
                                    '----------��������� ������
                                    'AddRequestRow(Declarations.MyRequestNum, Declarations.MyItemCode, MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
                                    '---��� Scala �� �������
                                    AddRequestRow(Declarations.MyRequestNum, "", MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
                                End If
                            End If
                        End If
                    End If
                End If
                MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------�������� ��������
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("��������� ������� ����� ������� �� ����� ���������� ���������.", vbOKOnly, "��������!")
    End Function

    Public Function ImportRequestDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ����� ������� �� ����� �� LibreOffice (������������)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySuppItemCode As String                  '��� ������ ����������
        Dim MyItemName As String                      '�������� ������
        Dim MyUOMStr As String                        '������� ���������
        Dim MyUOM As Integer                          '��� ������� ���������
        Dim MyQTY As Double                           '��� - ��
        Dim MyPrice As Double                         '���� �� 1
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String
        Dim MyVersion As String                       '������ ���������
        Dim MySQLStr As String                        'SQL ������
        Dim MyRez As Object
        Dim MyExcelCounter As Double                  '������� ����� Excel
        Dim MyItemCounter As Double                   '������� ��� - �� ��������� ����� ������ �����. ���� ������ ����������
        Dim MyScalaItemName As String                 '�������� ������ � Scala
        Dim MyScalaUOM As Integer                     '��� ������� ��������� � Scala
        Dim MyScalaSuppItemCode As String             '��� ������ ���������� � Scala
        Dim MyWeekQTY As Double                       '���� ��������

        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        oFileName = Replace(Declarations.ImportFileName, "\", "/")
        oFileName = "file:///" + oFileName
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)

        '---��������� ������ ����� Excel
        MyVersion = oSheet.getCellRangeByName("A1").String
        If MyVersion = "" Then
            MsgBox("� ������������� ����� Excel � ������ 'A1' �� ����������� ������ ����� Excel ", MsgBoxStyle.Critical, "��������!")
            oWorkBook.Close(True)
            Exit Function
        Else
            MySQLStr = "SELECT Version "
            MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel "
            MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������", vbCritical, "��������!")
                trycloseMyRec()
                oWorkBook.Close(True)
                Exit Function
            Else
                If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                    trycloseMyRec()
                Else
                    MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Declarations.MyRec.Fields("Version").Value & ".", vbCritical, "��������!")
                    trycloseMyRec()
                    oWorkBook.Close(True)
                    Exit Function
                End If
            End If
        End If

        '---�������� ������ �������� �� ������� (��� ������� ������)
        MyRez = MsgBox("������� ������ ������ �� �������?", MsgBoxStyle.YesNo, "��������!")
        If MyRez = vbYes Then
            MySQLStr = "DELETE FROM  tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (SupplSearchID = N'" & Declarations.MyRequestNum & "')  "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        MyExcelCounter = 11
        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------��� ������ Scala
            Declarations.MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            If Declarations.MyItemCode.Equals("") = False Then
                '---�������� ��� ����� ��� ���� � Scala
                MySQLStr = "SELECT COUNT(*) AS CC "
                MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MyItemCounter = 0
                Else
                    Declarations.MyRec.MoveFirst()
                    MyItemCounter = Declarations.MyRec.Fields("CC").Value
                    trycloseMyRec()
                End If
                If MyItemCounter = 0 Then
                    Declarations.MyItemCode = ""
                Else
                    '-----�������� ���������� �� ���������� ������
                    MySQLStr = "SELECT SC01135 AS UOM, LTRIM(RTRIM(LTRIM(RTRIM(SC01002)) + ' ' + LTRIM(RTRIM(SC01003)))) AS Name, SC01060 AS SuppItemID "
                    MySQLStr = MySQLStr & "FROM SC010300 "
                    MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "

                    InitMyConn(False)
                    InitMyRec(False, MySQLStr)

                    MyScalaUOM = CInt(Declarations.MyRec.Fields("UOM").Value)
                    MyScalaItemName = Declarations.MyRec.Fields("Name").Value
                    MyScalaSuppItemCode = Declarations.MyRec.Fields("SuppItemID").Value
                    trycloseMyRec()
                End If
            End If
            '------��� ������ ����������
            If Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String).Equals("") Then
                If Trim(MyItemCode) = "" Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = MyScalaSuppItemCode
                End If
            Else
                If Trim(MyItemCode) = "" Then
                    MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
                Else
                    MySuppItemCode = MyScalaSuppItemCode
                End If
            End If

            '------�������� ������
            If Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String).Equals("") Then
                If Trim(MyItemCode) = "" Then
                    MyItemName = ""
                Else
                    MyItemName = MyScalaItemName
                End If
            Else
                If Len(Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)) > 254 Then
                    MsgBox("������ D" & MyExcelCounter & " ������������ ������ �� ������ ��������� 254 �����.", MsgBoxStyle.YesNo, "��������!")
                    Exit While
                Else
                    If Trim(MyItemCode) = "" Then
                        MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
                    Else
                        MyItemName = MyScalaItemName
                    End If
                End If
            End If

            '------�������� ������������ ���� ������ ���������� + �������� ������ � �������
            MySQLStr = "SELECT COUNT(ID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_SupplSearchItems "
            MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
            MySQLStr = MySQLStr & "AND (ItemSuppID = N'" & Trim(MySuppItemCode) & "') "
            MySQLStr = MySQLStr & "AND (ItemName = N'" & Trim(MyItemName) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("������ D" & MyExcelCounter & " ������ �������� ������������ ���� ������ ������������� + �������� ������. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                Exit While
            Else
                If Declarations.MyRec.Fields("CC").Value = 0 Then
                    trycloseMyRec()
                Else
                    trycloseMyRec()
                    MsgBox("������ C" & MyExcelCounter & " ������ - ����� � ����� ������������� " & Trim(MySuppItemCode) & " � ��������� " & Trim(MyItemName) & " ��� ������������ � ������� �� ����� ����������.", MsgBoxStyle.Critical, "��������!")
                    Exit While
                End If
            End If

            '-------�������� ���������� ������ 3 - � �����
            '-----��� Scala �� �������
            If Trim(MySuppItemCode) = "" And Trim(MyItemName) = "" Then
                MsgBox("������ " & MyExcelCounter & " ����������� ������ ���� ������� ��� ������ ���������� ��� �������� ������!", MsgBoxStyle.Critical, "��������!")
                Exit While
            Else
                '------������� ���������
                MyUOMStr = oSheet.getCellRangeByName("E" & MyExcelCounter).String
                Dim MyRange As Object = oSheet.getCellrangeByName("O11:O18")
                Dim Search_Desc As Object = MyRange.createSearchDescriptor()
                Search_Desc.SearchString = MyUOMStr
                Dim Search_Result As Object = MyRange.findAll(Search_Desc)
                If Search_Result.Count < 1 Then
                    MyRez = MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' �� ������� �� ������ �����. ���������� ������� �������� �� ����������� ������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                Else
                    Dim Last_Occur As Object = Search_Result.getByIndex(Search_Result.Count - 1)
                    MyUOM = oSheet.getCellRangeByName("N" & Last_Occur.CellAddress.Row + 1).String
                End If
                '-----����������
                Try
                    MyQTY = oSheet.getCellRangeByName("F" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("������ F" & MyExcelCounter & " �������� '����������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '----------����
                Try
                    MyPrice = oSheet.getCellRangeByName("G" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("������ G" & MyExcelCounter & " �������� '���� ��� ���' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '----------���� ��������
                Try
                    MyWeekQTY = oSheet.getCellRangeByName("I" & MyExcelCounter).Value
                Catch ex As Exception
                    MyRez = MsgBox("������ I" & MyExcelCounter & " �������� '���� ��������' ������ ���� ��������� �������� ���������. �������� ���������� �������?", MsgBoxStyle.YesNo, "��������!")
                    If MyRez = vbYes Then
                        Exit While
                    End If
                End Try
                '----------��������� ������
                '---��� Scala �� �������
                AddRequestRow(Declarations.MyRequestNum, "", MySuppItemCode, MyItemName, MyUOM, MyQTY, MyPrice, MyWeekQTY)
            End If
            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("��������� ������� ����� ������� �� ����� ���������� ���������.", vbOKOnly, "��������!")
    End Function

    Public Sub AddRequestRow(ByVal MyRequest As Integer, ByVal MyItemCode As String, ByVal MySuppItemCode As String, ByVal MyItemName As String, _
        ByVal MyUOM As Integer, ByVal MyQTY As Double, ByVal MyPrice As Double, ByVal MyWeekQTY As Double)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � ������� �� ����� ����������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "INSERT INTO tbl_SupplSearchItems"
        MySQLStr = MySQLStr & "(SupplSearchID, ItemID, ItemSuppID, ItemName, UOM, QTY, LeadTimeWeek, Comments) "
        MySQLStr = MySQLStr & "VALUES ("
        MySQLStr = MySQLStr & MyRequest.ToString & ", "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MyItemCode), "'", "''") & "', "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MySuppItemCode), "'", "''") & "', "
        MySQLStr = MySQLStr & "N'" & Replace(Trim(MyItemName), "'", "''") & "', "
        MySQLStr = MySQLStr & MyUOM.ToString & ", "
        MySQLStr = MySQLStr & Replace(MyQty.ToString, ",", ".") & ", "
        If MyWeekQTY = 0 Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & Replace(MyWeekQTY.ToString, ",", ".") & ", "
        End If
        MySQLStr = MySQLStr & "N'' "
        MySQLStr = MySQLStr & ") "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Function CheckVersion(ByVal MyVer As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ Excel ����� - ����� �� � ��� ��������
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'������������ �����������') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("� Scala �� ����������� ������� ������ ����� Excel. ���������� � ��������������.", vbCritical, "��������!")
            CheckVersion = False
            trycloseMyRec()
            Exit Function
        Else
            If Trim(Declarations.MyRec.Fields("Version").Value) = Trim(MyVer) Then
                CheckVersion = True
                trycloseMyRec()
                Exit Function
            Else
                MsgBox("�� ��������� �������� � ������������ ������� ����� Excel. ���� �������� � ������� " & Trim(Declarations.MyRec.Fields("Version").Value) & ".", vbCritical, "��������!")
                CheckVersion = False
                trycloseMyRec()
                Exit Function
            End If
        End If
    End Function

    Public Function AddRow(ByVal MyOrder As String, ByVal MyStr As String, ByVal MyItemCode As String, ByVal MyItemName As String, _
        ByVal MyUOM As Integer, ByVal MyQTY As Double, ByVal MyPrice As Double, ByVal MyWeekQTY As Double, _
        ByVal MySuppItemCode As String, ByVal MyPriCost As Double)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ������ �� �������
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  '������������ �������� ����� 1
        Dim MyParam1 As ADODB.Parameter                 '������������ �������� ����� 2
        Dim MyParam2 As ADODB.Parameter                 '������������ �������� ����� 3
        Dim MyParam3 As ADODB.Parameter                 '������������ �������� ����� 4
        Dim MyParam4 As ADODB.Parameter                 '������������ �������� ����� 5
        Dim MyParam5 As ADODB.Parameter                 '������������ �������� ����� 6
        Dim MyParam6 As ADODB.Parameter                 '������������ �������� ����� 7
        Dim MyParam7 As ADODB.Parameter                 '������������ �������� ����� 8
        Dim MyParam8 As ADODB.Parameter                 '������������ �������� ����� 9
        Dim MySQLStr As String

        cmd.ActiveConnection = Declarations.MyConn
        cmd.CommandText = "spp_SalesWorkplace4_ImportRow"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.CommandTimeout = 300

        MyParam = cmd.CreateParameter("@MyOrderNum", 129, ADODB.ParameterDirectionEnum.adParamInput, 10)
        cmd.Parameters.Append(MyParam)
        MyParam.Value = MyOrder

        MyParam1 = cmd.CreateParameter("@MyStrNum", 129, ADODB.ParameterDirectionEnum.adParamInput, 6)
        cmd.Parameters.Append(MyParam1)
        MyParam1.Value = MyStr

        MyParam2 = cmd.CreateParameter("@MyItemCode", 129, ADODB.ParameterDirectionEnum.adParamInput, 35)
        cmd.Parameters.Append(MyParam2)
        MyParam2.Value = MyItemCode

        MyParam3 = cmd.CreateParameter("@MyItemName", 129, ADODB.ParameterDirectionEnum.adParamInput, 51)
        cmd.Parameters.Append(MyParam3)
        MyParam3.Value = MyItemName

        MyParam4 = cmd.CreateParameter("@MyUOM", 3, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam4)
        MyParam4.Value = MyUOM

        MyParam5 = cmd.CreateParameter("@MyQTY", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam5)
        MyParam5.Value = MyQTY

        MyParam6 = cmd.CreateParameter("@MyPrice", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam6)
        MyParam6.Value = MyPrice

        MyParam7 = cmd.CreateParameter("@MyWeekQTY", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam7)
        MyParam7.Value = MyWeekQTY

        MyParam8 = cmd.CreateParameter("@MyPriCost", 5, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam8)
        MyParam8.Value = MyPriCost

        cmd.Execute()

        '-----��������� ������ ��������
        'MySQLStr = "UPDATE tbl_OR030300 "
        'MySQLStr = MySQLStr & "SET WeekQTY = " & Replace(CStr(MyWeekQTY), ",", ".") & " "
        'MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
        'MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
        'InitMyConn(False)
        'Declarations.MyConn.Execute(MySQLStr)

        '-----���������� ���� ������ ����������
        MySQLStr = "UPDATE tbl_OR030300 "
        MySQLStr = MySQLStr & "SET SuppItemCode = N'" & MySuppItemCode & "' "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
        MySQLStr = MySQLStr & "AND (OR03002 = N'" & MyStr & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

    End Function

    Public Function CheckRights(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - �������� �� ������ ������ CRMManagers
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyCCPermission = False
            CheckRights = "���������"
        Else
            Declarations.MyCCPermission = True
            CheckRights = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 5")
        Declarations.MyCCPermission = False
        CheckRights = "���������"
    End Function

    Public Function CheckRights1(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - �������� �� ������ ������ CRMDirector
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyPermission = False
            CheckRights1 = "���������"
        Else
            Declarations.MyPermission = True
            CheckRights1 = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 6")
        Declarations.MyPermission = False
        CheckRights1 = "���������"
    End Function

    Public Function CheckRights2(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - �������� �� ������ ������ ProposalManager
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyCPPermission = False
            CheckRights2 = "���������"
        Else
            Declarations.MyCPPermission = True
            CheckRights2 = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 7")
        Declarations.MyCPPermission = False
        CheckRights2 = "���������"
    End Function

    Public Sub SendAddInfoReminder(ByVal MyItemCode As String, ByVal MySalesman As String, ByVal MyType As Integer)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� ����� ����������� ������ ������� � �������� ���������, ��������� � ���������
        '// MyType 0 - �������� 1 - �������� 2 - ��������
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "exec spp_SalesWorkplace4_SendReminder N'" & Trim(MyItemCode) & "', N'" & MySalesman & "', " & CStr(MyType)
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub SendInfoByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewRequestState As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� ����� ����������� ����������� �� ���������� � ��������
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String
        Dim MailTo() As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        MailTo = Split(EMailTo, ";")
        For i As Integer = 0 To UBound(MailTo)
            msg.To.Add(MailTo(i))
        Next
        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "����������� �� ��������� ������� �� �����"
        MyMsgStr = "��������� �������!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "��������� ��������� ���������� ������� �� �����:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " �� " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "��� �������: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "��������: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "����� ��������� �������: " & NewRequestState & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "� ���������," & Chr(13)
        MyMsgStr = MyMsgStr + "��� ""��������"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. �� ������ ������� �� ��������, ��� �������������� ��������. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

    Public Sub SendCommentByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewComment As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� ����� ������ ����������� �����������
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String
        Dim MailTo() As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        MailTo = Split(EMailTo, ";")
        For i As Integer = 0 To UBound(MailTo)
            msg.To.Add(MailTo(i))
        Next
        msg.From = New Net.Mail.MailAddress("reportserver@skandikagroup.ru")
        msg.Subject = "����������� � ����� ����������� ������� �� �����"
        MyMsgStr = "��������� �������!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "�������� ����� ����������� � ���������� ������� �� �����:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " �� " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "��� �������: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "��������: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "�������� �����������: " & NewComment & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "� ���������," & Chr(13)
        MyMsgStr = MyMsgStr + "��� ""��������"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. �� ������ ������� �� ��������, ��� �������������� ��������. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

    Public Function GetEmailFromDB(ByVal UserCode As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� Email �� ���� �������� / ���������...
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT RM66003 "
        MySQLStr = MySQLStr & "FROM RM.dbo.RM660100 "
        MySQLStr = MySQLStr & "WHERE (RM66001 = '" & Trim(UserCode) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            GetEmailFromDB = ""
        Else
            GetEmailFromDB = Trim(Declarations.MyRec.Fields("RM66003").Value.ToString)
        End If
        trycloseMyRec()
    End Function

    Public Function GetSrchManagerEmailFromDB() As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� Email ��������� �����������
        '// 
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MySTR As String
        Dim MyFlag As Integer

        MySTR = ""
        MyFlag = 0
        MySQLStr = "SELECT RM.dbo.RM660100.RM66003 "
        MySQLStr = MySQLStr & "FROM tbl_SupplSearch_Searchers INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON tbl_SupplSearch_Searchers.PurchID = RM.dbo.RM660100.RM66001 "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplSearch_Searchers.IsLeader = 1) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            GetSrchManagerEmailFromDB = ""
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                If Trim(Declarations.MyRec.Fields("RM66003").Value.ToString) <> "" Then
                    If MyFlag <> 0 Then
                        MySTR = MySTR & ";"
                    End If
                    MySTR = MySTR & Trim(Declarations.MyRec.Fields("RM66003").Value.ToString)
                End If
                Declarations.MyRec.MoveNext()
                MyFlag = MyFlag + 1
            End While
            GetSrchManagerEmailFromDB = Trim(MySTR)
        End If
    End Function

    Public Function CheckSalesman(ByVal SalesmanCode As String, ByVal CustomerCode As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��� �������� ������ ���� �� ���� �� ���� ������, ��� � ��������, �� ������� ��������� ������
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim CustomerCC As String
        Dim SalesmanCC As String
        Dim ExclQTY As Integer

        '---�������� - �� � ����������� �� ��������
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_Sales_Groups_CC_Exclude "
        MySQLStr = MySQLStr & "WHERE (Upper(UserName) = N'" & Declarations.UserCode.ToUpper() & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CheckSalesman = "���������� ���������� - �������� �������� � ���������� ��� ��� "
            Exit Function
        Else
            ExclQTY = Declarations.MyRec.Fields("CC").Value
        End If

        If ExclQTY > 0 Then '---�������� � ����������� �� ��������
            CheckSalesman = ""
            Exit Function
        End If

        If Trim(SalesmanCode) = "" Or Trim(CustomerCode) = "" Then
            CheckSalesman = ""
            Exit Function
        Else
            '--CC ��������
            MySQLStr = "SELECT SUBSTRING(ST01021, 7, 3) AS CC "
            MySQLStr = MySQLStr & "FROM ST010300 "
            MySQLStr = MySQLStr & "WHERE (ST01001 = N'" & Trim(SalesmanCode) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                SalesmanCC = "CC �������� �� ������"
            Else
                SalesmanCC = Declarations.MyRec.Fields("CC").Value.ToString
                trycloseMyRec()
            End If

            '--CC �������
            MySQLStr = "SELECT     SUBSTRING(ST010300.ST01021, 7, 3) AS CC "
            MySQLStr = MySQLStr & "FROM SL010300 INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON SL010300.SL01035 = ST010300.ST01001 "
            MySQLStr = MySQLStr & "WHERE (SL010300.SL01001 = N'" & Trim(CustomerCode) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                CustomerCC = "CC ������� �� ������"
            Else
                CustomerCC = Declarations.MyRec.Fields("CC").Value.ToString
                trycloseMyRec()
            End If

            '---������
            If SalesmanCC <> CustomerCC And CustomerCC <> "CC ������� �� ������" Then
                If My.Settings.CheckCC.ToUpper() = "��" Then
                    CheckSalesman = "���� ����� �������� ������ ���� ����� ��, ��� ���� ����� ��������, �� ������� ��������� ������. "
                    CheckSalesman = CheckSalesman & "���� ����� ��������: " & SalesmanCC & " ���� ����� �������: " & CustomerCC
                Else
                    CheckSalesman = ""
                End If
                Exit Function
            Else
                CheckSalesman = ""
                Exit Function
            End If
        End If
    End Function
End Module
