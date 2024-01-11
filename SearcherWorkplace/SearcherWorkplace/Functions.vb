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
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SQLCLS"
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

    Public Function UpdateRequestDataFromExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ����� ����������� �� Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim appXLSRC As Object
        Dim MyExcelCounter As Double                    '������� ����� Excel
        Dim MyItemCode As String                        '��� ������ Scala
        Dim MySuppItemCode As String                    '��� ������ ����������
        Dim MyItemName As String                        '�������� ������ � �������
        Dim MyItemCounter As Double                     '������� ��� - �� ��������� ����� ������ 
        Dim MyUOM As Integer                            '������� ���������
        Dim c As Object
        Dim MyQTY As Double
        Dim MyPrice As Double
        Dim MyWeekQTY As Double
        Dim MyCurr As Integer                           '��� ������

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(Declarations.ImportFileName)

        ExcelVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
        If CheckVersion(ExcelVersion) = True Then
            MyExcelCounter = 11
            MyCurr = 0

            While Not appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing _
                Or Not appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing

                '------��� ������ Scala
                If appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value = Nothing Then
                    MyItemCode = ""
                Else
                    If Trim(Declarations.MySupplierCode).Equals("") Then
                        MsgBox("������ B" & MyExcelCounter & ". �� ��������� �������� ��� ������ � Scala ��� ���������� " & Trim(Declarations.MySupplierName) & _
                            ", ������� � Scala �� �������. ������� �������� � Scala ���������� ��� �������� � ������ ����������� ��� ������ ���������� �� ���������� �� SCala.", _
                            MsgBoxStyle.Critical, "��������!")
                        Exit While
                    Else
                        MyItemCode = appXLSRC.Worksheets(1).Range("B" & MyExcelCounter).Value.ToString
                        '---�������� ��� ����� ��� ���� � Scala � ���������� ����������
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                        MySQLStr = MySQLStr & "AND (SC01058 = N'" & Trim(Declarations.MySupplierCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("������ B" & MyExcelCounter & ". ���������� �������� ���������� � ������� �� Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyItemCounter = Declarations.MyRec.Fields("CC").Value
                            trycloseMyRec()
                        End If
                        If MyItemCounter = 0 Then
                            MsgBox("������ B" & MyExcelCounter & ". ������ ���� ������ � ���������� " & Trim(Declarations.MySupplierName) & " � Scala ���.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        Else
                        End If
                    End If
                End If

                '------��� ������ ����������
                If appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value = Nothing Then
                    MySuppItemCode = ""
                Else
                    MySuppItemCode = appXLSRC.Worksheets(1).Range("C" & MyExcelCounter).Value.ToString
                End If

                '------�������� ������
                If appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value = Nothing Then
                    MyItemName = ""
                Else
                    MyItemName = appXLSRC.Worksheets(1).Range("D" & MyExcelCounter).Value.ToString
                End If

                '-------�������� ���������� ������ 3 - � �����
                If Trim(MyItemCode) = "" And (Trim(MySuppItemCode) = "" Or Trim(MyItemName) = "") Then
                    MsgBox("������ " & MyExcelCounter & " ����������� ������ ���� ������� ��� ��� ������ Scala, ��� ��� ������ ���������� + �������� ������!", MsgBoxStyle.Critical, "��������!")
                    Exit While
                Else

                    '------������� ���������
                    If Trim(MyItemCode).Equals("") Then
                        c = appXLSRC.Worksheets(1).Range("O11:O18").Find(appXLSRC.Worksheets(1).Range("E" & MyExcelCounter).Value, LookIn:=-4163)
                        If c.Text = "" Then
                            MsgBox("������ E" & MyExcelCounter & " �������� '������� ���������' �� ������� �� ������ �����. ���������� ������� �������� �� ����������� ������.", MsgBoxStyle.YesNo, "��������!")
                            Exit While
                        Else
                            MyUOM = appXLSRC.Worksheets(1).Rows(c.Row).Columns(c.Column - 1).Value
                        End If
                    Else
                        MySQLStr = "SELECT SC01135 AS UOM "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & Trim(MyItemCode) & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                            trycloseMyRec()
                            MsgBox("������ B" & MyExcelCounter & ". ���������� �������� ���������� �� ������� ��������� ����� ������ �� Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        Else
                            Declarations.MyRec.MoveFirst()
                            MyUOM = Declarations.MyRec.Fields("UOM").Value
                            trycloseMyRec()
                        End If
                    End If

                    '------����������
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("������ F" & MyExcelCounter & ". ���������� ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                        Exit While
                    Else
                        MyQTY = appXLSRC.Worksheets(1).Range("F" & MyExcelCounter).Value
                        If MyQTY = 0 Then
                            MsgBox("������ F" & MyExcelCounter & ". ���������� ������ ���� ������ ����.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        End If
                    End If
                    '------���������� ���� ��� ���
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("������ G" & MyExcelCounter & ". ���������� ���� ��� ��� ������ ���� ���������.", MsgBoxStyle.Critical, "��������!")
                        Exit While
                    Else
                        MyPrice = appXLSRC.Worksheets(1).Range("G" & MyExcelCounter).Value
                        If MyPrice = 0 Then
                            MsgBox("������ G" & MyExcelCounter & ". ���������� ���� ��� ��� ������ ���� ������ ����.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        End If
                    End If
                    '----------���� ��������
                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value) Is Double) Then
                        MsgBox("������ I" & MyExcelCounter & ". ���� �������� ������ ���� ��������.", MsgBoxStyle.Critical, "��������!")
                        Exit While
                    Else
                        MyWeekQTY = appXLSRC.Worksheets(1).Range("I" & MyExcelCounter).Value
                        If MyWeekQTY = 0 Then
                            MsgBox("������ I" & MyExcelCounter & ". ���� �������� ������ ���� ������ ����.", MsgBoxStyle.Critical, "��������!")
                            Exit While
                        End If
                    End If

                    '-----���������� ����������
                    ''-----�� ���� ������ �������������
                    'If MySuppItemCode.Equals("") = False Then
                    '    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    '    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    '    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    '    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    '    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    '    MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                    '    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    '    InitMyConn(False)
                    '    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                    ''-----�� ��������  ������
                    'If MyItemName.Equals("") = False Then
                    '    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    '    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    '    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    '    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    '    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    '    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    '    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    '    MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                    '    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    '    InitMyConn(False)
                    '    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                    '-----�� ���� ������ ������������� + ��������  ������
                    'If MySuppItemCode.Equals("") = False Then
                    MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                    MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                    MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                    MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                    MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                    MySQLStr = MySQLStr & ",CurrCode = 0 "
                    MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                    MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                    MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                    MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                    MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    'End If

                End If
                    MyExcelCounter = MyExcelCounter + 1
            End While
        End If
        '------------�������� ��������
        appXLSRC.DisplayAlerts = 0
        appXLSRC.Workbooks.Close()
        appXLSRC.DisplayAlerts = 1
        appXLSRC = Nothing
        MsgBox("��������� ���������� ������� ���������.", vbOKOnly, "��������!")
    End Function

    Public Function UpdateRequestDataFromLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ����� ����������� �� LibreOffice
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyItemCode As String                        '��� ������ Scala
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

        MyExcelCounter = 11
        While oSheet.getCellRangeByName("B" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("C" & MyExcelCounter).String.Equals("") = False Or _
            oSheet.getCellRangeByName("D" & MyExcelCounter).String.Equals("") = False
            '------��� ������ Scala
            MyItemCode = Trim(oSheet.getCellRangeByName("B" & MyExcelCounter).String)
            If MyItemCode.Equals("") = False Then
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
                    MyItemCode = ""
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
            MySuppItemCode = Trim(oSheet.getCellRangeByName("C" & MyExcelCounter).String)
            '------�������� ������
            MyItemName = Trim(oSheet.getCellRangeByName("D" & MyExcelCounter).String)
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
                '-----���������� ����������
                '-----�� ���� ������ ������������� + ��������  ������
                'If MySuppItemCode.Equals("") = False Then
                MySQLStr = "UPDATE tbl_SupplSearch_PropItems "
                MySQLStr = MySQLStr & "SET ItemCode = N'" & Replace(Trim(MyItemCode), "'", "''") & " '"
                MySQLStr = MySQLStr & ", UOM = " & MyUOM.ToString
                MySQLStr = MySQLStr & ", QTY = " & Replace(Replace(Replace(MyQTY.ToString, ",", "."), " ", ""), Chr(160), "")
                MySQLStr = MySQLStr & ", Price = " & Replace(Replace(Replace(MyPrice.ToString, ",", "."), " ", ""), Chr(160), "")
                MySQLStr = MySQLStr & ",CurrCode = 0 "
                MySQLStr = MySQLStr & ", LeadTimeWeek = " & Replace(Replace(Replace(MyWeekQTY.ToString, ",", "."), " ", ""), Chr(160), "") & " "
                MySQLStr = MySQLStr & "WHERE (SupplSearchID = " & Declarations.MyRequestNum.ToString & ") "
                MySQLStr = MySQLStr & "AND (ItemSuppCode = N'" & Replace(Trim(MySuppItemCode), "'", "''") & "')"
                MySQLStr = MySQLStr & "AND (ItemName  = N'" & Replace(Trim(MyItemName), "'", "''") & "') "
                MySQLStr = MySQLStr & "AND (SupplierID = " & Trim(Declarations.MySupplierID) & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                'End If
            End If

            MyExcelCounter = MyExcelCounter + 1
        End While
        oWorkBook.Close(True)
        MsgBox("��������� ���������� ������� ���������.", vbOKOnly, "��������!")
    End Function

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

    Public Sub SendInfoByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewRequestState As String, ByVal SearcherName As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� ����� ����������� ��������� �� ���������� � ��������
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
        MyMsgStr = MyMsgStr & "���������: " & SearcherName & Chr(13)
        MyMsgStr = MyMsgStr & "����� ��������� �������: " & NewRequestState & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "� ���������," & Chr(13)
        MyMsgStr = MyMsgStr + "��� ""��������"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. �� ������ ������� �� ��������, ��� �������������� ��������. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub

    Public Sub SendCommentByEmail(ByVal RequestID As Integer, ByVal RequestDate As String, ByVal EMailTo As String, ByVal ClientName As String, _
        ByVal SalesmanName As String, ByVal NewComment As String, ByVal SearcherName As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� ����� ����������� ��������� � ����� �����������
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
        msg.Subject = "����������� � ����� ����������� � ������� �� �����"
        MyMsgStr = "��������� �������!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "������� ����� ����������� � ���������� ������� �� �����:" & Chr(13)
        MyMsgStr = MyMsgStr & "ID: " & CStr(RequestID) & " �� " & RequestDate & Chr(13)
        MyMsgStr = MyMsgStr & "��� �������: " & ClientName & Chr(13)
        MyMsgStr = MyMsgStr & "��������: " & SalesmanName & Chr(13)
        MyMsgStr = MyMsgStr & "���������: " & SearcherName & Chr(13)
        MyMsgStr = MyMsgStr & "����� ����������� � �������: " & NewComment & Chr(10) & Chr(13)

        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "� ���������," & Chr(13)
        MyMsgStr = MyMsgStr + "��� ""��������"". " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr + "P.S. �� ������ ������� �� ��������, ��� �������������� ��������. "
        msg.Body = MyMsgStr

        smtp.Send(msg)
    End Sub
End Module
