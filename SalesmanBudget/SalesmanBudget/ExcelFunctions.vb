Module ExcelFunctions
    Public Sub UploadToExcel()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � Excel
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 3
        MyWRKBook = MyObj.Workbooks.Add
        MyWRKBook.Sheets(1).Name = "�������� ����"
        MyWRKBook.Sheets(2).Name = "��������� ������"
        MyWRKBook.Sheets(3).Name = "�������������"


        MyWRKBook.Sheets(3).Select()
        '----------------������ ������ - ����������� ��������----------------------
        MyWRKBook.Sheets(3).Cells.Locked = True

        '----------------�������� ������������� � 3 ����---------------------------
        UploadClassification(MyWRKBook)
        MyWRKBook.Sheets(3).Range("A1:A1").Select()


        MyWRKBook.Sheets(2).Select()
        '----------------�������� ������ ��������� �� 2 ����-----------------------
        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            UploadCommonHeader(MyWRKBook, "������ ��� ���������� ������� �� ��������")
        Else
            UploadCommonHeader(MyWRKBook, "������ ��� ���������� ������� �� ���� ������")
        End If

        '----------------������ ������ - ����������� ��������----------------------
        MyWRKBook.Sheets(2).Cells.Locked = True

        '----------------�������� ��������� ����� ������ �� 2 ����-----------------
        i = 5
        UploadTotalHeader(MyWRKBook, i)

        '----------------�������� ��������� 1 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1(MyWRKBook, i, "������� ���������� �������")

        '----------------�������� ����������� 1 ������� �� 2 ����------------------
        i = i + 1
        UploadActiveSales(MyWRKBook, i)


        '----------------�������� ��������� 2 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1(MyWRKBook, i, "�� ���������� �������")

        '----------------�������� ����������� 2 ������� �� 2 ����------------------
        i = i + 1
        UploadPassiveSales(MyWRKBook, i)

        '----------------�������� ��������� 3 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1(MyWRKBook, i, "����� �������")


        '----------------�������� ����������� 3 ������� �� 2 ����------------------
        i = i + 1
        UploadNewSales(MyWRKBook, i)

        '----------------������� ����� �� 2 ����-----------------------------------
        UploadFormulas2Sheet(MyWRKBook)
        MyWRKBook.Sheets(2).Range("A1:A1").Select()


        MyWRKBook.Sheets(1).Select()
        '----------------�������� ������ ��������� � 1 ����------------------------
        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            UploadCommonHeader(MyWRKBook, "���������� ������ ������� ��������")
        Else
            UploadCommonHeader(MyWRKBook, "���������� ������ ������� ���� ������")
        End If

        '----------------������ ������ - ����������� ��������----------------------
        MyWRKBook.Sheets(1).Cells.Locked = True

        '----------------�������� ��������� ��������������� ������ � 1 ����--------
        i = 5
        UploadGroupHeader(MyWRKBook, i)

        '----------------������� ����� �� 1 ����-----------------------------------
        UploadFormulas1Sheet(MyWRKBook, i)
        MyWRKBook.Sheets(1).Range("A1:A1").Select()

        '----------------������ ������ - �������� �������--------------------------
        PasswordProtectON(MyWRKBook)


        MyWRKBook.Sheets(2).Select()
        MyWRKBook.Sheets(2).Range("A1:A1").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Public Sub UploadToLO()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ � Libre Office
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer

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

        oWorkBook.getSheets().insertNewByName("�������� ����", 0)
        oWorkBook.getSheets().insertNewByName("��������� ������", 1)
        oWorkBook.getSheets().insertNewByName("�������������", 2)
        oWorkBook.getSheets().removeByName("����1")

        '=====================���� 3 �������������=================================
        oSheet = oWorkBook.getSheets().getByName("�������������")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        '----------------������ ������ - ����������� ��������----------------------
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '----------------�������� ������������� � 3 ����---------------------------
        UploadClassificationLO(oSheet, oServiceManager, oWorkBook, oDispatcher)
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        LOPasswordProtect(oSheet, "!pass2022", True)

        '=====================���� 2 ��������� ������==============================
        oSheet = oWorkBook.getSheets().getByName("��������� ������")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        '----------------�������� ������ ��������� �� 2 ����-----------------------
        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            UploadCommonHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, "������ ��� ���������� ������� �� ��������")
        Else
            UploadCommonHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, "������ ��� ���������� ������� �� ���� ������")
        End If
        '----------------������ ������ - ����������� ��������----------------------
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '----------------�������� ��������� ����� ������ �� 2 ����-----------------
        i = 5
        UploadTotalHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        '----------------�������� ��������� 1 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1LO(oSheet, oServiceManager, oWorkBook, oDispatcher, i, "������� ���������� �������")
        '----------------�������� ����������� 1 ������� �� 2 ����------------------
        i = i + 1
        UploadActiveSalesLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        '----------------�������� ��������� 2 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1LO(oSheet, oServiceManager, oWorkBook, oDispatcher, i, "�� ���������� �������")
        '----------------�������� ����������� 2 ������� �� 2 ����------------------
        i = i + 1
        UploadPassiveSalesLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        '----------------�������� ��������� 3 ������� �� 2 ����--------------------
        i = i + 2
        UploadHeader1LO(oSheet, oServiceManager, oWorkBook, oDispatcher, i, "����� �������")
        '----------------�������� ����������� 3 ������� �� 2 ����------------------
        i = i + 1
        UploadNewSalesLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        '----------------������� ����� �� 2 ����-----------------------------------
        UploadFormulas2SheetLO(oSheet, oServiceManager, oWorkBook, oDispatcher)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        LOPasswordProtect(oSheet, "!pass2022", True)

        '=====================���� 1 ����� ������==================================
        oSheet = oWorkBook.getSheets().getByName("�������� ����")
        oWorkBook.CurrentController.setActiveSheet(oSheet)

        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            UploadCommonHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, "���������� ������ ������� ��������")
        Else
            UploadCommonHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, "����������� ������ ������� ���� ������")
        End If
        '----------------������ ������ - ����������� ��������----------------------
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '----------------�������� ��������� ��������������� ������ � 1 ����--------
        i = 5
        UploadGroupHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        '----------------������� ����� �� 1 ����-----------------------------------
        UploadFormulas1SheetLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        LOPasswordProtect(oSheet, "!pass2022", True)

        oSheet = oWorkBook.getSheets().getByName("��������� ������")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub UploadClassification(ByRef MyWRKBook As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---------------���������------------
        MyWRKBook.ActiveSheet.Range("A1:B1").MergeCells = True
        MyWRKBook.ActiveSheet.Range("A1:B1") = "�������"
        MyWRKBook.ActiveSheet.Range("C1:D1").MergeCells = True
        MyWRKBook.ActiveSheet.Range("C1:D1") = "��� �������"
        MyWRKBook.ActiveSheet.Range("E1:F1").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E1:F1") = "�����"
        MyWRKBook.ActiveSheet.Range("G1:H1").MergeCells = True
        MyWRKBook.ActiveSheet.Range("G1:H1") = "IKA"

        MyWRKBook.ActiveSheet.Range("A1:H1").Select()
        MyWRKBook.ActiveSheet.Range("A1:H1").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A1:H1").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A1:H1").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:H1").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:H1").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:H1").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:H1").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:H1").Interior
            .ColorIndex = 36
            .TintAndShade = 0.7
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A1:H1").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A1:H1").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A1:H1").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 72
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 85
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 17
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 33

        '-------------������ ��������-------------------------------------------
        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM tbl_RexelIndustry "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            IndustryQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            IndustryQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("A2").CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        '-------------������ ����� ��������-------------------------------------
        MySQLStr = "SELECT RCGCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelCustomerGroup "
        MySQLStr = MySQLStr & "ORDER BY RCGCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            TypeQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            TypeQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("C2").CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        '-------------������ ������---------------------------------------------
        MySQLStr = "SELECT EMCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelEndMarkets "
        MySQLStr = MySQLStr & "ORDER BY EMCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MarketQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MarketQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("E2").CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        '-------------������ ����� IKA ��������---------------------------------
        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM tbl_RexelIKATypes "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            IKAQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            IKAQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("G2").CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Sub

    Private Sub UploadClassificationLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim oFrame As Object
        Dim i As Integer

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 1400
        oSheet.getColumns().getByName("B").Width = 14400
        oSheet.getColumns().getByName("C").Width = 1400
        oSheet.getColumns().getByName("D").Width = 17000
        oSheet.getColumns().getByName("E").Width = 1400
        oSheet.getColumns().getByName("F").Width = 3400
        oSheet.getColumns().getByName("G").Width = 1400
        oSheet.getColumns().getByName("H").Width = 6600

        '-----���
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "", RGB(204, 204, 204)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        '---------------���������------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "A1:B1")
        oSheet.getCellRangeByName("A1").String = "�������"
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "C1:D1")
        oSheet.getCellRangeByName("C1").String = "��� �������"
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E1:F1")
        oSheet.getCellRangeByName("E1").String = "�����"
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "G1:H1")
        oSheet.getCellRangeByName("G1").String = "IKA"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:H1", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B15:N15", 11)
        oSheet.getCellRangeByName("A1:H1").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("A1:H1").TopBorder = LineFormat
        oSheet.getCellRangeByName("A1:H1").RightBorder = LineFormat
        oSheet.getCellRangeByName("A1:H1").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A1:H1").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A1:H1").VertJustify = 2
        oSheet.getCellRangeByName("A1:H1").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A1:H1")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1:H1")

        '-------------������ ��������-------------------------------------------
        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIndustry "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            IndustryQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            IndustryQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 2
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("ID").Value
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("Name").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If

        '-------------������ ����� ��������-------------------------------------
        MySQLStr = "SELECT RCGCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelCustomerGroup "
        MySQLStr = MySQLStr & "ORDER BY RCGCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            TypeQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            TypeQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 2
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("RCGCode").Value
                oSheet.getCellRangeByName("D" & CStr(i)).String = Declarations.MyRec.Fields("RussianName").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If

        '-------------������ ������---------------------------------------------
        MySQLStr = "SELECT EMCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelEndMarkets "
        MySQLStr = MySQLStr & "ORDER BY EMCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MarketQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MarketQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 2
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("E" & CStr(i)).String = Declarations.MyRec.Fields("EMCode").Value
                oSheet.getCellRangeByName("F" & CStr(i)).String = Declarations.MyRec.Fields("RussianName").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If

        '-------------������ ����� IKA ��������---------------------------------
        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIKATypes "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            IKAQTY = 0
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            IKAQTY = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 2
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("G" & CStr(i)).String = Declarations.MyRec.Fields("ID").Value
                oSheet.getCellRangeByName("H" & CStr(i)).String = Declarations.MyRec.Fields("Name").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()
        End If
    End Sub

    Private Sub UploadCommonHeader(ByRef MyWRKBook As Object, ByVal HdrTxt As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A1") = HdrTxt
        MyWRKBook.ActiveSheet.Range("A1").Font.Size = 16
        MyWRKBook.ActiveSheet.Range("A1").Font.Bold = True
        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            MyWRKBook.ActiveSheet.Range("A2") = "�������� " & MainForm.ComboBox2.Text
        Else
            MyWRKBook.ActiveSheet.Range("A2") = "���� ����� " & MainForm.ComboBox2.Text
        End If
        MyWRKBook.ActiveSheet.Range("A2").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("A2").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3") = "�� " & CStr(CInt(MainForm.ComboBox1.Text) + 1) & " ���"
        MyWRKBook.ActiveSheet.Range("A3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("A3").Font.Bold = False

    End Sub

    Private Sub UploadCommonHeaderLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal HdrTxt As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        oSheet.getCellRangeByName("A1").String = HdrTxt
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 16)

        If MainForm.RadioButton1.Checked = True Then    '---������ �� ��������
            oSheet.getCellRangeByName("A2").String = "�������� " & MainForm.ComboBox2.Text
        Else
            oSheet.getCellRangeByName("A2").String = "���� ����� " & MainForm.ComboBox2.Text
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A2", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A2", 10)

        oSheet.getCellRangeByName("A3").String = "�� " & CStr(Now.Year + 1) & " ���"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A3", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A3", 10)
    End Sub

    Private Sub UploadHeader1(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal HdrTxt As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� ������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = HdrTxt
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Size = 10
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Color = -16777024
        UploadSubHeader(MyWRKBook, i)

        i = i + 1
        '---�������� ���������------
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "��� �������"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "������"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "����� �������"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "������� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "�����% ��� ��������"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "����� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "��������� �������"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "�����% ��������� ������"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "����� ��������� ������"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "������� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = "�����% ��� ��������"
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = "����� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("M" & CStr(i)) = "��������� �������"
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("N" & CStr(i)) = "�����% ��������� ������"
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("O" & CStr(i)) = "����� ��������� ������"
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 13
        MyWRKBook.ActiveSheet.Range("P" & CStr(i)) = "������� �� ��� �����"
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Range("Q" & CStr(i)) = "������� ��������������"
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Range("R" & CStr(i)) = "��� �������"
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Range("S" & CStr(i)) = "�����"
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 17
        MyWRKBook.ActiveSheet.Range("T" & CStr(i)) = "IKA"
        MyWRKBook.ActiveSheet.Columns("T:T").ColumnWidth = 33
        MyWRKBook.ActiveSheet.Range("U" & CStr(i)) = "������� �� ��� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("U:U").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Range("V" & CStr(i)) = "��������� ������� (���)"
        MyWRKBook.ActiveSheet.Columns("V:V").ColumnWidth = 15

        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":V" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With
    End Sub

    Private Sub UploadHeader1LO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer, ByVal HdrTxt As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� ������� � LibreOffice 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame

        oSheet.getCellRangeByName("A" & CStr(i)).String = HdrTxt
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), 10)
        oSheet.getCellRangeByName("A" & CStr(i)).CharColor = RGB(0, 0, 192)
        UploadSubHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        i = i + 1

        '---�������� ���������------
        oSheet.getCellRangeByName("A" & CStr(i)).String = "��� �������"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "������"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "����� �������"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "������� ��� ��������"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "�����% ��� ��������"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "����� ��� ��������"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "��������� �������"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "�����% ��������� ������"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "����� ��������� ������"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "������� ��� ��������"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "�����% ��� ��������"
        oSheet.getCellRangeByName("L" & CStr(i)).String = "����� ��� ��������"
        oSheet.getCellRangeByName("M" & CStr(i)).String = "��������� �������"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "�����% ��������� ������"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "����� ��������� ������"
        oSheet.getCellRangeByName("P" & CStr(i)).String = "������� �� ��� �����"
        oSheet.getCellRangeByName("Q" & CStr(i)).String = "������� ��������������"
        oSheet.getCellRangeByName("R" & CStr(i)).String = "��� �������"
        oSheet.getCellRangeByName("S" & CStr(i)).String = "�����"
        oSheet.getCellRangeByName("T" & CStr(i)).String = "IKA"
        oSheet.getCellRangeByName("U" & CStr(i)).String = "������� �� ��� ��� ��������"
        oSheet.getCellRangeByName("V" & CStr(i)).String = "��������� ������� (���)"

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":V" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":V" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":V" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":V" & CStr(i)).CellBackColor = RGB(174, 249, 255)  '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":V" & CStr(i), 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("A" & CStr(i) & ":V" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":V" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":V" & CStr(i))

    End Sub

    Private Sub UploadSubHeader(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� ����� ��������� ������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '---������------------------
        '---1----------
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)) = "���� " & CStr(CInt(MainForm.ComboBox1.Text) + 1)

        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Interior
            .ColorIndex = 5
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":I" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)) = "���� " & MainForm.ComboBox1.Text

        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":O" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        i = i + 1
        '---2-----------------------
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)) = "����������� �������"

        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Interior
            .ColorIndex = 5
            .TintAndShade = 0.7
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":F" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)) = "��������� �������"

        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Interior
            .ColorIndex = 5
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":I" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)) = "����������� �������"

        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.7
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":L" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).MergeCells = True
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)) = "��������� �������"

        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":O" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With
    End Sub

    Private Sub UploadSubHeaderLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� ����� ��������� ������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame

        '---������------------------
        '---1----------
        '---����
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("D" & CStr(i)).String = "���� " & CStr(Now.Year + 1)
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":I" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":I" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":I" & CStr(i), 8)
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).CellBackColor = RGB(255, 204, 204)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("D" & CStr(i) & ":I" & CStr(i)).HoriJustify = 2
        '---����
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":O" & CStr(i))
        oSheet.getCellRangeByName("J" & CStr(i)).String = "���� " & MainForm.ComboBox1.Text
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":O" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":O" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":O" & CStr(i), 8)
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).CellBackColor = RGB(179, 255, 179)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("J" & CStr(i) & ":O" & CStr(i)).HoriJustify = 2

        i = i + 1
        '---2-----------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":F" & CStr(i))
        oSheet.getCellRangeByName("D" & CStr(i)).String = "����������� �������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":F" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":F" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":F" & CStr(i), 8)
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).CellBackColor = RGB(255, 179, 179)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("D" & CStr(i) & ":F" & CStr(i)).HoriJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("G" & CStr(i)).String = "��������� �������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":I" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":I" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":I" & CStr(i), 8)
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).CellBackColor = RGB(255, 230, 230)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("G" & CStr(i) & ":I" & CStr(i)).HoriJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":L" & CStr(i))
        oSheet.getCellRangeByName("J" & CStr(i)).String = "����������� �������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":L" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":L" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "J" & CStr(i) & ":L" & CStr(i), 8)
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).CellBackColor = RGB(140, 255, 140)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("J" & CStr(i) & ":L" & CStr(i)).HoriJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":O" & CStr(i))
        oSheet.getCellRangeByName("M" & CStr(i)).String = "��������� �������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":O" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":O" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":O" & CStr(i), 8)
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).CellBackColor = RGB(217, 255, 217)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("M" & CStr(i) & ":O" & CStr(i)).HoriJustify = 2
    End Sub

    Private Sub UploadTotalHeader(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� � ������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '--------������� ���������---------------------------------------
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "�������� ����������"
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Size = 12
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Color = -11489280
        UploadSubHeader(MyWRKBook, i)
        i = i + 1

        '--------�������� ���������-------------------------------------
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "������� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "�����% ��� ��������"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "����� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "��������� �������"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "�����% ��������� ������"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "����� ��������� ������"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "������� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = "�����% ��� ��������"
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = "����� ��� ��������"
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("M" & CStr(i)) = "��������� �������"
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("N" & CStr(i)) = "�����% ��������� ������"
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("O" & CStr(i)) = "����� ��������� ������"

        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With
        i = i + 1

        StrTotalStart = i
        '--------������� ���������--------------------------------------
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)) = "������� ���������� �������"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 1) & ":C" & CStr(i + 1)) = "�� ���������� �������"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i + 2) & ":C" & CStr(i + 2)) = "����� �������"

        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Select()
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(6).LineStyle = -4142
        'MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Interior
            .Color = -16727809
            .TintAndShade = 0.7
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        '------------����� � ����-----------------------------------
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Select()
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 3)).NumberFormat = "#,##0.00"

        i = i + 3

        '----------�����--------------------------------------------
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)) = "�����"

        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(6).LineStyle = -4142
        'MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Interior
            .Color = -16727809
            .TintAndShade = 0.4
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i + 2)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).Interior
            .Color = -16727809
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '---------������� "����� �����"--------------------------------
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":O" & CStr(i)).FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)"
        i = i + 1
    End Sub

    Private Sub UploadTotalHeaderLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ��������� � ������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 2400
        oSheet.getColumns().getByName("B").Width = 5000
        oSheet.getColumns().getByName("C").Width = 8000
        oSheet.getColumns().getByName("D").Width = 2400
        oSheet.getColumns().getByName("E").Width = 2400
        oSheet.getColumns().getByName("F").Width = 2400
        oSheet.getColumns().getByName("G").Width = 2400
        oSheet.getColumns().getByName("H").Width = 2400
        oSheet.getColumns().getByName("I").Width = 2400
        oSheet.getColumns().getByName("J").Width = 2400
        oSheet.getColumns().getByName("K").Width = 2400
        oSheet.getColumns().getByName("L").Width = 2400
        oSheet.getColumns().getByName("M").Width = 2400
        oSheet.getColumns().getByName("N").Width = 2400
        oSheet.getColumns().getByName("O").Width = 2400
        oSheet.getColumns().getByName("P").Width = 3000
        oSheet.getColumns().getByName("Q").Width = 8000
        oSheet.getColumns().getByName("R").Width = 8000
        oSheet.getColumns().getByName("S").Width = 3400
        oSheet.getColumns().getByName("T").Width = 6600
        oSheet.getColumns().getByName("U").Width = 3000
        oSheet.getColumns().getByName("V").Width = 3000

        '-----���
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "", RGB(204, 204, 204)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        '--------������� ���������---------------------------------------
        oSheet.getCellRangeByName("A" & CStr(i)).String = "�������� ����������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), 12)
        oSheet.getCellRangeByName("A" & CStr(i)).CharColor = RGB(0, 176, 80)
        UploadSubHeaderLO(oSheet, oServiceManager, oWorkBook, oDispatcher, i)
        i = i + 1

        '--------�������� ���������-------------------------------------
        oSheet.getCellRangeByName("D" & CStr(i)).String = "������� ��� ��������"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "�����% ��� ��������"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "����� ��� ��������"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "��������� �������"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "�����% ��������� ������"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "����� ��������� ������"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "������� ��� ��������"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "�����% ��� ��������"
        oSheet.getCellRangeByName("L" & CStr(i)).String = "����� ��� ��������"
        oSheet.getCellRangeByName("M" & CStr(i)).String = "��������� �������"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "�����% ��������� ������"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "����� ��������� ������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i))
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).TopBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).RightBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).BottomBorder = LineFormat
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("D" & CStr(i) & ":O" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i))
        i = i + 1

        StrTotalStart = i
        '--------������� ���������--------------------------------------
        oSheet.getCellRangeByName("C" & CStr(i)).String = "������� ���������� �������"
        oSheet.getCellRangeByName("C" & CStr(i + 1)).String = "�� ���������� �������"
        oSheet.getCellRangeByName("C" & CStr(i + 2)).String = "����� �������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i + 2), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i + 2), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i + 2))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i + 2)).CellBackColor = RGB(232, 250, 255)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "C" & CStr(i) & ":C" & CStr(i + 2), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i + 2)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i + 2)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i + 2))

        '------------���� + ����� � ����--------------------------------
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & CStr(i) & ":O" & CStr(i + 2), RGB(232, 232, 232)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "D" & CStr(i) & ":I" & CStr(i + 2), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "J" & CStr(i) & ":O" & CStr(i + 2), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i + 2), 4)
        i = i + 3

        '----------�����--------------------------------------------
        oSheet.getCellRangeByName("C" & CStr(i)).String = "�����"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).CellBackColor = RGB(162, 232, 255)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "C" & CStr(i) & ":C" & CStr(i), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("C" & CStr(i) & ":C" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "C" & CStr(i) & ":C" & CStr(i))

        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & CStr(i) & ":O" & CStr(i), RGB(232, 250, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "D" & CStr(i) & ":I" & CStr(i), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "J" & CStr(i) & ":O" & CStr(i), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "D" & CStr(i) & ":O" & CStr(i), 4)

        '---------������� "����� �����"--------------------------------
        'LOSetFormula(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & CStr(i) & ":O" & CStr(i), "=SUM(R[-3]C:R[-1]C)", 2)
        oSheet.getCellRangeByName("D" & CStr(i)).FormulaLocal = "=SUM(D" & CStr(i - 3) & ":D" & CStr(i - 1)
        oSheet.getCellRangeByName("E" & CStr(i)).FormulaLocal = "=IF(D" & CStr(i) & "=0;0;F" & CStr(i) & "/D" & CStr(i) & "*100)"
        oSheet.getCellRangeByName("F" & CStr(i)).FormulaLocal = "=SUM(F" & CStr(i - 3) & ":F" & CStr(i - 1)
        oSheet.getCellRangeByName("G" & CStr(i)).FormulaLocal = "=SUM(G" & CStr(i - 3) & ":G" & CStr(i - 1)
        oSheet.getCellRangeByName("H" & CStr(i)).FormulaLocal = "=IF(G" & CStr(i) & "=0;0;I" & CStr(i) & "/G" & CStr(i) & "*100)"
        oSheet.getCellRangeByName("I" & CStr(i)).FormulaLocal = "=SUM(I" & CStr(i - 3) & ":I" & CStr(i - 1)
        oSheet.getCellRangeByName("J" & CStr(i)).FormulaLocal = "=SUM(J" & CStr(i - 3) & ":J" & CStr(i - 1)
        oSheet.getCellRangeByName("K" & CStr(i)).FormulaLocal = "=IF(J" & CStr(i) & "=0;0;L" & CStr(i) & "/J" & CStr(i) & "*100)"
        oSheet.getCellRangeByName("L" & CStr(i)).FormulaLocal = "=SUM(L" & CStr(i - 3) & ":L" & CStr(i - 1)
        oSheet.getCellRangeByName("M" & CStr(i)).FormulaLocal = "=SUM(M" & CStr(i - 3) & ":M" & CStr(i - 1)
        oSheet.getCellRangeByName("N" & CStr(i)).FormulaLocal = "=IF(M" & CStr(i) & "=0;0;O" & CStr(i) & "/M" & CStr(i) & "*100)"
        oSheet.getCellRangeByName("O" & CStr(i)).FormulaLocal = "=SUM(O" & CStr(i - 3) & ":O" & CStr(i - 1)
        i = i + 1
    End Sub

    Private Sub UploadGroupHeader(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ��� ��������������� ������ � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "����� ������ � ��������"
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Size = 14
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)).Font.Bold = True
        i = i + 1

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "�������"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 72
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "������� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "�� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "����� �������"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "�����"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 15

        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "��� �������"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 85
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "������� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "�� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "����� �������"
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = "�����"
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 15

        MyWRKBook.ActiveSheet.Range("M" & CStr(i)) = "IKA"
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 33
        MyWRKBook.ActiveSheet.Range("N" & CStr(i)) = "������� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("O" & CStr(i)) = "�� ���������� �������"
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("P" & CStr(i)) = "����� �������"
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Q" & CStr(i)) = "�����"
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 15

        '-------�������-------------------------
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Interior
            .ColorIndex = 5
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        '-------��� �������---------------------
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Interior
            .ColorIndex = 10
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        '-------IKA-------------------------
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + 2)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Interior
            .ColorIndex = 3
            .TintAndShade = 0.8
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With
        i = i + 1

    End Sub

    Private Sub UploadGroupHeaderLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ��� ��������������� ������ � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 13800
        oSheet.getColumns().getByName("B").Width = 2300
        oSheet.getColumns().getByName("C").Width = 2300
        oSheet.getColumns().getByName("D").Width = 2300
        oSheet.getColumns().getByName("E").Width = 2875
        oSheet.getColumns().getByName("F").Width = 1100
        oSheet.getColumns().getByName("G").Width = 13800
        oSheet.getColumns().getByName("H").Width = 2300
        oSheet.getColumns().getByName("I").Width = 2300
        oSheet.getColumns().getByName("J").Width = 2300
        oSheet.getColumns().getByName("K").Width = 2875
        oSheet.getColumns().getByName("L").Width = 1100
        oSheet.getColumns().getByName("M").Width = 6325
        oSheet.getColumns().getByName("N").Width = 2300
        oSheet.getColumns().getByName("O").Width = 2300
        oSheet.getColumns().getByName("P").Width = 2300
        oSheet.getColumns().getByName("Q").Width = 2875

        '-----���
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "", RGB(204, 204, 204)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����

        '--------������� ���������---------------------------------------
        oSheet.getCellRangeByName("A" & CStr(i)).String = "����� ������ � ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i), 14)
        oSheet.getCellRangeByName("A" & CStr(i)).CharColor = RGB(0, 0, 0) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        i = i + 1

        '--------
        oSheet.getCellRangeByName("A" & CStr(i)).String = "�������"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "������� ���������� �������"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "�� ���������� �������"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "����� �������"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "�����"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).CellBackColor = RGB(255, 204, 204)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":E" & CStr(i), 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":E" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i))

        oSheet.getCellRangeByName("G" & CStr(i)).String = "��� �������"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "������� ���������� �������"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "�� ���������� �������"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "����� �������"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "�����"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i))
        oSheet.getCellRangeByName("G" & CStr(i) & ":K" & CStr(i)).CellBackColor = RGB(179, 255, 179)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "G" & CStr(i) & ":K" & CStr(i), 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("G" & CStr(i) & ":K" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("G" & CStr(i) & ":K" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i))

        oSheet.getCellRangeByName("M" & CStr(i)).String = "IKA"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "������� ���������� �������"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "�� ���������� �������"
        oSheet.getCellRangeByName("P" & CStr(i)).String = "����� �������"
        oSheet.getCellRangeByName("Q" & CStr(i)).String = "�����"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i))
        oSheet.getCellRangeByName("M" & CStr(i) & ":Q" & CStr(i)).CellBackColor = RGB(204, 204, 255)    '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBorders(oServiceManager, oSheet, "M" & CStr(i) & ":Q" & CStr(i), 70, 70, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        oSheet.getCellRangeByName("M" & CStr(i) & ":Q" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("M" & CStr(i) & ":Q" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i))
        i = i + 1


    End Sub

    Private Sub UploadActiveSales(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� �������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        ActiveAreaStart = i
        'MySQLStr = "Exec spp_Sales_For_Budget_R2 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 1 "
        If MainForm.RadioButton1.Checked = True Then    '---�������������� �� ��������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 1, N'', 0 "
        Else                                            '---�������������� �� ���� ������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'', 1, N'" & MainForm.ComboBox2.SelectedValue & "', 1 "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            ActiveAreaFinish = i
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            ActiveAreaFinish = i + Declarations.MyRec.RecordCount - 1
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
        i = ActiveAreaFinish

        '------------------��������������-----------------------------------
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '--------����������� ��� ��������������-----------------------------------------
        With MyWRKBook.ActiveSheet.Range("D" & CStr(ActiveAreaStart) & ":E" & CStr(ActiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(ActiveAreaStart) & ":E" & CStr(ActiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("G" & CStr(ActiveAreaStart) & ":H" & CStr(ActiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("G" & CStr(ActiveAreaStart) & ":H" & CStr(ActiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("V" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("V" & CStr(ActiveAreaStart) & ":V" & CStr(ActiveAreaFinish)).Locked = False

        '--------------------�������------------------------------------------------------
        MyWRKBook.ActiveSheet.Range("F" & CStr(ActiveAreaStart) & ":F" & CStr(ActiveAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"
        MyWRKBook.ActiveSheet.Range("I" & CStr(ActiveAreaStart) & ":I" & CStr(ActiveAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"

        '--------------------�������� ������----------------------------------------------
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$B$2:$B$" & IndustryQTY + 1)
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$D$2:$D$" & TypeQTY + 1)
        MyWRKBook.ActiveSheet.Range("R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$F$2:$F$" & MarketQTY + 1)
        MyWRKBook.ActiveSheet.Range("S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$H$2:$H$" & IKAQTY + 1)
        MyWRKBook.ActiveSheet.Range("T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish)).Validation.ShowError = True
    End Sub

    Private Sub UploadActiveSalesLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� �������� �������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        ActiveAreaStart = i
        If MainForm.RadioButton1.Checked = True Then    '---�������������� �� ��������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 1, N'', 0 "
        Else                                            '---�������������� �� ���� ������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'', 1, N'" & MainForm.ComboBox2.SelectedValue & "', 1 "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            ActiveAreaFinish = i
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("CustomerCode").Value
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("CustomerName").Value
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("CustomerAddress").Value
                oSheet.getCellRangeByName("F" & CStr(i)).FormulaLocal = "=D" & CStr(i) & " * E" & CStr(i) & "/100 "
                oSheet.getCellRangeByName("I" & CStr(i)).FormulaLocal = "=G" & CStr(i) & " * H" & CStr(i) & "/100 "
                oSheet.getCellRangeByName("J" & CStr(i)).Value = Declarations.MyRec.Fields("VolumeNotProject").Value
                oSheet.getCellRangeByName("K" & CStr(i)).Value = Declarations.MyRec.Fields("MarginPerCentNotProject").Value
                oSheet.getCellRangeByName("L" & CStr(i)).Value = Declarations.MyRec.Fields("MarginNotProject").Value
                oSheet.getCellRangeByName("M" & CStr(i)).Value = Declarations.MyRec.Fields("VolumeProject").Value
                oSheet.getCellRangeByName("N" & CStr(i)).Value = Declarations.MyRec.Fields("MarginPerCentProject").Value
                oSheet.getCellRangeByName("O" & CStr(i)).Value = Declarations.MyRec.Fields("MarginProject").Value
                oSheet.getCellRangeByName("P" & CStr(i)).Value = Declarations.MyRec.Fields("TotalSales").Value
                oSheet.getCellRangeByName("Q" & CStr(i)).String = Declarations.MyRec.Fields("ClientIndustry").Value
                oSheet.getCellRangeByName("R" & CStr(i)).String = Declarations.MyRec.Fields("ClientType").Value
                oSheet.getCellRangeByName("S" & CStr(i)).String = Declarations.MyRec.Fields("ClientMarket").Value
                oSheet.getCellRangeByName("T" & CStr(i)).String = Declarations.MyRec.Fields("ClientIKA").Value
                oSheet.getCellRangeByName("U" & CStr(i)).Value = Declarations.MyRec.Fields("TotalYearSales").Value
                oSheet.getCellRangeByName("V" & CStr(i)).Value = Declarations.MyRec.Fields("Potencial").Value

                ActiveAreaFinish = i
                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        '------------------��������������-----------------------------------
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "A" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), 4)

        '--------����������� ��� ��������������-----------------------------------------
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & CStr(ActiveAreaStart) & ":E" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "G" & CStr(ActiveAreaStart) & ":H" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "Q" & CStr(ActiveAreaStart) & ":T" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "V" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "D" & CStr(ActiveAreaStart) & ":E" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "G" & CStr(ActiveAreaStart) & ":H" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "Q" & CStr(ActiveAreaStart) & ":T" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "V" & CStr(ActiveAreaStart) & ":V" & CStr(i - 1), False)

        '--------------------�������� ������----------------------------------------------
        LOSetValidation(oSheet, "Q" & CStr(ActiveAreaStart) & ":Q" & CStr(ActiveAreaFinish - 1), "=$�������������.$B$2:$B$" & IndustryQTY + 1)
        LOSetValidation(oSheet, "R" & CStr(ActiveAreaStart) & ":R" & CStr(ActiveAreaFinish - 1), "=$�������������.$D$2:$D$" & TypeQTY + 1)
        LOSetValidation(oSheet, "S" & CStr(ActiveAreaStart) & ":S" & CStr(ActiveAreaFinish - 1), "=$�������������.$F$2:$F$" & MarketQTY + 1)
        LOSetValidation(oSheet, "T" & CStr(ActiveAreaStart) & ":T" & CStr(ActiveAreaFinish - 1), "=$�������������.$H$2:$H$" & IKAQTY + 1)

    End Sub

    Private Sub UploadPassiveSales(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���������� �������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        PassiveAreaStart = i
        'MySQLStr = "Exec spp_Sales_For_Budget_R2 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 0 "
        If MainForm.RadioButton1.Checked = True Then    '---�������������� �� ��������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 0, N'', 0 "
        Else                                            '---�������������� �� ���� ������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'', 0, N'" & MainForm.ComboBox2.SelectedValue & "', 1 "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PassiveAreaFinish = i
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            PassiveAreaFinish = i + Declarations.MyRec.RecordCount - 1
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
        i = PassiveAreaFinish

        '------------------��������������-----------------------------------
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '--------����������� ��� ��������������-----------------------------------------
        With MyWRKBook.ActiveSheet.Range("D" & CStr(PassiveAreaStart) & ":E" & CStr(PassiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("D" & CStr(PassiveAreaStart) & ":E" & CStr(PassiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("G" & CStr(PassiveAreaStart) & ":H" & CStr(PassiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("G" & CStr(PassiveAreaStart) & ":H" & CStr(PassiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("V" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("V" & CStr(PassiveAreaStart) & ":V" & CStr(PassiveAreaFinish)).Locked = False

        '--------------------�������------------------------------------------------------
        MyWRKBook.ActiveSheet.Range("F" & CStr(PassiveAreaStart) & ":F" & CStr(PassiveAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"
        MyWRKBook.ActiveSheet.Range("I" & CStr(PassiveAreaStart) & ":I" & CStr(PassiveAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"

        '--------------------�������� ������----------------------------------------------
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$B$2:$B$" & IndustryQTY + 1)
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$D$2:$D$" & TypeQTY + 1)
        MyWRKBook.ActiveSheet.Range("R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$F$2:$F$" & MarketQTY + 1)
        MyWRKBook.ActiveSheet.Range("S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$H$2:$H$" & IKAQTY + 1)
        MyWRKBook.ActiveSheet.Range("T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish)).Validation.ShowError = True
    End Sub

    Private Sub UploadPassiveSalesLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���������� �������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame
        PassiveAreaStart = i
        'MySQLStr = "Exec spp_Sales_For_Budget_R2 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 0 "
        If MainForm.RadioButton1.Checked = True Then    '---�������������� �� ��������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'" & MainForm.ComboBox2.SelectedValue & "', 0, N'', 0 "
        Else                                            '---�������������� �� ���� ������
            MySQLStr = "Exec spp_Sales_For_Budget_R3 N'" & MainForm.ComboBox1.Text & "', N'', 0, N'" & MainForm.ComboBox2.SelectedValue & "', 1 "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            PassiveAreaFinish = i
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Not Declarations.MyRec.EOF
                oSheet.getCellRangeByName("A" & CStr(i)).String = Declarations.MyRec.Fields("CustomerCode").Value
                oSheet.getCellRangeByName("B" & CStr(i)).String = Declarations.MyRec.Fields("CustomerName").Value
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("CustomerAddress").Value
                oSheet.getCellRangeByName("F" & CStr(i)).FormulaLocal = "=D" & CStr(i) & " * E" & CStr(i) & "/100 "
                oSheet.getCellRangeByName("I" & CStr(i)).FormulaLocal = "=G" & CStr(i) & " * H" & CStr(i) & "/100 "
                oSheet.getCellRangeByName("J" & CStr(i)).Value = Declarations.MyRec.Fields("VolumeNotProject").Value
                oSheet.getCellRangeByName("K" & CStr(i)).Value = Declarations.MyRec.Fields("MarginPerCentNotProject").Value
                oSheet.getCellRangeByName("L" & CStr(i)).Value = Declarations.MyRec.Fields("MarginNotProject").Value
                oSheet.getCellRangeByName("M" & CStr(i)).Value = Declarations.MyRec.Fields("VolumeProject").Value
                oSheet.getCellRangeByName("N" & CStr(i)).Value = Declarations.MyRec.Fields("MarginPerCentProject").Value
                oSheet.getCellRangeByName("O" & CStr(i)).Value = Declarations.MyRec.Fields("MarginProject").Value
                oSheet.getCellRangeByName("P" & CStr(i)).Value = Declarations.MyRec.Fields("TotalSales").Value
                oSheet.getCellRangeByName("Q" & CStr(i)).String = Declarations.MyRec.Fields("ClientIndustry").Value
                oSheet.getCellRangeByName("R" & CStr(i)).String = Declarations.MyRec.Fields("ClientType").Value
                oSheet.getCellRangeByName("S" & CStr(i)).String = Declarations.MyRec.Fields("ClientMarket").Value
                oSheet.getCellRangeByName("T" & CStr(i)).String = Declarations.MyRec.Fields("ClientIKA").Value
                oSheet.getCellRangeByName("U" & CStr(i)).Value = Declarations.MyRec.Fields("TotalYearSales").Value
                oSheet.getCellRangeByName("V" & CStr(i)).Value = Declarations.MyRec.Fields("Potencial").Value

                PassiveAreaFinish = i
                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If

        '------------------��������������-----------------------------------
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "A" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), 4)

        '--------����������� ��� ��������������-----------------------------------------
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "D" & CStr(PassiveAreaStart) & ":E" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "G" & CStr(PassiveAreaStart) & ":H" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "Q" & CStr(PassiveAreaStart) & ":T" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "V" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "D" & CStr(PassiveAreaStart) & ":E" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "G" & CStr(PassiveAreaStart) & ":H" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "Q" & CStr(PassiveAreaStart) & ":T" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "V" & CStr(PassiveAreaStart) & ":V" & CStr(i - 1), False)

        '--------------------�������� ������----------------------------------------------
        LOSetValidation(oSheet, "Q" & CStr(PassiveAreaStart) & ":Q" & CStr(PassiveAreaFinish - 1), "=$�������������.$B$2:$B$" & IndustryQTY + 1)
        LOSetValidation(oSheet, "R" & CStr(PassiveAreaStart) & ":R" & CStr(PassiveAreaFinish - 1), "=$�������������.$D$2:$D$" & TypeQTY + 1)
        LOSetValidation(oSheet, "S" & CStr(PassiveAreaStart) & ":S" & CStr(PassiveAreaFinish - 1), "=$�������������.$F$2:$F$" & MarketQTY + 1)
        LOSetValidation(oSheet, "T" & CStr(PassiveAreaStart) & ":T" & CStr(PassiveAreaFinish - 1), "=$�������������.$H$2:$H$" & IKAQTY + 1)
    End Sub

    Private Sub UploadNewSales(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ����� �������� � Excel 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        NewAreaStart = i
        NewAreaFinish = i + 99
        i = NewAreaFinish

        '------------------��������������-----------------------------------
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '--------����������� ��� ��������������-----------------------------------------
        With MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":E" & CStr(NewAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(NewAreaStart) & ":E" & CStr(NewAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("G" & CStr(NewAreaStart) & ":H" & CStr(NewAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("G" & CStr(NewAreaStart) & ":H" & CStr(NewAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Locked = False

        With MyWRKBook.ActiveSheet.Range("V" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Interior
            .Color = 16777215
            '.TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("V" & CStr(NewAreaStart) & ":V" & CStr(NewAreaFinish)).Locked = False

        '--------------------�������------------------------------------------------------
        MyWRKBook.ActiveSheet.Range("F" & CStr(NewAreaStart) & ":F" & CStr(NewAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"
        MyWRKBook.ActiveSheet.Range("I" & CStr(NewAreaStart) & ":I" & CStr(NewAreaFinish)).FormulaR1C1 = "=RC[-2]*RC[-1]/100"

        '--------------------�������� ������----------------------------------------------
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$B$2:$B$" & IndustryQTY + 1)
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$D$2:$D$" & TypeQTY + 1)
        MyWRKBook.ActiveSheet.Range("R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$F$2:$F$" & MarketQTY + 1)
        MyWRKBook.ActiveSheet.Range("S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish)).Validation.ShowError = True

        MyWRKBook.ActiveSheet.Range("T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=�������������!$H$2:$H$" & IKAQTY + 1)
        MyWRKBook.ActiveSheet.Range("T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Validation.IgnoreBlank = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Validation.InCellDropdown = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Validation.ShowInput = True
        MyWRKBook.ActiveSheet.Range("T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish)).Validation.ShowError = True
    End Sub

    Private Sub UploadNewSalesLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ����� �������� � Libre Office 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object

        oFrame = oWorkBook.getCurrentController.getFrame

        NewAreaStart = i
        For j As Integer = i To i + 99
            oSheet.getCellRangeByName("F" & CStr(i)).FormulaLocal = "=D" & CStr(i) & " * E" & CStr(i) & "/100 "
            oSheet.getCellRangeByName("I" & CStr(i)).FormulaLocal = "=G" & CStr(i) & " * H" & CStr(i) & "/100 "
            i = i + 1
        Next j
        NewAreaFinish = i - 1

        '------------------��������������-----------------------------------
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(NewAreaStart) & ":V" & CStr(i - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(NewAreaStart) & ":V" & CStr(i - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(NewAreaStart) & ":V" & CStr(i - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(NewAreaStart) & ":V" & CStr(i - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "A" & CStr(NewAreaStart) & ":V" & CStr(i - 1), 4)

        '--------����������� ��� ��������������-----------------------------------------
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "A" & CStr(NewAreaStart) & ":E" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "G" & CStr(NewAreaStart) & ":H" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "Q" & CStr(NewAreaStart) & ":T" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetBGColor(oServiceManager, oDispatcher, oFrame, oWorkBook, oSheet, "V" & CStr(NewAreaStart) & ":V" & CStr(i - 1), RGB(255, 255, 255)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A" & CStr(NewAreaStart) & ":E" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "G" & CStr(NewAreaStart) & ":H" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "Q" & CStr(NewAreaStart) & ":T" & CStr(i - 1), False)
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "V" & CStr(NewAreaStart) & ":V" & CStr(i - 1), False)

        '--------------------�������� ������----------------------------------------------
        LOSetValidation(oSheet, "Q" & CStr(NewAreaStart) & ":Q" & CStr(NewAreaFinish), "=$�������������.$B$2:$B$" & IndustryQTY + 1)
        LOSetValidation(oSheet, "R" & CStr(NewAreaStart) & ":R" & CStr(NewAreaFinish), "=$�������������.$D$2:$D$" & TypeQTY + 1)
        LOSetValidation(oSheet, "S" & CStr(NewAreaStart) & ":S" & CStr(NewAreaFinish), "=$�������������.$F$2:$F$" & MarketQTY + 1)
        LOSetValidation(oSheet, "T" & CStr(NewAreaStart) & ":T" & CStr(NewAreaFinish), "=$�������������.$H$2:$H$" & IKAQTY + 1)
    End Sub

    Private Sub UploadFormulas2Sheet(ByRef MyWRKBook As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ����� �� 2 ���� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '-----����
        '-----����������� �������-----
        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("D8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("D9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("D10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("F8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("F9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("F10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '------------------------����� %----------------------------------
        MyWRKBook.ActiveSheet.Range("E8").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("E9").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("E10").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("E11").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"

        '-----��������� �������-----
        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("G8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("G9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("G10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("I8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("I9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("I10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '------------------------����� %----------------------------------
        MyWRKBook.ActiveSheet.Range("H8").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("H9").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("H10").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("H11").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"

        '-----���� ���������� ���
        '-----����������� �������-----
        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("J8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("J9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("J10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("L8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("L9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("L10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '------------------------����� %----------------------------------
        MyWRKBook.ActiveSheet.Range("K8").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("K9").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("K10").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("H11").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"

        '-----��������� �������-----
        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("M8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("M9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("M10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '-------------------------�����-----------------------------------
        MyWRKBook.ActiveSheet.Range("O8").FormulaR1C1 = "=SUM(R[" & CStr(ActiveAreaStart - 8) & "]C:R[" & CStr(ActiveAreaFinish - 8) & "]C)"
        MyWRKBook.ActiveSheet.Range("O9").FormulaR1C1 = "=SUM(R[" & CStr(PassiveAreaStart - 9) & "]C:R[" & CStr(PassiveAreaFinish - 9) & "]C)"
        MyWRKBook.ActiveSheet.Range("O10").FormulaR1C1 = "=SUM(R[" & CStr(NewAreaStart - 10) & "]C:R[" & CStr(NewAreaFinish - 10) & "]C)"

        '------------------------����� %----------------------------------
        MyWRKBook.ActiveSheet.Range("N8").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("N9").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("N10").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"
        MyWRKBook.ActiveSheet.Range("H11").FormulaR1C1 = "=IF(RC[-1]=0,0,RC[1]/RC[-1]*100)"

    End Sub

    Private Sub UploadFormulas2SheetLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ����� �� 2 ���� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '========================== ���� =================================
        '-----����������� �������-----
        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("D8").FormulaLocal = "=SUM(D" & CStr(ActiveAreaStart) & ":D" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("D9").FormulaLocal = "=SUM(D" & CStr(PassiveAreaStart) & ":D" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("D10").FormulaLocal = "=SUM(D" & CStr(NewAreaStart) & ":D" & CStr(NewAreaFinish) & ")"

        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("F8").FormulaLocal = "=SUM(F" & CStr(ActiveAreaStart) & ":F" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("F9").FormulaLocal = "=SUM(F" & CStr(PassiveAreaStart) & ":F" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("F10").FormulaLocal = "=SUM(F" & CStr(NewAreaStart) & ":F" & CStr(NewAreaFinish) & ")"

        '------------------------����� %----------------------------------
        oSheet.getCellRangeByName("E8").FormulaLocal = "=IF(D8=0;0;F8/D8*100)"
        oSheet.getCellRangeByName("E9").FormulaLocal = "=IF(D9=0;0;F9/D9*100)"
        oSheet.getCellRangeByName("E10").FormulaLocal = "=IF(D10=0;0;F10/D10*100)"

        '-----��������� �������-----
        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("G8").FormulaLocal = "=SUM(G" & CStr(ActiveAreaStart) & ":G" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("G9").FormulaLocal = "=SUM(G" & CStr(PassiveAreaStart) & ":G" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("G10").FormulaLocal = "=SUM(G" & CStr(NewAreaStart) & ":G" & CStr(NewAreaFinish) & ")"

        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("I8").FormulaLocal = "=SUM(I" & CStr(ActiveAreaStart) & ":I" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("I9").FormulaLocal = "=SUM(I" & CStr(PassiveAreaStart) & ":I" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("I10").FormulaLocal = "=SUM(I" & CStr(NewAreaStart) & ":I" & CStr(NewAreaFinish) & ")"

        '------------------------����� %----------------------------------
        oSheet.getCellRangeByName("H8").FormulaLocal = "=IF(G8=0;0;I8/G8*100)"
        oSheet.getCellRangeByName("H9").FormulaLocal = "=IF(G9=0;0;I9/G9*100)"
        oSheet.getCellRangeByName("H10").FormulaLocal = "=IF(G10=0;0;I10/G10*100)"

        '================== ���� ���������� ��� ==========================
        '-----����������� �������-----
        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("J8").FormulaLocal = "=SUM(J" & CStr(ActiveAreaStart) & ":J" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("J9").FormulaLocal = "=SUM(J" & CStr(PassiveAreaStart) & ":J" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("J10").FormulaLocal = "=SUM(J" & CStr(NewAreaStart) & ":J" & CStr(NewAreaFinish) & ")"

        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("L8").FormulaLocal = "=SUM(L" & CStr(ActiveAreaStart) & ":L" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("L9").FormulaLocal = "=SUM(L" & CStr(PassiveAreaStart) & ":L" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("L10").FormulaLocal = "=SUM(L" & CStr(NewAreaStart) & ":L" & CStr(NewAreaFinish) & ")"

        '------------------------����� %----------------------------------
        oSheet.getCellRangeByName("K8").FormulaLocal = "=IF(J8=0;0;L8/J8*100)"
        oSheet.getCellRangeByName("K9").FormulaLocal = "=IF(J9=0;0;L9/J9*100)"
        oSheet.getCellRangeByName("K10").FormulaLocal = "=IF(J10=0;0;L10/J10*100)"

        '-----��������� �������-----
        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("M8").FormulaLocal = "=SUM(M" & CStr(ActiveAreaStart) & ":M" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("M9").FormulaLocal = "=SUM(M" & CStr(PassiveAreaStart) & ":M" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("M10").FormulaLocal = "=SUM(M" & CStr(NewAreaStart) & ":M" & CStr(NewAreaFinish) & ")"

        '-------------------------�����-----------------------------------
        oSheet.getCellRangeByName("O8").FormulaLocal = "=SUM(O" & CStr(ActiveAreaStart) & ":O" & CStr(ActiveAreaFinish) & ")"
        oSheet.getCellRangeByName("O9").FormulaLocal = "=SUM(O" & CStr(PassiveAreaStart) & ":O" & CStr(PassiveAreaFinish) & ")"
        oSheet.getCellRangeByName("O10").FormulaLocal = "=SUM(O" & CStr(NewAreaStart) & ":O" & CStr(NewAreaFinish) & ")"

        '------------------------����� %----------------------------------
        oSheet.getCellRangeByName("N8").FormulaLocal = "=IF(M8=0;0;O8/M8*100)"
        oSheet.getCellRangeByName("N9").FormulaLocal = "=IF(M9=0;0;O9/M9*100)"
        oSheet.getCellRangeByName("N10").FormulaLocal = "=IF(M10=0;0;O10/M10*100)"

    End Sub

    Private Sub UploadFormulas1Sheet(ByRef MyWRKBook As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ����� �� 1 ���� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-------------������ ��������-------------------------------------------
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIndustry "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '------�������
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + IndustryQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C17:R" & CStr(ActiveAreaFinish) & "C17,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C4:R" & CStr(ActiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C17:R" & CStr(ActiveAreaFinish) & "C17,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C7:R" & CStr(ActiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i) & ":C" & CStr(i + IndustryQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C17:R" & CStr(PassiveAreaFinish) & "C17,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C4:R" & CStr(PassiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C17:R" & CStr(PassiveAreaFinish) & "C17,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C7:R" & CStr(PassiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("D" & CStr(i) & ":D" & CStr(i + IndustryQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C17:R" & CStr(NewAreaFinish) & "C17,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C4:R" & CStr(NewAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C17:R" & CStr(NewAreaFinish) & "C17,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C7:R" & CStr(NewAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("E" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1)).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"


        '-------------������ ����� ��������-------------------------------------
        MySQLStr = "SELECT RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelCustomerGroup "
        MySQLStr = MySQLStr & "ORDER BY RCGCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("G" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '------�������
        MyWRKBook.ActiveSheet.Range("H" & CStr(i) & ":H" & CStr(i + TypeQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C18:R" & CStr(ActiveAreaFinish) & "C18,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C4:R" & CStr(ActiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C18:R" & CStr(ActiveAreaFinish) & "C18,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C7:R" & CStr(ActiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("I" & CStr(i) & ":I" & CStr(i + TypeQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C18:R" & CStr(PassiveAreaFinish) & "C18,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C4:R" & CStr(PassiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C18:R" & CStr(PassiveAreaFinish) & "C18,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C7:R" & CStr(PassiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("J" & CStr(i) & ":J" & CStr(i + TypeQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C18:R" & CStr(NewAreaFinish) & "C18,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C4:R" & CStr(NewAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C18:R" & CStr(NewAreaFinish) & "C18,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C7:R" & CStr(NewAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("K" & CStr(i) & ":K" & CStr(i + TypeQTY - 1)).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"

        '-------------������ ����� IKA ��������---------------------------------
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIKATypes "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("M" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If

        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).NumberFormat = "#,##0.00"
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Font.Size = 8

        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).Interior
            .Color = 0
            .TintAndShade = 0.95
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        '------�������
        MyWRKBook.ActiveSheet.Range("N" & CStr(i) & ":N" & CStr(i + IKAQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C20:R" & CStr(ActiveAreaFinish) & "C20,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C4:R" & CStr(ActiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(ActiveAreaStart) & "C20:R" & CStr(ActiveAreaFinish) & "C20,RC[-1],'��������� ������'!R" & CStr(ActiveAreaStart) & "C7:R" & CStr(ActiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("O" & CStr(i) & ":O" & CStr(i + IKAQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C20:R" & CStr(PassiveAreaFinish) & "C20,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C4:R" & CStr(PassiveAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(PassiveAreaStart) & "C20:R" & CStr(PassiveAreaFinish) & "C20,RC[-2],'��������� ������'!R" & CStr(PassiveAreaStart) & "C7:R" & CStr(PassiveAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("P" & CStr(i) & ":P" & CStr(i + IKAQTY - 1)).FormulaR1C1 = "=SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C20:R" & CStr(NewAreaFinish) & "C20,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C4:R" & CStr(NewAreaFinish) & "C4) + SUMIF('��������� ������'!R" & CStr(NewAreaStart) & "C20:R" & CStr(NewAreaFinish) & "C20,RC[-3],'��������� ������'!R" & CStr(NewAreaStart) & "C7:R" & CStr(NewAreaFinish) & "C7)"
        MyWRKBook.ActiveSheet.Range("Q" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1)).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"

    End Sub

    Private Sub UploadFormulas1SheetLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByRef i As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ����� �� 1 ���� 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim oFrame As Object
        Dim j As Integer

        oFrame = oWorkBook.getCurrentController.getFrame

        '-------------������ ��������-------------------------------------------
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIndustry "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            j = i
            While Not Declarations.MyRec.EOF
                '-----��������
                oSheet.getCellRangeByName("A" & CStr(j)).String = Declarations.MyRec.Fields("Name").Value
                '-----�������
                oSheet.getCellRangeByName("B" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$Q$" & CStr(ActiveAreaStart) & ":$Q$" & CStr(ActiveAreaFinish) & ";A" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(ActiveAreaStart) & ":$D$" & CStr(ActiveAreaFinish) & ") + SUMIF('��������� ������'.$Q$" & CStr(ActiveAreaStart) & _
                    ":$Q$" & CStr(ActiveAreaFinish) & ";A" & CStr(j) & ";'��������� ������'.$G$" & CStr(ActiveAreaStart) & ":$G$" & CStr(ActiveAreaFinish) & ")"
                oSheet.getCellRangeByName("C" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$Q$" & CStr(PassiveAreaStart) & ":$Q$" & CStr(PassiveAreaFinish) & ";A" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(PassiveAreaStart) & ":$D$" & CStr(PassiveAreaFinish) & ") + SUMIF('��������� ������'.$Q$" & CStr(PassiveAreaStart) & _
                    ":$Q$" & CStr(PassiveAreaFinish) & ";A" & CStr(j) & ";'��������� ������'.$G$" & CStr(PassiveAreaStart) & ":$G$" & CStr(PassiveAreaFinish) & ")"
                oSheet.getCellRangeByName("D" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$Q$" & CStr(NewAreaStart) & ":$Q$" & CStr(NewAreaFinish) & ";A" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(NewAreaStart) & ":$D$" & CStr(NewAreaFinish) & ") + SUMIF('��������� ������'.$Q$" & CStr(NewAreaStart) & _
                    ":$Q$" & CStr(NewAreaFinish) & ";A" & CStr(j) & ";'��������� ������'.$G$" & CStr(NewAreaStart) & ":$G$" & CStr(NewAreaFinish) & ")"
                oSheet.getCellRangeByName("E" & CStr(j)).FormulaLocal = "=SUM(B" & CStr(j) & ":D" & CStr(j) & ")"
                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '-----��������������
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":E" & CStr(i + IndustryQTY - 1), 4)

        '-------------������ ����� ��������-------------------------------------
        MySQLStr = "SELECT RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelCustomerGroup "
        MySQLStr = MySQLStr & "ORDER BY RCGCode "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            j = i
            While Not Declarations.MyRec.EOF
                '-----��������
                oSheet.getCellRangeByName("G" & CStr(j)).String = Declarations.MyRec.Fields("RussianName").Value
                '-----�������
                oSheet.getCellRangeByName("H" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$R$" & CStr(ActiveAreaStart) & ":$R$" & CStr(ActiveAreaFinish) & ";G" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(ActiveAreaStart) & ":$D$" & CStr(ActiveAreaFinish) & ") + SUMIF('��������� ������'.$R$" & CStr(ActiveAreaStart) & _
                    ":$R$" & CStr(ActiveAreaFinish) & ";G" & CStr(j) & ";'��������� ������'.$G$" & CStr(ActiveAreaStart) & ":$G$" & CStr(ActiveAreaFinish) & ")"
                oSheet.getCellRangeByName("I" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$R$" & CStr(PassiveAreaStart) & ":$R$" & CStr(PassiveAreaFinish) & ";G" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(PassiveAreaStart) & ":$D$" & CStr(PassiveAreaFinish) & ") + SUMIF('��������� ������'.$R$" & CStr(PassiveAreaStart) & _
                    ":$R$" & CStr(PassiveAreaFinish) & ";G" & CStr(j) & ";'��������� ������'.$G$" & CStr(PassiveAreaStart) & ":$G$" & CStr(PassiveAreaFinish) & ")"
                oSheet.getCellRangeByName("J" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$R$" & CStr(NewAreaStart) & ":$R$" & CStr(NewAreaFinish) & ";G" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(NewAreaStart) & ":$D$" & CStr(NewAreaFinish) & ") + SUMIF('��������� ������'.$R$" & CStr(NewAreaStart) & _
                    ":$R$" & CStr(NewAreaFinish) & ";G" & CStr(j) & ";'��������� ������'.$G$" & CStr(NewAreaStart) & ":$G$" & CStr(NewAreaFinish) & ")"
                oSheet.getCellRangeByName("K" & CStr(j)).FormulaLocal = "=SUM(H" & CStr(j) & ":J" & CStr(j) & ")"
                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '-----��������������
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "G" & CStr(i) & ":K" & CStr(i + TypeQTY - 1), 4)

        '-------------������ ����� IKA ��������---------------------------------
        MySQLStr = "SELECT Name "
        MySQLStr = MySQLStr & "FROM tbl_RexelIKATypes "
        MySQLStr = MySQLStr & "ORDER BY ID "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            j = i
            While Not Declarations.MyRec.EOF
                '-----��������
                oSheet.getCellRangeByName("M" & CStr(j)).String = Declarations.MyRec.Fields("Name").Value
                '-----�������
                oSheet.getCellRangeByName("N" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$T$" & CStr(ActiveAreaStart) & ":$T$" & CStr(ActiveAreaFinish) & ";M" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(ActiveAreaStart) & ":$D$" & CStr(ActiveAreaFinish) & ") + SUMIF('��������� ������'.$T$" & CStr(ActiveAreaStart) & _
                    ":$T$" & CStr(ActiveAreaFinish) & ";M" & CStr(j) & ";'��������� ������'.$G$" & CStr(ActiveAreaStart) & ":$G$" & CStr(ActiveAreaFinish) & ")"
                oSheet.getCellRangeByName("O" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$T$" & CStr(PassiveAreaStart) & ":$T$" & CStr(PassiveAreaFinish) & ";M" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(PassiveAreaStart) & ":$D$" & CStr(PassiveAreaFinish) & ") + SUMIF('��������� ������'.$T$" & CStr(PassiveAreaStart) & _
                    ":$T$" & CStr(PassiveAreaFinish) & ";M" & CStr(j) & ";'��������� ������'.$G$" & CStr(PassiveAreaStart) & ":$G$" & CStr(PassiveAreaFinish) & ")"
                oSheet.getCellRangeByName("P" & CStr(j)).FormulaLocal = "=SUMIF('��������� ������'.$T$" & CStr(NewAreaStart) & ":$T$" & CStr(NewAreaFinish) & ";M" & CStr(j) & _
                    ";'��������� ������'.$D$" & CStr(NewAreaStart) & ":$D$" & CStr(NewAreaFinish) & ") + SUMIF('��������� ������'.$T$" & CStr(NewAreaStart) & _
                    ":$T$" & CStr(NewAreaFinish) & ";M" & CStr(j) & ";'��������� ������'.$G$" & CStr(NewAreaStart) & ":$G$" & CStr(NewAreaFinish) & ")"
                oSheet.getCellRangeByName("Q" & CStr(j)).FormulaLocal = "=SUM(N" & CStr(j) & ":P" & CStr(j) & ")"
                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '-----��������������
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1), 8)
        LOSetBorders(oServiceManager, oSheet, "M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1), 70, 20, RGB(0, 0, 0)) '---��� ����� �������� ������� - �� ����� ���� ��� �� RGB � BGR!!!! ������ ���� �����
        LOWrapText(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1))
        LOFormatCells(oServiceManager, oDispatcher, oFrame, "M" & CStr(i) & ":Q" & CStr(i + IKAQTY - 1), 4)
    End Sub

    Private Sub PasswordProtectON(ByRef MyWRKBook As Object)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������ ������ 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.Sheets(1).Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)
        MyWRKBook.Sheets(2).Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)
        MyWRKBook.Sheets(3).Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)
    End Sub
End Module
