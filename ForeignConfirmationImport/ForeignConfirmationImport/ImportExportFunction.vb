Module ImportExportFunction

    Public Function UploadOrderToExcel(ByVal ComOrder As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ����������� ������ � Excel
        '//  
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              '������� �����

        If Trim(ComOrder) = "" Then
            MsgBox("���������� ������ ����� ����������� ������ �� �������.", MsgBoxStyle.Critical, "��������!")
            Main.TextBox1.Select()
            Exit Function
        Else
            ComOrder = Right("0000000000" & ComOrder, 10)
        End If
        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        ExportOrderHeaderToExcel(MyWRKBook, ComOrder, i)
        ExportOrderBodyToExcel(MyWRKBook, ComOrder, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing

    End Function

    Public Function UploadOrderToLO(ByVal ComOrder As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ����������� ������ � LibreOffice
        '//  
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer                              '������� �����
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim Counter As Integer
        Dim oFrame As Object

        If Trim(ComOrder) = "" Then
            MsgBox("���������� ������ ����� ����������� ������ �� �������.", MsgBoxStyle.Critical, "��������!")
            Main.TextBox1.Select()
            Exit Function
        Else
            ComOrder = Right("0000000000" & ComOrder, 10)
        End If

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

        ExportOrderHeaderToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)
        ExportOrderBodyToLO(oSheet, oServiceManager, oWorkBook, oDispatcher, ComOrder, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()

    End Function


    Private Sub ExportOrderHeaderToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ��������� ����������� ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String

        '------���� �������� � �����--------------
        MyWRKBook.ActiveSheet.Range("B2:I2").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B2") = "��� ""��������"""
        MyWRKBook.ActiveSheet.Range("B2:I2").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B2:I2").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B2:I2").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B2:I2").WrapText = True

        MyWRKBook.ActiveSheet.Range("B3:D3").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B3") = "Address/�����:"
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:D3").Font.Bold = True

        MyWRKBook.ActiveSheet.Range("E3:I3").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E3") = "������,195027, �����-���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        MyWRKBook.ActiveSheet.Range("E3:I3").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B3:I3").WrapText = True
        MyWRKBook.ActiveSheet.Rows("3:3").RowHeight = 30
        MyWRKBook.ActiveSheet.Range("B3:I3").VerticalAlignment = -4108

        '-------����� � ���� ������ �� �������--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate, SupplierCode, WH "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("B4:I4").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MySupplierCode = ""
            Declarations.MyWH = ""
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / ����� �� ������� � " & ComOrder & " ��  "
            MyWRKBook.ActiveSheet.Range("E1").NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("E1") = ComOrder
            trycloseMyRec()
        Else
            Declarations.MySupplierCode = Declarations.MyRec.Fields("SupplierCode").Value
            Declarations.MyWH = Declarations.MyRec.Fields("WH").Value
            MyWRKBook.ActiveSheet.Range("B4") = "Purchase Order / ����� �� ������� � " & ComOrder & " ��  " & Declarations.MyRec.Fields("OrderDate").Value
            MyWRKBook.ActiveSheet.Range("E1").NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("E1") = ComOrder
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B4:I4").Font.Name = "Arial"
        MyWRKBook.ActiveSheet.Range("B4:I4").Font.Size = 12
        MyWRKBook.ActiveSheet.Range("B4:I4").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("B4:I4").WrapText = True
        MyWRKBook.ActiveSheet.Range("B4:I4").HorizontalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("5:5").RowHeight = 5

        '-----------���������-----------------------------
        MyWRKBook.ActiveSheet.Range("B6:D6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B6") = "Supplier / ���������:"
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:D6").Font.Bold = True

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E6:I6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E6") = SupplierName & Chr(10) & SupplierAddress
        MyWRKBook.ActiveSheet.Range("E6:I6").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B6:I6").WrapText = True
        MyWRKBook.ActiveSheet.Rows("6:6").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B6:I6").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("7:7").RowHeight = 5

        '---------����� ��������--------------------------------
        MyWRKBook.ActiveSheet.Range("B8:D8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B8") = "Delivery Address / ����� ��������"
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:D8").Font.Bold = True

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E8:I8").MergeCells = True
        MyWRKBook.ActiveSheet.Range("E8") = DelName & Chr(10) & DelAddr
        MyWRKBook.ActiveSheet.Range("E8:I8").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B8:I8").WrapText = True
        MyWRKBook.ActiveSheet.Rows("8:8").RowHeight = 45
        MyWRKBook.ActiveSheet.Range("B8:I8").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("9:9").RowHeight = 5

        '---------������� ��������------------------------------
        MyWRKBook.ActiveSheet.Range("B10:D10").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B10") = "Terms Of Delivery / ������� ��������"
        MyWRKBook.ActiveSheet.Range("B10:D10").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'RUS')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E10:I10").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E10") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E10:I10").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B10:I10").WrapText = True
        MyWRKBook.ActiveSheet.Range("B10:I10").VerticalAlignment = -4108

        '------������� ������------------------------------
        MyWRKBook.ActiveSheet.Range("B11:D11").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B11") = "Terms Of Payment / ������� ������"
        MyWRKBook.ActiveSheet.Range("B11:D11").Font.Size = 7

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'RUS')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E11:I11").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E11") = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E11:I11").Font.Size = 8
        MyWRKBook.ActiveSheet.Range("B11:I11").WrapText = True
        MyWRKBook.ActiveSheet.Range("B11:I11").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("12:12").RowHeight = 5

        '-------��������------------------------------------
        MyWRKBook.ActiveSheet.Range("B13:D13").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B13") = "Purchaser / ��������:"
        MyWRKBook.ActiveSheet.Range("B13:D13").Font.Size = 10

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        MyWRKBook.ActiveSheet.Range("E13:I13").MergeCells = True
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("E13") = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("E13:I13").Font.Size = 10
        MyWRKBook.ActiveSheet.Range("B13:I13").WrapText = True
        MyWRKBook.ActiveSheet.Range("B13:I13").VerticalAlignment = -4108

        MyWRKBook.ActiveSheet.Rows("14:14").RowHeight = 5
        '-------��������� �������---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B15") = "N �/�"
        MyWRKBook.ActiveSheet.Range("C15") = "Supp Item Code / ��� ������ �������"
        MyWRKBook.ActiveSheet.Range("D15") = "Item Code / ��� ������"
        MyWRKBook.ActiveSheet.Range("E15") = "Item Name / ������������ ������"
        MyWRKBook.ActiveSheet.Range("F15") = "UOM / ��. �����-�"
        MyWRKBook.ActiveSheet.Range("G15") = "�-� �� �������"
        MyWRKBook.ActiveSheet.Range("H15") = "N ������"
        MyWRKBook.ActiveSheet.Range("I15") = "QTY / ����������"
        MyWRKBook.ActiveSheet.Range("J15") = "% ����������� ����� "
        MyWRKBook.ActiveSheet.Range("K15") = "Price / ����, " & CurrName
        MyWRKBook.ActiveSheet.Range("L15") = "���� �� ���������� "
        MyWRKBook.ActiveSheet.Range("M15") = "New Price / ����� ����, " & CurrName
        MyWRKBook.ActiveSheet.Range("N15") = "����� ���� �� ���������� "
        MyWRKBook.ActiveSheet.Range("B15:N15").Font.Size = 7
        MyWRKBook.ActiveSheet.Range("B15:N15").WrapText = True
        MyWRKBook.ActiveSheet.Range("B15:N15").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B15:N15").HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Rows("15:15").RowHeight = 40
        MyWRKBook.ActiveSheet.Range("B15:N15").Select()
        MyWRKBook.ActiveSheet.Range("B15:N15").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B15:N15").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("B15:N15").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:N15").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:N15").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:N15").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:N15").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B15:N15").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With


        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 1
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 3
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 9
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 9
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 7
        MyWRKBook.ActiveSheet.Columns("K:N").ColumnWidth = 12
        i = 16
    End Sub

    Private Sub ExportOrderHeaderToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ��������� ����������� ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SupplierName As String
        Dim SupplierAddress As String
        Dim DelName As String
        Dim DelAddr As String
        Dim CurrName As String

        oFrame = oWorkBook.getCurrentController.getFrame
        '-----������ �������
        oSheet.getColumns().getByName("A").Width = 400
        oSheet.getColumns().getByName("B").Width = 600
        oSheet.getColumns().getByName("C").Width = 1800
        oSheet.getColumns().getByName("D").Width = 1800
        oSheet.getColumns().getByName("E").Width = 6000
        oSheet.getColumns().getByName("F").Width = 1000
        oSheet.getColumns().getByName("G").Width = 2400
        oSheet.getColumns().getByName("H").Width = 1600
        oSheet.getColumns().getByName("I").Width = 1400
        oSheet.getColumns().getByName("J").Width = 1400
        oSheet.getColumns().getByName("K").Width = 2400

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B2:I2")
        oSheet.getCellRangeByName("B2").String = "��� ""��������"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 12)
        oSheet.getCellRangeByName("B2").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3").String = "Address/�����:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B3:D3", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B3:D3")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B3:D3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B3:D3")
        oSheet.getCellRangeByName("B3:D3").VertJustify = 2

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E3:I3")
        oSheet.getCellRangeByName("E3").String = "������,195027, �����-���������, ������� ��., �.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E3:I3", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E3:I3", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E3:I3")
        oSheet.getCellRangeByName("E3:I3").VertJustify = 2

        '-------����� � ���� ������ �� �������--------------
        MySQLStr = "SELECT CONVERT(nvarchar(30),OrderDate,103) AS OrderDate, SupplierCode, WH "
        MySQLStr = MySQLStr & "FROM tbl_PurchaseWorkplace_ConsolidatedOrders WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ID = N'" & ComOrder & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B4:I4")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MySupplierCode = ""
            Declarations.MyWH = ""
            oSheet.getCellRangeByName("B4").String = "Purchase Order / ����� �� ������� � " & ComOrder
            oSheet.getCellRangeByName("E1").String = ComOrder
            trycloseMyRec()
        Else
            Declarations.MySupplierCode = Declarations.MyRec.Fields("SupplierCode").Value
            Declarations.MyWH = Declarations.MyRec.Fields("WH").Value
            oSheet.getCellRangeByName("B4").String = "Purchase Order / ����� �� ������� � " & ComOrder & " ��  " & Declarations.MyRec.Fields("OrderDate").Value
            oSheet.getCellRangeByName("E1").String = ComOrder
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B4:I4", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B4:I4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B4:I4", 12)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B4:I4")
        oSheet.getCellRangeByName("B4:I4").VertJustify = 2
        oSheet.getCellRangeByName("A5").Rows.Height = 200

        '-----------���������-----------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6").String = "Supplier / ���������:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B6:D6", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B6:D6")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B6:D6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B6:D6")
        oSheet.getCellRangeByName("B6:D6").VertJustify = 2

        MySQLStr = "SELECT PL01002 AS SuppName, LTRIM(RTRIM(LTRIM(RTRIM(PL01003)) + ' ' + LTRIM(RTRIM(PL01004)) + ' ' + LTRIM(RTRIM(PL01005)))) AS SuppAddress "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            SupplierName = ""
            SupplierAddress = ""
            trycloseMyRec()
        Else
            SupplierName = Declarations.MyRec.Fields("SuppName").Value
            SupplierAddress = Declarations.MyRec.Fields("SuppAddress").Value
            trycloseMyRec()
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6").String = SupplierName & Chr(10) & SupplierAddress
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E6:I6", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E6:I6", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E6:I6")
        oSheet.getCellRangeByName("E6:I6").VertJustify = 2
        oSheet.getCellRangeByName("A7").Rows.Height = 200

        '---------����� ��������--------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8").String = "Delivery Address / ����� ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B8:D8", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B8:D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B8:D8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B8:D8")
        oSheet.getCellRangeByName("B8:D8").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(ConsignorOfGoodsName)) AS DelName, LTRIM(RTRIM(ConsignorOfGoodsAddr)) AS DelAddr "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & Declarations.MyWH & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DelName = ""
            DelAddr = ""
            trycloseMyRec()
        Else
            DelName = Declarations.MyRec.Fields("DelName").Value
            DelAddr = Declarations.MyRec.Fields("DelAddr").Value
            trycloseMyRec()
        End If
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8").String = DelName & Chr(10) & DelAddr
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E8:I8", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E8:I8", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E8:I8")
        oSheet.getCellRangeByName("E8:I8").VertJustify = 2
        oSheet.getCellRangeByName("A9").Rows.Height = 200

        '---------������� ��������------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10").String = "Delivery Address / ����� ��������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B10:D10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B10:D10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B10:D10")
        oSheet.getCellRangeByName("B10:D10").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'1') AND (PL23002 = N'RUS')) AS View_1 ON CONVERT(int, PL010300.PL01029) = CONVERT(int, View_1.PL23003) "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E10:I10")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E10").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E10:I10", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E10:I10")
        oSheet.getCellRangeByName("E10:I10").VertJustify = 2

        '------������� ������------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11").String = "Terms Of Payment / ������� ������"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B11:D11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B11:D11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B11:D11")
        oSheet.getCellRangeByName("B11:D11").VertJustify = 2

        MySQLStr = "SELECT View_1.PL23004 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT PL23001, PL23002, PL23003, PL23004 "
        MySQLStr = MySQLStr & "FROM PL230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (PL23001 = N'0') AND (PL23002 = N'RUS')) AS View_1 ON PL010300.PL01028 = View_1.PL23003 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E11:I11")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E11").String = Declarations.MyRec.Fields("PL23004").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E11:I11", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E11:I11", 7)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E11:I11")
        oSheet.getCellRangeByName("E11:I11").VertJustify = 2
        oSheet.getCellRangeByName("A12").Rows.Height = 200

        '-------��������------------------------------------
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13").String = "Purchaser / ��������:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B13:D13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B13:D13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B13:D13")
        oSheet.getCellRangeByName("B13:D13").VertJustify = 2

        MySQLStr = "SELECT LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(View_1.SYPD001, ''))) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, ''))))) AS Purchaser "
        MySQLStr = MySQLStr & "FROM tbl_SupplierCard0300 WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SYPD001, SYPD002, SYPD003 "
        MySQLStr = MySQLStr & "FROM SYPD0300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYPD002 = N'ENG')) AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) "
        MySQLStr = MySQLStr & "WHERE (tbl_SupplierCard0300.PL01001 = N'" & Declarations.MySupplierCode & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "E13:I13")
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            oSheet.getCellRangeByName("E13").String = Declarations.MyRec.Fields("Purchaser").Value
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "E13:I13", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "E13:I13", 10)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "E13:I13")
        oSheet.getCellRangeByName("E13:I13").VertJustify = 2
        oSheet.getCellRangeByName("A14").Rows.Height = 200

        '-------��������� �������---------------------------
        MySQLStr = "SELECT SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON PL010300.PL01026 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (PL010300.PL01001 = N'" & Declarations.MySupplierCode & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            CurrName = "RUB"
            trycloseMyRec()
        Else
            CurrName = Declarations.MyRec.Fields("SYCD009").Value
            trycloseMyRec()
        End If

        oSheet.getCellRangeByName("B15").String = "N �/�"
        oSheet.getCellRangeByName("C15").String = "Supp Item Code / ��� ������ �������"
        oSheet.getCellRangeByName("D15").String = "Item Code / ��� ������"
        oSheet.getCellRangeByName("E15").String = "Item Name / ������������ ������"
        oSheet.getCellRangeByName("F15").String = "UOM / ��. �����-�"
        oSheet.getCellRangeByName("G15").String = "�-� �� �������"
        oSheet.getCellRangeByName("H15").String = "N ������"
        oSheet.getCellRangeByName("I15").String = "QTY / ����������"
        oSheet.getCellRangeByName("J15").String = "% ����������� ����� "
        oSheet.getCellRangeByName("K15").String = "Price / ����, " & CurrName
        oSheet.getCellRangeByName("L15").String = "���� �� ���������� "
        oSheet.getCellRangeByName("M15").String = "New Price / ����� ����, " & CurrName
        oSheet.getCellRangeByName("N15").String = "����� ���� �� ���������� "

        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B15:N15", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B15:N15", 7)
        oSheet.getCellRangeByName("B15:N15").CellBackColor = 16775598
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 70
        oSheet.getCellRangeByName("B15:N15").TopBorder = LineFormat
        oSheet.getCellRangeByName("B15:N15").RightBorder = LineFormat
        oSheet.getCellRangeByName("B15:N15").LeftBorder = LineFormat
        oSheet.getCellRangeByName("B15:N15").BottomBorder = LineFormat
        oSheet.getCellRangeByName("B15:N15").VertJustify = 2
        oSheet.getCellRangeByName("B15:N15").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B15:N15")

        i = 16
    End Sub


    Private Sub ExportOrderBodyToExcel(ByRef MyWRKBook As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���� ����������� ������ � Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim j As Integer                        '������� �����

        j = i
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[PC01001] [nvarchar] (10) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03002] [nvarchar] (6) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01057] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[PC03019] [numeric](18, 8) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_ForeignPurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder WITH(NOLOCK) "
        MySQLStr = MySQLStr & "Order BY SC01060, "
        MySQLStr = MySQLStr & "PC01001, "
        MySQLStr = MySQLStr & "PC03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                '-----N �/�
                MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = i - j + 1
                '-----��� ������
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = Declarations.MyRec.Fields("SC01060").Value
                '-----��� ������ ����������
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = Declarations.MyRec.Fields("PC03005").Value
                '-----�������� ������
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = Declarations.MyRec.Fields("PC03006").Value
                '-----������� ��������� ������
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----N ������ �� �������
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = Declarations.MyRec.Fields("PC01001").Value
                '-----N ������ ������ �� �������
                MyWRKBook.ActiveSheet.Range("H" & CStr(i)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = Declarations.MyRec.Fields("PC03002").Value
                '-----���������� ������
                MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = Declarations.MyRec.Fields("QTY").Value
                '-----% ����������� �����
                MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = Declarations.MyRec.Fields("SC01057").Value
                '-----����
                MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = Declarations.MyRec.Fields("Price").Value
                '-----���� �� ����������
                MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = Declarations.MyRec.Fields("PC03019").Value


                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Select()
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(5).LineStyle = -4142
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(6).LineStyle = -4142
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(7)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(8)
                .LineStyle = 1
                .Weight = 4
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(9)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(10)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(11)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Borders(12)
                .LineStyle = 1
                .Weight = 2
                .ColorIndex = -4105
            End With

            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Select()
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Font
                .Name = "Arial"
                .Size = 7
            End With

            MyWRKBook.ActiveSheet.Range("C" & CStr(j) & ":D" & CStr(i - 1)).Font.Bold = True
            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).WrapText = True

            MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":N" & CStr(i - 1)).Select()
            With MyWRKBook.ActiveSheet.Range("B" & CStr(j) & ":L" & CStr(i - 1)).Interior
                .Color = 15395562
                .Pattern = 1
                .PatternColorIndex = -4105
            End With
        End If
    End Sub

    Private Sub ExportOrderBodyToLO(ByRef oSheet As Object, ByRef oServiceManager As Object, ByRef oWorkBook As Object, _
        ByRef oDispatcher As Object, ByVal ComOrder As String, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���� ����������� ������ � LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim j As Integer                        '������� �����

        oFrame = oWorkBook.getCurrentController.getFrame
        j = i
        MySQLStr = "IF exists(select * from tempdb..sysobjects where "
        MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
        MySQLStr = MySQLStr & "and xtype = N'U') "
        MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "CREATE TABLE #_MyPCOrder( "
        MySQLStr = MySQLStr & "[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[QTY] [numeric](20, 8), "
        MySQLStr = MySQLStr & "[PC03009] [int], "
        MySQLStr = MySQLStr & "[PC03009_Name][nvarchar](10), "
        MySQLStr = MySQLStr & "[Price] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[PC01001] [nvarchar] (10) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[PC03002] [nvarchar] (6) COLLATE Cyrillic_General_BIN, "
        MySQLStr = MySQLStr & "[SC01057] [numeric](28, 8), "
        MySQLStr = MySQLStr & "[PC03019] [numeric](18, 8) "
        MySQLStr = MySQLStr & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "EXEC spp_PurchaseWorkplace_ForeignPurchaseGroupOrderPreparation N'" & ComOrder & "' "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM #_MyPCOrder WITH(NOLOCK) "
        MySQLStr = MySQLStr & "Order BY SC01060, "
        MySQLStr = MySQLStr & "PC01001, "
        MySQLStr = MySQLStr & "PC03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                '-----N �/�
                oSheet.getCellRangeByName("B" & CStr(i)).value = i - j + 1
                '-----��� ������
                oSheet.getCellRangeByName("C" & CStr(i)).String = Declarations.MyRec.Fields("SC01060").Value
                '-----��� ������ ����������
                oSheet.getCellRangeByName("D" & CStr(i)).String = Declarations.MyRec.Fields("PC03005").Value
                '-----�������� ������
                oSheet.getCellRangeByName("E" & CStr(i)).String = Declarations.MyRec.Fields("PC03006").Value
                '-----������� ��������� ������
                oSheet.getCellRangeByName("F" & CStr(i)).String = Declarations.MyRec.Fields("PC03009_Name").Value
                '-----N ������ �� �������
                oSheet.getCellRangeByName("G" & CStr(i)).String = Declarations.MyRec.Fields("PC01001").Value
                '-----N ������ ������ �� �������
                oSheet.getCellRangeByName("H" & CStr(i)).String = Declarations.MyRec.Fields("PC03002").Value
                '-----���������� ������
                oSheet.getCellRangeByName("I" & CStr(i)).value = Declarations.MyRec.Fields("QTY").Value
                '-----% ����������� �����
                oSheet.getCellRangeByName("J" & CStr(i)).value = Declarations.MyRec.Fields("SC01057").Value
                '-----����
                oSheet.getCellRangeByName("K" & CStr(i)).value = Declarations.MyRec.Fields("Price").Value
                '-----���� �� ����������
                oSheet.getCellRangeByName("L" & CStr(i)).value = Declarations.MyRec.Fields("PC03019").Value
                Declarations.MyRec.MoveNext()
                i = i + 1
            End While
            trycloseMyRec()

            MySQLStr = "IF exists(select * from tempdb..sysobjects where "
            MySQLStr = MySQLStr & "id = object_id(N'tempdb..#_MyPCOrder')  "
            MySQLStr = MySQLStr & "and xtype = N'U') "
            MySQLStr = MySQLStr & "DROP TABLE #_MyPCOrder "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":N" & CStr(i - 1), "Arial")
            LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":N" & CStr(i - 1), 7)
            oSheet.getCellRangeByName("B" & CStr(j) & ":L" & CStr(i - 1)).CellBackColor = 14540253
            Dim LineFormat As Object
            LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
            LineFormat.LineStyle = 0
            LineFormat.LineWidth = 20
            oSheet.getCellRangeByName("B" & CStr(j) & ":N" & CStr(i - 1)).TopBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":N" & CStr(i - 1)).RightBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":N" & CStr(i - 1)).LeftBorder = LineFormat
            oSheet.getCellRangeByName("B" & CStr(j) & ":N" & CStr(i - 1)).BottomBorder = LineFormat
            LOWrapText(oServiceManager, oDispatcher, oFrame, "B" & CStr(j) & ":N" & CStr(i - 1))
            LOFontSetBold(oServiceManager, oDispatcher, oFrame, "C" & CStr(j) & ":D" & CStr(i - 1))
            LOFormatCells(oServiceManager, oDispatcher, oFrame, "I" & CStr(j) & ":N" & CStr(i - 1), 4)
        End If
    End Sub

    Public Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� Excel ����� ��� �� ����������� ������ �� �������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim appXLSRC As Object
        Dim MyERRStr As String                      '��������� �� �������
        Dim i As Double                             '������� �����
        Dim MyItemN As String                       '��� ������ � Scala
        Dim My�POrderN As String                    '���������� ����� ������ �� ������� � Scala
        Dim MyPOrderN As String                     '������������ ����� ������ �� ������� � Scala
        Dim MyRowN As String                        '����� ������ � ������������ ������ �� �������
        Dim MyQTY As Double                         '��� - �� � ������������ ������ �� �������
        Dim MyNewPrice As Double                    '����� ����
        Dim MyNewPriceAtQTY As Double               '����� ���� �� ����������
        Dim MySQLStr As String
        Dim MyCC As Integer                         '�������

        If Main.OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (Main.OpenFileDialog1.FileName = "") Then
            Else
                MyERRStr = ""
                System.Windows.Forms.Application.DoEvents()
                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(Main.OpenFileDialog1.FileName)
                '---����� ����������� ������ �� �������
                If appXLSRC.Worksheets(1).Range("E1").Value = Nothing Then
                    MyERRStr = MyERRStr & "� ������ E1 �� ���������� ����� ����������� ������ �� �������." & Chr(13)
                Else
                    My�POrderN = appXLSRC.Worksheets(1).Range("E1").Value().ToString
                    My�POrderN = Right("0000000000" & My�POrderN, 10)
                    i = 16
                    While Not appXLSRC.Worksheets(1).Range("D" & i).Value = Nothing
                        MyItemN = appXLSRC.Worksheets(1).Range("D" & CStr(i)).Value.ToString
                        If (appXLSRC.Worksheets(1).Range("G" & CStr(i)).Value = Nothing) Then
                            MyERRStr = MyERRStr & "������ G" & CStr(i) & " ����� ������������� ������ �� ������� � Scala ������ ���� ��������." & Chr(13)
                        Else
                            MyPOrderN = appXLSRC.Worksheets(1).Range("G" & CStr(i)).Value.ToString
                            If (appXLSRC.Worksheets(1).Range("H" & CStr(i)).Value = Nothing) Then
                                MyERRStr = MyERRStr & "������ H" & CStr(i) & " ����� ������ ������������� ������ �� ������� � Scala ������ ���� ��������." & Chr(13)
                            Else
                                MyRowN = appXLSRC.Worksheets(1).Range("H" & CStr(i)).Value.ToString
                                If (appXLSRC.Worksheets(1).Range("I" & CStr(i)).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("I" & CStr(i)).Value) Is Double) Then
                                    MyERRStr = MyERRStr & "������ I" & CStr(i) & " ���������� � ������ ������ ������ ���� ���������." & Chr(13)
                                Else
                                    If (Not TypeOf (appXLSRC.Worksheets(1).Range("I" & CStr(i)).Value) Is Double) Then
                                        MyERRStr = MyERRStr & "������ I" & CStr(i) & " ���������� � ������ ������ ������ ���� �������� ���������." & Chr(13)
                                    Else
                                        MyQTY = appXLSRC.Worksheets(1).Range("I" & CStr(i)).Value
                                        If MyQTY = 0 Then
                                            MyERRStr = MyERRStr & "������ I" & CStr(i) & " ��� ���������� � ������ ������ 0 ���������� ��� �� ������������." & Chr(13)
                                        Else
                                            If (appXLSRC.Worksheets(1).Range("M" & CStr(i)).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("M" & CStr(i)).Value) Is Double) Then
                                                MyERRStr = MyERRStr & "������ M" & CStr(i) & " ����� ���� ������ ���� ���������." & Chr(13)
                                            Else
                                                If (Not TypeOf (appXLSRC.Worksheets(1).Range("M" & CStr(i)).Value) Is Double) Then
                                                    MyERRStr = MyERRStr & "������ M" & CStr(i) & " ����� ���� ������ ���� �������� ���������." & Chr(13)
                                                Else
                                                    MyNewPrice = appXLSRC.Worksheets(1).Range("M" & CStr(i)).Value
                                                    If (appXLSRC.Worksheets(1).Range("N" & CStr(i)).Value = Nothing And Not TypeOf (appXLSRC.Worksheets(1).Range("N" & CStr(i)).Value) Is Double) Then
                                                        MyERRStr = MyERRStr & "������ N" & CStr(i) & " ����� ���� �� ���������� ������ ���� ���������." & Chr(13)
                                                    Else
                                                        If (Not TypeOf (appXLSRC.Worksheets(1).Range("M" & CStr(i)).Value) Is Double) Then
                                                            MyERRStr = MyERRStr & "������ N" & CStr(i) & " ����� ���� �� ���������� ������ ���� �������� ���������." & Chr(13)
                                                        Else
                                                            MyNewPriceAtQTY = appXLSRC.Worksheets(1).Range("N" & CStr(i)).Value

                                                            '---�������� - ���� �� ����� ������ � ����� ������ � ����� ����� ������
                                                            MySQLStr = "SELECT COUNT(*) AS CC "
                                                            MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) "
                                                            MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & Trim(MyPOrderN) & "') AND "
                                                            MySQLStr = MySQLStr & "(PC03002 = N'" & Trim(MyRowN) & "') AND "
                                                            MySQLStr = MySQLStr & "(PC03005 = N'" & Trim(MyItemN) & "') "
                                                            InitMyConn(False)
                                                            InitMyRec(False, MySQLStr)
                                                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                                trycloseMyRec()
                                                                MyERRStr = MyERRStr & "������ " & CStr(i) & " ���������� �������� ������ � ������� ������ ������ � ������. ���������� � ��������������." & Chr(13)
                                                            Else
                                                                MyCC = Declarations.MyRec.Fields("CC").Value
                                                                trycloseMyRec()
                                                                If MyCC = 0 Then '---��� ����� ������ � ������ 
                                                                    MyERRStr = MyERRStr & "������ " & CStr(i) & ". � Scala �� ���������� ������ � ������ �������." & Chr(13)
                                                                Else
                                                                    '---�������� - ��� �� �� ������� ������ ���������� �������
                                                                    MySQLStr = "SELECT COUNT(PC19001) as CC "
                                                                    MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) "
                                                                    MySQLStr = MySQLStr & "WHERE (PC19007 <> 0) AND "
                                                                    MySQLStr = MySQLStr & "(PC19001 = N'" & Trim(MyPOrderN) & "') "
                                                                    InitMyConn(False)
                                                                    InitMyRec(False, MySQLStr)
                                                                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                                        trycloseMyRec()
                                                                        MyERRStr = MyERRStr & "������ " & CStr(i) & " ���������� �������� ������ � ���������� �������� �� ������� ������. ���������� � ��������������." & Chr(13)
                                                                    Else
                                                                        MyCC = Declarations.MyRec.Fields("CC").Value
                                                                        trycloseMyRec()
                                                                        If MyCC <> 0 Then '---���� ���������� ������� �� ������ �� ������� 
                                                                            MyERRStr = MyERRStr & "������ " & CStr(i) & " ����� �� ������� " & Trim(MyPOrderN) & ". �� ������� ������ ���� ���������� �������, �� �� �������� ���������� ��� ����������." & Chr(13)
                                                                        Else
                                                                            '---���������� ���� (�� ���������� 1!!)
                                                                            MySQLStr = "UPDATE PC030300 "
                                                                            MySQLStr = MySQLStr & "SET PC03008 = " & Replace(CStr(MyNewPrice), ",", ".") & ", "
                                                                            MySQLStr = MySQLStr & "PC03019 = " & Replace(CStr(MyNewPriceAtQTY), ",", ".") & " "
                                                                            MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & Trim(MyPOrderN) & "') AND "
                                                                            MySQLStr = MySQLStr & "(PC03002 = N'" & Trim(MyRowN) & "') AND "
                                                                            MySQLStr = MySQLStr & "(PC03005 = N'" & Trim(MyItemN) & "') AND "
                                                                            MySQLStr = MySQLStr & "(PC03010 <> 0) "
                                                                            InitMyConn(False)
                                                                            Declarations.MyConn.Execute(MySQLStr)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        i = i + 1
                    End While
                    '---���������� ���� ������� �� �������, �������� � ����������
                    MySQLStr = "UPDATE PC010300 "
                    MySQLStr = MySQLStr & "SET PC01020 = View_1.Summ "
                    MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
                    MySQLStr = MySQLStr & "(SELECT PC030300.PC03001, SUM(ROUND(PC030300.PC03008, 2) * PC030300.PC03010 / PC030300.PC03019) AS Summ "
                    MySQLStr = MySQLStr & "FROM PC030300 INNER JOIN "
                    MySQLStr = MySQLStr & "PC010300 AS PC010300_1 ON PC030300.PC03001 = PC010300_1.PC01001 "
                    MySQLStr = MySQLStr & "WHERE (PC010300_1.PC01052 = N'" & My�POrderN & "') "
                    MySQLStr = MySQLStr & "GROUP BY PC030300.PC03001) AS View_1 ON PC010300.PC01001 = View_1.PC03001 "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                End If
                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                If MyERRStr = "" Then '---��� ������
                    MsgBox("������ ������������������ ��� ��� ������� �� ������� ����������", MsgBoxStyle.OkOnly, "��������!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "�� ����� ������� ������������������ ��� ��� ������� �� ������� ���� ������ " & Chr(13) & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
            End If
        End If
    End Sub

    Public Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �� LibreOffice ����� ��� �� ����������� ������ �� �������  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyERRStr As String                      '��������� �� �������
        Dim i As Double                             '������� �����
        Dim MyItemN As String                       '��� ������ � Scala
        Dim My�POrderN As String                    '���������� ����� ������ �� ������� � Scala
        Dim MyPOrderN As String                     '������������ ����� ������ �� ������� � Scala
        Dim MyRowN As String                        '����� ������ � ������������ ������ �� �������
        Dim MyQTY As Double                         '��� - �� � ������������ ������ �� �������
        Dim MyNewPrice As Double                    '����� ����
        Dim MyNewPriceAtQTY As Double               '����� ���� �� ����������
        Dim MySQLStr As String
        Dim MyCC As Integer                         '�������
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        If Main.OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (Main.OpenFileDialog2.FileName = "") Then
            Else
                MyERRStr = ""
                System.Windows.Forms.Application.DoEvents()
                Try
                    LOSetNotation(0)
                    oServiceManager = CreateObject("com.sun.star.ServiceManager")
                    oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                    oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                    oFileName = Replace(Main.OpenFileDialog2.FileName, "\", "/")
                    oFileName = "file:///" + oFileName
                    Dim arg(1)
                    arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                    arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                    oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                    oSheet = oWorkBook.getSheets().getByIndex(0)
                    If oSheet.getCellRangeByName("E1").String.Equals("") Then
                        MsgBox("� ������ E1 �� ���������� ����� ����������� ������ �� �������.", MsgBoxStyle.OkOnly, "��������!")
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        My�POrderN = Right("0000000000" & Trim(oSheet.getCellRangeByName("E1").String), 10)
                        i = 16
                        While Not oSheet.getCellRangeByName("D" & i).String.Equals("")
                            MyItemN = oSheet.getCellRangeByName("D" & CStr(i)).String
                            MyPOrderN = oSheet.getCellRangeByName("G" & CStr(i)).String
                            If MyPOrderN.Equals("") Then
                                MyERRStr = MyERRStr & "������ G" & CStr(i) & " ����� ������������� ������ �� ������� � Scala ������ ���� ��������." & Chr(13)
                            Else
                                MyRowN = oSheet.getCellRangeByName("H" & CStr(i)).String
                                If MyRowN.Equals("") Then
                                    MyERRStr = MyERRStr & "������ H" & CStr(i) & " ����� ������ ������������� ������ �� ������� � Scala ������ ���� ��������." & Chr(13)
                                Else
                                    Try
                                        MyQTY = oSheet.getCellRangeByName("I" & CStr(i)).Value
                                        If MyQTY = 0 Then
                                            MyERRStr = MyERRStr & "������ I" & CStr(i) & " ��� ���������� � ������ ������ 0 ���������� ��� �� ������������." & Chr(13)
                                        Else
                                            Try
                                                MyNewPrice = oSheet.getCellRangeByName("M" & CStr(i)).Value
                                                Try
                                                    MyNewPriceAtQTY = oSheet.getCellRangeByName("N" & CStr(i)).Value
                                                    '---�������� - ���� �� ����� ������ � ����� ������ � ����� ����� ������
                                                    MySQLStr = "SELECT COUNT(*) AS CC "
                                                    MySQLStr = MySQLStr & "FROM PC030300 WITH(NOLOCK) "
                                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & Trim(MyPOrderN) & "') AND "
                                                    MySQLStr = MySQLStr & "(PC03002 = N'" & Trim(MyRowN) & "') AND "
                                                    MySQLStr = MySQLStr & "(PC03005 = N'" & Trim(MyItemN) & "') "
                                                    InitMyConn(False)
                                                    InitMyRec(False, MySQLStr)
                                                    If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                        trycloseMyRec()
                                                        MyERRStr = MyERRStr & "������ " & CStr(i) & " ���������� �������� ������ � ������� ������ ������ � ������. ���������� � ��������������." & Chr(13)
                                                    Else
                                                        MyCC = Declarations.MyRec.Fields("CC").Value
                                                        trycloseMyRec()
                                                        If MyCC = 0 Then '---��� ����� ������ � ������ 
                                                            MyERRStr = MyERRStr & "������ " & CStr(i) & ". � Scala �� ���������� ������ � ������ �������." & Chr(13)
                                                        Else
                                                            '---�������� - ��� �� �� ������� ������ ���������� �������
                                                            MySQLStr = "SELECT COUNT(PC19001) as CC "
                                                            MySQLStr = MySQLStr & "FROM PC190300 WITH(NOLOCK) "
                                                            MySQLStr = MySQLStr & "WHERE (PC19007 <> 0) AND "
                                                            MySQLStr = MySQLStr & "(PC19001 = N'" & Trim(MyPOrderN) & "') "
                                                            InitMyConn(False)
                                                            InitMyRec(False, MySQLStr)
                                                            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                                                                trycloseMyRec()
                                                                MyERRStr = MyERRStr & "������ " & CStr(i) & " ���������� �������� ������ � ���������� �������� �� ������� ������. ���������� � ��������������." & Chr(13)
                                                            Else
                                                                MyCC = Declarations.MyRec.Fields("CC").Value
                                                                trycloseMyRec()
                                                                If MyCC <> 0 Then '---���� ���������� ������� �� ������ �� ������� 
                                                                    MyERRStr = MyERRStr & "������ " & CStr(i) & " ����� �� ������� " & Trim(MyPOrderN) & ". �� ������� ������ ���� ���������� �������, �� �� �������� ���������� ��� ����������." & Chr(13)
                                                                Else
                                                                    '---���������� ���� (�� ���������� 1!!)
                                                                    MySQLStr = "UPDATE PC030300 "
                                                                    MySQLStr = MySQLStr & "SET PC03008 = " & Replace(CStr(MyNewPrice), ",", ".") & ", "
                                                                    MySQLStr = MySQLStr & "PC03019 = " & Replace(CStr(MyNewPriceAtQTY), ",", ".") & " "
                                                                    MySQLStr = MySQLStr & "WHERE (PC03001 = N'" & Trim(MyPOrderN) & "') AND "
                                                                    MySQLStr = MySQLStr & "(PC03002 = N'" & Trim(MyRowN) & "') AND "
                                                                    MySQLStr = MySQLStr & "(PC03005 = N'" & Trim(MyItemN) & "') AND "
                                                                    MySQLStr = MySQLStr & "(PC03010 <> 0) "
                                                                    InitMyConn(False)
                                                                    Declarations.MyConn.Execute(MySQLStr)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Catch ex As Exception
                                                    MyERRStr = MyERRStr & "������ N" & CStr(i) & " ����� ���� �� ���������� ������ ���� �������� ���������." & Chr(13)
                                                End Try
                                            Catch ex As Exception
                                                MyERRStr = MyERRStr & "������ M" & CStr(i) & " ����� ���� ������ ���� �������� ���������." & Chr(13)
                                            End Try
                                        End If
                                    Catch ex As Exception
                                        MyERRStr = MyERRStr & "������ I" & CStr(i) & " ���������� � ������ ������ ������ ���� �������� ���������." & Chr(13)
                                    End Try
                                End If
                            End If
                            i = i + 1
                        End While
                        '---���������� ���� ������� �� �������, �������� � ����������
                        MySQLStr = "UPDATE PC010300 "
                        MySQLStr = MySQLStr & "SET PC01020 = View_1.Summ "
                        MySQLStr = MySQLStr & "FROM PC010300 INNER JOIN "
                        MySQLStr = MySQLStr & "(SELECT PC030300.PC03001, SUM(ROUND(PC030300.PC03008, 2) * PC030300.PC03010 / PC030300.PC03019) AS Summ "
                        MySQLStr = MySQLStr & "FROM PC030300 INNER JOIN "
                        MySQLStr = MySQLStr & "PC010300 AS PC010300_1 ON PC030300.PC03001 = PC010300_1.PC01001 "
                        MySQLStr = MySQLStr & "WHERE (PC010300_1.PC01052 = N'" & My�POrderN & "') "
                        MySQLStr = MySQLStr & "GROUP BY PC030300.PC03001) AS View_1 ON PC010300.PC01001 = View_1.PC03001 "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If
                Catch ex As Exception
                    MsgBox("������ : " & ex.Message, MsgBoxStyle.Critical, "��������!")
                End Try
                oWorkBook.Close(True)
                If MyERRStr = "" Then '---��� ������
                    MsgBox("������ ������������������ ��� ��� ������� �� ������� ����������", MsgBoxStyle.OkOnly, "��������!")
                Else
                    MyErrorForm = New ErrorForm
                    MyERRStr = "�� ����� ������� ������������������ ��� ��� ������� �� ������� ���� ������ " & Chr(13) & Chr(13) & MyERRStr
                    MyErrorForm.MyMsg = MyERRStr
                    MyErrorForm.ShowDialog()
                End If
            End If
        End If
    End Sub
End Module
