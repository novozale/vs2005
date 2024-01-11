Module UploadToECommercFunctions
    Public Function UploadSenGobenToExcel(ByVal MyPath As String, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД информации для электронной торговли с Сен Гобеном в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        '=================Выгрузка картинок========================================
        UploadSenGoben_Pictures(MyPath, MyRange)


        '===============Выгрузка excel=============================================
        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 4
        MyWRKBook = MyObj.Workbooks.Add
        MyWRKBook.Sheets(1).Name = "Header"
        MyWRKBook.Sheets(2).Name = "Products"
        MyWRKBook.Sheets(3).Name = "Contract Pricing"
        MyWRKBook.Sheets(4).Name = "Attributes"

        MyWRKBook.Sheets(1).Select()
        '----------------Выгрузка закладки заголовок--------------------------------
        UploadSenGoben_HeaderExcel(MyWRKBook)


        MyWRKBook.Sheets(2).Select()
        i = 1
        '----------------Выгрузка продуктов ----------------------------------------
        UploadSenGoben_ProductHeaderExcel(MyWRKBook, i)
        UploadSenGoben_ProductBodyExcel(MyWRKBook, i, MyRange)


        MyWRKBook.Sheets(3).Select()
        i = 1
        '----------------Выгрузка цен ----------------------------------------------
        UploadSenGoben_PriceHeaderExcel(MyWRKBook, i)
        UploadSenGoben_PriceBodyExcel(MyWRKBook, i, MyRange)


        MyWRKBook.Sheets(4).Select()
        i = 1
        '----------------Выгрузка аттрибутов----------------------------------------
        UploadSenGoben_AttributesHeaderExcel(MyWRKBook, i)
        UploadSenGoben_AttributesBodyEXCEL(MyWRKBook, i, MyRange)


        MyWRKBook.Sheets(1).Select()
        MyWRKBook.Sheets(1).Range("A1:A1").Select()

        MyObj.Application.DisplayAlerts = False
        MyWRKBook.SaveAs(MyPath & "\ProductsForSaintGobain.xlsx")

        MyObj.Application.DisplayAlerts = True

        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Function

    Public Function UploadSenGobenToLO(ByVal MyPath As String, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка из БД информации для электронной торговли с Сен Гобеном в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer

        '=================Выгрузка картинок========================================
        UploadSenGoben_Pictures(MyPath, MyRange)

        '===============Выгрузка LibreOffice=======================================
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

        oWorkBook.getSheets().insertNewByName("Header", 0)
        oWorkBook.getSheets().insertNewByName("Products", 1)
        oWorkBook.getSheets().insertNewByName("Contract Pricing", 2)
        oWorkBook.getSheets().insertNewByName("Attributes", 3)
        oWorkBook.getSheets().removeByName("Лист1")

        '----------------Выгрузка закладки заголовок--------------------------------
        oSheet = oWorkBook.getSheets().getByName("Header")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        UploadSenGoben_HeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame)

        '----------------Выгрузка продуктов ----------------------------------------
        i = 1
        oSheet = oWorkBook.getSheets().getByName("Products")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        UploadSenGoben_ProductHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        UploadSenGoben_ProductBodyLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyRange)

        '----------------Выгрузка цен ----------------------------------------------
        i = 1
        oSheet = oWorkBook.getSheets().getByName("Contract Pricing")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        UploadSenGoben_PriceHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        UploadSenGoben_PriceBodyLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyRange)

        '----------------Выгрузка аттрибутов----------------------------------------
        i = 1
        oSheet = oWorkBook.getSheets().getByName("Attributes")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        UploadSenGoben_AttributesHeaderLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        UploadSenGoben_AttributesBodyLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i, MyRange)

        '----------------окончание процедуры----------------------------------------
        oSheet = oWorkBook.getSheets().getByName("Header")
        oWorkBook.CurrentController.setActiveSheet(oSheet)
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim arg2() As Object
        ReDim arg2(1)
        arg2(0) = mAkePropertyValue("URL", "file:///" & Replace(MyPath & "\ProductsForSaintGobain.ods", "\", "/"), oServiceManager)
        arg2(1) = mAkePropertyValue("FilterName", "calc8", oServiceManager)
        oDispatcher.executeDispatch(oFrame, ".uno:SaveAs", "", 0, arg2)
        oWorkBook.Close(True)
    End Function

    Private Sub UploadSenGoben_HeaderExcel(ByRef MyWRKBook As Object)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка страницы "header" для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 80

        MyWRKBook.ActiveSheet.Range("A1") = "Remarks"
        MyWRKBook.ActiveSheet.Range("A2") = "Contacts"
        MyWRKBook.ActiveSheet.Range("A5") = "Comments"
        MyWRKBook.ActiveSheet.Range("A2:A4").MergeCells = True

        MyWRKBook.ActiveSheet.Range("B1") = "Enter any comments. This will not be seen unless someone refers back to this file for troubleshooting."
        MyWRKBook.ActiveSheet.Range("B2") = "In the United States:   If you have any questions regarding the template or application, please contact Perfect Commerce Content Team:  DL-SupplierServices@perfect.com "
        MyWRKBook.ActiveSheet.Range("B3") = "In Europe:   If you have any questions regarding the template or application, please contact Perfect Commerce Content Team:  DL-SupplierSupportEurope@perfect.com"
        MyWRKBook.ActiveSheet.Range("B4") = " If you have any Customer-specific content questions, please contact Customer directly.  [This field may be replaced with appropriate Customer email contact information.]"

        MyWRKBook.ActiveSheet.Range("A1:A5").Select()
        MyWRKBook.ActiveSheet.Range("A1:A5").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A1:A5").Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A1:A5").WrapText = True
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A1:A5").Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A1:A5").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A1:A5").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A1:A5").Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B1:B4").Select()
        MyWRKBook.ActiveSheet.Range("B1:B4").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B1:B4").Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B1:B4").WrapText = True
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B1:B4").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("B1:B4").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B1:B4").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("B1:B4").Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Range("B5:B5").Select()
        MyWRKBook.ActiveSheet.Range("B5:B5").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B5:B5").Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("B5:B5").WrapText = True
        With MyWRKBook.ActiveSheet.Range("B5:B5").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B5:B5").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B5:B5").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B5:B5").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B5:B5").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("B5:B5").Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("B5:B5").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("B5:B5").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("B5:B5").Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = False
        End With
    End Sub

    Private Sub UploadSenGoben_HeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка страницы "header" для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 15200

        oSheet.getCellRangeByName("A1").String = "Remarks"
        oSheet.getCellRangeByName("A2").String = "Contacts"
        oSheet.getCellRangeByName("A5").String = "Comments"
        LOMergeCells(oServiceManager, oDispatcher, oFrame, "A2:A4")
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1:A5", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1:A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1:A5", 8)
        oSheet.getCellRangeByName("A1:A5").VertJustify = 2
        oSheet.getCellRangeByName("A1:A5").HoriJustify = 2
        oSheet.getCellRangeByName("A1:A5").CellBackColor = RGB(230, 255, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A1:A5", 70, 20, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        oSheet.getCellRangeByName("B1").String = "Enter any comments. This will not be seen unless someone refers back to this file for troubleshooting."
        oSheet.getCellRangeByName("B2").String = "In the United States:   If you have any questions regarding the template or application, please contact Perfect Commerce Content Team:  DL-SupplierServices@perfect.com"
        oSheet.getCellRangeByName("B3").String = "In Europe:   If you have any questions regarding the template or application, please contact Perfect Commerce Content Team:  DL-SupplierSupportEurope@perfect.com"
        oSheet.getCellRangeByName("B4").String = "If you have any Customer-specific content questions, please contact Customer directly.  [This field may be replaced with appropriate Customer email contact information.]"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B1:B5", "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B1:B5", 8)
        oSheet.getCellRangeByName("B1:B5").VertJustify = 2
        oSheet.getCellRangeByName("B1:B5").HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B1:B5")
        oSheet.getCellRangeByName("B5:B5").CellBackColor = RGB(230, 255, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "B1:B4", 70, 20, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "B5:B5", 70, 20, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

    End Sub

    Private Sub UploadSenGoben_ProductHeaderExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "Action"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 0
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "PartNum"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "ShortDesc_en"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 0
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "LongDesc_en"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 0
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "ShortDesc_ru"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 25
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "LongDesc_ru"
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "UNSPSCSGA"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "UoM"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "Keywords_ru"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "MfrName"
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = "MfrPartNum"
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = "BuyerPartNum"
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("M" & CStr(i)) = "PackagingQuantity"
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("N" & CStr(i)) = "PackagingUoM"
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("O" & CStr(i)) = "Icon1"
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("P" & CStr(i)) = "Icon2"
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Q" & CStr(i)) = "Icon3"
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("R" & CStr(i)) = "Icon4"
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("S" & CStr(i)) = "Icon5"
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("T" & CStr(i)) = "ThumbName"
        MyWRKBook.ActiveSheet.Columns("T:T").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("U" & CStr(i)) = "PicName"
        MyWRKBook.ActiveSheet.Columns("U:U").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("V" & CStr(i)) = "PicName2"
        MyWRKBook.ActiveSheet.Columns("V:V").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("W" & CStr(i)) = "PicName3"
        MyWRKBook.ActiveSheet.Columns("W:W").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("X" & CStr(i)) = "PicName4"
        MyWRKBook.ActiveSheet.Columns("X:X").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Y" & CStr(i)) = "PicName5"
        MyWRKBook.ActiveSheet.Columns("Y:Y").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Z" & CStr(i)) = "SupplierURL"
        MyWRKBook.ActiveSheet.Columns("Z:Z").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AA" & CStr(i)) = "SupplierURLDesc_ru"
        MyWRKBook.ActiveSheet.Columns("AA:AA").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AB" & CStr(i)) = "TechSpec"
        MyWRKBook.ActiveSheet.Columns("AB:AB").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AC" & CStr(i)) = "TechSpecDesc_ru"
        MyWRKBook.ActiveSheet.Columns("AC:AC").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AD" & CStr(i)) = "MSDS"
        MyWRKBook.ActiveSheet.Columns("AD:AD").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AE" & CStr(i)) = "MSDSDesc_ru"
        MyWRKBook.ActiveSheet.Columns("AE:AE").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AF" & CStr(i)) = "PreferredItemFlag"
        MyWRKBook.ActiveSheet.Columns("AF:AF").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AG" & CStr(i)) = "EditableField"
        MyWRKBook.ActiveSheet.Columns("AG:AG").ColumnWidth = 12


        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AG" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_ProductHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2280
        oSheet.getColumns().getByName("A").IsVisible = False
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2280
        oSheet.getColumns().getByName("C").IsVisible = False
        oSheet.getColumns().getByName("D").Width = 2280
        oSheet.getColumns().getByName("D").IsVisible = False
        oSheet.getColumns().getByName("E").Width = 4750
        oSheet.getColumns().getByName("F").Width = 7600
        oSheet.getColumns().getByName("G").Width = 2280
        oSheet.getColumns().getByName("H").Width = 2280
        oSheet.getColumns().getByName("I").Width = 2280
        oSheet.getColumns().getByName("J").Width = 2280
        oSheet.getColumns().getByName("K").Width = 2280
        oSheet.getColumns().getByName("L").Width = 2280
        oSheet.getColumns().getByName("M").Width = 2280
        oSheet.getColumns().getByName("N").Width = 2280
        oSheet.getColumns().getByName("O").Width = 2280
        oSheet.getColumns().getByName("P").Width = 2280
        oSheet.getColumns().getByName("Q").Width = 2280
        oSheet.getColumns().getByName("R").Width = 2280
        oSheet.getColumns().getByName("S").Width = 2280
        oSheet.getColumns().getByName("T").Width = 2280
        oSheet.getColumns().getByName("U").Width = 2280
        oSheet.getColumns().getByName("V").Width = 2280
        oSheet.getColumns().getByName("W").Width = 2280
        oSheet.getColumns().getByName("X").Width = 2280
        oSheet.getColumns().getByName("Y").Width = 2280
        oSheet.getColumns().getByName("Z").Width = 2280
        oSheet.getColumns().getByName("AA").Width = 2280
        oSheet.getColumns().getByName("AB").Width = 2280
        oSheet.getColumns().getByName("AC").Width = 2280
        oSheet.getColumns().getByName("AD").Width = 2280
        oSheet.getColumns().getByName("AE").Width = 2280
        oSheet.getColumns().getByName("AF").Width = 2280
        oSheet.getColumns().getByName("AG").Width = 2280

        oSheet.getCellRangeByName("A" & CStr(i)).String = "" '"Action"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "PartNum"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "" '"ShortDesc_en"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "" '"LongDesc_en"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "ShortDesc_ru"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "LongDesc_ru"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "UNSPSCSGA"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "UoM"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Keywords_ru"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "MfrName"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "MfrPartNum"
        oSheet.getCellRangeByName("L" & CStr(i)).String = "BuyerPartNum"
        oSheet.getCellRangeByName("M" & CStr(i)).String = "PackagingQuantity"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "PackagingUoM"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "Icon1"
        oSheet.getCellRangeByName("P" & CStr(i)).String = "Icon2"
        oSheet.getCellRangeByName("Q" & CStr(i)).String = "Icon3"
        oSheet.getCellRangeByName("R" & CStr(i)).String = "Icon4"
        oSheet.getCellRangeByName("S" & CStr(i)).String = "Icon5"
        oSheet.getCellRangeByName("T" & CStr(i)).String = "ThumbName"
        oSheet.getCellRangeByName("U" & CStr(i)).String = "PicName"
        oSheet.getCellRangeByName("V" & CStr(i)).String = "PicName2"
        oSheet.getCellRangeByName("W" & CStr(i)).String = "PicName3"
        oSheet.getCellRangeByName("X" & CStr(i)).String = "PicName4"
        oSheet.getCellRangeByName("Y" & CStr(i)).String = "PicName5"
        oSheet.getCellRangeByName("Z" & CStr(i)).String = "SupplierURL"
        oSheet.getCellRangeByName("AA" & CStr(i)).String = "SupplierURLDesc_ru"
        oSheet.getCellRangeByName("AB" & CStr(i)).String = "TechSpec"
        oSheet.getCellRangeByName("AC" & CStr(i)).String = "TechSpecDesc_ru"
        oSheet.getCellRangeByName("AD" & CStr(i)).String = "MSDS"
        oSheet.getCellRangeByName("AE" & CStr(i)).String = "MSDSDesc_ru"
        oSheet.getCellRangeByName("AF" & CStr(i)).String = "PreferredItemFlag"
        oSheet.getCellRangeByName("AG" & CStr(i)).String = "EditableField"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AG" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AG" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AG" & CStr(i), 8)
        oSheet.getCellRangeByName("A" & CStr(i) & ":AG" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":AG" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AG" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":AG" & CStr(i)).CellBackColor = RGB(230, 255, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":AG" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_ProductBodyExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка списка продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Exec spp_WEB_SaintGobain_ProductsUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        Try
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
            Else
                MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
                trycloseMyRec()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UploadSenGoben_ProductBodyLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка списка продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyLORange As Object

        MySQLStr = "Exec spp_WEB_SaintGobain_ProductsUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(20)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArr(7) = Declarations.MyRec.Fields(7).Value
                MyArr(8) = Declarations.MyRec.Fields(8).Value
                MyArr(9) = Declarations.MyRec.Fields(9).Value
                MyArr(10) = Declarations.MyRec.Fields(10).Value
                MyArr(11) = Declarations.MyRec.Fields(11).Value
                MyArr(12) = CDbl(Declarations.MyRec.Fields(12).Value)
                MyArr(13) = Declarations.MyRec.Fields(13).Value
                MyArr(14) = Declarations.MyRec.Fields(14).Value
                MyArr(15) = Declarations.MyRec.Fields(15).Value
                MyArr(16) = Declarations.MyRec.Fields(16).Value
                MyArr(17) = Declarations.MyRec.Fields(17).Value
                MyArr(18) = Declarations.MyRec.Fields(18).Value
                MyArr(19) = Declarations.MyRec.Fields(19).Value
                MyArr(20) = Declarations.MyRec.Fields(20).Value
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyLORange = oSheet.getCellRangeByName("A" & CStr(i) & ":U" & CStr(i + MyL))
            MyLORange.setDataArray(MyArrStr)
        End If
    End Sub

    Private Sub UploadSenGoben_PriceHeaderExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка цен продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "Action"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 0
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "PartNum"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "ContractPrice"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "CurrencyCode"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "EffectiveDate"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "ExpireDate"
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "LeadTime"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "PriceUnit"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "ValidFromQuantity"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "MinOrderQuantity"
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("K" & CStr(i)) = "MaxOrderQuantity"
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("L" & CStr(i)) = "QuantityOrderInterval"
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("M" & CStr(i)) = "Tier2Quantity"
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("N" & CStr(i)) = "Tier2Price"
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("O" & CStr(i)) = "EffectiveDate2"
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("P" & CStr(i)) = "ExpireDate2"
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Q" & CStr(i)) = "Tier3Quantity"
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("R" & CStr(i)) = "Tier3Price"
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("S" & CStr(i)) = "EffectiveDate3"
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("T" & CStr(i)) = "ExpireDate3"
        MyWRKBook.ActiveSheet.Columns("T:T").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("U" & CStr(i)) = "Tier4Quantity"
        MyWRKBook.ActiveSheet.Columns("U:U").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("V" & CStr(i)) = "Tier4Price"
        MyWRKBook.ActiveSheet.Columns("V:V").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("W" & CStr(i)) = "EffectiveDate4"
        MyWRKBook.ActiveSheet.Columns("W:W").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("X" & CStr(i)) = "ExpireDate4"
        MyWRKBook.ActiveSheet.Columns("X:X").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Y" & CStr(i)) = "Tier5Quantity"
        MyWRKBook.ActiveSheet.Columns("Y:Y").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("Z" & CStr(i)) = "Tier5Price"
        MyWRKBook.ActiveSheet.Columns("Z:Z").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AA" & CStr(i)) = "EffectiveDate5"
        MyWRKBook.ActiveSheet.Columns("AA:AA").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("AB" & CStr(i)) = "ExpireDate5"
        MyWRKBook.ActiveSheet.Columns("AB:AB").ColumnWidth = 12


        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":AB" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_PriceHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка цен продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2280
        oSheet.getColumns().getByName("A").IsVisible = False
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2280
        oSheet.getColumns().getByName("D").Width = 2280
        oSheet.getColumns().getByName("E").Width = 2280
        oSheet.getColumns().getByName("F").Width = 2280
        oSheet.getColumns().getByName("G").Width = 2280
        oSheet.getColumns().getByName("H").Width = 2280
        oSheet.getColumns().getByName("I").Width = 2280
        oSheet.getColumns().getByName("J").Width = 2280
        oSheet.getColumns().getByName("K").Width = 2280
        oSheet.getColumns().getByName("L").Width = 2280
        oSheet.getColumns().getByName("M").Width = 2280
        oSheet.getColumns().getByName("N").Width = 2280
        oSheet.getColumns().getByName("O").Width = 2280
        oSheet.getColumns().getByName("P").Width = 2280
        oSheet.getColumns().getByName("Q").Width = 2280
        oSheet.getColumns().getByName("R").Width = 2280
        oSheet.getColumns().getByName("S").Width = 2280
        oSheet.getColumns().getByName("T").Width = 2280
        oSheet.getColumns().getByName("U").Width = 2280
        oSheet.getColumns().getByName("V").Width = 2280
        oSheet.getColumns().getByName("W").Width = 2280
        oSheet.getColumns().getByName("X").Width = 2280
        oSheet.getColumns().getByName("Y").Width = 2280
        oSheet.getColumns().getByName("Z").Width = 2280
        oSheet.getColumns().getByName("AA").Width = 2280
        oSheet.getColumns().getByName("AB").Width = 2280

        oSheet.getCellRangeByName("A" & CStr(i)).String = "Action"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "PartNum"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "ContractPrice"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "CurrencyCode"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "EffectiveDate"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "ExpireDate"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "LeadTime"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "PriceUnit"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "ValidFromQuantity"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "MinOrderQuantity"
        oSheet.getCellRangeByName("K" & CStr(i)).String = "MaxOrderQuantity"
        oSheet.getCellRangeByName("L" & CStr(i)).String = "QuantityOrderInterval"
        oSheet.getCellRangeByName("M" & CStr(i)).String = "Tier2Quantity"
        oSheet.getCellRangeByName("N" & CStr(i)).String = "Tier2Price"
        oSheet.getCellRangeByName("O" & CStr(i)).String = "EffectiveDate2"
        oSheet.getCellRangeByName("P" & CStr(i)).String = "ExpireDate2"
        oSheet.getCellRangeByName("Q" & CStr(i)).String = "Tier3Quantity"
        oSheet.getCellRangeByName("R" & CStr(i)).String = "Tier3Price"
        oSheet.getCellRangeByName("S" & CStr(i)).String = "EffectiveDate3"
        oSheet.getCellRangeByName("T" & CStr(i)).String = "ExpireDate3"
        oSheet.getCellRangeByName("U" & CStr(i)).String = "Tier4Quantity"
        oSheet.getCellRangeByName("V" & CStr(i)).String = "Tier4Price"
        oSheet.getCellRangeByName("W" & CStr(i)).String = "EffectiveDate4"
        oSheet.getCellRangeByName("X" & CStr(i)).String = "ExpireDate4"
        oSheet.getCellRangeByName("Y" & CStr(i)).String = "Tier5Quantity"
        oSheet.getCellRangeByName("Z" & CStr(i)).String = "Tier5Price"
        oSheet.getCellRangeByName("AA" & CStr(i)).String = "EffectiveDate5"
        oSheet.getCellRangeByName("AB" & CStr(i)).String = "ExpireDate5"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AB" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AB" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AB" & CStr(i), 8)
        oSheet.getCellRangeByName("A" & CStr(i) & ":AB" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":AB" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":AB" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":AB" & CStr(i)).CellBackColor = RGB(230, 255, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":AB" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_PriceBodyExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела цен продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Exec spp_WEB_SaintGobain_PriceUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        Try
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
            Else
                MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
                trycloseMyRec()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UploadSenGoben_PriceBodyLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела цен продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyLORange As Object

        MySQLStr = "Exec spp_WEB_SaintGobain_PriceUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(9)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = CDbl(Declarations.MyRec.Fields(2).Value)
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArr(7) = Declarations.MyRec.Fields(7).Value
                MyArr(8) = Declarations.MyRec.Fields(8).Value
                MyArr(9) = CInt(Declarations.MyRec.Fields(9).Value)
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()


            End While
            MyLORange = oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i + MyL))
            MyLORange.setDataArray(MyArrStr)
        End If
    End Sub

    Private Sub UploadSenGoben_AttributesHeaderExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка аттрибутов продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "Action"
        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 0
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "PartNum"
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "Class_stg999Bdf"
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "stg999view"
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "FixMultiLangAttribute_stg999_User_Defined_Field_1_ru"
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "FixMultiLangAttribute_stg999_User_Defined_Field_2_ru"
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "FixMultiLangAttribute_stg999_User_Defined_Field_3_ru"
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "FixMultiLangAttribute_stg999_User_Defined_Field_4_ru"
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "FixMultiLangAttribute_stg999_User_Defined_Field_5_ru"
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 50


        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(6).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).WrapText = True
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Interior
            .Color = 65535
            .TintAndShade = 0.9
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":I" & CStr(i)).Font
            .Name = "Arial"
            .Size = 8
            .Color = 0
            .Bold = True
        End With

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_AttributesHeaderLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка аттрибутов продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 2280
        oSheet.getColumns().getByName("A").IsVisible = False
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2280
        oSheet.getColumns().getByName("D").Width = 2280
        oSheet.getColumns().getByName("E").Width = 2280
        oSheet.getColumns().getByName("F").Width = 2280
        oSheet.getColumns().getByName("G").Width = 2280
        oSheet.getColumns().getByName("H").Width = 2280
        oSheet.getColumns().getByName("I").Width = 9500

        oSheet.getCellRangeByName("A" & CStr(i)).String = "Action"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "PartNum"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Class_stg999Bdf"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "stg999view"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "FixMultiLangAttribute_stg999_User_Defined_Field_1_ru"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "FixMultiLangAttribute_stg999_User_Defined_Field_2_ru"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "FixMultiLangAttribute_stg999_User_Defined_Field_3_ru"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "FixMultiLangAttribute_stg999_User_Defined_Field_4_ru"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "FixMultiLangAttribute_stg999_User_Defined_Field_5_ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i), 8)
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":I" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i)).CellBackColor = RGB(230, 255, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":I" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        i = i + 1
    End Sub

    Private Sub UploadSenGoben_AttributesBodyExcel(ByRef MyWRKBook As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела аттрибутов продуктов для Сен Гобен в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Exec spp_WEB_SaintGobain_AttributesUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        Try
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
            Else
                MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
                trycloseMyRec()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UploadSenGoben_AttributesBodyLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка тела аттрибутов продуктов для Сен Гобен в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyLORange As Object

        MySQLStr = "Exec spp_WEB_SaintGobain_AttributesUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(8)
                MyArr(0) = Declarations.MyRec.Fields(0).Value
                MyArr(1) = Declarations.MyRec.Fields(1).Value
                MyArr(2) = Declarations.MyRec.Fields(2).Value
                MyArr(3) = Declarations.MyRec.Fields(3).Value
                MyArr(4) = Declarations.MyRec.Fields(4).Value
                MyArr(5) = Declarations.MyRec.Fields(5).Value
                MyArr(6) = Declarations.MyRec.Fields(6).Value
                MyArr(7) = Declarations.MyRec.Fields(7).Value
                MyArr(8) = Declarations.MyRec.Fields(8).Value

                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyLORange = oSheet.getCellRangeByName("A" & CStr(i) & ":I" & CStr(i + MyL))
            MyLORange.setDataArray(MyArrStr)
        End If
    End Sub

    Private Sub UploadSenGoben_Pictures(ByVal MyPath As String, ByVal MyRange As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка картинок для Сен Гобен в определенный каталог
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        MySQLStr = "exec spp_WEB_SaintGobain_PicturesUpload " & CStr(MyRange)
        InitMyConn(False)
        InitMyRec(False, MySQLStr)

        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("Для данной выгрузки нет ни одной картинки.", MsgBoxStyle.Critical, "Внимание!")
            Exit Sub
        Else
            Declarations.MyRec.MoveLast()
            MyUploadInfoToSaintGobain.Label5.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()

            i = 0
            While Not Declarations.MyRec.EOF = True
                DownloadOnePicture(MyPath, Declarations.MyRec.Fields("ScalaItemCode").Value, Declarations.MyRec.Fields("Picture").Value)
                i = i + 1
                MyUploadInfoToSaintGobain.Label6.Text = i
                Application.DoEvents()
                Declarations.MyRec.MoveNext()
            End While
        End If
    End Sub

    Private Sub DownloadOnePicture(ByVal MyCatalog As String, ByVal MyPictureName As String, ByVal MyPictureByte As Byte())
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура выгрузки одной картинки в каталог
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim stream As New IO.MemoryStream(MyPictureByte)
        Dim picture As Image

        Try
            picture = Image.FromStream(stream)
            picture.Save(MyCatalog + "\" + MyPictureName + ".jpg")
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Внимание!")
        End Try
    End Sub
End Module
