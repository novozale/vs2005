Module LOFunctions

    Public Sub LOFontSetBold(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ������ ������ ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(0)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "Bold"
        args1(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:Bold", "", 0, args1)
    End Sub

    Public Sub LOFontSetFamilyName(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String, ByVal MyFont As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ������ ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(5)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "CharFontName.StyleName"
        args1(0).Value = "�������"
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "CharFontName.Pitch"
        args1(1).Value = 2
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "CharFontName.CharSet"
        args1(2).Value = 0
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "CharFontName.Family"
        args1(3).Value = 5
        args1(4) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(4).Name = "CharFontName.FamilyName"
        args1(4).Value = MyFont
        oDispatcher.executeDispatch(oFrame, ".uno:CharFontName", "", 0, args1)
    End Sub

    Public Sub LOFontSetSize(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String, ByVal MySize As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ������� ������ ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args1() As Object
        ReDim args1(2)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "FontHeight.Height"
        args1(0).Value = MySize
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "FontHeight.Prop"
        args1(1).Value = 100
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "FontHeight.Diff"
        args1(2).Value = 0
        oDispatcher.executeDispatch(oFrame, ".uno:FontHeight", "", 0, args1)
    End Sub

    Public Sub LOFormatCells(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String, ByVal MyFormat As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ ��������� �����
        '//     100 �����
        '//     4   ����� 2 ����� ����� �������
        '//     0   ����� ��� ������� �����
        '//     3   ����� ��� ������� ����� � �������������
        '//     36  ���� � ������� dd.MM.yyyy
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "NumberFormatValue"
        args2(0).Value = MyFormat
        oDispatcher.executeDispatch(oFrame, ".uno:NumberFormatValue", "", 0, args2)
    End Sub

    Public Sub LOMergeCells(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        oDispatcher.executeDispatch(oFrame, ".uno:MergeCells", "", 0, args)
    End Sub

    Public Sub LOWrapText(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� �������� �������� ��� ���������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = MyRange
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        Dim args2() As Object
        ReDim args2(0)
        args2(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args2(0).Name = "WrapText"
        args2(0).Value = True
        oDispatcher.executeDispatch(oFrame, ".uno:WrapText", "", 0, args2)
    End Sub

    Public Function mAkePropertyValue(ByVal cName, ByVal uValue, ByRef oServiceManager) As Object
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���������� ��� LO
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim oPropertyValue As Object

        oPropertyValue = oServiceManager.Bridge_getStruct("com.sun.star.beans.PropertyValue")
        oPropertyValue.Name = cName
        oPropertyValue.Value = uValue

        mAkePropertyValue = oPropertyValue
        oPropertyValue = Nothing
    End Function

    Public Function mAkeSortValue(ByVal FieldN, ByVal SortAscending, ByRef oServiceManager) As Object
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���������� ��� ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim oSortValue As Object
        oSortValue = oServiceManager.Bridge_getStruct("com.sun.star.util.SortField")
        oSortValue.Field = FieldN
        oSortValue.SortAscending = SortAscending
        mAkeSortValue = oSortValue
        oSortValue = Nothing
    End Function

    Public Sub LOSetCellProtection(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByVal MyRange As String, ByVal MyProtection As Boolean)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� �������� ������ �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyRange.Equals("") Then
            '-----���� ����
            Dim args() As Object
            ReDim args(3)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "Protection.Locked"
            args(0).Value = MyProtection
            args(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(1).Name = "Protection.FormulasHidden"
            args(1).Value = False
            args(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(2).Name = "Protection.Hidden"
            args(2).Value = False
            args(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(3).Name = "Protection.HiddenInPrintout"
            args(3).Value = False
            oDispatcher.executeDispatch(oFrame, ".uno:Protection", "", 0, args)
        Else
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = MyRange
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

            Dim args1() As Object
            ReDim args1(3)
            args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(0).Name = "Protection.Locked"
            args1(0).Value = MyProtection
            args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(1).Name = "Protection.FormulasHidden"
            args1(1).Value = False
            args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(2).Name = "Protection.Hidden"
            args1(2).Value = False
            args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args1(3).Name = "Protection.HiddenInPrintout"
            args1(3).Value = False
            oDispatcher.executeDispatch(oFrame, ".uno:Protection", "", 0, args1)
        End If
    End Sub

    Public Sub LOSetBorders(ByRef oServiceManager As Object, ByRef oSheet As Object, _
        ByVal MyRange As String, ByVal MyOutThickness As Integer, ByVal MyIntThickness As Integer, _
        ByVal MyColor As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ���� ����� ��������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim BasicBorder As Object
        Dim oBorder As Object
        Dim oRange As Object

        oRange = oSheet.getCellRangeByName(MyRange)
        oBorder = oRange.TableBorder
        BasicBorder = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine")

        BasicBorder.Color = MyColor
        BasicBorder.OuterLineWidth = MyIntThickness
        oBorder.VerticalLine = BasicBorder
        oBorder.HorizontalLine = BasicBorder

        BasicBorder.OuterLineWidth = MyOutThickness
        oBorder.LeftLine = BasicBorder
        oBorder.TopLine = BasicBorder
        oBorder.RightLine = BasicBorder
        oBorder.BottomLine = BasicBorder

        oRange.TableBorder = oBorder
    End Sub

    Public Sub LOSetBGColor(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oFrame As Object, ByRef oWorkbook As Object, ByRef oSheet As Object, _
        ByVal MyRange As String, ByVal MyColor As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ����� ���� ��� ��������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Not MyRange.Equals("") Then
            Dim args() As Object
            ReDim args(0)
            args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
            args(0).Name = "ToPoint"
            args(0).Value = MyRange
            oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        Else
            oWorkbook.getCurrentController.select(oSheet)
        End If

        Dim args1() As Object
        ReDim args1(3)
        args1(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(0).Name = "BackgroundPattern.Transparent"
        args1(0).Value = False
        args1(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(1).Name = "BackgroundPattern.BackColor"
        args1(1).Value = MyColor
        args1(2) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(2).Name = "BackgroundPattern.URL"
        args1(2).Value = ""
        args1(3) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args1(3).Name = "BackgroundPattern.Filtername"
        args1(3).Value = ""
        oDispatcher.executeDispatch(oFrame, ".uno:BackgroundPattern", "", 0, args1)
    End Sub

    Public Sub LOSetFormula(ByRef oServiceManager As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef oWorkBook As Object, ByVal MyRange As String, ByVal MyFormula As String, _
        ByVal AddressConvention As Long)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� ��� ��������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////
        '//AddressConvention:
        '//     0 - Calc A1
        '//     1 - Excel A1
        '//     2 - Excel R1C1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim tokenizeFormula As Object
        Dim oRange As Object
        Dim oAddress As Object
        Dim i As Integer
        Dim j As Integer
        Dim MyCa As Object

        MyCa = oServiceManager.Bridge_GetStruct("com.sun.star.table.CellAddress")
        oRange = oSheet.getCellRangeByName(MyRange)
        oAddress = oRange.RangeAddress

        Dim hParser As Object
        hParser = oWorkBook.CreateInstance("com.sun.star.sheet.FormulaParser")

        For i = oAddress.StartColumn To oAddress.EndColumn
            For j = oAddress.StartRow To oAddress.EndRow
                oRange = oSheet.getCellByPosition(i, j)
                MyCa = oRange.CellAddress
                hParser.FormulaConvention = AddressConvention
                tokenizeFormula = hParser.ParseFormula(MyFormula, MyCa)
                oRange.SetTokens(tokenizeFormula)
            Next j
        Next i
    End Sub

    Public Sub LOSetNotation(ByVal AddressConvention As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ������� ��� ������ � Libre Office
        '//
        '////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////
        '//AddressConvention:
        '//     0 - Calc A1
        '//     1 - Excel A1
        '//     2 - Excel R1C1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object

        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)


        Dim aProps() As Object
        ReDim aProps(1)
        aProps(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        aProps(0).Name = "nodepath"
        aProps(0).Value = "org.openoffice.Office.Calc/Formula/Syntax"
        aProps(1) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        aProps(1).Name = "enableasync"
        aProps(1).Value = False

        Dim oSM = CreateObject("com.sun.star.ServiceManager")
        Dim oConfig As Object
        oConfig = oSM.createInstance("com.sun.star.configuration.ConfigurationProvider")
        Dim oFormulaSyntax As Object
        oFormulaSyntax = oConfig.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess", aProps)
        oFormulaSyntax.replaceByName("Grammar", AddressConvention)
        oFormulaSyntax.commitChanges()
        oConfig.flush()

        oWorkBook.Close(True)
    End Sub

    Public Sub LOSetValidation(ByRef oSheet As Object, _
        ByVal MyRange As String, ByVal MyFormula As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� �������� ��� �������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oRange As Object
        Dim oValidation As Object

        oRange = oSheet.getCellRangeByName(MyRange)
        oValidation = oRange.Validation
        oValidation.Type = 6
        oValidation.ShowList = 1
        oValidation.Operator = 1
        oValidation.ShowErrorMessage = True

        oValidation.setFormula1(MyFormula)
        oRange.Validation = oValidation
    End Sub

    Public Sub LOPasswordProtect(ByRef oSheet As Object, _
        ByVal MyPass As String, ByVal MyOnOff As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� ��� ������ ������ �����
        '// MyOnOff:    True    - ����������
        '//             False   - �����
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MyOnOff Then
            oSheet.protect(MyPass)
        Else
            oSheet.unprotect(MyPass)
        End If
    End Sub
End Module
