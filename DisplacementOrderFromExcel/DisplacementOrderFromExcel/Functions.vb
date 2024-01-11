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

    Public Function SetBlock(ByVal MyOrderNum As String, ByVal MyIsExternal As Integer)
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ��� �������� ������ � 1 ��� ���������� ���������� � ������� tbl_SalesmanWorkplace2Block
        '// ��� ��������� �������. ���� ���������� �� ��������� � ������� ����������� ������� -
        '// ������� �� �������������
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        InitMyConn(False)
        MySQLStr = "exec spp_System_SetBlock N'" & Trim(MyOrderNum) & "', " & CStr(MyIsExternal)
        Declarations.MyConn.Execute(MySQLStr)
    End Function

    Public Function RemoveBlock()
        '////////////////////////////////////////////////////////////////////////////////////////
        '// ������ ���������� � ������� tbl_SalesmanWorkplace2Block
        '//
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        InitMyConn(False)
        MySQLStr = "exec spp_System_RemoveBlock "
        Declarations.MyConn.Execute(MySQLStr)
    End Function

    Public Function IsRawMaterialsWH(ByVal MyWH As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////////////
        '// 
        '// �������� - �������� �� ������ ����� ������� ������������� �����
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT COUNT(DISTINCT tbl_WarehouseAccessoiresID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_WarehouseAccessoires0300 "
        MySQLStr = MySQLStr & "WHERE (SC23001 = N'" & MyWH & "') AND "
        MySQLStr = MySQLStr & "(IsRawMaterials = 1) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
            IsRawMaterialsWH = False
        Else
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                trycloseMyRec()
                IsRawMaterialsWH = False
            Else
                trycloseMyRec()
                IsRawMaterialsWH = True
            End If
        End If

    End Function
End Module
