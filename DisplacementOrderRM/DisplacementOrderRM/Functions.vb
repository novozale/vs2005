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

    Public Function CheckRights(ByVal UserID As String, ByVal RoleName As String) As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '// �������� ���� ������������ - ����������� �� � ������������ ������
        '// 
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        On Error GoTo MyCatch
        MySQLStr = "SELECT DISTINCT ScalaSystemDB.dbo.ScaRoles.RoleName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUserToOrgnode ON ScalaSystemDB.dbo.ScaUsers.UserID = ScalaSystemDB.dbo.ScaUserToOrgnode.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoleToOrgnode ON ScalaSystemDB.dbo.ScaUserToOrgnode.OrgnodeID = ScalaSystemDB.dbo.ScaRoleToOrgnode.OrgnodeID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaRoles ON ScalaSystemDB.dbo.ScaRoleToOrgnode.RoleID = ScalaSystemDB.dbo.ScaRoles.RoleID "
        MySQLStr = MySQLStr & "WHERE (Upper(ScalaSystemDB.dbo.ScaUsers.UserName) = Upper('" & UserID & "')) AND (ScalaSystemDB.dbo.ScaRoles.RoleName = N'" & RoleName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Declarations.MyPermission = False
            CheckRights = "���������"
        Else
            Declarations.MyPermission = True
            CheckRights = "���������"
        End If
        trycloseMyRec()
        Exit Function
MyCatch:
        MsgBox(Err.Description, vbCritical, "������ Functions 5")
        CheckRights = "���������"
    End Function
End Module
