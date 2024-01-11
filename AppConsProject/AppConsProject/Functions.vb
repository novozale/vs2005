
Module Functions

    Public Sub InitMyConn(ByVal IsSystem As Boolean)

        'If Conn Is Nothing Then
        '    Conn = New ADODB.Connection
        '    Conn.CursorLocation = 3
        '    Conn.CommandTimeout = 600
        '    Conn.ConnectionTimeout = 300
        '    If Declarations.NETConnStr = "" Then
        '        Declarations.NETConnStr = "User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SPBDVL3"
        '        Declarations.ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=sqladmin;Initial Catalog=ScaDataDB;Data Source=SPBDVL3"
        '    End If
        '    If IsSystem = True Then
        '        Conn.Open(Replace(Declarations.ConnStr, "ScaDataDB", "ScalaSystemDB"))
        '    Else
        '        Conn.Open(Declarations.ConnStr)
        '    End If
        'End If

        Dim Scala As New SfwIII.Application

        On Error GoTo MyCatch
        If Conn Is Nothing Then
            Conn = New ADODB.Connection
            Conn.CursorLocation = 3
            Conn.CommandTimeout = 600
            Conn.ConnectionTimeout = 300
            If Declarations.ConnStr = "" Then
                Declarations.ConnStr = Scala.ActiveProcess.UserContext.GetConnectionString(1)
                Declarations.NETConnStr = Replace(Declarations.ConnStr, "Provider=SQLOLEDB;", "")
                Declarations.NETConnStr = Declarations.NETConnStr & ";Timeout=0;"
            End If
            If IsSystem = True Then
                Conn.Open(Replace(Declarations.ConnStr, "ScaDataDB", "ScalaSystemDB"))
            Else
                Conn.Open(Declarations.ConnStr)
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
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 1")

    End Sub

    Public Sub trycloseMyRec()
 
        On Error Resume Next
        Rec.Close()
    End Sub

    Public Sub InitMyRec(ByVal IsSystem As Boolean, ByVal sql As String)

        Dim Err

        On Error GoTo MyCatch
        InitMyConn(IsSystem)
        If Rec Is Nothing Then
            Rec = New ADODB.Recordset
        End If
        trycloseRec()
        Rec.Open(sql, Conn)
        If Conn.Errors.Count > 0 Then
            For Each Err In Conn.Errors
                Err.Raise(Err.Number, Err.Source, Err.Description)
            Next Err
        End If
        Exit Sub
MyCatch:
        MsgBox(Err.Description, vbCritical, "Ошибка Functions 2")
    End Sub
    Public Sub trycloseRec()
    
        On Error Resume Next
        Rec.Close()
    End Sub
End Module
