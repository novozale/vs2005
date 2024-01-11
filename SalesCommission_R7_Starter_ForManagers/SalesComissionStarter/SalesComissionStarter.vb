Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports Microsoft.VisualBasic.FileIO
Imports SalesComissionStarter.spbprd2
Imports System.Collections.ObjectModel


Public Class SalesComissionStarter
    Public MyYear1 As String

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MyStartDate.Value = DateAdd(DateInterval.Month, -1, DateAdd(DateInterval.Day, 1 - Microsoft.VisualBasic.DateAndTime.Day(Now()), Now()))
        MyFinDate.Value = DateAdd(DateInterval.Day, 0 - Microsoft.VisualBasic.DateAndTime.Day(Now()), Now())
        MyCatalog.Text = My.Settings.ExportPath
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '----проверка заполненности полей---------

        '-----------------------------------------
        'MyYear1 = CStr(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.DateAndTime.Year(MyFinDate.Value), 2))
        MyYear1 = CStr(Microsoft.VisualBasic.Right(Microsoft.VisualBasic.DateAndTime.Year(Today()), 2))
        ExportData(MyStartDate.Value, MyFinDate.Value, MyYear1, MyFormat.SelectedItem, MyCatalog.Text)
        MsgBox("Закончено формирование отчетов", MsgBoxStyle.Information, "Внимание!")
    End Sub

    Private Sub ExportData(ByVal MyStartDate1 As Date, ByVal MyFinDate1 As Date, ByVal MyYear1 As String, ByVal MyFormat1 As String, ByVal MyCatalog1 As String)
        Dim myConnection As SqlConnection = Nothing
        Dim mySalesmans As SqlDataReader = Nothing
        Dim mySalesmansCNT As SqlDataReader = Nothing
        Dim MyWorkSTR As String                     'для общего каталога
        Dim count As Double
        Dim PBar As Double
        Dim execInfo As New ExecutionInfo

        '------готовим каталог для записи-----------------------------------------------
        MyWorkSTR = MyCatalog1 + "\" + "SalesComission_FROM_" + CStr(DatePart(DateInterval.Year, MyStartDate1))
        MyWorkSTR = MyWorkSTR + "-" + CStr(DatePart(DateInterval.Month, MyStartDate1)) + "-" + CStr(DatePart(DateInterval.Day, MyStartDate1))
        MyWorkSTR = MyWorkSTR + "_TO_" + CStr(DatePart(DateInterval.Year, MyFinDate1)) + "-" + CStr(DatePart(DateInterval.Month, MyFinDate1))
        MyWorkSTR = MyWorkSTR + "-" + CStr(DatePart(DateInterval.Day, MyFinDate1)) + "_Produced_" + CStr(DatePart(DateInterval.Year, Now()))
        MyWorkSTR = MyWorkSTR + "-" + CStr(DatePart(DateInterval.Month, Now())) + "-" + CStr(DatePart(DateInterval.Day, Now))
        MyWorkSTR = MyWorkSTR + "_AT_" + CStr(DatePart(DateInterval.Hour, Now())) + "-" + CStr(DatePart(DateInterval.Minute, Now()))
        My.Computer.FileSystem.CreateDirectory(MyWorkSTR)


        '------Выгружаем общий отчет----------------------------------------------------
        Dim rs As New spbprd2.ReportExecutionService
        rs.Credentials = System.Net.CredentialCache.DefaultCredentials
        rs.Url = My.Settings.SalesComissionStarter_spbprd22_ReportExecutionService
        ' Render arguments.
        Dim result As Byte() = Nothing
        Dim reportPath As String = My.Settings.CommonReport
        Dim format As String = MyFormat.Text
        Dim historyID As String = Nothing

        ' Prepare report parameter.
        Dim parameters(2) As ParameterValue
        parameters(0) = New ParameterValue()
        parameters(0).Name = "StartDate"
        parameters(0).Value = Microsoft.VisualBasic.FormatDateTime(MyStartDate1, DateFormat.ShortDate)
        parameters(1) = New ParameterValue()
        parameters(1).Name = "FinishDate"
        parameters(1).Value = Microsoft.VisualBasic.FormatDateTime(MyFinDate1, DateFormat.ShortDate)
        parameters(2) = New ParameterValue()
        parameters(2).Name = "Salesman"
        parameters(2).Value = "---"

        Dim encoding As String = String.Empty
        Dim mimeType As String = String.Empty
        Dim warnings As Warning() = Nothing
        Dim streamIDs As String() = Nothing
        Dim deviceInfo As String = Nothing
        Dim Extencion As String = Nothing
        Dim MyLng As String = "ru-RU"
        execInfo = rs.LoadReport(reportPath, historyID)
        rs.SetExecutionParameters(parameters, MyLng)
        rs.Timeout = -1
        result = rs.Render(format, deviceInfo, Extencion, mimeType, encoding, warnings, streamIDs)

        Using stream As FileStream = File.OpenWrite(MyWorkSTR + "\Общий отчет" + GetFilterString(format))
            stream.Write(result, 0, result.Length)
        End Using


        '------Читаем список продавцов----------------------------------------------
        myConnection = New SqlConnection(My.Settings.Connection)
        Dim mySalesmanCommandCNT As SqlCommand = myConnection.CreateCommand()
        mySalesmanCommandCNT.CommandText = "SELECT COUNT(DISTINCT ST010300.ST01001) AS СС "
        mySalesmanCommandCNT.CommandText = mySalesmanCommandCNT.CommandText + "FROM ST010300 INNER JOIN "
        mySalesmanCommandCNT.CommandText = mySalesmanCommandCNT.CommandText + "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName "
        mySalesmanCommandCNT.CommandText = mySalesmanCommandCNT.CommandText + "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
        mySalesmanCommandCNT.CommandType = CommandType.Text
        myConnection.Open()
        mySalesmansCNT = mySalesmanCommandCNT.ExecuteReader()
        mySalesmansCNT.Read()
        count = 100 / (mySalesmansCNT.GetValue(0) + 1)
        mySalesmansCNT.Close()
        Dim mySalesmanCommand As SqlCommand = myConnection.CreateCommand()
        'mySalesmanCommand.CommandText = "Select DISTINCT  dbo.ST010300.ST01001 AS Code, dbo.ST010300.ST01002 AS Name, "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ISNULL(dbo.GL0303" + MyYear1 + ".GL03002, '') AS CCCode, ISNULL(dbo.GL0303" + MyYear1 + ".GL03003, 'Не определен') AS CCName "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "FROM  dbo.ST010300 INNER JOIN "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName LEFT OUTER JOIN  "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "dbo.GL0303" + MyYear1 + " ON SUBSTRING(dbo.ST010300.ST01021, 7, 6) = dbo.GL0303" + MyYear1 + ".GL03002 "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "WHERE (dbo.GL0303" + MyYear1 + ".GL03001 = N'B') AND ((ST010300.ST01001 LIKE N'S%' OR "
        'mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ST010300.ST01001 LIKE N'P%')) AND (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "

        mySalesmanCommand.CommandText = "Select DISTINCT "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ST010300.ST01001 AS Code, ST010300.ST01002 AS Name, ISNULL(GL0303" + MyYear1 + ".GL03002, '') AS CCCode, "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ISNULL(GL0303" + MyYear1 + ".GL03003, 'Не определен') AS CCName "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "FROM ST010300 INNER JOIN "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "tbl_Sales_Groups ON ST010300.ST01001 = tbl_Sales_Groups.SalesmanCode LEFT OUTER JOIN "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "GL0303" + MyYear1 + " ON SUBSTRING(ST010300.ST01021, 7, 6) = GL0303" + MyYear1 + ".GL03002 "
        mySalesmanCommand.CommandText = mySalesmanCommand.CommandText + "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND (GL0303" + MyYear1 + ".GL03001 = N'B') "
        mySalesmanCommand.CommandType = CommandType.Text
        mySalesmans = mySalesmanCommand.ExecuteReader()
        PBar = 0
        ProgressBar1.Value = 0
        Application.DoEvents()

        '------Выгружаем отчеты по продавцам----------------------------------------
        Dim rs1 As New spbprd2.ReportExecutionService
        rs1.Credentials = System.Net.CredentialCache.DefaultCredentials
        rs1.Url = My.Settings.SalesComissionStarter_spbprd22_ReportExecutionService
        ' Render arguments.
        result = Nothing
        reportPath = My.Settings.DetailReport
        format = MyFormat.Text
        historyID = Nothing
        ' Prepare report parameter.
        parameters(0) = New ParameterValue()
        parameters(0).Name = "StartDate"
        parameters(0).Value = Microsoft.VisualBasic.FormatDateTime(MyStartDate1, DateFormat.ShortDate)
        parameters(1) = New ParameterValue()
        parameters(1).Name = "FinishDate"
        parameters(1).Value = Microsoft.VisualBasic.FormatDateTime(MyFinDate1, DateFormat.ShortDate)
        parameters(2) = New ParameterValue()
        parameters(2).Name = "Salesman"
        'parameters(2).Value = MyYear1

        While mySalesmans.Read()
            result = Nothing
            encoding = String.Empty
            mimeType = String.Empty
            warnings = Nothing
            streamIDs = Nothing
            deviceInfo = Nothing
            Extencion = Nothing
            MyLng = "ru-RU"
            parameters(2).Value = mySalesmans.GetValue(0)
            rs1.LoadReport(reportPath, historyID)
            rs1.SetExecutionParameters(parameters, MyLng)
            rs1.Timeout = -1
            result = rs1.Render(format, deviceInfo, Extencion, mimeType, encoding, warnings, streamIDs)
            Using stream As FileStream = File.OpenWrite(MyWorkSTR + "\" + RTrim(mySalesmans.GetValue(2)) + "-" + RTrim(mySalesmans.GetValue(3)) + "-" + mySalesmans.GetValue(0) + "-" + mySalesmans.GetValue(1) + "-отчет" + GetFilterString(format))
                stream.Write(result, 0, result.Length)
            End Using
            PBar = PBar + count
            ProgressBar1.Value = PBar
            Application.DoEvents()
        End While
        myConnection.Close()
        ProgressBar1.Value = 0
    End Sub

    Private Function GetFilterString(ByVal MyStr) As String
        Select Case MyStr
            Case "MHTML"
                Return ".mhtml"
            Case "PDF"
                Return ".pdf"
            Case "IMAGE"
                Return ".tif"
            Case "EXCEL"
                Return ".xls"
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FolderBrowserDialog1 = New FolderBrowserDialog
        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dr As DialogResult = FolderBrowserDialog1.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            MyCatalog.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
