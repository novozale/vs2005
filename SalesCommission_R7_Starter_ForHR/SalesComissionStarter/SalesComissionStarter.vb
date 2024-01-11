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

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MyMonth.SelectedIndex = CInt(Microsoft.VisualBasic.DateAndTime.Month(DateAdd(DateInterval.Month, -1, Now())) - 1)
        MyCatalogR.Text = My.Settings.DistributionCatalogPath

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FolderBrowserDialog1 = New FolderBrowserDialog
        'FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.SelectedPath = My.Settings.InitCatalog
        Dim dr As DialogResult = FolderBrowserDialog1.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            MyCatalogR.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1 = New OpenFileDialog()
        OpenFileDialog1.InitialDirectory = My.Settings.InitCatalog
        OpenFileDialog1.Filter = "*.xls|*.xls"
        Dim dr As DialogResult = OpenFileDialog1.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            MyFileR.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        AllocateData(MyMonth.Text, MyFileR.Text, MyCatalogR.Text)
        MsgBox("��������� ������������� �������", MsgBoxStyle.Information, "��������!")
    End Sub

    Private Sub AllocateData(ByVal MyMonth As String, ByVal MyFile As String, ByVal MyCatalog As String)
        Dim appXLSRC As Object                             'Excel ������ - ����� �����
        Dim appXLDST As Object                             'Excel ������ - ������������� ����� �� ��������
        Dim Counter As Double
        Dim PBar As Double
        Dim Salesman As String
        Dim cc As Object
        Dim LineNumber As Integer

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(MyFile)
        appXLDST = CreateObject("Excel.Application")
        appXLDST.Application.displayalerts = False

        Counter = 100 / (My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls").Count)
        PBar = 0
        ProgressBar2.Value = 0
        Application.DoEvents()
        '---������ ������, �� ������� �������� ������-----------------------------------
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls")
            appXLDST.Workbooks.Open(foundFile)
            '---�������� ������ ����� �� ����� �� �������-------------------------------
            Try
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Delete()
            Catch ex As Exception
            End Try
            Try
                appXLDST.Worksheets("������").copy(appXLDST.Worksheets("������"))
                appXLDST.Worksheets("������ (2)").name = GetRowNumber(MyMonth)
                appXLDST.ActiveWorkbook.Close(SaveChanges:=True)
                '---���� ���������----------------------------------------------------------
                appXLDST.Workbooks.Open(foundFile)
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("A2") = Microsoft.VisualBasic.Year(Today())
                appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D2") = MyMonth
                Salesman = "S"
                LineNumber = 9
                Do While ((Microsoft.VisualBasic.Left(Salesman, 1) = "S") Or (Microsoft.VisualBasic.Left(Salesman, 1) = "P"))
                    LineNumber = LineNumber + 1
                    Salesman = appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("C" + CStr(LineNumber)).Value
                    cc = appXLSRC.Worksheets(1).Range("C3:C1000").Find(Salesman)
                    '---������� ��������--------------------------------------------------------
                    If Not cc Is Nothing Then
                        '--- ���� � ������
                        If (appXLSRC.Worksheets(1).Range("D" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D" + CStr(LineNumber)) = CStr(appXLSRC.Worksheets(1).Range("D" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("D" + CStr(LineNumber)) = ""
                        End If

                        '--- % ������� � ������
                        If (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("E" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("E" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("F" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("F" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("F" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("F" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("G" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("G" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("G" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("G" + CStr(LineNumber)) = 0
                        End If

                        '--- ��������� �� ������� - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("H" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("H" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("H" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("H" + CStr(LineNumber)) = 0
                        End If

                        '--- ��������� �� ������� - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("I" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("I" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("I" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("I" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ �� 12 �� 16% ���. - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("J" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("J" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("J" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("J" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ �� 12 �� 16% ���. - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("K" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("K" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("K" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("K" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ �� 16 �� 22% ���. - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("L" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("L" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("L" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("L" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ �� 16 �� 22% ���. - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("M" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("M" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("M" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("M" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ �� 22 �� 27% ���. - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("N" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("N" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("N" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("N" + CStr(LineNumber)) = 0
                        End If
                        '--- ���������� �������� � ������ �� 22 �� 27% ���. - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("O" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("O" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("O" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("o" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������� � ������ ����� 27% - ���������� ��� ����������, ���.
                        If (appXLSRC.Worksheets(1).Range("P" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("P" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("P" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("P" + CStr(LineNumber)) = 0
                        End If
                        '--- ���������� �������� � ������ ����� 27% - �����, ���.
                        If (appXLSRC.Worksheets(1).Range("Q" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("Q" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("Q" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("Q" + CStr(LineNumber)) = 0
                        End If

                        '--- ����������� ���������� ����� ������
                        If (appXLSRC.Worksheets(1).Range("V" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("W" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("V" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("W" + CStr(LineNumber)) = 0
                        End If

                        '--- ����������� ����������
                        If (appXLSRC.Worksheets(1).Range("Z" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("X" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("Z" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("X" + CStr(LineNumber)) = 0
                        End If

                        '--- �������� �������������� �� ����� ��������
                        If (appXLSRC.Worksheets(1).Range("AA" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AB" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AA" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AA" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AB" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AA" + CStr(LineNumber)) = 0
                        End If

                        '--- ����� ������ ����� �������� � ������ ������������, ���.
                        If (appXLSRC.Worksheets(1).Range("AD" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AB" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AD" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AB" + CStr(LineNumber)) = 0
                        End If

                        '--- �������� �� ������� ���������� ������
                        If (appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AF" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AD" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AF" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AE" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) / 100
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AD" + CStr(LineNumber)) = 0
                        End If

                        '---�������� �� ������� ����� Web-����
                        If (appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) / 100
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = 0
                        End If


                        '--- ���� ������, ���.
                        If (appXLSRC.Worksheets(1).Range("S" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AI" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("S" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AI" + CStr(LineNumber)) = 0
                        End If

                        '--- ���� ������, ���.
                        If (appXLSRC.Worksheets(1).Range("T" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AJ" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("T" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AJ" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� ������������ ���������� ����
                        If (appXLSRC.Worksheets(1).Range("W" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AL" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("W" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AL" + CStr(LineNumber)) = 0
                        End If

                        '--- ���������� �������
                        If (appXLSRC.Worksheets(1).Range("X" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AM" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("X" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AM" + CStr(LineNumber)) = 0
                        End If

                        '--- ����� ��������
                        If (appXLSRC.Worksheets(1).Range("R" + CStr(cc.Row)).Value) <> Nothing Then
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AO" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("R" + CStr(cc.Row)).Value)
                        Else
                            appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AO" + CStr(LineNumber)) = 0
                        End If
                    End If
                    Salesman = appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("C" + CStr(LineNumber + 1)).Value
                Loop
                '---��������� �������� ��������---------------------------------------------

            Catch ex As Exception
            End Try
            appXLDST.ActiveWorkbook.Close(SaveChanges:=True)
            PBar = PBar + Counter
            ProgressBar2.Value = PBar
            Application.DoEvents()
        Next
        appXLSRC.ActiveWorkbook.Close(SaveChanges:=True)
        appXLSRC.Quit()
        appXLSRC = Nothing
        appXLDST.Application.displayalerts = False
        appXLDST.Quit()
        appXLDST = Nothing

    End Sub

    Private Sub TotalData(ByVal MyMonth As String, ByVal MyFile As String, ByVal MyCatalog As String)
        Dim appXLSRC As Object                             'Excel ������ - ����� �����
        Dim appXLDST As Object                             'Excel ������ - ������������� ����� �� ��������
        Dim Counter As Double
        Dim PBar As Double
        Dim Salesman As String
        Dim cc As Object
        Dim LineNumber As Integer
        Dim Count As Integer

        appXLSRC = CreateObject("Excel.Application")
        appXLSRC.Workbooks.Open(MyFile)
        appXLDST = CreateObject("Excel.Application")
        appXLDST.Application.displayalerts = False

        Counter = 100 / (My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls").Count)
        PBar = 0
        ProgressBar2.Value = 0
        Application.DoEvents()
        '---������ ������, �� ������� �������� ������-----------------------------------
        For Each foundFile As String In My.Computer.FileSystem.GetFiles(MyCatalog, FileIO.SearchOption.SearchAllSubDirectories, "*.xls")
            appXLDST.Workbooks.Open(foundFile)
            Try
                '---������ ��������� ��������-----------------------------------------------
                '---�������� ������ ����� "Total" �� "�������"
                appXLDST.Application.displayalerts = False
                appXLDST.Worksheets("Total").Delete()
                appXLDST.Worksheets("������").copy(appXLDST.sheets(1))
                appXLDST.Application.displayalerts = False
                appXLDST.Worksheets("������ (2)").name = "Total"
                appXLDST.Worksheets("Total").Range("A2") = Microsoft.VisualBasic.Year(Today())
                appXLDST.Worksheets("Total").Range("D2") = "� ������ ����"
                '---������
                Salesman = "S"
                LineNumber = 9
                Do While (Microsoft.VisualBasic.Left(Salesman, 1) = "S" Or (Microsoft.VisualBasic.Left(Salesman, 1) = "P")) '������� ���� ��������� �� ����� "Total" ������� � 10-�� ������
                    LineNumber = LineNumber + 1                     '������� �� ��������� ������
                    appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber) + ":Y" + CStr(LineNumber)).value = 0 '---������� ����� �� ��������� ������
                    Salesman = appXLDST.Worksheets("Total").Range("C" + CStr(LineNumber)).Value
                    Count = 1
                    Do While Count <= appXLDST.Worksheets.count         '������� ���� ������ � ������� ����� ��� ������������
                        If (appXLDST.Worksheets(Count).Name <> "Total") And (appXLDST.Worksheets(Count).Name <> "������") Then
                            cc = appXLDST.Worksheets(Count).Range("C3:C1000").Find(Salesman)
                            If Not cc Is Nothing Then
                                '--- ���� � ������
                                If (appXLDST.Worksheets(Count).Range("D" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("D" + CStr(LineNumber)) = CStr(appXLDST.Worksheets(Count).Range("D" + CStr(cc.Row)).Value)
                                End If

                                '--- % ������� � ������
                                If (appXLDST.Worksheets(Count).Range("E" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("E" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("F" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("F" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("F" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("F" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("G" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("G" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("G" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("G" + CStr(cc.Row)).Value)
                                End If

                                '--- ��������� �� ������� - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("H" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("H" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("H" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("H" + CStr(cc.Row)).Value)
                                End If

                                '--- ��������� �� ������� - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("I" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("I" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("I" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("I" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ �� 12 �� 16% ���. - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("J" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("J" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("J" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("J" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ �� 12 �� 16% ���. - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("K" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("K" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("K" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("K" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ �� 16 �� 22% ���. - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("L" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("L" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("L" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("L" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ �� 16 �� 22% ���. - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("M" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("M" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("M" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("M" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ �� 22 �� 27% ���. - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("N" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("N" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("N" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("N" + CStr(cc.Row)).Value)
                                End If
                                '--- ���������� �������� � ������ �� 22 �� 27% ���. - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("O" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("O" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("O" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("O" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ ����� 27% - ���������� ��� ����������, ���.
                                If (appXLDST.Worksheets(Count).Range("P" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("P" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("P" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("P" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������� � ������ ����� 27% - �����, ���.
                                If (appXLDST.Worksheets(Count).Range("Q" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("Q" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("Q" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("Q" + CStr(cc.Row)).Value)
                                End If

                                '--- ����������� ���������� ����� ������
                                If (appXLDST.Worksheets(Count).Range("W" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("W" + CStr(cc.Row)).Value)
                                End If

                                '--- ����������� ����������
                                If (appXLDST.Worksheets(Count).Range("X" + CStr(cc.Row)).Value) <> Nothing Then
                                    '--appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("X" + CStr(cc.Row)).Value)
                                End If

                                '--- �������� � ������ �������������
                                If (appXLDST.Worksheets(Count).Range("Y" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("Y" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("Y" + CStr(cc.Row)).Value)
                                End If

                                '--- �������� �������������� �� ����� ��������
                                If (appXLDST.Worksheets(Count).Range("AA" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AA" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AA" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AA" + CStr(cc.Row)).Value)
                                End If

                                '--- �������� �� ������ ����� ��������
                                If (appXLDST.Worksheets(Count).Range("AB" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AB" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AB" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AB" + CStr(cc.Row)).Value)
                                End If

                                '--- �������� �� ������� ���������� ������
                                If (appXLDST.Worksheets(Count).Range("AD" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AD" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AD" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AD" + CStr(cc.Row)).Value)
                                End If

                                '---�������� �� ������� ����� Web-����
                                If (appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) <> Nothing And (appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = CDbl(appXLSRC.Worksheets(1).Range("AG" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("E" + CStr(cc.Row)).Value) * CDbl(appXLSRC.Worksheets(1).Range("AH" + CStr(cc.Row)).Value) / 100
                                Else
                                    appXLDST.Worksheets(GetRowNumber(MyMonth)).Range("AE" + CStr(LineNumber)) = 0
                                End If

                                '--- ���� ������, ���.
                                If (appXLDST.Worksheets(Count).Range("AI" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AI" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AH" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AH" + CStr(cc.Row)).Value)
                                End If

                                '--- ���� ������, ���.
                                If (appXLDST.Worksheets(Count).Range("AJ" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AJ" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AI" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AI" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� ������������ ���������� ����
                                If (appXLDST.Worksheets(Count).Range("AL" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AL" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AK" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AK" + CStr(cc.Row)).Value)
                                End If

                                '--- ���������� �������
                                If (appXLDST.Worksheets(Count).Range("AM" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AM" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AL" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AL" + CStr(cc.Row)).Value)
                                End If

                                '--- ����� ��������
                                If (appXLDST.Worksheets(Count).Range("AO" + CStr(cc.Row)).Value) <> Nothing Then
                                    appXLDST.Worksheets("Total").Range("AO" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("AN" + CStr(LineNumber)).value) + CDbl(appXLDST.Worksheets(Count).Range("AN" + CStr(cc.Row)).Value)
                                End If

                            End If
                        End If
                        Count = Count + 1       '������� �� ��������� ����
                    Loop
                    '--- ���������� ������� ��������
                    'appXLDST.Worksheets("Total").Range("AO" + CStr(LineNumber)) = appXLDST.Worksheets.count()

                    '--- % ������� � ������
                    'appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("E" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)

                    '--- ����������� ���������� ����� ������
                    'appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("W" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)

                    '--- ����������� ���������� ��������
                    'appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)) = CDbl(appXLDST.Worksheets("Total").Range("X" + CStr(LineNumber)).value) / (appXLDST.Worksheets.count() - 2)


                    '--- ������� � ���������� ��������
                    Salesman = appXLDST.Worksheets("Total").Range("C" + CStr(LineNumber + 1)).Value
                Loop
            Catch ex As Exception
            End Try
            appXLDST.ActiveWorkbook.Close(SaveChanges:=True)

            PBar = PBar + Counter
            ProgressBar2.Value = PBar
            Application.DoEvents()
        Next
        appXLSRC.ActiveWorkbook.Close(SaveChanges:=True)
        appXLSRC.Quit()
        appXLSRC = Nothing
        appXLDST.Quit()
        appXLDST = Nothing

    End Sub

    Private Function GetRowNumber(ByVal MyStr) As String
        Select Case LTrim(RTrim(MyStr))
            Case "������"
                Return "01"
            Case "�������"
                Return "02"
            Case "����"
                Return "03"
            Case "������"
                Return "04"
            Case "���"
                Return "05"
            Case "����"
                Return "06"
            Case "����"
                Return "07"
            Case "������"
                Return "08"
            Case "��������"
                Return "09"
            Case "�������"
                Return "10"
            Case "������"
                Return "11"
            Case "�������"
                Return "12"
            Case Else
                Return "1012"
        End Select
    End Function

    Public Sub New()
        InitializeComponent()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TotalData(MyMonth.Text, MyFileR.Text, MyCatalogR.Text)
        MsgBox("�������� �������� �������", MsgBoxStyle.Information, "��������!")
    End Sub

End Class
