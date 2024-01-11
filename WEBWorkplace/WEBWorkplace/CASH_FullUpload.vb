Imports System.IO
Public Class CASH_FullUpload

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� �������� ��� �������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCatalog As String

        MyCatalog = GetFolderPath()
        If MyCatalog = "" Then      '--������ ������
        Else
            TextBox1.Text = MyCatalog
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CASH_FullUpload_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = My.Settings.CASHCatalog
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������� �������� ��������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) <> "" Then
            If Directory.Exists(Trim(TextBox1.Text)) = False Then
                MsgBox("��������� ������� �� ����������. ������� ���������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            Me.Cursor = Cursors.WaitCursor
            FullUpload()
            Me.Cursor = Cursors.Default
            MsgBox("�������� ������ ���������.", MsgBoxStyle.Information, "��������!")
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" Then
            MsgBox("������� ��� �������� ������ ���� ��������. ������� ���������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
            TextBox1.Select()
            CheckData = False
            Exit Function
        End If

        If Directory.Exists(Trim(TextBox1.Text)) = False Then
            MsgBox("��������� ������� �� ����������. ������� ���������� ��� ��������.", MsgBoxStyle.Critical, "��������!")
            TextBox1.Select()
            CheckData = False
            Exit Function
        End If
        CheckData = True
    End Function

    Private Sub FullUpload()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� ������ �������� ������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyFile As String
        Dim MyFlagFile As String
        Dim MyWrkStr As String

        MyFile = Trim(TextBox1.Text) + "\" + "goods.txt"
        MyFlagFile = Trim(TextBox1.Text) + "\" + "goods_flag.txt"

        '-----------������ ������ �� Scala
        MySQLStr = "exec spp_CASH_Items_FromScala 1 "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----------������� ��������
        If File.Exists(MyFile) = True Then
            Try
                File.Delete(MyFile)
            Catch ex As Exception
                MsgBox("���������� �������� ��������� �������. ���������� ����� ��� ���������� � ��������������. " + ex.Message, MsgBoxStyle.Critical, "��������!")
                Exit Sub
            End Try
        End If
        If File.Exists(MyFlagFile) = True Then
            Try
                File.Delete(MyFlagFile)
            Catch ex As Exception
                MsgBox("���������� �������� ��������� �������. ���������� ����� ��� ���������� � ��������������. " + ex.Message, MsgBoxStyle.Critical, "��������!")
                Exit Sub
            End Try
        End If

        '----------�������� ����� � ����������
        Dim f As New StreamWriter(MyFile, False, System.Text.Encoding.GetEncoding(1251))
        '---��������� ������
        MyWrkStr = "$$$ADDTAXRATES" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;��� 0%;;;0" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "2;��� 10%;;;10" + vbCrLf
        f.Write(MyWrkStr)
        If Now < CDate("01/01/2019") Then
            MyWrkStr = "3;��� 18%;;;18" + vbCrLf
        Else
            MyWrkStr = "3;��� 20%;;;20" + vbCrLf
        End If
        f.Write(MyWrkStr)
        MyWrkStr = "4;��� ���;;;0" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "5;��������� ������ 10/110;;;10" + vbCrLf
        f.Write(MyWrkStr)
        If Now < CDate("01/01/2019") Then
            MyWrkStr = "6;��������� ������ 18/118;;;18" + vbCrLf
        Else
            MyWrkStr = "6;��������� ������ 20/120;;;20" + vbCrLf
        End If
        f.Write(MyWrkStr)
        '---��������� ������
        MyWrkStr = "$$$ADDTAXGROUPRATES" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;2;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;3;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;4;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;5;3" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "1;10;1" + vbCrLf
        f.Write(MyWrkStr)
        '---�������� ���� �������
        MyWrkStr = "$$$DELETEALLWARES" + vbCrLf
        f.Write(MyWrkStr)
        '---�������� ���� �������
        MyWrkStr = "$$$ADDQUANTITY" + vbCrLf
        f.Write(MyWrkStr)

        'MyWrkStr = "1;;�������������������;�������������������;;;;;;;;;;;;;0;;;;;;;;;;;;;;;;;" + vbCrLf
        'f.Write(MyWrkStr)

        MyWrkStr = "00000002;;;00000002 �������������������;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000002;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000003;;;00000003 ������;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000003;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000004;;;00000004 �������� �������� �.�;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000004;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000006;;;00000006 ��� (��, ��������);0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000006;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)
        MyWrkStr = "00000007;;;00000007 ������ ������;0.00000000;;;0;0;;;;;;;;1;;;;;;2;;;00000007;;;;;;;;" + vbCrLf
        f.Write(MyWrkStr)

        MySQLStr = "exec spp_CASH_Items_FromDB 1, 1, 1, 0 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            MySQLStr = Declarations.MyRec.GetString(, , ";", vbCrLf)
            f.Write(MySQLStr)
        End If
        trycloseMyRec()
        '---�������� �������� ������������


        '---�������� �����
        f.Close()
        '---�������� ����� - �����
        Dim f1 As New StreamWriter(MyFlagFile, False, System.Text.Encoding.GetEncoding(1251))
        f1.Close()
        '---���������� ������ � ��
        MySQLStr = "DELETE FROM tbl_CASH_Items "
        MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2)"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "UPDATE tbl_CASH_Items "
        MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0, ScalaStatus = 0 "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            Me.Cursor = Cursors.WaitCursor
            ChangeUpload()
            Me.Cursor = Cursors.Default
            MsgBox("�������� ������ ���������.", MsgBoxStyle.Information, "��������!")
        End If
    End Sub

    Private Sub ChangeUpload()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� �������� ��������� ������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyFile As String
        Dim MyFlagFile As String
        Dim MyWrkStr As String

        MyFile = Trim(TextBox1.Text) + "\" + "goods.txt"
        MyFlagFile = Trim(TextBox1.Text) + "\" + "goods_flag.txt"

        '-----------������ ������ �� Scala
        MySQLStr = "exec spp_CASH_Items_FromScala 1 "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----------������� ��������
        If File.Exists(MyFile) = True Then
            Try
                File.Delete(MyFile)
            Catch ex As Exception
                MsgBox("���������� �������� ��������� �������. ���������� ����� ��� ���������� � ��������������. " + ex.Message, MsgBoxStyle.Critical, "��������!")
                Exit Sub
            End Try
        End If
        If File.Exists(MyFlagFile) = True Then
            Try
                File.Delete(MyFlagFile)
            Catch ex As Exception
                MsgBox("���������� �������� ��������� �������. ���������� ����� ��� ���������� � ��������������. " + ex.Message, MsgBoxStyle.Critical, "��������!")
                Exit Sub
            End Try
        End If

        '----------�������� ����� � ����������
        Dim f As New StreamWriter(MyFile, False, System.Text.Encoding.GetEncoding(1251))
        '---�������� �������------------
        MyWrkStr = "$$$DELETEWARESBYWARECODE" + vbCrLf
        f.Write(MyWrkStr)
        '---���������
        MySQLStr = "exec spp_CASH_Items_FromDB 0, 1, 2, 1 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            MySQLStr = Declarations.MyRec.GetString(, , ";", vbCrLf)
            f.Write(MySQLStr)
        End If
        trycloseMyRec()
        '---����������
        MySQLStr = "exec spp_CASH_Items_FromDB 0, 1, 3, 1 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            MySQLStr = Declarations.MyRec.GetString(, , ";", vbCrLf)
            f.Write(MySQLStr)
        End If
        trycloseMyRec()

        '---�������� �������------------
        MyWrkStr = "$$$ADDQUANTITY" + vbCrLf
        f.Write(MyWrkStr)

        '---�����
        MySQLStr = "exec spp_CASH_Items_FromDB 0, 1, 1, 0 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            MySQLStr = Declarations.MyRec.GetString(, , ";", vbCrLf)
            f.Write(MySQLStr)
        End If
        trycloseMyRec()
        '---����������
        MySQLStr = "exec spp_CASH_Items_FromDB 0, 1, 3, 0 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            MySQLStr = Declarations.MyRec.GetString(, , ";", vbCrLf)
            f.Write(MySQLStr)
        End If
        trycloseMyRec()

        '---�������� �����
        f.Close()
        '---�������� ����� - �����
        Dim f1 As New StreamWriter(MyFlagFile, False, System.Text.Encoding.GetEncoding(1251))
        f1.Close()
        '---���������� ������ � ��
        MySQLStr = "DELETE FROM tbl_CASH_Items "
        MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2)"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        MySQLStr = "UPDATE tbl_CASH_Items "
        MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0, ScalaStatus = 0 "
        MySQLStr = MySQLStr & "WHERE(ScalaStatus <> 0) "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class