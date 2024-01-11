Imports System.Xml

Public Class Main

    Private Sub Main_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� ��������� �������� ��� ������ �� ����������
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            appXLSRC.DisplayAlerts = 0
            appXLSRC.Workbooks.Close()
            appXLSRC.DisplayAlerts = 1
            appXLSRC.Quit()
            appXLSRC = Nothing
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Main_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������ �������� ���� �� ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub Main_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��� ������� ���������� ��������� - ���, ��������, ������������ � �.�.
        '// ����� ���� ������� ������ ����������� ������� ������������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        '������� ������
        Dim MyDs As New DataSet                       '

        '---��������� �������
        Try
            Dim Scala As New SfwIII.Application

            declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            declarations.ScalaDate = CDate(Scala.ActiveFrame.Parent.ScalaDate)


            MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
            MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
            MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & declarations.UserCode & "')) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If declarations.MyRec.EOF = True And declarations.MyRec.BOF = True Then
                MsgBox("�� ������ ��� ��������, ��������������� ������ �� ���� � Scala. ���������� � ��������������.", MsgBoxStyle.Critical, "��������!")
                trycloseMyRec()
                Application.Exit()
            Else
                declarations.SalesmanCode = declarations.MyRec.Fields("SC").Value
                declarations.SalesmanName = declarations.MyRec.Fields("FullName").Value
                trycloseMyRec()
            End If
        Catch
            MsgBox("��������� ������ ����������� ������ �� ���� Scala", MsgBoxStyle.Critical, "��������!")
            Application.Exit()
        End Try

        '---��������
        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        CheckButtonState()
    End Sub

    Private Sub CheckButtonState()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� � ����������� ��������� ������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If ComboBox1.SelectedItem = "" Then
            button2.Enabled = False
        Else
            button2.Enabled = True
        End If
        button3.Enabled = False
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����� ������ ����������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        CheckButtonState()
    End Sub

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �����
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Application.Exit()
    End Sub

    Private Sub button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� Invoice - �����, ����������� �������� ���������� ��
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        TextBox1.Text = ""
        textBox3.Text = ""
        textBox4.Text = ""
        textBox5.Text = ""
        label6.Text = ""
        button3.Enabled = False
        progressBar1.Value = 0
        OpenInvoiceFile()
    End Sub

    Private Sub button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button3.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� Invoice - ����� � Scala
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckDate() = False Then  '---���� � Scala �� ��������� � ������� ����� �� ����������
            MsgBox("��������� ���� � Scala �� ��������� � ������� ����� �� ����������. ��������� � Scala ������� ���� � ������ ����� ����� ����������� ������.", MsgBoxStyle.Critical, "��������!")
            Exit Sub
        End If

        UploadInvoiceFile()
    End Sub

    Private Function CheckDate() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� ���� � Scala - ��������� �� � ������������
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If Math.Abs(DateDiff(DateInterval.Day, declarations.ScalaDate, Now())) >= 1 Then
            CheckDate = False
        Else
            CheckDate = True
        End If
    End Function
End Class
