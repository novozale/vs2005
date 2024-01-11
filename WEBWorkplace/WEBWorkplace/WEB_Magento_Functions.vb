Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Data.SqlClient

Module WEB_Magento_Functions
    Public Sub UploadInfo_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� ���������� �� ���� Magento
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MyGlobalStr = ""
        '-----------------���������� ���������� �� Scala--------------------------------------
        MySQLStr = "exec spp_WEB_ALL_FromScala "
        InitMyConn(False)
        'Declarations.MyConn.Execute(MySQLStr)

        '-----------------�������� ��������� (����� � �������� �������)-----------------------
        UploadInfo_Categories_ToMagento(MyUploadType)

        '-----------------�������� ����� ��������� "�������������"----------------------------
        UploadInfo_ManufacturerList_ToMagento(MyUploadType)

        '-----------------�������� ����� ��������� "������ �������������"---------------------
        UploadInfo_CountriesList_ToMagento(MyUploadType)

        '-----------------�������� �������----------------------------------------------------
        UploadInfo_Products_ToMagento(MyUploadType)

    End Sub

    Public Sub UploadInfo_Categories_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �� ���������� �� ����� Magento
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------�������� �������������� ���������-------------------------------------
            Delete_NotUsedCategories_FromMagento()
        End If

        '---------------�������� ��������� ---------------------------------------------------
        Upload_Categories_ToMagento(MyUploadType)

    End Sub

    Public Sub Delete_NotUsedCategories_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������������� ��������� (������������) �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Categories_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ���������� �� ���� Magento
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        If MyUploadType = 0 Then        '------������ ��������
            '-------------�������� ������� � ����������, ���������� �� ��������, �� �� ����������� �� Magento--------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 2"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------�������� ������� � �������, ���������� �� ��������, �� �� ����������� �� Magento-----------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 1"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------�������� ��������� 2 ������ (��������� ���������) ���������� �� ��������-------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 1, 2"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    Delete_Category_FromMagento(Declarations.MyRec.Fields("MagentoCode").Value, 2)
                    Declarations.MyRec.MoveNext()
                End While
            End If
            trycloseMyRec()

            '-------------�������� ��������� 1 ������ (������ ���������) ���������� �� ��������----------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 1, 1"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    Delete_Category_FromMagento(Declarations.MyRec.Fields("MagentoCode").Value, 1)
                    Declarations.MyRec.MoveNext()
                End While
            End If
            trycloseMyRec()

            '-------------�������� ��������� 1 ������ (������ ���������)----------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 3, 1"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label3.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Category_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MagentoCode").Value, Declarations.MyRec.Fields("MyJSON").Value, 1)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label2.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox1.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

            '-------------�������� ��������� 2 ������ (��������� ���������)-------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 3, 2"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label14.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Category_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MagentoCode").Value, Declarations.MyRec.Fields("MyJSON").Value, 2)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label15.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox5.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

        Else                            '------�������� ������ ���������� ������
            '-------------�������� ������� � ����������, ���������� �� ��������, �� �� ����������� �� Magento--------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 2"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------�������� ������� � �������, ���������� �� ��������, �� �� ����������� �� Magento-----------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 1"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------�������� ��������� 2 ������ (��������� ���������) ���������� �� ��������-------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 1, 2"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    Delete_Category_FromMagento(Declarations.MyRec.Fields("MagentoCode").Value, 2)
                    Declarations.MyRec.MoveNext()
                End While
            End If
            trycloseMyRec()
            '-------------�������� ��������� 1 ������ (������ ���������) ���������� �� ��������----------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 1, 1"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF = False
                    Delete_Category_FromMagento(Declarations.MyRec.Fields("MagentoCode").Value, 1)
                    Declarations.MyRec.MoveNext()
                End While
            End If
            trycloseMyRec()
            '-------------�������� ��������� 1 ������ (������ ���������)----------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 2, 1"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label3.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Category_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MagentoCode").Value, Declarations.MyRec.Fields("MyJSON").Value, 1)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label2.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox1.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()
            '-------------�������� ��������� 2 ������ (��������� ���������)-------------------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 2, 2"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label14.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Category_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MagentoCode").Value, Declarations.MyRec.Fields("MyJSON").Value, 2)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label15.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox5.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

        End If
    End Sub

    Public Sub Delete_Category_FromMagento(ByVal MyCategoryID As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ��������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRep As String
        Dim MyWR As HttpWebResponse

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/categories/" & Trim(MyCategoryID))
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"

            Try
                MyRep = MC.UploadString(MyUrl, "DELETE", "")
                Delete_Category_FromMagento_Confirm(MyCategoryID, MyLevel)
            Catch ex As WebException
                MyWR = ex.Response
                If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                    Delete_Category_FromMagento_Confirm(MyCategoryID, MyLevel)
                Else
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("�������� ��������� " + MyCategoryID + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "�������� ��������� " + MyCategoryID + " ---> " + ex.Message + Chr(13) + Chr(10)
                End If
            End Try
        End Using
    End Sub

    Public Sub CreateUpdate_Category_InMagento(ByVal MyCode As String, ByVal MyCategoryID As String, ByVal MyCategoryJSon As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (����������) ��������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyWR As HttpWebResponse
        Dim MyStr As String
        Dim MyRep As String
        Dim NewCategoryID As String

        If Trim(MyCategoryID) = "" Then     '---��������
            Using MC As New WebClient
                Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/categories/")
                MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                MC.Headers(HttpRequestHeader.ContentType) = "application/json"
                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                Try
                    MyRep = MC.UploadString(MyUrl, "POST", MyStr)
                    '---��������� ���������� � �������� ���������
                    NewCategoryID = GetCategoryID(MyRep)
                    If Trim(NewCategoryID) <> "" Then
                        CreateUpdate_Category_InMagento_Confirm(MyCode, Trim(NewCategoryID), MyLevel)
                    End If
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("�������� ��������� " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "�������� ��������� " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                End Try
            End Using
        Else
            '---����������
            Using MC As New WebClient
                Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/categories/" & Trim(MyCategoryID))
                MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                MC.Headers(HttpRequestHeader.ContentType) = "application/json"
                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                Try
                    MyRep = MC.UploadString(MyUrl, "PUT", MyStr)
                    '---��������� ���������� �� ���������� ���������
                    CreateUpdate_Category_InMagento_Confirm(MyCode, MyCategoryID, MyLevel)
                Catch ex As WebException
                    Try
                        MyWR = ex.Response
                        If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                            '---���� ��������� �� ������� � �������� �� ������� - �� ������� ������.
                            Using MC1 As New WebClient
                                Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/categories/" & Trim(MyCategoryID))
                                MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                                MC1.Headers(HttpRequestHeader.ContentType) = "application/json"
                                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                                Try
                                    MyRep = MC1.UploadString(MyUrl1, "POST", MyStr)
                                    '---��������� ���������� � �������� ���������
                                    NewCategoryID = GetCategoryID(MyRep)
                                    If Trim(NewCategoryID) <> "" Then
                                        CreateUpdate_Category_InMagento_Confirm(MyCode, Trim(NewCategoryID), MyLevel)
                                    End If
                                Catch ex1 As WebException
                                    If My.Settings.MyDebug = "YES" Then
                                        MsgBox("�������� ��������� ��� ���������� " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "��������!")
                                    End If
                                    MyGlobalStr = MyGlobalStr + "�������� ��������� ��� ���������� " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                                End Try
                            End Using
                        Else
                            If My.Settings.MyDebug = "YES" Then
                                MsgBox("���������� ��������� " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                            End If
                            MyGlobalStr = MyGlobalStr + "���������� ��������� " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                        End If
                    Catch ex2 As Exception
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("��������� ������ 1 ---> " + ex2.Message, MsgBoxStyle.Information, "��������!")
                        End If
                        MyGlobalStr = MyGlobalStr + "��������� ������ 1 ---> " + ex2.Message + Chr(13) + Chr(10)
                    End Try
                End Try
            End Using
        End If
    End Sub

    Public Sub Delete_Category_FromMagento_Confirm(ByVal MyCategoryID As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyLevel = 1 Then     '-----������ ���������
            '---������� ������
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemGroup_MG_Correspondence ON tbl_WEB_ItemGroup.Code = tbl_WEB_ItemGroup_MG_Correspondence.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_ItemGroup_MG_Correspondence.MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        Else                    '-----��������� ���������
            '---������� ���������
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup_MG_Correspondence ON tbl_WEB_ItemSubGroup.SubgroupID = tbl_WEB_ItemSubGroup_MG_Correspondence.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_ItemSubGroup_MG_Correspondence.MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup_MG_Correspondence"
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
    End Sub

    Public Sub CreateUpdate_Category_InMagento_Confirm(ByVal MyCode As String, ByVal MyCategoryID As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� (����������) ��������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyLevel = 1 Then     '-----������ ���������
            '---������� ������ ��������
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') OR "
            MySQLStr = MySQLStr & "(MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������ ��������
            MySQLStr = "INSERT INTO tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "(Code, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyCode) & "', N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������ ����������
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0 "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        Else                    '-----��������� ���������
            '---������� ������ ��������
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') OR "
            MySQLStr = MySQLStr & "(MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������ ��������
            MySQLStr = "INSERT INTO tbl_WEB_ItemSubGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "(Code, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyCode) & "', N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ��������� ����������
            MySQLStr = "UPDATE tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0 "
            MySQLStr = MySQLStr & "WHERE (SubgroupID = N'" & Trim(MyCode) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
    End Sub

    Public Function GetCategoryID(ByVal MyJSON As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ID �� ������ �������� JSON 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Try
            Dim json As JObject = JObject.Parse(MyJSON)
            GetCategoryID = json.GetValue("id").ToString()
        Catch ex As Exception
            GetCategoryID = ""
        End Try
    End Function

    Public Sub UploadInfo_ManufacturerList_ToMagento(ByVal MyUploadType As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �������������� � ����� ��������� "proizvoditel" 
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------�������� �������������� �������������� �� ����� ��������� "proizvoditel"-
            Delete_NotUsedManufacturers_FromMagento()
        End If

        '---------------�������� ������ �������������� � ����� ��������� "proizvoditel"--
        Upload_Manufacturers_ToMagento()
    End Sub

    Public Sub Delete_NotUsedManufacturers_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������������� �������������� �� ����� ��������� "proizvoditel"
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Manufacturers_ToMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �������������� � ����� ��������� "proizvoditel"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        '----�������� ��������������, ���������� �� ��������, �� �� ����������� � Magento-
        MySQLStr = "exec spp_WEB_Manufacturers_MG_FromDB 0"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----�������� ��������������, ���������� �� ��������----------------------------
        MySQLStr = "exec spp_WEB_Manufacturers_MG_FromDB 1"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                Delete_Manufacturer_FromMagento(Declarations.MyRec.Fields("MagentoCode").Value)
                Declarations.MyRec.MoveNext()
            End While
        End If
        trycloseMyRec()

        '----------------�������� ��������������----------------------------------------
        MySQLStr = "exec spp_WEB_Manufacturers_MG_FromDB 2"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            Declarations.MyRec.MoveLast()
            MyUploadDataToMagento.Label5.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                Create_Manufacturer_InMagento(Declarations.MyRec.Fields("Name").Value, Declarations.MyRec.Fields("MyJSON").Value)
                Declarations.MyRec.MoveNext()
                i = i + 1
                MyUploadDataToMagento.Label6.Text = i
                Application.DoEvents()
            End While
        End If
        MyUploadDataToMagento.GroupBox2.BackColor = Color.LightGreen
        Application.DoEvents()
        trycloseMyRec()
    End Sub

    Public Sub Delete_Manufacturer_FromMagento(ByVal MyManufacturerID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRep As String
        Dim MyWR As HttpWebResponse

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/attributes/proizvoditel/options/" & Trim(MyManufacturerID))
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"

            Try
                MyRep = MC.UploadString(MyUrl, "DELETE", "")
                Delete_Manufacturer_FromMagento_Confirm(MyManufacturerID)
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        Delete_Manufacturer_FromMagento_Confirm(MyManufacturerID)
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("�������� ������������� " + MyManufacturerID + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                        End If
                        MyGlobalStr = MyGlobalStr + "�������� ������������� " + MyManufacturerID + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("��������� ������ 2 ---> " + ex2.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "�������� ������������� " + MyManufacturerID + " ---> " + ex.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub


    Public Sub Create_Manufacturer_InMagento(ByVal MyName As String, ByVal MyManufacturerJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyStr As String
        Dim MyRep As String

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/attributes/proizvoditel/options/")
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyManufacturerJSon)))

            Try
                MyRep = MC.UploadString(MyUrl, "POST", MyStr)
                '---��������� ���������� � �������� �������������
                Create_Manufacturer_InMagento_Confirm(MyName)
            Catch ex As WebException
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("�������� ������������� " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                End If
                MyGlobalStr = MyGlobalStr + "�������� ������������� " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub Delete_Manufacturer_FromMagento_Confirm(ByVal MyManufacturerID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� ������������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---������� ������
        MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Manufacturers INNER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers_MG_Correspondence ON CASE WHEN Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) "
        MySQLStr = MySQLStr & "= '' THEN tbl_WEB_Manufacturers.Name ELSE Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) "
        MySQLStr = MySQLStr & "END = tbl_WEB_Manufacturers_MG_Correspondence.ManufacturerName "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Manufacturers_MG_Correspondence.MagentoCode = N'" & Trim(MyManufacturerID) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---������� ������
        MySQLStr = "DELETE FROM tbl_WEB_Manufacturers_MG_Correspondence "
        MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyManufacturerID) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub Create_Manufacturer_InMagento_Confirm(ByVal MyName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� ������������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyManufacturerID As String
        Dim MySQLStr As String

        MyManufacturerID = ""
        MyManufacturerID = GetManufacturerIDByName(MyName)
        If Trim(MyManufacturerID) <> "" Then
            '---������� ������ ��������
            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyManufacturerID) & "') "
            MySQLStr = MySQLStr & "OR (ManufacturerName = N'" & Trim(MyName) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������ ��������
            MySQLStr = "INSERT INTO tbl_WEB_Manufacturers_MG_Correspondence "
            MySQLStr = MySQLStr & "(ManufacturerName, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyName) & "'"
            MySQLStr = MySQLStr & ", N'" & Trim(MyManufacturerID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� �������������� ����������
            MySQLStr = "UPDATE tbl_WEB_Manufacturers "
            MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0 "
            MySQLStr = MySQLStr & "WHERE (CASE WHEN Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) = '' "
            MySQLStr = MySQLStr & "THEN tbl_WEB_Manufacturers.Name ELSE Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) END = N'" & Trim(MyName) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Public Function GetManufacturerIDByName(ByVal MyName As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ID ������������� �� ����� � ����� Magento
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRep As String

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/attributes/proizvoditel/options")
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + "g7rvo6kkef82uwv5isvwokk3n2ohicyh"
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"

            GetManufacturerIDByName = ""

            Try
                MyRep = MC.DownloadString(MyUrl)
                MyRep = "{""data"":" + MyRep + "}"
                Dim json As JObject = JObject.Parse(MyRep)
                Dim jarray As JArray = json.SelectToken("data")
                For i As Integer = 0 To jarray.Count - 1
                    If Trim(jarray(i).Item("label").ToString) = Trim(MyName) Then
                        GetManufacturerIDByName = Trim(jarray(i).Item("value").ToString)
                        Exit For
                    End If
                Next i
            Catch ex As WebException
                GetManufacturerIDByName = ""
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("��������� id ������������� " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                End If
                MyGlobalStr = MyGlobalStr + "��������� id ������������� " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Function

    Public Sub UploadInfo_CountriesList_ToMagento(ByVal MyUploadType As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����� � ����� ��������� "strana_proishozhdenija" 
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------�������� �������������� �������������� �� ����� ��������� "strana_proishozhdenija"-
            Delete_NotUsedCountries_FromMagento()
        End If

        '---------------�������� ������ �������������� � ����� ��������� "strana_proishozhdenija"--
        Upload_Countries_ToMagento()
    End Sub

    Public Sub Delete_NotUsedCountries_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� �������������� ����� �� ����� ��������� "strana_proishozhdenija"
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Countries_ToMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ ����� � ����� ��������� "strana_proishozhdenija"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        '----------------�������� �����--------------------------------------------------
        MySQLStr = "exec spp_WEB_Countries_MG_FromDB "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            Declarations.MyRec.MoveLast()
            MyUploadDataToMagento.Label8.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                Create_Country_InMagento(Declarations.MyRec.Fields("Name").Value, Declarations.MyRec.Fields("MyJSON").Value)
                Declarations.MyRec.MoveNext()
                i = i + 1
                MyUploadDataToMagento.Label9.Text = i
                Application.DoEvents()
            End While
        End If
        MyUploadDataToMagento.GroupBox3.BackColor = Color.LightGreen
        Application.DoEvents()
        trycloseMyRec()
    End Sub

    Public Sub Create_Country_InMagento(ByVal MyName As String, ByVal MyManufacturerJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ������ �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyStr As String
        Dim MyRep As String

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/attributes/strana_proishozhdenija/options/")
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyManufacturerJSon)))

            Try
                MyRep = MC.UploadString(MyUrl, "POST", MyStr)
                '---��������� ���������� � �������� ������
                Create_Country_InMagento_Confirm(MyName)
            Catch ex As WebException
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("�������� ������ " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                End If
                MyGlobalStr = MyGlobalStr + "�������� ������ " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub Create_Country_InMagento_Confirm(ByVal MyName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCountryID As String
        Dim MySQLStr As String

        MyCountryID = ""
        MyCountryID = GetCountryIDByName(MyName)
        If Trim(MyCountryID) <> "" Then
            '---������� ������ ��������
            MySQLStr = "DELETE FROM tbl_WEB_Countries_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCountryID) & "') "
            MySQLStr = MySQLStr & "OR (CountryName = N'" & Replace(Trim(MyName), "'", "''") & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---������� ������ ��������
            MySQLStr = "INSERT INTO tbl_WEB_Countries_MG_Correspondence "
            MySQLStr = MySQLStr & "(CountryName, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Replace(Trim(MyName), "'", "''") & "'"
            MySQLStr = MySQLStr & ", N'" & Trim(MyCountryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Public Function GetCountryIDByName(ByVal MyName As String) As String
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� ID ������ �� ����� � ����� Magento
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRep As String

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/attributes/strana_proishozhdenija/options")
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + "g7rvo6kkef82uwv5isvwokk3n2ohicyh"
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"

            GetCountryIDByName = ""

            Try
                MyRep = MC.DownloadString(MyUrl)
                MyRep = "{""data"":" + MyRep + "}"
                Dim json As JObject = JObject.Parse(MyRep)
                Dim jarray As JArray = json.SelectToken("data")
                For i As Integer = 0 To jarray.Count - 1
                    If Trim(jarray(i).Item("label").ToString) = Trim(MyName) Then
                        GetCountryIDByName = Trim(jarray(i).Item("value").ToString)
                        Exit For
                    End If
                Next i
            Catch ex As WebException
                GetCountryIDByName = ""
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("��������� ID ������ " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                End If
                MyGlobalStr = MyGlobalStr + "��������� ID ������ " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Function

    Public Sub UploadInfo_Products_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ���������� �� ������� �� ����� Magento
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------����������� ����������� ��� �������������� ���������------------------
            Hide_NotUsedProducts_InMagento()
        End If

        '---------------�������� ��������� ---------------------------------------------------
        Upload_Products_ToMagento(MyUploadType)

    End Sub

    Public Sub Hide_NotUsedProducts_InMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ������� � ��������� �������������� ��������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Products_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ���������� �� ��������� �� ���� Magento
        '// MyUploadType =  0 ������ ��������
        '//                 1 ������ ����� ����������
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        If MyUploadType = 0 Then        '------������ ��������
            '-------------������� � ��������� �������, ���������� �� ��������-----------------
            MySQLStr = "exec spp_WEB_Items_MG_FromDB 0 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label11.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    Hide_Product_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MyJSON").Value)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label12.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox4.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()
            '--------------�������� �������---------------------------------------------------
            MySQLStr = "exec spp_WEB_Items_MG_FromDB 2 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label17.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Product_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MyJSON").Value)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label18.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox6.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

        Else                            '------�������� ������ ���������� ������
            '-------------������� � ��������� �������, ���������� �� ��������-----------------
            MySQLStr = "exec spp_WEB_Items_MG_FromDB 0 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label11.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    Hide_Product_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MyJSON").Value)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label12.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox4.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

            '--------------�������� �������---------------------------------------------------
            MySQLStr = "exec spp_WEB_Items_MG_FromDB 1 "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Else
                Declarations.MyRec.MoveLast()
                MyUploadDataToMagento.Label17.Text = Declarations.MyRec.RecordCount
                Declarations.MyRec.MoveFirst()
                i = 0
                While Declarations.MyRec.EOF = False
                    CreateUpdate_Product_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MyJSON").Value)
                    Declarations.MyRec.MoveNext()
                    i = i + 1
                    MyUploadDataToMagento.Label18.Text = i
                    Application.DoEvents()
                End While
            End If
            MyUploadDataToMagento.GroupBox6.BackColor = Color.LightGreen
            Application.DoEvents()
            trycloseMyRec()

        End If
    End Sub

    Public Sub Hide_Product_InMagento(ByVal MyCode As String, ByVal MyJSON As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ����������� �������� "���������" ��� ���������� � Scala ������
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyStr As String
        Dim MyRep As String
        Dim MyWR As HttpWebResponse

        Using MC As New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode))
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyJSON)))

            Try
                MyRep = MC.UploadString(MyUrl, "PUT", MyStr)
                '---��������� ���������� �� �������� ������
                Hide_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        Hide_Product_InMagento_Confirm(MyCode)
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("�������� (�����������) ������ " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                        End If
                        MyGlobalStr = MyGlobalStr + "�������� (�����������) ������ " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("��������� ������ 3 ---> " + ex2.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "��������� ������ 3 ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub Hide_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� ������ 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_WEB_Items "
        MySQLStr = MySQLStr & "WHERE (ScalaStatus = 2) "
        MySQLStr = MySQLStr & "AND (Code = N'" & Trim(MyCode) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub CreateUpdate_Product_InMagento(ByVal MyCode As String, ByVal MyJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (����������) �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---------------�������� (����������) ��������-----------------------------------
        CreateUpdate_ProductCore_InMagento(MyCode, MyJSon)

        '---------------�������� (����������) �������� ��������--------------------------
        CreateUpdate_Picture_InMagento(MyCode)

        '---------------�������� (����������) ������ ��������----------------------------
        Update_ExtendedPrice_InMagento(MyCode)

    End Sub

    Public Sub CreateUpdate_ProductCore_InMagento(ByVal MyCode As String, ByVal MyJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (����������) ������ � �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyWR As HttpWebResponse
        Dim MyStr As String
        Dim MyRep As String

        Using MC = New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode))
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyJSon)))

            Try
                MyRep = MC.UploadString(MyUrl, "PUT", MyStr)
                '---��������� ���������� �� ���������� ��������
                'CreateUpdate_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        '----��� ������� �� ������ - �� �������
                        Using MC1 = New WebClient
                            Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/products/")
                            MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                            MC1.Headers(HttpRequestHeader.ContentType) = "application/json"
                            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyJSon)))

                            Try
                                MyRep = MC1.UploadString(MyUrl1, "POST", MyStr)
                                '---��������� ���������� � �������� ��������
                                'CreateUpdate_Product_InMagento_Confirm(MyCode)
                            Catch ex1 As WebException
                                If My.Settings.MyDebug = "YES" Then
                                    MsgBox("�������� ������ " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "��������!")
                                End If
                                MyGlobalStr = MyGlobalStr + "�������� ������ " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                            End Try
                        End Using
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("���������� ������ " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                        End If
                        MyGlobalStr = MyGlobalStr + "���������� ������ " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("�������� ������ 1 " + MyCode + " ---> " + ex2.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "�������� ������ 1 " + MyCode + " ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub CreateUpdate_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� �������� (����������) �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_Items "
        MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0, ScalaStatus = 0 "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub CreateUpdate_Picture_InMagento(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (����������) ������ � �������� �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyWR As HttpWebResponse
        Dim MyRep As String
        Dim MyErrFlag As Integer

        Using MC = New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode) + "/media")
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"


            Try
                '-------��������� ������ �������� ��� ������� ��������
                MyRep = MC.DownloadString(MyUrl)
                MyRep = "{""data"":" + MyRep + "}"
                Dim json As JObject = JObject.Parse(MyRep)
                Dim jarray As JArray = json.SelectToken("data")
                MyErrFlag = 0
                For i As Integer = 0 To jarray.Count - 1
                    If Trim(jarray(i).Item("id").ToString) <> "" Then
                        '--------�������� ������ ��������
                        Using MC1 = New WebClient
                            Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode) + "/media/" + Trim(jarray(i).Item("id").ToString))
                            MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                            MC1.Headers(HttpRequestHeader.ContentType) = "application/json"

                            Try
                                MyRep = MC1.UploadString(MyUrl1, "DELETE", "")
                            Catch ex1 As WebException
                                MyErrFlag = 1
                                If My.Settings.MyDebug = "YES" Then
                                    MsgBox("�������� �������� � ������ " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "��������!")
                                End If
                                MyGlobalStr = MyGlobalStr + "�������� �������� � ������ " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                            End Try
                        End Using
                    End If
                Next i
                If MyErrFlag = 0 Then
                    '---------������� ��������
                    UploadNew_Picture_ToMagento(MyCode)
                End If
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        '---------������� ��������
                        UploadNew_Picture_ToMagento(MyCode)
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("��������� ���������� � �������� � ������ " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                        End If
                        MyGlobalStr = MyGlobalStr + "��������� ���������� � �������� � ������ " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("��������� ������ 4 ---> " + ex2.Message, MsgBoxStyle.Information, "��������!")
                    End If
                    MyGlobalStr = MyGlobalStr + "��������� ������ 4 ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub UploadNew_Picture_ToMagento(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� ����� �������� �������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim rdr As SqlDataReader
        Dim MyRep As String
        Dim MyStr As String

        MySQLStr = "SELECT Picture "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Pictures "
        MySQLStr = MySQLStr & "WHERE (ScalaItemCode = N'" & Trim(MyCode) & "') "
        Using MyConn As SqlConnection = New SqlConnection(Declarations.MyNETConnStr)
            Try
                MyConn.Open()
                Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                    rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                    While (rdr.Read())
                        '---------------�������� ����� �������� �� ���� Magento----------
                        MyStr = "{"
                        MyStr = MyStr & """entry"":{"
                        MyStr = MyStr & """id"":""0"", "
                        MyStr = MyStr & """media_type"":""image"","
                        MyStr = MyStr & """label"":""" & MyCode & ""","
                        MyStr = MyStr & """position"":""1"","
                        MyStr = MyStr & """disabled"":""false"","
                        MyStr = MyStr & """types"":[""image"",""small_image"",""thumbnail"",""swatch_image""],"
                        MyStr = MyStr & """file"":"""","
                        MyStr = MyStr & """content"": {"
                        MyStr = MyStr & """type"":""image/jpeg"","
                        MyStr = MyStr & """name"":""" & MyCode & ".jpg"","
                        MyStr = MyStr & """base64EncodedData"":""" & Convert.ToBase64String(rdr.GetValue(0)) & """"
                        MyStr = MyStr & "}"
                        MyStr = MyStr & "}"
                        MyStr = MyStr & "}"

                        Using MC As New WebClient
                            Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode) + "/media")
                            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
                            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyStr)))

                            Try
                                MyRep = MC.UploadString(MyUrl1, "POST", MyStr)
                            Catch ex1 As WebException
                                If My.Settings.MyDebug = "YES" Then
                                    MsgBox("�������� �������� � ������ " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "��������!")
                                End If
                                MyGlobalStr = MyGlobalStr + "�������� �������� � ������ " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                            End Try
                        End Using
                    End While
                End Using
            Catch ex As Exception
            End Try
        End Using
    End Sub

    Public Sub Update_ExtendedPrice_InMagento(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// �������� (����������) ������������ ������ �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

    End Sub

    Public Sub UploadPictures_ToMagento(ByVal MySupplierCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� �������� ���������� ���������� ��� ���� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        MyGlobalStr = ""

        MySQLStr = "exec spp_WEB_Pictures_MG_FromDB N'" + Trim(MySupplierCode) + "' "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            Declarations.MyRec.MoveLast()
            MyUploadPicturesToMagento.Label3.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                CreateUpdate_Picture_InMagento(Declarations.MyRec.Fields("Code").Value)
                Declarations.MyRec.MoveNext()
                i = i + 1
                MyUploadPicturesToMagento.Label2.Text = i
                Application.DoEvents()
            End While
        End If
        MyUploadPicturesToMagento.GroupBox1.BackColor = Color.LightGreen
        Application.DoEvents()
        trycloseMyRec()
    End Sub

    Public Sub UploadAvailability_ToMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� �������� ���������� � ����������� ������� �� ���� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        MyGlobalStr = ""

        MySQLStr = "spp_WEB_ItemAvailability_MG_FromDB "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
        Else
            Declarations.MyRec.MoveLast()
            MyUploadAvailabilityToMagento.Label3.Text = Declarations.MyRec.RecordCount
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                Update_Availability_InMagento(Declarations.MyRec.Fields("Code").Value, Declarations.MyRec.Fields("MyJSon").Value)
                Declarations.MyRec.MoveNext()
                i = i + 1
                MyUploadAvailabilityToMagento.Label2.Text = i
                Application.DoEvents()
            End While
        End If
        MyUploadAvailabilityToMagento.GroupBox1.BackColor = Color.LightGreen
        Application.DoEvents()
        trycloseMyRec()
    End Sub

    Public Sub Update_Availability_InMagento(ByVal MyCode As String, ByVal MyJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ���������� ������ � �������� �� ����� Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyWR As HttpWebResponse
        Dim MyStr As String
        Dim MyRep As String

        Using MC = New WebClient
            Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode))
            MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
            MC.Headers(HttpRequestHeader.ContentType) = "application/json"
            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyJSon)))

            Try
                MyRep = MC.UploadString(MyUrl, "PUT", MyStr)
                '---��������� ���������� �� ���������� ��������
                UpdateAvailability_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                MyWR = ex.Response
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("���������� ����������� ������ " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "��������!")
                End If
                MyGlobalStr = MyGlobalStr + "���������� ����������� ������ " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub UpdateAvailability_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// ��������� � �� Scala ���������� �� �������� ���������� ���������� � ����������� �������� 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_WEB_ItemAvailability_MG_Correspondence "
        MySQLStr = MySQLStr & "SET WEBStatus = 0 "
        MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Module
