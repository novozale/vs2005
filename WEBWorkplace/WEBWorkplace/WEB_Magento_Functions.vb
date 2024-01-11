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
        '// Выгрузка новой информации на сайт Magento
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MyGlobalStr = ""
        '-----------------Обновление информации из Scala--------------------------------------
        MySQLStr = "exec spp_WEB_ALL_FromScala "
        InitMyConn(False)
        'Declarations.MyConn.Execute(MySQLStr)

        '-----------------Выгрузка категорий (групп и подгрупп товаров)-----------------------
        UploadInfo_Categories_ToMagento(MyUploadType)

        '-----------------Выгрузка опций аттрибута "производитель"----------------------------
        UploadInfo_ManufacturerList_ToMagento(MyUploadType)

        '-----------------Выгрузка опций аттрибута "страна происхождения"---------------------
        UploadInfo_CountriesList_ToMagento(MyUploadType)

        '-----------------Выгрузка товаров----------------------------------------------------
        UploadInfo_Products_ToMagento(MyUploadType)

    End Sub

    Public Sub UploadInfo_Categories_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации по категориям на сайте Magento
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------Удаление неиспользуемых категорий-------------------------------------
            Delete_NotUsedCategories_FromMagento()
        End If

        '---------------загрузка категорий ---------------------------------------------------
        Upload_Categories_ToMagento(MyUploadType)

    End Sub

    Public Sub Delete_NotUsedCategories_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление неиспользуемых категорий (подкатегорий) на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Categories_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации по категориям на сайт Magento
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        If MyUploadType = 0 Then        '------полная выгрузка
            '-------------удаление записей о подгруппах, помеченных на удаление, но не выгруженных на Magento--------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 2"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------удаление записей о группах, помеченных на удаление, но не выгруженных на Magento-----------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 1"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------удаление категорий 2 уровня (подгруппы продуктов) помеченных на удаление-------------------
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

            '-------------удаление категорий 1 уровня (группы продуктов) помеченных на удаление----------------------
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

            '-------------загрузка категорий 1 уровня (группы продуктов)----------------------
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

            '-------------загрузка категорий 2 уровня (подгруппы продуктов)-------------------
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

        Else                            '------выгрузка только измененных данных
            '-------------удаление записей о подгруппах, помеченных на удаление, но не выгруженных на Magento--------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 2"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------удаление записей о группах, помеченных на удаление, но не выгруженных на Magento-----------
            MySQLStr = "exec spp_WEB_ItemGroups_MG_FromDB 4, 1"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '-------------удаление категорий 2 уровня (подгруппы продуктов) помеченных на удаление-------------------
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
            '-------------удаление категорий 1 уровня (группы продуктов) помеченных на удаление----------------------
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
            '-------------загрузка категорий 1 уровня (группы продуктов)----------------------
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
            '-------------загрузка категорий 2 уровня (подгруппы продуктов)-------------------
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
        '// Удаление категории на сайте Magento 
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
                        MsgBox("Удаление категории " + MyCategoryID + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "Удаление категории " + MyCategoryID + " ---> " + ex.Message + Chr(13) + Chr(10)
                End If
            End Try
        End Using
    End Sub

    Public Sub CreateUpdate_Category_InMagento(ByVal MyCode As String, ByVal MyCategoryID As String, ByVal MyCategoryJSon As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание (обновление) категории на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyWR As HttpWebResponse
        Dim MyStr As String
        Dim MyRep As String
        Dim NewCategoryID As String

        If Trim(MyCategoryID) = "" Then     '---создание
            Using MC As New WebClient
                Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/categories/")
                MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                MC.Headers(HttpRequestHeader.ContentType) = "application/json"
                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                Try
                    MyRep = MC.UploadString(MyUrl, "POST", MyStr)
                    '---сохраняем информацию о создании категории
                    NewCategoryID = GetCategoryID(MyRep)
                    If Trim(NewCategoryID) <> "" Then
                        CreateUpdate_Category_InMagento_Confirm(MyCode, Trim(NewCategoryID), MyLevel)
                    End If
                Catch ex As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("Создание категории " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "Создание категории " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                End Try
            End Using
        Else
            '---обновление
            Using MC As New WebClient
                Dim MyUrl As New Uri("http://spbprd7/index.php/rest/V1/categories/" & Trim(MyCategoryID))
                MC.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                MC.Headers(HttpRequestHeader.ContentType) = "application/json"
                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                Try
                    MyRep = MC.UploadString(MyUrl, "PUT", MyStr)
                    '---сохраняем информацию об обновлении категории
                    CreateUpdate_Category_InMagento_Confirm(MyCode, MyCategoryID, MyLevel)
                Catch ex As WebException
                    Try
                        MyWR = ex.Response
                        If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                            '---Если категория не найдена и обновить не удалось - то создаем заново.
                            Using MC1 As New WebClient
                                Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/categories/" & Trim(MyCategoryID))
                                MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                                MC1.Headers(HttpRequestHeader.ContentType) = "application/json"
                                MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyCategoryJSon)))

                                Try
                                    MyRep = MC1.UploadString(MyUrl1, "POST", MyStr)
                                    '---сохраняем информацию о создании категории
                                    NewCategoryID = GetCategoryID(MyRep)
                                    If Trim(NewCategoryID) <> "" Then
                                        CreateUpdate_Category_InMagento_Confirm(MyCode, Trim(NewCategoryID), MyLevel)
                                    End If
                                Catch ex1 As WebException
                                    If My.Settings.MyDebug = "YES" Then
                                        MsgBox("Создание категории при обновлении " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "Внимание!")
                                    End If
                                    MyGlobalStr = MyGlobalStr + "Создание категории при обновлении " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                                End Try
                            End Using
                        Else
                            If My.Settings.MyDebug = "YES" Then
                                MsgBox("Обновление категории " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                            End If
                            MyGlobalStr = MyGlobalStr + "Обновление категории " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                        End If
                    Catch ex2 As Exception
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("получение ответа 1 ---> " + ex2.Message, MsgBoxStyle.Information, "Внимание!")
                        End If
                        MyGlobalStr = MyGlobalStr + "получение ответа 1 ---> " + ex2.Message + Chr(13) + Chr(10)
                    End Try
                End Try
            End Using
        End If
    End Sub

    Public Sub Delete_Category_FromMagento_Confirm(ByVal MyCategoryID As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном удалении категории 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyLevel = 1 Then     '-----Группа продуктов
            '---Таблица группы
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemGroup INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemGroup_MG_Correspondence ON tbl_WEB_ItemGroup.Code = tbl_WEB_ItemGroup_MG_Correspondence.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_ItemGroup_MG_Correspondence.MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        Else                    '-----Подгруппа продуктов
            '---Таблица подгруппы
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup "
            MySQLStr = MySQLStr & "FROM tbl_WEB_ItemSubGroup INNER JOIN "
            MySQLStr = MySQLStr & "tbl_WEB_ItemSubGroup_MG_Correspondence ON tbl_WEB_ItemSubGroup.SubgroupID = tbl_WEB_ItemSubGroup_MG_Correspondence.Code "
            MySQLStr = MySQLStr & "WHERE (tbl_WEB_ItemSubGroup_MG_Correspondence.MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup_MG_Correspondence"
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCategoryID) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
    End Sub

    Public Sub CreateUpdate_Category_InMagento_Confirm(ByVal MyCode As String, ByVal MyCategoryID As String, ByVal MyLevel As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном создании (обновлении) категории 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If MyLevel = 1 Then     '-----Группа продуктов
            '---таблица связки удаление
            MySQLStr = "DELETE FROM tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') OR "
            MySQLStr = MySQLStr & "(MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки создание
            MySQLStr = "INSERT INTO tbl_WEB_ItemGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "(Code, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyCode) & "', N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---Таблица группы обновление
            MySQLStr = "UPDATE tbl_WEB_ItemGroup "
            MySQLStr = MySQLStr & "SET RMStatus = 0, WEBStatus = 0 "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        Else                    '-----Подгруппа продуктов
            '---таблица связки удаление
            MySQLStr = "DELETE FROM tbl_WEB_ItemSubGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (Code = N'" & Trim(MyCode) & "') OR "
            MySQLStr = MySQLStr & "(MagentoCode = N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки создание
            MySQLStr = "INSERT INTO tbl_WEB_ItemSubGroup_MG_Correspondence "
            MySQLStr = MySQLStr & "(Code, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyCode) & "', N'" & Trim(MyCategoryID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---Таблица подгруппы обновление
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
        '// Получение ID из строки возврата JSON 
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
        '// Загрузка списка производителей в опции аттрибута "proizvoditel" 
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------Удаление неиспользуемых производителей из опций аттрибута "proizvoditel"-
            Delete_NotUsedManufacturers_FromMagento()
        End If

        '---------------загрузка списка производителей в опции аттрибута "proizvoditel"--
        Upload_Manufacturers_ToMagento()
    End Sub

    Public Sub Delete_NotUsedManufacturers_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление неиспользуемых производителей из опций аттрибута "proizvoditel"
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Manufacturers_ToMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка списка производителей в опции аттрибута "proizvoditel"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        '----Удаление производителей, помеченных на удаление, но не выгруженных в Magento-
        MySQLStr = "exec spp_WEB_Manufacturers_MG_FromDB 0"
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----Удаление производителей, помеченных на удаление----------------------------
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

        '----------------Создание производителей----------------------------------------
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
        '// Удаление производителя на сайте Magento 
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
                            MsgBox("Удаление производителя " + MyManufacturerID + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                        End If
                        MyGlobalStr = MyGlobalStr + "Удаление производителя " + MyManufacturerID + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("получение ответа 2 ---> " + ex2.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "Удаление производителя " + MyManufacturerID + " ---> " + ex.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub


    Public Sub Create_Manufacturer_InMagento(ByVal MyName As String, ByVal MyManufacturerJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание производителя на сайте Magento 
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
                '---сохраняем информацию о создании производителя
                Create_Manufacturer_InMagento_Confirm(MyName)
            Catch ex As WebException
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("Создание производителя " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                End If
                MyGlobalStr = MyGlobalStr + "Создание производителя " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub Delete_Manufacturer_FromMagento_Confirm(ByVal MyManufacturerID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном удалении производителя 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Таблица группы
        MySQLStr = "DELETE FROM tbl_WEB_Manufacturers "
        MySQLStr = MySQLStr & "FROM tbl_WEB_Manufacturers INNER JOIN "
        MySQLStr = MySQLStr & "tbl_WEB_Manufacturers_MG_Correspondence ON CASE WHEN Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) "
        MySQLStr = MySQLStr & "= '' THEN tbl_WEB_Manufacturers.Name ELSE Ltrim(Rtrim(tbl_WEB_Manufacturers.WEBName)) "
        MySQLStr = MySQLStr & "END = tbl_WEB_Manufacturers_MG_Correspondence.ManufacturerName "
        MySQLStr = MySQLStr & "WHERE (tbl_WEB_Manufacturers_MG_Correspondence.MagentoCode = N'" & Trim(MyManufacturerID) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---таблица связки
        MySQLStr = "DELETE FROM tbl_WEB_Manufacturers_MG_Correspondence "
        MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyManufacturerID) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Public Sub Create_Manufacturer_InMagento_Confirm(ByVal MyName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном создании производителя 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyManufacturerID As String
        Dim MySQLStr As String

        MyManufacturerID = ""
        MyManufacturerID = GetManufacturerIDByName(MyName)
        If Trim(MyManufacturerID) <> "" Then
            '---таблица связки удаление
            MySQLStr = "DELETE FROM tbl_WEB_Manufacturers_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyManufacturerID) & "') "
            MySQLStr = MySQLStr & "OR (ManufacturerName = N'" & Trim(MyName) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки создание
            MySQLStr = "INSERT INTO tbl_WEB_Manufacturers_MG_Correspondence "
            MySQLStr = MySQLStr & "(ManufacturerName, MagentoCode) "
            MySQLStr = MySQLStr & "VALUES (N'" & Trim(MyName) & "'"
            MySQLStr = MySQLStr & ", N'" & Trim(MyManufacturerID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица производителей обновление
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
        '// получение ID производителя по имени с сайта Magento
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
                    MsgBox("получение id производителя " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                End If
                MyGlobalStr = MyGlobalStr + "получение id производителя " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Function

    Public Sub UploadInfo_CountriesList_ToMagento(ByVal MyUploadType As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка стран в опции аттрибута "strana_proishozhdenija" 
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------Удаление неиспользуемых производителей из опций аттрибута "strana_proishozhdenija"-
            Delete_NotUsedCountries_FromMagento()
        End If

        '---------------загрузка списка производителей в опции аттрибута "strana_proishozhdenija"--
        Upload_Countries_ToMagento()
    End Sub

    Public Sub Delete_NotUsedCountries_FromMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление неиспользуемых стран из опций аттрибута "strana_proishozhdenija"
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Countries_ToMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка списка стран в опции аттрибута "strana_proishozhdenija"
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        '----------------Создание стран--------------------------------------------------
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
        '// Создание страны на сайте Magento 
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
                '---сохраняем информацию о создании страны
                Create_Country_InMagento_Confirm(MyName)
            Catch ex As WebException
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("Создание страны " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                End If
                MyGlobalStr = MyGlobalStr + "Создание страны " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub Create_Country_InMagento_Confirm(ByVal MyName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном создании страны 
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCountryID As String
        Dim MySQLStr As String

        MyCountryID = ""
        MyCountryID = GetCountryIDByName(MyName)
        If Trim(MyCountryID) <> "" Then
            '---таблица связки удаление
            MySQLStr = "DELETE FROM tbl_WEB_Countries_MG_Correspondence "
            MySQLStr = MySQLStr & "WHERE (MagentoCode = N'" & Trim(MyCountryID) & "') "
            MySQLStr = MySQLStr & "OR (CountryName = N'" & Replace(Trim(MyName), "'", "''") & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---таблица связки создание
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
        '// получение ID страны по имени с сайта Magento
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
                    MsgBox("Получение ID страны " + MyName + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                End If
                MyGlobalStr = MyGlobalStr + "Получение ID страны " + MyName + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Function

    Public Sub UploadInfo_Products_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление информации по товарам на сайте Magento
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '/////////////////////////////////////////////////////////////////////////////////////

        If MyUploadType = 0 Then
            '-----------Выставление невидимости для неиспользуемых продуктов------------------
            Hide_NotUsedProducts_InMagento()
        End If

        '---------------загрузка категорий ---------------------------------------------------
        Upload_Products_ToMagento(MyUploadType)

    End Sub

    Public Sub Hide_NotUsedProducts_InMagento()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перевод в невидимые неиспользуемых продуктов на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////


    End Sub

    Public Sub Upload_Products_ToMagento(ByVal MyUploadType As Integer)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка информации по продуктам на сайт Magento
        '// MyUploadType =  0 полная выгрузка
        '//                 1 только новая информация
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim i As Integer

        If MyUploadType = 0 Then        '------полная выгрузка
            '-------------перевод в невидимые товаров, помеченных на удаление-----------------
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
            '--------------загрузка товаров---------------------------------------------------
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

        Else                            '------выгрузка только измененных данных
            '-------------перевод в невидимые товаров, помеченных на удаление-----------------
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

            '--------------загрузка товаров---------------------------------------------------
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
        '// Выставление признака "невидимый" для удаленного в Scala товара
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
                '---сохраняем информацию об удалении товара
                Hide_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        Hide_Product_InMagento_Confirm(MyCode)
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("Удаление (невидимость) товара " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                        End If
                        MyGlobalStr = MyGlobalStr + "Удаление (невидимость) товара " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("получение ответа 3 ---> " + ex2.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "получение ответа 3 ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub Hide_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном удалении запаса 
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
        '// Создание (обновление) продукта на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---------------Создание (обновление) продукта-----------------------------------
        CreateUpdate_ProductCore_InMagento(MyCode, MyJSon)

        '---------------Создание (обновление) картинки продукта--------------------------
        CreateUpdate_Picture_InMagento(MyCode)

        '---------------Создание (обновление) прайса продукта----------------------------
        Update_ExtendedPrice_InMagento(MyCode)

    End Sub

    Public Sub CreateUpdate_ProductCore_InMagento(ByVal MyCode As String, ByVal MyJSon As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание (обновление) записи о продукте на сайте Magento 
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
                '---сохраняем информацию об обновлении продукта
                'CreateUpdate_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        '----сли продукт не найден - то создаем
                        Using MC1 = New WebClient
                            Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/products/")
                            MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                            MC1.Headers(HttpRequestHeader.ContentType) = "application/json"
                            MyStr = Encoding.GetEncoding("Windows-1251").GetString(Encoding.Convert(Encoding.GetEncoding("Windows-1251"), Encoding.GetEncoding("UTF-8"), Encoding.GetEncoding("Windows-1251").GetBytes(MyJSon)))

                            Try
                                MyRep = MC1.UploadString(MyUrl1, "POST", MyStr)
                                '---сохраняем информацию о создании продукта
                                'CreateUpdate_Product_InMagento_Confirm(MyCode)
                            Catch ex1 As WebException
                                If My.Settings.MyDebug = "YES" Then
                                    MsgBox("Создание товара " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "Внимание!")
                                End If
                                MyGlobalStr = MyGlobalStr + "Создание товара " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                            End Try
                        End Using
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("Обновление товара " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                        End If
                        MyGlobalStr = MyGlobalStr + "Обновление товара " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("Создание товара 1 " + MyCode + " ---> " + ex2.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "Создание товара 1 " + MyCode + " ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub CreateUpdate_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном создании (обновлении) продукта 
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
        '// Создание (обновление) записи о картинке продукта на сайте Magento 
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
                '-------Получение списка картинок для данного продукта
                MyRep = MC.DownloadString(MyUrl)
                MyRep = "{""data"":" + MyRep + "}"
                Dim json As JObject = JObject.Parse(MyRep)
                Dim jarray As JArray = json.SelectToken("data")
                MyErrFlag = 0
                For i As Integer = 0 To jarray.Count - 1
                    If Trim(jarray(i).Item("id").ToString) <> "" Then
                        '--------Удаление старых картинок
                        Using MC1 = New WebClient
                            Dim MyUrl1 As New Uri("http://spbprd7/index.php/rest/V1/products/" + Trim(MyCode) + "/media/" + Trim(jarray(i).Item("id").ToString))
                            MC1.Headers(HttpRequestHeader.Authorization) = "Bearer " + Declarations.MyAccessToken
                            MC1.Headers(HttpRequestHeader.ContentType) = "application/json"

                            Try
                                MyRep = MC1.UploadString(MyUrl1, "DELETE", "")
                            Catch ex1 As WebException
                                MyErrFlag = 1
                                If My.Settings.MyDebug = "YES" Then
                                    MsgBox("Удаление картинки к товару " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "Внимание!")
                                End If
                                MyGlobalStr = MyGlobalStr + "Удаление картинки к товару " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
                            End Try
                        End Using
                    End If
                Next i
                If MyErrFlag = 0 Then
                    '---------Создаем картинку
                    UploadNew_Picture_ToMagento(MyCode)
                End If
            Catch ex As WebException
                Try
                    MyWR = ex.Response
                    If MyWR.StatusCode = System.Net.HttpStatusCode.NotFound Then
                        '---------Создаем картинку
                        UploadNew_Picture_ToMagento(MyCode)
                    Else
                        If My.Settings.MyDebug = "YES" Then
                            MsgBox("Получение информации о картинке к товару " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                        End If
                        MyGlobalStr = MyGlobalStr + "Получение информации о картинке к товару " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
                    End If
                Catch ex2 As Exception
                    If My.Settings.MyDebug = "YES" Then
                        MsgBox("получение ответа 4 ---> " + ex2.Message, MsgBoxStyle.Information, "Внимание!")
                    End If
                    MyGlobalStr = MyGlobalStr + "получение ответа 4 ---> " + ex2.Message + Chr(13) + Chr(10)
                End Try
            End Try
        End Using
    End Sub

    Public Sub UploadNew_Picture_ToMagento(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка новой картинки продукта на сайт Magento 
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
                        '---------------Загрузка новой картинки на сайт Magento----------
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
                                    MsgBox("Создание картинки к товару " + MyCode + " ---> " + ex1.Message, MsgBoxStyle.Information, "Внимание!")
                                End If
                                MyGlobalStr = MyGlobalStr + "Создание картинки к товару " + MyCode + " ---> " + ex1.Message + Chr(13) + Chr(10)
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
        '// Создание (обновление) расширенного прайса продукта на сайте Magento 
        '//
        '////////////////////////////////////////////////////////////////////////////////

    End Sub

    Public Sub UploadPictures_ToMagento(ByVal MySupplierCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура загрузки картинок выбранного поставщика или всех на сайт Magento 
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
        '// Процедура загрузки информации о доступности товаров на сайт Magento 
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
        '// Обновление записи о продукте на сайте Magento 
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
                '---сохраняем информацию об обновлении продукта
                UpdateAvailability_Product_InMagento_Confirm(MyCode)
            Catch ex As WebException
                MyWR = ex.Response
                If My.Settings.MyDebug = "YES" Then
                    MsgBox("Обновление доступности товара " + MyCode + " ---> " + ex.Message, MsgBoxStyle.Information, "Внимание!")
                End If
                MyGlobalStr = MyGlobalStr + "Обновление доступности товара " + MyCode + " ---> " + ex.Message + Chr(13) + Chr(10)
            End Try
        End Using
    End Sub

    Public Sub UpdateAvailability_Product_InMagento_Confirm(ByVal MyCode As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Занесение в БД Scala информации об успешном обновлении информации о доступности продукта 
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
