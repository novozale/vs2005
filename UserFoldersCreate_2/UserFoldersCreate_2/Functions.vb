Imports System
Imports System.IO
Imports System.Security.AccessControl
Imports System.DirectoryServices
Imports System.Text
Imports System.Diagnostics


Module Functions
    Dim MyDomain As String

    Public Function WriteAppLog(ByVal MyMsg As String, ByVal MyType As EventLogEntryType) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ф-ция записи информации в системные логи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim EventLog1 As New System.Diagnostics.EventLog()
        Dim EventInstance1 As New EventInstance(1, 6)

        If Not EventLog.SourceExists(My.Settings.AppName) Then
            EventLog.CreateEventSource(My.Settings.AppName, My.Settings.AppLogType)
        End If
        EventLog1.Source = My.Settings.AppName
        EventInstance1.EntryType = MyType
        EventLog1.WriteEvent(EventInstance1, New String() {MyMsg})
        WriteAppLog = True
    End Function

    Public Function UserFolderCreation() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ф-ция просмотра АД, поиска пользователей и создания папок, соответствующих пользователям
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim rootDse As DirectoryEntry
        Dim MyEntry As DirectoryEntry
        Dim MyDsearch As DirectorySearcher
        Dim myResultPropColl As ResultPropertyCollection
        Dim myKey As String
        Dim MyLogin As String
        Dim MyGroup() As String
        Dim MyGRPCnt As Int64
        Dim MyGRPNum As Int64
        Dim i As Int64

        Try
            MyGRPNum = 0
            MyGRPCnt = My.Settings.UserGroups.Count
            ReDim MyGroup(MyGRPCnt - 1)
            For i = 0 To MyGRPCnt - 1
                MyGroup(i) = My.Settings.UserGroups(i)
            Next
            rootDse = New DirectoryEntry("LDAP://rootDSE")
            MyEntry = New DirectoryEntry("LDAP://" + DirectCast(rootDse.Properties("defaultNamingContext").Value, String))
            MyDomain = Right(MyEntry.Name, Len(MyEntry.Name) - 3)
            MyDsearch = New DirectorySearcher(MyEntry)

            MyDsearch.Filter = "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user))"
            MyDsearch.SearchScope = SearchScope.Subtree
            For Each result As SearchResult In MyDsearch.FindAll()
                If Not (IsNothing(result)) Then
                    myResultPropColl = result.Properties
                    For Each myKey In myResultPropColl.PropertyNames
                        If myKey = "samaccountname" Then
                            MyLogin = myResultPropColl(myKey)(0)
                            If BelongToGroup(MyLogin, MyGroup, MyGRPCnt, MyGRPNum) = True Then
                                CreateFolder(MyLogin, MyGroup(MyGRPNum))
                            End If
                        End If
                    Next
                End If
            Next
            UserFolderCreation = True
        Catch ex As Exception
            UserFolderCreation = False
            WriteAppLog("Ошибка при создании пользовательских каталогов " & ex.Message, EventLogEntryType.Error)
        End Try
    End Function

    Public Function BelongToGroup(ByVal MyLogin As String, ByRef MyGroup() As String, ByVal MyGRPCnt As Int64, ByRef MyGRPNum As Int64) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Находим, принадлежит ли пользователь к группам, указанным в конфиге, и если да - то к какой
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim rootDse As DirectoryEntry
        Dim MyEntry As DirectoryEntry
        Dim MyDsearch As DirectorySearcher
        Dim groupCount As Int64
        Dim GroupArr As Array
        Dim GroupName As String
        Dim Counter As Int64
        Dim i As Int64
        Dim j As Int64

        Try
            rootDse = New DirectoryEntry("LDAP://rootDSE")
            MyEntry = New DirectoryEntry("LDAP://" + DirectCast(rootDse.Properties("defaultNamingContext").Value, String))
            MyDsearch = New DirectorySearcher(MyEntry)

            MyDsearch.Filter = "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(objectCategory=user)(samaccountname=" + MyLogin + "))"
            MyDsearch.PropertiesToLoad.Add("memberOf")
            MyDsearch.SearchScope = SearchScope.Subtree
            Dim result As SearchResult = MyDsearch.FindOne()
            If Not (IsNothing(result)) Then
                Try
                    groupCount = result.Properties("memberOf").Count
                Catch ex As NullReferenceException
                    groupCount = 0
                End Try
                If groupCount > 0 Then
                    For Counter = 0 To groupCount - 1
                        GroupName = ""
                        GroupName = CStr(result.Properties("memberOf")(Counter))
                        GroupArr = Split(GroupName, ",")
                        For i = 0 To GroupArr.Length - 1
                            If Not (IsNothing(GroupArr(i))) Then
                                If Left(GroupArr(i), 3) = "OU=" Then
                                    For j = 0 To MyGRPCnt - 1
                                        If MyGroup(j) = Mid(GroupArr(i), 4, Len(GroupArr(i)) - 3) Then
                                            MyGRPNum = j
                                            If My.Settings.MyDebug = "YES" Then
                                                WriteAppLog("Пользователь " + MyLogin + " Группа " + MyGroup(MyGRPNum) + " В домене " + MyEntry.Path, EventLogEntryType.Information)
                                            End If
                                            BelongToGroup = True
                                            Exit Function
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    Next
                End If
            End If
            BelongToGroup = False
        Catch ex As Exception
            BelongToGroup = False
            WriteAppLog("Ошибка при создании пользовательских каталогов " & ex.Message, EventLogEntryType.Error)
        End Try
    End Function

    Public Function CreateFolder(ByVal MyLogin As String, ByVal MyGroup As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создаем каталог пользователя по соответствующему пути
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyPath As String
        Dim CatalogUser As DirectoryInfo
        Dim CatalogParent As DirectoryInfo

        MyPath = ""
        Try
            MyPath = My.Settings.Item(MyGroup)
            CatalogParent = New DirectoryInfo(MyPath)
        Catch ex As Exception
            WriteAppLog("В конфигурации не задан путь создания пользовательского каталога для контейнера " & MyGroup & ". Обратитесь к администратору", EventLogEntryType.Error)
            Exit Function
        End Try

        CatalogUser = New DirectoryInfo(MyPath + MyLogin)
        Try
            If Not CatalogUser.Exists Then
                Dim ds As DirectorySecurity = New DirectorySecurity
                ds.AddAccessRule(New FileSystemAccessRule(MyDomain + "\Администраторы домена", FileSystemRights.FullControl, InheritanceFlags.ObjectInherit + InheritanceFlags.ContainerInherit, _
                    PropagationFlags.None, AccessControlType.Allow))
                ds.AddAccessRule(New FileSystemAccessRule(MyDomain + "\Administrator", FileSystemRights.FullControl, InheritanceFlags.ObjectInherit + InheritanceFlags.ContainerInherit, _
                    PropagationFlags.None, AccessControlType.Allow))
                ds.AddAccessRule(New FileSystemAccessRule(MyDomain + "\" + MyLogin, FileSystemRights.Modify, InheritanceFlags.ObjectInherit + InheritanceFlags.ContainerInherit, _
                    PropagationFlags.None, AccessControlType.Allow))
                ds.SetAccessRuleProtection(True, True)
                CatalogParent.CreateSubdirectory(MyLogin, ds)
                If My.Settings.MyDebug = "YES" Then
                    WriteAppLog("Создан каталог " + CatalogUser.FullName, EventLogEntryType.Information)
                End If
            Else
                If My.Settings.MyDebug = "YES" Then
                    WriteAppLog("Уже существует каталог " & CatalogUser.FullName, EventLogEntryType.Information)
                End If
            End If
        Catch ex As Exception
            If My.Settings.MyDebug = "YES" Then
                WriteAppLog("Ошибка создания каталога  " & MyPath + MyLogin & ". " & ex.Message, EventLogEntryType.Error)
            End If
        End Try
    End Function
End Module
