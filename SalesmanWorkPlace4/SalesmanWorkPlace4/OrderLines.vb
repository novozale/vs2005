Public Class OrderLines

    Private Sub OrderLines_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна редактирования строк предложения - загружаем исходную информацию
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка

        Declarations.MyOrderNum = Trim(MyEditHeader.Label3.Text)
        Label6.Text = Declarations.MyOrderNum

        '---параметры валюты
        MySQLStr = "SELECT tbl_OR010300.OR01028 AS CurrencyCode, "
        MySQLStr = MySQLStr & "tbl_OR010300.OR01067 AS CurrencyValue, "
        MySQLStr = MySQLStr & "SYCD0100.SYCD009 AS CurrencyName "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON tbl_OR010300.OR01028 = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01001 = N'" & Declarations.MyOrderNum & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Declarations.CurrencyCode = 0
            Declarations.CurrencyName = "RUB"
            Declarations.CurrencyValueOrder = 1
        Else
            Declarations.CurrencyCode = Declarations.MyRec.Fields("CurrencyCode").Value
            Declarations.CurrencyName = Declarations.MyRec.Fields("CurrencyName").Value
            Declarations.CurrencyValue = Declarations.MyRec.Fields("CurrencyValue").Value
            If Declarations.CurrencyValueOrder = 0 Then
                Declarations.CurrencyValueOrder = 1
            End If
        End If
        trycloseMyRec()

        '--и курс валюты на текущий день
        MySQLStr = "SELECT SYCH006 "
        MySQLStr = MySQLStr & "FROM SYCH0100 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SYCH001 = " & Declarations.CurrencyCode & ") And "
        MySQLStr = MySQLStr & "(SYCH004 <= GETDATE()) And "
        MySQLStr = MySQLStr & "(SYCH005 > GETDATE())"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        Declarations.CurrencyValue = Declarations.MyRec.Fields("SYCH006").Value
        trycloseMyRec()

        Label8.Text = Declarations.CurrencyName
        Label9.Text = Declarations.CurrencyValue

        '--минимальные значения маржи
        MySQLStr = "SELECT tbl_MarginLimitMatrixDetails.MarginLevelFrom, tbl_MarginLimitMatrixDetails.MarginLevelTo "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CustomerCard0300 ON tbl_OR010300.OR01003 = tbl_CustomerCard0300.SL01001 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_MarginLimitMatrix ON tbl_CustomerCard0300.MarginLimitLevel = tbl_MarginLimitMatrix.ID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_MarginLimitMatrixDetails ON tbl_MarginLimitMatrix.ID = tbl_MarginLimitMatrixDetails.ID "
        MySQLStr = MySQLStr & "WHERE (tbl_MarginLimitMatrixDetails.CheckLevel = N'ShipmentsWithLowMarginLevel1') "
        MySQLStr = MySQLStr & "AND (tbl_OR010300.OR01001 = N'" & Declarations.MyOrderNum & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            '---не определили - берем базовый вариант (ID = 0)
            MySQLStr = "SELECT MarginLevelFrom, MarginLevelTo "
            MySQLStr = MySQLStr & "FROM tbl_MarginLimitMatrixDetails "
            MySQLStr = MySQLStr & "WHERE (CheckLevel = N'ShipmentsWithLowMarginLevel1') AND (ID = 0) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                '---не определили - берем значения по умолчанию
                Label33.Text = 20
                Declarations.MinMarginLevelManager = 20
                Label34.Text = 15
                Declarations.MinMarginLevelDirector = 15
            Else
                Label33.Text = Declarations.MyRec.Fields("MarginLevelFrom").Value
                Declarations.MinMarginLevelManager = Declarations.MyRec.Fields("MarginLevelFrom").Value
                Label34.Text = Declarations.MyRec.Fields("MarginLevelTo").Value
                Declarations.MinMarginLevelDirector = Declarations.MyRec.Fields("MarginLevelTo").Value
                trycloseMyRec()
            End If
        Else
            Label33.Text = Declarations.MyRec.Fields("MarginLevelFrom").Value
            Declarations.MinMarginLevelManager = Declarations.MyRec.Fields("MarginLevelFrom").Value
            Label34.Text = Declarations.MyRec.Fields("MarginLevelTo").Value
            Declarations.MinMarginLevelDirector = Declarations.MyRec.Fields("MarginLevelTo").Value
            trycloseMyRec()
        End If

        Button8.Text = "Только дост"
        GetProductList()

    End Sub

    Private Function GetProductList()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ф-ция готовит список продуктов Электроскандии для последующего выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Declarations.CustomerNumber = MyEditHeader.TextBox1.Text

        DataPreparation(0, 0)
        OrderPreparation()
        Label3.Text = ""
        ChangeButtonsStatus()
    End Function

    Private Function DataPreparation(ByVal MyParam As Integer, ByVal MyParamAct As Integer)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim PriceNum As String                   'прайс - лист
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet
        Dim WhName As String                     'название склада

        '---прайс - лист
        MySQLStr = "SELECT OR01051 "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Declarations.MyOrderNum & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        PriceNum = Declarations.MyRec.Fields("OR01051").Value
        trycloseMyRec()

        '---номер склада
        MySQLStr = "SELECT SC23001, SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23001 IN "
        MySQLStr = MySQLStr & "(SELECT OR01050 "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 "
        MySQLStr = MySQLStr & "WHERE (OR01001 = N'" & Declarations.MyOrderNum & "')))"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
            MySQLStr = "SELECT SC23001, SC23002 "
            MySQLStr = MySQLStr & "FROM SC230300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (SC23001 = '01') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
        End If
        Declarations.WHNum = Declarations.MyRec.Fields("SC23001").Value
        WhName = Declarations.MyRec.Fields("SC23002").Value
        trycloseMyRec()

        'MySQLStr = "SELECT tbl_WEB_Pictures.PictureSmall, "
        MySQLStr = "SELECT NULL, "
        MySQLStr = MySQLStr & "SC010300.SC01001 AS ID, "
        MySQLStr = MySQLStr & "SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
        MySQLStr = MySQLStr & "Round(ISNULL(t2.SC39005, 0) / " & Replace(CStr(Declarations.CurrencyValue), ",", ".") & ",2) AS Price, "
        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppID, "
        MySQLStr = MySQLStr & "View_1.txt AS UnitName, "
        MySQLStr = MySQLStr & "SC010300.SC01042 AS TotalQty, "
        MySQLStr = MySQLStr & "ISNULL(t1.WhQty, 0) AS WhQty, "
        MySQLStr = MySQLStr & "ISNULL(t1.WhAvl, 0) AS WhAvl, "
        MySQLStr = MySQLStr & "SC010300.SC01058 AS SuppCode, "
        MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') AS SuppName, "
        '---неликвиды
        MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(SC01" & Declarations.CompanyID & "00.SC01058)) = 'FIN001_U' OR "
        MySQLStr = MySQLStr & "Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'неликвид' THEN '+' ELSE '' END AS Dead, "

        MySQLStr = MySQLStr & "LTRIM(RTRIM(ISNULL(tbl_ItemCard" & Declarations.CompanyID & "00.ManufacturerItemCode, ''))) AS ManufacturerItemCode, "
        '---акция
        MySQLStr = MySQLStr & "CASE WHEN Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'акция' THEN '+' ELSE '' END AS Action, "
        MySQLStr = MySQLStr & "ISNULL(View_20.SC04004, N'') AS FullName "
        MySQLStr = MySQLStr & "FROM SC010300 WITH (NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC09" & Declarations.CompanyID & "00 AS SC090300_1 WITH (NOLOCK) ) AS View_1 ON "
        MySQLStr = MySQLStr & "SC01" & Declarations.CompanyID & "00.SC01135 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC04001, SC04004 "
        MySQLStr = MySQLStr & "FROM SC040300 "
        MySQLStr = MySQLStr & "WHERE (SC04002 = N'ST') AND (SC04003 = N'RUS')) AS View_20 ON SC010300.SC01001 = View_20.SC04001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT ScalaCode, ActionOrSales "
        MySQLStr = MySQLStr & "FROM tbl_ActionsAndSales "
        MySQLStr = MySQLStr & "WHERE (ActionFinished = 0) AND (DateStart <= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0)) "
        MySQLStr = MySQLStr & "AND (DateFinish >= DATEADD(day, DATEDIFF(day, 0, GETDATE()), 0))) AS View_4 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = View_4.ScalaCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL01" & Declarations.CompanyID & "00 ON SC01" & Declarations.CompanyID & "00.SC01058 = PL01" & Declarations.CompanyID & "00.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_ItemCard" & Declarations.CompanyID & "00 ON SC01" & Declarations.CompanyID & "00.SC01001 = tbl_ItemCard" & Declarations.CompanyID & "00.SC01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC03001, SC03003 AS WhQty, "
        MySQLStr = MySQLStr & "SC03003 - SC03004 - SC03005 AS WhAvl "
        MySQLStr = MySQLStr & "FROM SC03" & Declarations.CompanyID & "00 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) AS t1 ON "
        MySQLStr = MySQLStr & "SC01" & Declarations.CompanyID & "00.SC01001 = t1.SC03001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC39001, SC39005 "
        MySQLStr = MySQLStr & "FROM SC39" & Declarations.CompanyID & "00 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC39002 = N'00')) AS t2 ON "
        MySQLStr = MySQLStr & "SC01" & Declarations.CompanyID & "00.SC01001 = t2.SC39001 "
        MySQLStr = MySQLStr & "WHERE (SC01" & Declarations.CompanyID & "00.SC01001 <> N'00000000') "
        MySQLStr = MySQLStr & "AND (tbl_ItemCard" & Declarations.CompanyID & "00.IsBlocked = N'0') "
        MySQLStr = MySQLStr & "AND (Ltrim(Rtrim(SC010300.SC01066)) <> N'8')"
        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & Trim(TextBox1.Text) & "') "
        End If
        If MyParam = 1 Then '---выводятся только доступные продукты
            MySQLStr = MySQLStr & " AND (WhAvl > 0) "
        End If
        If MyParamAct = 1 Then '---выводятся только акционные / распродажные продукты
            MySQLStr = MySQLStr & " AND ((CASE WHEN Ltrim(Rtrim(SC01" & Declarations.CompanyID & "00.SC01058)) = 'FIN001_U' OR Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'неликвид' THEN '1' ELSE '' END = '1') "
            MySQLStr = MySQLStr & "OR (CASE WHEN Ltrim(Rtrim(ISNULL(View_4.ActionOrSales, ''))) = 'акция' THEN '1' ELSE '' END = '1')) "
        End If
        MySQLStr = MySQLStr & "ORDER BY dbo.SC010300.SC01001  "

        DataGridView1.RowTemplate.MinimumHeight = 35

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Кар тин ка"
        DataGridView1.Columns(0).Width = 35
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "ID продукта"
        DataGridView1.Columns(1).Width = 110
        DataGridView1.Columns(2).HeaderText = "Имя продукта"
        DataGridView1.Columns(2).Width = 300
        DataGridView1.Columns(3).HeaderText = "Прайс"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "Код тов постав"
        DataGridView1.Columns(4).Width = 110
        DataGridView1.Columns(5).HeaderText = "Ед изм"
        DataGridView1.Columns(5).Width = 40
        DataGridView1.Columns(6).HeaderText = "Всего на складах"
        DataGridView1.Columns(6).Width = 115
        DataGridView1.Columns(7).HeaderText = "Ост" + " " + Declarations.WHNum + " " + WhName
        DataGridView1.Columns(7).Width = 115
        DataGridView1.Columns(8).HeaderText = "Дост" + " " + Declarations.WHNum + " " + WhName
        DataGridView1.Columns(8).Width = 115
        DataGridView1.Columns(9).HeaderText = "Поставщик ID"
        DataGridView1.Columns(9).Width = 70
        DataGridView1.Columns(10).HeaderText = "Поставщик"
        DataGridView1.Columns(10).Width = 200
        DataGridView1.Columns(11).HeaderText = "Неликвид"
        DataGridView1.Columns(11).Width = 70
        DataGridView1.Columns(12).HeaderText = "Код тов произв"
        DataGridView1.Columns(12).Width = 150
        DataGridView1.Columns(13).HeaderText = "Акция"
        DataGridView1.Columns(13).Width = 70
        DataGridView1.Columns(14).HeaderText = "Полное название"
        DataGridView1.Columns(14).Width = 800

    End Function

    Private Function OrderPreparation()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование строк заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLstr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet
        Dim DeliverySumm As Double               'стоимость доставки
        Dim OrderSum As Double                   'сумма заказа
        Dim PropKoeff As Double                  'коэффициент - на сколько снижаем стоимость строки для учета доставки

        '---стоимость доставки
        MySQLstr = "SELECT DeliverySum "
        MySQLstr = MySQLstr & "FROM tbl_SW4SalesHdr_AddInfo "
        MySQLstr = MySQLstr & "WHERE (OrderID = N'" & Trim(Label6.Text) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLstr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            DeliverySumm = 0
        Else
            Declarations.MyRec.MoveFirst()
            DeliverySumm = Declarations.MyRec.Fields("DeliverySum").Value
        End If
        trycloseMyRec()
        '---Стоимость доставки в валюте заказа
        DeliverySumm = DeliverySumm / Declarations.CurrencyValue

        '---стоимость заказа
        MySQLstr = "SELECT View_0.CorrSum AS OrderSum "
        MySQLstr = MySQLstr & "FROM  tbl_OR01" & Declarations.CompanyID & "00 INNER JOIN "
        MySQLstr = MySQLstr & "(SELECT TOP (100) PERCENT tbl_OR01" & Declarations.CompanyID & "00_1.OR01001, "
        MySQLstr = MySQLstr & "SUM(ROUND(tbl_OR03" & Declarations.CompanyID & "00.OR03008 - tbl_OR03" & Declarations.CompanyID & "00.OR03008 * CONVERT(float, REPLACE(tbl_OR03" & Declarations.CompanyID & "00.OR03018, ',', '.')) / 100 / tbl_OR030300.OR03022, 2) "
        MySQLstr = MySQLstr & "* tbl_OR03" & Declarations.CompanyID & "00.OR03011) AS CorrSum "
        MySQLstr = MySQLstr & "FROM tbl_OR01" & Declarations.CompanyID & "00 AS tbl_OR010300_1 INNER JOIN "
        MySQLstr = MySQLstr & "tbl_OR03" & Declarations.CompanyID & "00 ON tbl_OR01" & Declarations.CompanyID & "00_1.OR01001 = tbl_OR03" & Declarations.CompanyID & "00.OR03001 "
        MySQLstr = MySQLstr & "WHERE (tbl_OR03" & Declarations.CompanyID & "00.OR03003 = N'000000') "
        MySQLstr = MySQLstr & "GROUP BY tbl_OR01" & Declarations.CompanyID & "00_1.OR01001 "
        MySQLstr = MySQLstr & "ORDER BY tbl_OR01" & Declarations.CompanyID & "00_1.OR01001) AS View_0 ON tbl_OR01" & Declarations.CompanyID & "00.OR01001 = View_0.OR01001 "
        MySQLstr = MySQLstr & "WHERE (tbl_OR01" & Declarations.CompanyID & "00.OR01001 = N'" & Trim(Label6.Text) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLstr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            OrderSum = 0
        Else
            Declarations.MyRec.MoveFirst()
            OrderSum = Declarations.MyRec.Fields("OrderSum").Value
        End If
        trycloseMyRec()

        '---коэфф снижения стоимости
        If OrderSum = 0 Then
            PropKoeff = 0.999999
        Else
            PropKoeff = DeliverySumm / OrderSum
        End If

        'MySQLstr = "SELECT  tbl_WEB_Pictures.PictureSmall, tbl_OR030300.OR03002 AS StrNum, "
        'MySQLstr = MySQLstr & "tbl_OR030300.OR03005 AS Code, "
        'MySQLstr = MySQLstr & "tbl_OR030300.OR03006 + tbl_OR030300.OR03007 AS Name, "
        'MySQLstr = MySQLstr & "ROUND(ISNULL(t1.SC39005, 0) / " & Replace(Declarations.CurrencyValue, ",", ".") & ", 2) AS Price, "
        'MySQLstr = MySQLstr & "ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & ", 2)AS PriCost, "
        'MySQLstr = MySQLstr & "ROUND(tbl_OR030300.OR03008, 2) AS Sum, "
        'MySQLstr = MySQLstr & "ROUND(tbl_OR030300.OR03011, 3) AS Qty, "
        'MySQLstr = MySQLstr & "ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2) AS PriCostSum, "
        'MySQLstr = MySQLstr & "ROUND(ROUND(tbl_OR030300.OR03008, 2) * tbl_OR030300.OR03011, 2) AS SumSum, "
        'MySQLstr = MySQLstr & "REPLACE(tbl_OR030300.OR03018, '.', ',') AS Disc, "
        'MySQLstr = MySQLstr & "ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) AS SumSumDisc, "
        'MySQLstr = MySQLstr & "ROUND(CASE WHEN ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) = 0 THEN 0 ELSE (ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) - ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2)) / ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) * 100 END, 2) AS Margin, "
        'MySQLstr = MySQLstr & "ROUND(CASE WHEN ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) = 0 THEN 0 ELSE (ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) - ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2)) / ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) * 100 END, 2) AS DelMargin, "
        'MySQLstr = MySQLstr & "CONVERT(float, CASE WHEN LTrim(RTrim(Isnull(SY290300.SY29003, '0'))) = '' THEN '0' ELSE Isnull(SY290300.SY29003, '0') END) AS VAT, "
        'MySQLstr = MySQLstr & "ROUND(ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28,8),REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) * (100 + CONVERT(float, CASE WHEN LTrim(RTrim(Isnull(SY290300.SY29003, '0'))) = '' THEN '0' ELSE Isnull(SY290300.SY29003, '0') END)) / 100, 2) AS SumSumDiscVAT, "
        'MySQLstr = MySQLstr & "ISNULL(PL010300.PL01001, N'') AS SupplierID, "
        'MySQLstr = MySQLstr & "ISNULL(PL010300.PL01002, N'') AS SupplierName, "
        'MySQLstr = MySQLstr & "CASE WHEN ISNULL(dbo.SC010300.SC01055, 0) = 0 THEN '' ELSE '+' END AS InPrice, "
        'MySQLstr = MySQLstr & "CASE WHEN ISNULL(tbl_ItemCard0300.IsBlocked, 1) = 1 THEN '+' ELSE '' END AS IsBlocked, "
        'MySQLstr = MySQLstr & "ISNULL(tbl_SupplierCard0300.Purchaser, N'') AS PurchaserCode, "
        'MySQLstr = MySQLstr & "ISNULL(View_1.SYPD003, N'') AS PurchaserName, "
        'MySQLstr = MySQLstr & "SC010300.SC01037 AS ItemGroup, "
        ''MySQLstr = MySQLstr & "tbl_OR030300.OR03037 "
        'MySQLstr = MySQLstr & "tbl_OR030300.WeekQTY, "
        'MySQLstr = MySQLstr & "ISNULL(SC010300.SC01060, ISNULL(tbl_OR030300.SuppItemCode, N'')) AS SuppItemCode "
        'MySQLstr = MySQLstr & "FROM tbl_SupplierCard0300 WITH (NOLOCK) LEFT OUTER JOIN "
        'MySQLstr = MySQLstr & "(SELECT SYPD001, SYPD003 "
        'MySQLstr = MySQLstr & "FROM SYPD0300 WITH (NOLOCK) "
        'MySQLstr = MySQLstr & "WHERE (SYPD002 = N'RUS')) AS View_1 ON "
        'MySQLstr = MySQLstr & "UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) RIGHT OUTER JOIN "
        'MySQLstr = MySQLstr & "tbl_ItemCard0300 RIGHT OUTER JOIN "
        'MySQLstr = MySQLstr & "SC010300 WITH (NOLOCK) LEFT OUTER JOIN "
        'MySQLstr = MySQLstr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        'MySQLstr = MySQLstr & "(SELECT SC03001, SC03057 "
        'MySQLstr = MySQLstr & "FROM SC030300 WITH (NOLOCK) "
        'MySQLstr = MySQLstr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) AS View_2 ON SC010300.SC01001 = View_2.SC03001 ON "
        'MySQLstr = MySQLstr & "tbl_ItemCard0300.SC01001 = SC010300.SC01001 RIGHT OUTER JOIN "
        'MySQLstr = MySQLstr & "SY290300 RIGHT OUTER JOIN "
        'MySQLstr = MySQLstr & "tbl_WEB_Pictures RIGHT OUTER JOIN "
        'MySQLstr = MySQLstr & "tbl_OR030300 WITH (NOLOCK) ON tbl_WEB_Pictures.ScalaItemCode = tbl_OR030300.OR03005 ON SY290300.SY29001 = tbl_OR030300.OR03061 ON "
        'MySQLstr = MySQLstr & "SC010300.SC01001 = tbl_OR030300.OR03005 LEFT OUTER JOIN "
        'MySQLstr = MySQLstr & "(SELECT SC39001, SC39005 "
        'MySQLstr = MySQLstr & "FROM SC390300 WITH (NOLOCK) "
        'MySQLstr = MySQLstr & "WHERE (SC39002 = N'00')) AS t1 ON tbl_OR030300.OR03005 = t1.SC39001 ON "
        'MySQLstr = MySQLstr & "tbl_SupplierCard0300.PL01001 = PL010300.PL01001 "
        'MySQLstr = MySQLstr & "WHERE (tbl_OR030300.OR03001 = N'" & Declarations.MyOrderNum & "') "
        'MySQLstr = MySQLstr & "AND (tbl_OR030300.OR03003 = '000000') "
        'MySQLstr = MySQLstr & "ORDER BY StrNum "

        MySQLstr = "SELECT tbl_WEB_Pictures.PictureSmall, "
        MySQLstr = MySQLstr & "tbl_OR030300.OR03002 AS StrNum, "
        MySQLstr = MySQLstr & "tbl_OR030300.OR03005 AS Code, "
        MySQLstr = MySQLstr & "tbl_OR030300.OR03006 + tbl_OR030300.OR03007 AS Name, "
        MySQLstr = MySQLstr & "ROUND(ISNULL(t1.SC39005, 0) / " & Replace(Declarations.CurrencyValue, ",", ".") & ", 2) AS Price, "
        MySQLstr = MySQLstr & "ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & ", 2) AS PriCost, "
        MySQLstr = MySQLstr & "ROUND(tbl_OR030300.OR03008, 2) AS Sum, "
        MySQLstr = MySQLstr & "ROUND(tbl_OR030300.OR03011, 3) AS Qty, "
        MySQLstr = MySQLstr & "ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2) AS PriCostSum, "
        MySQLstr = MySQLstr & "ROUND(ROUND(tbl_OR030300.OR03008, 2) * tbl_OR030300.OR03011, 2) AS SumSum, "
        MySQLstr = MySQLstr & "REPLACE(tbl_OR030300.OR03018, '.', ',') AS Disc, "
        MySQLstr = MySQLstr & "ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) AS SumSumDisc, "
        MySQLstr = MySQLstr & "ROUND(CASE WHEN ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) = 0 THEN 0 ELSE (ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) - ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2)) / ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) * 100 END, 2) AS Margin, "
        MySQLstr = MySQLstr & "ROUND(CASE WHEN ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) = 0 THEN 0 ELSE (ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) - ROUND(CASE ISNULL(SC010300.SC01053, 0) WHEN 0 THEN tbl_OR030300.OR03009 ELSE SC010300.SC01053 END / " & Replace(Declarations.CurrencyValue, ",", ".") & " * tbl_OR030300.OR03011, 2)) / ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011 * (1 - " & Replace(CStr(PropKoeff), ",", ".") & "), 2) * 100 END, 2) AS DelMargin, "
        MySQLstr = MySQLstr & "CONVERT(float, CASE WHEN LTrim(RTrim(Isnull(SY290300.SY29003, '0'))) = '' THEN '0' ELSE Isnull(SY290300.SY29003, '0') END) AS VAT, "
        MySQLstr = MySQLstr & "ROUND(ROUND(ROUND(tbl_OR030300.OR03008 - tbl_OR030300.OR03008 * CONVERT(numeric(28, 8), REPLACE(tbl_OR030300.OR03018, ',', '.')) / 100, 2) * tbl_OR030300.OR03011, 2) * (100 + CONVERT(float, CASE WHEN LTrim(RTrim(Isnull(SY290300.SY29003, '0'))) = '' THEN '0' ELSE Isnull(SY290300.SY29003, '0') END)) / 100, 2) AS SumSumDiscVAT, "
        MySQLstr = MySQLstr & "ISNULL(PL010300_1.PL01001, ISNULL(tbl_OR030300.SuppCode, N'')) AS SupplierID, "
        MySQLstr = MySQLstr & "ISNULL(PL010300_1.PL01002, ISNULL(PL010300.PL01002, ISNULL(tbl_OR030300.SuppName, N''))) AS SupplierName, "
        MySQLstr = MySQLstr & "CASE WHEN ISNULL(dbo.SC010300.SC01055, 0) = 0 THEN '' ELSE '+' END AS InPrice, "
        MySQLstr = MySQLstr & "CASE WHEN ISNULL(tbl_ItemCard0300.IsBlocked, 1) = 1 THEN '+' ELSE '' END AS IsBlocked, "
        MySQLstr = MySQLstr & "ISNULL(tbl_SupplierCard0300.Purchaser, N'') AS PurchaserCode, "
        MySQLstr = MySQLstr & "ISNULL(View_1.SYPD003, N'') AS PurchaserName, "
        MySQLstr = MySQLstr & "SC010300.SC01037 AS ItemGroup, "
        MySQLstr = MySQLstr & "tbl_OR030300.WeekQTY, "
        MySQLstr = MySQLstr & "ISNULL(SC010300.SC01060, ISNULL(tbl_OR030300.SuppItemCode, N'')) AS SuppItemCode "
        MySQLstr = MySQLstr & "FROM (SELECT SC39001, SC39005 "
        MySQLstr = MySQLstr & "FROM SC390300 WITH (NOLOCK) "
        MySQLstr = MySQLstr & "WHERE (SC39002 = N'00')) AS t1 RIGHT OUTER JOIN "
        MySQLstr = MySQLstr & "SY290300 RIGHT OUTER JOIN "
        MySQLstr = MySQLstr & "tbl_WEB_Pictures RIGHT OUTER JOIN "
        MySQLstr = MySQLstr & "PL010300 RIGHT OUTER JOIN "
        MySQLstr = MySQLstr & "tbl_OR030300 WITH (NOLOCK) ON PL010300.PL01001 = tbl_OR030300.SuppCode ON "
        MySQLstr = MySQLstr & "tbl_WEB_Pictures.ScalaItemCode = tbl_OR030300.OR03005 ON "
        MySQLstr = MySQLstr & "SY290300.SY29001 = tbl_OR030300.OR03061 LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "tbl_ItemCard0300 RIGHT OUTER JOIN "
        MySQLstr = MySQLstr & "SC010300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "PL010300 AS PL010300_1 ON SC010300.SC01058 = PL010300_1.PL01001 LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "(SELECT SC03001, SC03057 "
        MySQLstr = MySQLstr & "FROM SC030300 WITH (NOLOCK) "
        MySQLstr = MySQLstr & "WHERE (SC03002 = N'" & Declarations.WHNum & "')) AS View_2 ON "
        MySQLstr = MySQLstr & "SC010300.SC01001 = View_2.SC03001 ON "
        MySQLstr = MySQLstr & "tbl_ItemCard0300.SC01001 = SC010300.SC01001 ON "
        MySQLstr = MySQLstr & "tbl_OR030300.OR03005 = SC010300.SC01001 ON t1.SC39001 = tbl_OR030300.OR03005 LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "tbl_SupplierCard0300 WITH (NOLOCK) LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "(SELECT SYPD001, SYPD003 "
        MySQLstr = MySQLstr & "FROM SYPD0300 WITH (NOLOCK) "
        MySQLstr = MySQLstr & "WHERE (SYPD002 = N'RUS')) AS View_1 ON UPPER(tbl_SupplierCard0300.Purchaser) = UPPER(View_1.SYPD001) ON "
        MySQLstr = MySQLstr & "PL010300_1.PL01001 = tbl_SupplierCard0300.PL01001 "
        MySQLstr = MySQLstr & "WHERE (tbl_OR030300.OR03001 = N'" & Declarations.MyOrderNum & "') AND "
        MySQLstr = MySQLstr & "(tbl_OR030300.OR03003 = '000000') "
        MySQLstr = MySQLstr & "ORDER BY StrNum "

        DataGridView1.RowTemplate.MinimumHeight = 35

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLstr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView2.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView2.Columns(0).HeaderText = "Кар тин ка"
        DataGridView2.Columns(0).Width = 35
        DataGridView2.Columns(1).HeaderText = "N строки"
        DataGridView2.Columns(1).Width = 50
        DataGridView2.Columns(2).HeaderText = "ID продукта"
        DataGridView2.Columns(2).Width = 80
        DataGridView2.Columns(3).HeaderText = "Имя продукта"
        DataGridView2.Columns(3).Width = 200
        DataGridView2.Columns(4).HeaderText = "Прайс за 1"
        DataGridView2.Columns(4).Width = 60
        DataGridView2.Columns(5).HeaderText = "с/стоим за 1"
        DataGridView2.Columns(5).Width = 60
        DataGridView2.Columns(6).HeaderText = "Цена за 1"
        DataGridView2.Columns(6).Width = 60
        DataGridView2.Columns(7).HeaderText = "Кол-во"
        DataGridView2.Columns(7).Width = 60
        DataGridView2.Columns(8).HeaderText = "с/стоим"
        DataGridView2.Columns(8).Width = 60
        DataGridView2.Columns(9).HeaderText = "Стоим"
        DataGridView2.Columns(9).Width = 60
        DataGridView2.Columns(10).HeaderText = "Скидка %"
        DataGridView2.Columns(10).Width = 60
        DataGridView2.Columns(11).HeaderText = "Стоим со скидк"
        DataGridView2.Columns(11).Width = 60
        DataGridView2.Columns(12).HeaderText = "Маржа"
        DataGridView2.Columns(12).Width = 60
        DataGridView2.Columns(13).HeaderText = "Маржа с дост."
        DataGridView2.Columns(13).Width = 60
        DataGridView2.Columns(14).HeaderText = "НДС (%)"
        DataGridView2.Columns(14).Width = 60
        DataGridView2.Columns(15).HeaderText = "Стоим с НДС"
        DataGridView2.Columns(15).Width = 70
        DataGridView2.Columns(16).HeaderText = "Поставщик ID"
        DataGridView2.Columns(16).Width = 70
        DataGridView2.Columns(17).HeaderText = "Поставщик"
        DataGridView2.Columns(17).Width = 200
        DataGridView2.Columns(18).HeaderText = "Есть прайс на закупку"
        DataGridView2.Columns(18).Width = 50
        DataGridView2.Columns(19).HeaderText = "Блокирован"
        DataGridView2.Columns(19).Width = 48
        DataGridView2.Columns(20).HeaderText = "Закупщик ID"
        DataGridView2.Columns(20).Width = 58
        DataGridView2.Columns(21).HeaderText = "Закупщик"
        DataGridView2.Columns(21).Width = 150
        DataGridView2.Columns(22).HeaderText = "Гр"
        DataGridView2.Columns(22).Width = 30
        DataGridView2.Columns(23).HeaderText = "Срок поставки (нед.)"
        DataGridView2.Columns(23).DefaultCellStyle.Format = "###,##0.##"
        DataGridView2.Columns(23).Width = 70
        DataGridView2.Columns(24).HeaderText = "Код товара поставщика"
        DataGridView2.Columns(24).Width = 80

        TotalRecount()
    End Function

    Private Sub ChangeButtonsStatus()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена состояния кнопки "найти в продуктах" в зависимости от того, какой товар подсвечен
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button15.Enabled = False
            Button22.Enabled = False
        Else
            Button15.Enabled = True
            Button22.Enabled = True
        End If

        If DataGridView1.SelectedRows.Count = 0 Then
            Button5.Enabled = False
            Button7.Enabled = False
        Else
            Button5.Enabled = True
            Button7.Enabled = True
        End If
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// копирование содержимого ячейки в буфер при начале редактирования
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        My.Computer.Clipboard.Clear()
        My.Computer.Clipboard.SetText(DataGridView1.CurrentCell.Value)
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выделения неликвидов и акций
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(13).Value.ToString = "+" Then
            row.DefaultCellStyle.Font = New Font(row.InheritedStyle.Font, FontStyle.Bold)
            row.DefaultCellStyle.ForeColor = Color.DarkGreen
        End If
        If row.Cells(11).Value.ToString = "+" Then
            row.DefaultCellStyle.Font = New Font(row.InheritedStyle.Font, FontStyle.Bold)
            row.DefaultCellStyle.ForeColor = Color.Red
        End If

    End Sub

    Private Sub DataGridView2_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выделение заблокированных продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView2.Rows(e.RowIndex)
        '---Заблокированные запасы
        If Trim(row.Cells(19).Value.ToString) = "+" Then
            row.DefaultCellStyle.BackColor = Color.Red
        End If
        '---Нет прайс - листа на закупку
        If Trim(row.Cells(18).Value.ToString) = "" Then
            row.DefaultCellStyle.BackColor = Color.Red
        End If
        '---Нет поставщика
        If Trim(row.Cells(17).Value.ToString) = "" Then
            row.DefaultCellStyle.BackColor = Color.Red
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком поставщиков
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "OrderLines"
        MySupplierSelect.ShowDialog()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Обновление списка продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        TextBox2.Text = ""
        TextBox3.Text = ""
        RefreshProductList()
    End Sub

    Public Function RefreshProductList()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// ф-ция обновляет список продуктов Электроскандии для последующего выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If Trim(Button8.Text) = "Только дост" Then
            If Button24.Text = "Только акции" Then
                DataPreparation(0, 0)
            Else
                DataPreparation(0, 1)
            End If
        Else
            If Button24.Text = "Только акции" Then
                DataPreparation(1, 0)
            Else
                DataPreparation(1, 1)
            End If
        End If
        ChangeButtonsStatus()
    End Function

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код поставщика - Обновляем список продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox1.Modified = True Then
            TextBox2.Text = ""
            TextBox3.Text = ""
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            RefreshProductList()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код поставщика - находим и подписываем его название
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If TextBox1.Modified = True Then
            If Trim(TextBox1.Text) = "" Then
                Label3.Text = ""
            Else
                MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH (NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(MyOrderLines.TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MsgBox("Вы ввели неверный код поставщика. Введите корректный или воспользуйтесь поиском.", vbCritical, "Внимание!")
                    Label3.Text = ""
                    e.Cancel = True
                    Exit Sub
                Else
                    Label3.Text = Declarations.MyRec.Fields("PL01002").Value & " " & Declarations.MyRec.Fields("PL01003").Value
                    trycloseMyRec()
                End If
            End If
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// только доступные для заказа товары будут отображены
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If Button8.Text = "Только дост" Then '---оставляем только доступные к заказу продукты
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Button24.Text = "Все прод." Then
                DataPreparation(1, 1)
            Else
                DataPreparation(1, 0)
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button8.Text = "Все прод."
        Else                                             '---показываем все продукты
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Button24.Text = "Все прод." Then
                DataPreparation(0, 1)
            Else
                DataPreparation(0, 0)
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button8.Text = "Только дост"
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего по критерию запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего по критерию поставщика
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        System.Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                        System.Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
            Next i
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсвечивание всех подходящих по критерию запасов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(12, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Снять выдел."
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Подсветить все"
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "Подсветить все"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию запасов в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            MyItemSelectList = New ItemSelectList
            MyItemSelectList.MySrcWin = "OrderLines"
            MyItemSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Открытие окна с детальной информацией по складам
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyShowWHDetails = New ShowWHDetails
        MyShowWHDetails.MyItem = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        MyShowWHDetails.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком запасов, альтернативных данному
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyALTItems = New ALTItems
        MyALTItems.MyItem = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString())
        MyALTItems.MySrcWin = "OrderLines"
        MyALTItems.ShowDialog()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Находим продукт из заказа в списке продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(1, i).Value.ToString) = Trim(Me.DataGridView2.SelectedRows.Item(0).Cells(2).Value.ToString()) Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                System.Windows.Forms.Cursor.Current = Cursors.Default
                Exit Sub
            End If
        Next
        System.Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Function TotalRecount()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перерасчет итоговых показателей заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MySum As Double
        Dim MyCost As Double
        Dim MyMargin As Double
        Dim MyVATSum As Double
        Dim MyProm As Double
        Dim i As Integer
        Dim MyDelSum As Double
        Dim MyDelMargin As Double

        MySum = 0
        MyCost = 0
        MyVATSum = 0
        For i = 0 To DataGridView2.Rows.Count - 1
            MySum = MySum + CDbl(DataGridView2.Item(11, i).Value.ToString)
            MyCost = MyCost + CDbl(DataGridView2.Item(8, i).Value.ToString)
            'MyProm = (CDbl(DataGridView2.Item(5, i).Value.ToString) - CDbl(DataGridView2.Item(5, i).Value.ToString) * CDbl(DataGridView2.Item(9, i).Value.ToString) / 100) * CDbl(DataGridView2.Item(6, i).Value.ToString) * CDbl(DataGridView2.Item(13, i).Value.ToString) / 100
            MyProm = CDbl(DataGridView2.Item(11, i).Value.ToString) * CDbl(DataGridView2.Item(14, i).Value.ToString) / 100
            MyVATSum = MyVATSum + MyProm
        Next i

        If MySum = 0 Then
            MyMargin = 0
        Else
            MyMargin = Math.Round((MySum - MyCost) / MySum * 100, 3)
        End If
        Label12.Text = CStr(Math.Round(MyCost, 2))
        Label13.Text = CStr(Math.Round(MySum, 2))
        Label18.Text = CStr(MyMargin)
        Label16.Text = CStr(Math.Round(MySum + MyVATSum, 2))
        '----------Сумма доставки и итоговые суммы с учетом стоимости доставки
        '---Сумма доставки
        MySQLStr = "SELECT DeliverySum "
        MySQLStr = MySQLStr & "FROM tbl_SW4SalesHdr_AddInfo "
        MySQLStr = MySQLStr & "WHERE (OrderID = N'" & Trim(Label6.Text) & "')"
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            Label22.Text = 0
            MyDelSum = 0
        Else
            Declarations.MyRec.MoveFirst()
            Label22.Text = CStr(Declarations.MyRec.Fields("DeliverySum").Value)
            MyDelSum = Declarations.MyRec.Fields("DeliverySum").Value
        End If
        trycloseMyRec()
        '--стоимость без доставки
        Label24.Text = CStr(Math.Round(MySum - (MyDelSum / Declarations.CurrencyValue), 2))
        '--Итоговая маржа (с доставкой)
        If MySum = 0 Then
            MyDelMargin = 0
        Else
            MyDelMargin = Math.Round((MySum - (MyDelSum / Declarations.CurrencyValue) - MyCost) / Math.Abs(MySum - (MyDelSum / Declarations.CurrencyValue)) * 100, 3)
        End If
        Label28.Text = CStr(MyDelMargin)
    End Function

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            ExportToLO()
        Else
            ExportToExcel()
        End If
        

    End Sub

    Private Sub ExportToExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---заголовки
        MyWRKBook.ActiveSheet.Range("B1") = "Предложение о поставке номер " & Label6.Text & "   Валюта " & Declarations.CurrencyName & "   Курс " & Declarations.CurrencyValue
        MyWRKBook.ActiveSheet.Range("A3") = "N строки"
        MyWRKBook.ActiveSheet.Range("B3") = "ID продукта"
        MyWRKBook.ActiveSheet.Range("C3") = "Имя продукта"
        MyWRKBook.ActiveSheet.Range("D3") = "Прайс за 1"
        MyWRKBook.ActiveSheet.Range("E3") = "с/стоим за 1"
        MyWRKBook.ActiveSheet.Range("F3") = "Цена за 1"
        MyWRKBook.ActiveSheet.Range("G3") = "Кол-во"
        MyWRKBook.ActiveSheet.Range("H3") = "с/стоим"
        MyWRKBook.ActiveSheet.Range("I3") = "Стоим"
        MyWRKBook.ActiveSheet.Range("J3") = "Скидка %"
        MyWRKBook.ActiveSheet.Range("K3") = "Стоим со скидк"
        MyWRKBook.ActiveSheet.Range("L3") = "Маржа"
        MyWRKBook.ActiveSheet.Range("M3") = "НДС (%)"
        MyWRKBook.ActiveSheet.Range("N3") = "Стоим. с НДС"
        MyWRKBook.ActiveSheet.Range("O3") = "Поставщик ID"
        MyWRKBook.ActiveSheet.Range("P3") = "Поставщик"
        MyWRKBook.ActiveSheet.Range("Q3") = "Есть прайс на закупку"
        MyWRKBook.ActiveSheet.Range("R3") = "Блокирован"
        MyWRKBook.ActiveSheet.Range("S3") = "Закупщик ID"
        MyWRKBook.ActiveSheet.Range("T3") = "Закупщик"

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("K:K").ColumnWidth = 15
        MyWRKBook.ActiveSheet.Columns("L:L").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("M:M").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("N:N").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("O:O").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("P:P").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("Q:Q").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("R:R").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("S:S").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("T:T").ColumnWidth = 30


        '---Форматирование заголовка
        MyWRKBook.ActiveSheet.Range("A3:T3").Select()
        MyWRKBook.ActiveSheet.Range("A3:T3").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A3:T3").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A3:T3").Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:T3").Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:T3").Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:T3").Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:T3").Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A3:T3").Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("B1").Select()
        MyWRKBook.ActiveSheet.Range("B1").Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A3:T3").Select()
        MyWRKBook.ActiveSheet.Range("A3:T3").Font.Bold = True

        '---таблица
        For i = 0 To DataGridView2.Rows.Count - 1
            MyWRKBook.ActiveSheet.Range("A" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("A" & CStr(i + 4)) = DataGridView2.Item(1, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 4)) = DataGridView2.Item(2, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("C" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("C" & CStr(i + 4)) = DataGridView2.Item(3, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("D" & CStr(i + 4)) = CDbl(DataGridView2.Item(4, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("E" & CStr(i + 4)) = CDbl(DataGridView2.Item(5, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("F" & CStr(i + 4)) = CDbl(DataGridView2.Item(6, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("G" & CStr(i + 4)) = CDbl(DataGridView2.Item(7, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("H" & CStr(i + 4)) = CDbl(DataGridView2.Item(8, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("I" & CStr(i + 4)) = CDbl(DataGridView2.Item(9, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("J" & CStr(i + 4)) = CDbl(DataGridView2.Item(10, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("K" & CStr(i + 4)) = CDbl(DataGridView2.Item(11, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("L" & CStr(i + 4)) = CDbl(DataGridView2.Item(12, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("M" & CStr(i + 4)) = CDbl(DataGridView2.Item(13, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("N" & CStr(i + 4)) = CDbl(DataGridView2.Item(14, i).Value.ToString)
            MyWRKBook.ActiveSheet.Range("O" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("O" & CStr(i + 4)) = DataGridView2.Item(15, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("P" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("P" & CStr(i + 4)) = "'" & DataGridView2.Item(16, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("Q" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("Q" & CStr(i + 4)) = "'" & DataGridView2.Item(17, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("R" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("R" & CStr(i + 4)) = "'" & DataGridView2.Item(18, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("S" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("S" & CStr(i + 4)) = "'" & DataGridView2.Item(20, i).Value.ToString
            MyWRKBook.ActiveSheet.Range("T" & CStr(i + 4)).NumberFormat = "@"
            MyWRKBook.ActiveSheet.Range("T" & CStr(i + 4)) = "'" & DataGridView2.Item(21, i).Value.ToString
        Next i
        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing
    End Sub

    Private Sub ExportToLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---ширина колонок
        oSheet.getColumns().getByName("A").Width = 2490
        oSheet.getColumns().getByName("B").Width = 2490
        oSheet.getColumns().getByName("C").Width = 7980
        oSheet.getColumns().getByName("D").Width = 2490
        oSheet.getColumns().getByName("E").Width = 2490
        oSheet.getColumns().getByName("F").Width = 2490
        oSheet.getColumns().getByName("G").Width = 2490
        oSheet.getColumns().getByName("H").Width = 2490
        oSheet.getColumns().getByName("I").Width = 2490
        oSheet.getColumns().getByName("J").Width = 2490
        oSheet.getColumns().getByName("K").Width = 3080
        oSheet.getColumns().getByName("L").Width = 2490
        oSheet.getColumns().getByName("M").Width = 2490
        oSheet.getColumns().getByName("N").Width = 2490
        oSheet.getColumns().getByName("O").Width = 2490
        oSheet.getColumns().getByName("P").Width = 6020
        oSheet.getColumns().getByName("Q").Width = 4060
        oSheet.getColumns().getByName("R").Width = 2490
        oSheet.getColumns().getByName("S").Width = 2490
        oSheet.getColumns().getByName("T").Width = 6020

        '-----колонки
        oSheet.getCellRangeByName("A1").String = "Предложение о поставке номер " & Label6.Text & "   Валюта " & Declarations.CurrencyName & "   Курс " & Declarations.CurrencyValue
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A1")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)

        oSheet.getCellRangeByName("A3").String = "N строки"
        oSheet.getCellRangeByName("B3").String = "ID продукта"
        oSheet.getCellRangeByName("C3").String = "Имя продукта"
        oSheet.getCellRangeByName("D3").String = "Прайс за 1"
        oSheet.getCellRangeByName("E3").String = "с/стоим за 1"
        oSheet.getCellRangeByName("F3").String = "Цена за 1"
        oSheet.getCellRangeByName("G3").String = "Кол-во"
        oSheet.getCellRangeByName("H3").String = "с/стоим"
        oSheet.getCellRangeByName("I3").String = "Стоим"
        oSheet.getCellRangeByName("J3").String = "Скидка %"
        oSheet.getCellRangeByName("K3").String = "Стоим со скидк"
        oSheet.getCellRangeByName("L3").String = "Маржа"
        oSheet.getCellRangeByName("M3").String = "НДС (%)"
        oSheet.getCellRangeByName("N3").String = "Стоим. с НДС"
        oSheet.getCellRangeByName("O3").String = "Поставщик ID"
        oSheet.getCellRangeByName("P3").String = "Поставщик"
        oSheet.getCellRangeByName("Q3").String = "Есть прайс на закупку"
        oSheet.getCellRangeByName("R3").String = "Блокирован"
        oSheet.getCellRangeByName("S3").String = "Закупщик ID"
        oSheet.getCellRangeByName("T3").String = "Закупщик"

        Dim i As Integer
        i = 3
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":T" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":T" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":T" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":T" & CStr(i)).CellBackColor = RGB(174, 249, 255)  '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":T" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("A" & CStr(i) & ":T" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":T" & CStr(i)).HoriJustify = 2
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":T" & CStr(i))

        '---таблица
        For i = 0 To DataGridView2.Rows.Count - 1
            oSheet.getCellRangeByName("A" & CStr(i + 4)).String = DataGridView2.Item(1, i).Value.ToString
            oSheet.getCellRangeByName("B" & CStr(i + 4)).String = DataGridView2.Item(2, i).Value.ToString
            oSheet.getCellRangeByName("C" & CStr(i + 4)).String = DataGridView2.Item(3, i).Value.ToString
            oSheet.getCellRangeByName("D" & CStr(i + 4)).Value = DataGridView2.Item(4, i).Value
            oSheet.getCellRangeByName("E" & CStr(i + 4)).Value = DataGridView2.Item(5, i).Value
            oSheet.getCellRangeByName("F" & CStr(i + 4)).Value = DataGridView2.Item(6, i).Value
            oSheet.getCellRangeByName("G" & CStr(i + 4)).Value = DataGridView2.Item(7, i).Value
            oSheet.getCellRangeByName("H" & CStr(i + 4)).Value = DataGridView2.Item(8, i).Value
            oSheet.getCellRangeByName("I" & CStr(i + 4)).Value = DataGridView2.Item(9, i).Value
            oSheet.getCellRangeByName("J" & CStr(i + 4)).Value = DataGridView2.Item(10, i).Value
            oSheet.getCellRangeByName("K" & CStr(i + 4)).Value = DataGridView2.Item(11, i).Value
            oSheet.getCellRangeByName("L" & CStr(i + 4)).Value = DataGridView2.Item(12, i).Value
            oSheet.getCellRangeByName("M" & CStr(i + 4)).Value = DataGridView2.Item(14, i).Value
            oSheet.getCellRangeByName("N" & CStr(i + 4)).Value = DataGridView2.Item(15, i).Value
            oSheet.getCellRangeByName("O" & CStr(i + 4)).String = DataGridView2.Item(16, i).Value.ToString
            oSheet.getCellRangeByName("P" & CStr(i + 4)).String = DataGridView2.Item(17, i).Value.ToString
            oSheet.getCellRangeByName("Q" & CStr(i + 4)).String = DataGridView2.Item(18, i).Value.ToString
            oSheet.getCellRangeByName("R" & CStr(i + 4)).String = DataGridView2.Item(19, i).Value.ToString
            oSheet.getCellRangeByName("S" & CStr(i + 4)).String = DataGridView2.Item(20, i).Value.ToString
            oSheet.getCellRangeByName("T" & CStr(i + 4)).String = DataGridView2.Item(21, i).Value.ToString
        Next i

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// По двойному щелчку добавляем запас в заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        AddItemToOrder()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// По щелчку добавляем запас в заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        AddItemToOrder()
    End Sub

    Private Sub AddItemToOrder()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура добавления запаса в заказ
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          'рабочая строка
        Dim MyOrder As String                           'Номер предложения о покупке
        Dim MyStr As String                             'Номер строки
        Dim MyName1 As String                           'Наименование запаса строка 1
        Dim MyName2 As String                           'Наименование запаса строка 2
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        Dim MyParam9 As ADODB.Parameter                 'передаваемый параметр номер 10
        Dim MyParam10 As ADODB.Parameter                'передаваемый параметр номер 11
        Dim MyParam11 As ADODB.Parameter                'передаваемый параметр номер 12

        Declarations.MySuccess = False
        MyAddToOrder = New AddToOrder
        MyAddToOrder.ShowDialog()
        If Declarations.MySuccess = False Then
            Exit Sub
        Else '---добавление строки в заказ
            MyOrder = Me.Label6.Text                    'Номер предложения
            MyStr = GetNewStrNum()                      'номер строки заказа
            MyName1 = Microsoft.VisualBasic.Left(Declarations.MyItemName, 25)       'Наименование запаса строка 1
            MyName2 = Microsoft.VisualBasic.Mid(Declarations.MyItemName, 26, 25)    'Наименование запаса строка 2

            cmd.ActiveConnection = Declarations.MyConn
            cmd.CommandText = "spp_SalesWorkplace4_AddOrder"
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            cmd.CommandTimeout = 300

            MyParam = cmd.CreateParameter("@Order", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
            cmd.Parameters.Append(MyParam)
            MyParam.Value = MyOrder

            MyParam1 = cmd.CreateParameter("@Str", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
            cmd.Parameters.Append(MyParam1)
            MyParam1.Value = MyStr

            MyParam2 = cmd.CreateParameter("@Code", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
            cmd.Parameters.Append(MyParam2)
            MyParam2.Value = Declarations.MyItemID

            MyParam3 = cmd.CreateParameter("@Name1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
            cmd.Parameters.Append(MyParam3)
            MyParam3.Value = MyName1

            MyParam4 = cmd.CreateParameter("@Name2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
            cmd.Parameters.Append(MyParam4)
            MyParam4.Value = MyName2

            MyParam5 = cmd.CreateParameter("@Cost", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam5)
            MyParam5.Value = Declarations.MySum

            MyParam6 = cmd.CreateParameter("@CostIntr", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam6)
            MyParam6.Value = Declarations.MySS

            MyParam7 = cmd.CreateParameter("@Qty", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam7)
            MyParam7.Value = Declarations.MyQty

            MyParam8 = cmd.CreateParameter("@Wh", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 2)
            cmd.Parameters.Append(MyParam8)
            MyParam8.Value = Declarations.WHNum

            MyParam9 = cmd.CreateParameter("@Unit", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam9)
            MyParam9.Value = Declarations.MyUOM

            MyParam10 = cmd.CreateParameter("@WeekQTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam10)
            MyParam10.Value = Declarations.WeekQTY

            MyParam11 = cmd.CreateParameter("@DelWeekQTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
            cmd.Parameters.Append(MyParam11)
            MyParam11.Value = Declarations.DelWeekQTY

            Try
                cmd.Execute()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            '---------добавление кода товара поставщика
            MySQLStr = "UPDATE tbl_OR030300 "
            MySQLStr = MySQLStr & "SET SuppItemCode = N'" & Trim(Declarations.MyItemSuppID) & "' "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
            MySQLStr = MySQLStr & "AND (OR03002 = N'" & MyStr & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---работа с датами поставки
            'If Declarations.DeliveryDateFlag = 0 Then '---Только для одной строки
            '    '---Обновление данных в строке
            '    MySQLStr = "UPDATE tbl_OR030300 "
            '    MySQLStr = MySQLStr & "SET OR03037 = CONVERT(DATETIME, '" & Format(Declarations.DeliveryDate, "dd/MM/yyyy") & "', 103) "
            '    MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
            '    MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
            '    InitMyConn(False)
            '    Declarations.MyConn.Execute(MySQLStr)
            '    '---Обновление данных в заголовке
            'Else   '---для всех строк заказа
            '    MySQLStr = "UPDATE tbl_OR030300 "
            '    MySQLStr = MySQLStr & "SET OR03037 = CONVERT(DATETIME, '" & Format(Declarations.DeliveryDate, "dd/MM/yyyy") & "', 103) "
            '    MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') "
            '    InitMyConn(False)
            '    Declarations.MyConn.Execute(MySQLStr)
            'End If
            'MySQLStr = "UPDATE tbl_OR030300 "
            'MySQLStr = MySQLStr & "SET WeekQTY = " & Replace(CStr(Declarations.WeekQTY), ",", ".") & " "
            'MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
            'MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
            'InitMyConn(False)
            'Declarations.MyConn.Execute(MySQLStr)


            'MySQLStr = "Update tbl_OR010300 "
            'MySQLStr = MySQLStr & "Set OR01016 = View_1.CC, "
            'MySQLStr = MySQLStr & "ReadyDate = View_1.CC "
            'MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
            'MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
            'MySQLStr = MySQLStr & "FROM tbl_OR030300 "
            'MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') "
            'MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON tbl_OR010300.OR01001 = View_1.OR03001 "
            'InitMyConn(False)
            'Declarations.MyConn.Execute(MySQLStr)

            '---------добавление кода поставщика и названия
            MySQLStr = "Update tbl_OR030300  "
            MySQLStr = MySQLStr & "Set SuppCode = N'" & Trim(Declarations.MySuppID) & "', "
            MySQLStr = MySQLStr & "SuppName = N'" & Trim(Declarations.MySuppName) & "' "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') AND "
            MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            OrderPreparation()
            TotalRecount()
        End If
    End Sub

    Private Function GetNewStrNum() As String
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение следующего номера строки в предложении
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyStrNum As String                          'следующий номер строки
        Dim MySQLStr As String                          'рабочая строка
        Dim MyRez As Double

        If DataGridView2.Rows.Count = 0 Then
            MyStrNum = Microsoft.VisualBasic.Right("000000" & "10", 6)
        Else
            MySQLStr = "SELECT MAX(CONVERT(float, OR03002)) / 10 AS STRNUM "
            MySQLStr = MySQLStr & "FROM tbl_OR030300 WITH (NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(Me.Label6.Text) & "')"
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MyStrNum = Microsoft.VisualBasic.Right("000000" & "10", 6)
            End If
            MyRez = Declarations.MyRec.Fields("STRNUM").Value
            trycloseMyRec()
            MyStrNum = Microsoft.VisualBasic.Right("000000" & CStr((MyRez + 1) * 10), 6)
        End If
        GetNewStrNum = MyStrNum
    End Function

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление выделенной строки из заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                              'рабочая срока

        If DataGridView2.SelectedRows.Count > 0 Then
            MySQLStr = "EXEC ScaDataDB.dbo.spp_SalesWorkplace4_DelOrder N'"
            MySQLStr = MySQLStr & Trim(Label6.Text) & "', N'"
            MySQLStr = MySQLStr & Trim(Me.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString()) & "' "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            OrderPreparation()
            TotalRecount()
        End If
    End Sub

    Private Sub DataGridView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// редактирование выделенной строки
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                          'рабочая строка
        Dim MyOrder As String                           'Номер предложения о покупке
        Dim MyStr As String                             'Номер строки
        Dim MyName1 As String                           'Наименование запаса строка 1
        Dim MyName2 As String                           'Наименование запаса строка 2
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        Dim MyParam9 As ADODB.Parameter                 'передаваемый параметр номер 10
        Dim MyParam10 As ADODB.Parameter                'передаваемый параметр номер 11

        If DataGridView2.SelectedRows.Count > 0 Then
            Declarations.MySuccess = False
            MyEditInOrder = New EditInOrder
            MyEditInOrder.MyItem = Trim(Me.DataGridView2.SelectedRows.Item(0).Cells(2).Value.ToString())
            MyEditInOrder.MyOrder = Trim(Label6.Text)
            MyStr = Trim(Me.DataGridView2.SelectedRows.Item(0).Cells(1).Value.ToString()) 'номер строки заказа
            MyEditInOrder.ShowDialog()
            If Declarations.MySuccess = False Then
                MyOrder = Me.Label6.Text                    'Номер предложения

                MySQLStr = "Update tbl_OR010300 "
                MySQLStr = MySQLStr & "Set OR01016 = View_1.CC, "
                MySQLStr = MySQLStr & "ReadyDate = View_1.CC "
                MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
                MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_OR030300 "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') "
                MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON tbl_OR010300.OR01001 = View_1.OR03001 "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                OrderPreparation()
                TotalRecount()
                For i As Integer = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2.Rows(i).Cells(1).Value.ToString = MyStr Then
                        DataGridView2.Rows(i).Selected = True
                    End If
                Next
                Exit Sub
            Else '---изменение строки в заказе
                MyOrder = Me.Label6.Text                    'Номер предложения

                MyName1 = Microsoft.VisualBasic.Left(Declarations.MyItemName, 25)       'Наименование запаса строка 1
                MyName2 = Microsoft.VisualBasic.Mid(Declarations.MyItemName, 26, 25)    'Наименование запаса строка 2

                cmd.ActiveConnection = Declarations.MyConn
                cmd.CommandText = "spp_SalesWorkplace4_EditOrder"
                cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmd.CommandTimeout = 300

                MyParam = cmd.CreateParameter("@Order", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
                cmd.Parameters.Append(MyParam)
                MyParam.Value = MyOrder

                MyParam1 = cmd.CreateParameter("@Str", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
                cmd.Parameters.Append(MyParam1)
                MyParam1.Value = MyStr

                MyParam2 = cmd.CreateParameter("@MyItemID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
                cmd.Parameters.Append(MyParam2)
                MyParam2.Value = Declarations.MyItemID

                MyParam3 = cmd.CreateParameter("@MyName1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
                cmd.Parameters.Append(MyParam3)
                MyParam3.Value = MyName1

                MyParam4 = cmd.CreateParameter("@MyName2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
                cmd.Parameters.Append(MyParam4)
                MyParam4.Value = MyName2

                MyParam5 = cmd.CreateParameter("@Cost", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam5)
                MyParam5.Value = Declarations.MySum

                MyParam6 = cmd.CreateParameter("@CostIntr", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam6)
                MyParam6.Value = Declarations.MySS

                MyParam7 = cmd.CreateParameter("@Qty", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam7)
                MyParam7.Value = Declarations.MyQty

                MyParam8 = cmd.CreateParameter("@Unit", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam8)
                MyParam8.Value = Declarations.MyUOM

                MyParam9 = cmd.CreateParameter("@Discount", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
                cmd.Parameters.Append(MyParam9)
                MyParam9.Value = Replace(Declarations.MyDiscount, ",", ".")

                MyParam10 = cmd.CreateParameter("@EditOrRecalc", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
                cmd.Parameters.Append(MyParam10)
                MyParam10.Value = 0

                Try
                    cmd.Execute()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                '---работа с датами доставки
                'If Declarations.DeliveryDateFlag = 0 Then '---Только для одной строки
                '    '---Обновление данных в строке
                '    MySQLStr = "UPDATE tbl_OR030300 "
                '    MySQLStr = MySQLStr & "SET OR03037 = CONVERT(DATETIME, '" & Format(Declarations.DeliveryDate, "dd/MM/yyyy") & "', 103) "
                '    MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
                '    MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
                '    InitMyConn(False)
                '    Declarations.MyConn.Execute(MySQLStr)
                '    '---Обновление данных в заголовке
                'Else   '---для всех строк заказа
                '    MySQLStr = "UPDATE tbl_OR030300 "
                '    MySQLStr = MySQLStr & "SET OR03037 = CONVERT(DATETIME, '" & Format(Declarations.DeliveryDate, "dd/MM/yyyy") & "', 103) "
                '    MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') "
                '    InitMyConn(False)
                '    Declarations.MyConn.Execute(MySQLStr)
                'End If

                MySQLStr = "UPDATE tbl_OR030300 "
                MySQLStr = MySQLStr & "SET WeekQTY = " & Replace(CStr(Declarations.WeekQTY), ",", ".") & ", "
                MySQLStr = MySQLStr & "DelWeekQTY = " & Replace(CStr(Declarations.DelWeekQTY), ",", ".") & " "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') AND "
                MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---------добавление кода товара поставщика
                MySQLStr = "UPDATE tbl_OR030300 "
                MySQLStr = MySQLStr & "SET SuppItemCode = N'" & Trim(Declarations.MyItemSuppID) & "' "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') "
                MySQLStr = MySQLStr & "AND (OR03002 = N'" & MyStr & "')"
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                'MySQLStr = "Update tbl_OR010300 "
                'MySQLStr = MySQLStr & "Set OR01016 = View_1.CC, "
                'MySQLStr = MySQLStr & "ReadyDate = View_1.CC "
                'MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
                'MySQLStr = MySQLStr & "(SELECT OR03001, MIN(OR03037) AS CC "
                'MySQLStr = MySQLStr & "FROM tbl_OR030300 "
                'MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Trim(MyOrder) & "') "
                'MySQLStr = MySQLStr & "GROUP BY OR03001) AS View_1 ON tbl_OR010300.OR01001 = View_1.OR03001 "
                'InitMyConn(False)
                'Declarations.MyConn.Execute(MySQLStr)

                '---------добавление кода поставщика и названия
                MySQLStr = "Update tbl_OR030300  "
                MySQLStr = MySQLStr & "Set SuppCode = N'" & Trim(Declarations.MySuppID) & "', "
                MySQLStr = MySQLStr & "SuppName = N'" & Trim(Declarations.MySuppName) & "' "
                MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & MyOrder & "') AND "
                MySQLStr = MySQLStr & "(OR03002 = N'" & MyStr & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                OrderPreparation()
                TotalRecount()
                For i As Integer = 0 To DataGridView2.Rows.Count - 1
                    If DataGridView2.Rows(i).Cells(1).Value.ToString = MyStr Then
                        DataGridView2.Rows(i).Selected = True
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Перерасчет заказа в соответствии с заданными скидкой и маржой
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyMargin As Double
        Dim MyOrder As String                           'Номер заказа
        Dim MyStr As String                             'Номер строки
        Dim MyDisc As Double                            'Скидка
        Dim MyFullDiscCost As String                    'стоим. всего кол - ва товара со скидкой
        Dim MyFullCost As String                        'стоим. всего кол - ва товара
        Dim MyCost As Double                            'Цена товара
        Dim cmd As New ADODB.Command
        Dim MyParam As ADODB.Parameter                  'передаваемый параметр номер 1
        Dim MyParam1 As ADODB.Parameter                 'передаваемый параметр номер 2
        Dim MyParam2 As ADODB.Parameter                 'передаваемый параметр номер 3
        Dim MyParam3 As ADODB.Parameter                 'передаваемый параметр номер 4
        Dim MyParam4 As ADODB.Parameter                 'передаваемый параметр номер 5
        Dim MyParam5 As ADODB.Parameter                 'передаваемый параметр номер 6
        Dim MyParam6 As ADODB.Parameter                 'передаваемый параметр номер 7
        Dim MyParam7 As ADODB.Parameter                 'передаваемый параметр номер 8
        Dim MyParam8 As ADODB.Parameter                 'передаваемый параметр номер 9
        Dim MyParam9 As ADODB.Parameter                 'передаваемый параметр номер 10
        Dim MyParam10 As ADODB.Parameter                'передаваемый параметр номер 11
        Dim MyPropCoeff As Double                       'коэфф пропорциональности при введенной стоимости доставки
        Dim MyStrTRCost As Double                       'стоимость доставки строки заказ (пропорционально себестоимости)
        Dim MyFullDiscTRCost As Double                  'стоим. всего кол - ва товара со скидкой без стоимости доставки
        Dim MyDiscCost As Double                        'Цена товара со скидкой

        MyDisc = 0
        MyMargin = 0
        '---проверка заполнения
        If TextBox4.Text <> "" Then
            MyDisc = Math.Round(CDbl(TextBox4.Text), 2)
        End If
        If TextBox5.Text <> "" Then
            MyMargin = CDbl(TextBox5.Text)
        End If

        '--коэфф пропорциональности (при введенной стоимости доставки)
        If CDbl(Label12.Text) = 0 Then
            MyPropCoeff = 999
        Else
            MyPropCoeff = (CDbl(Label22.Text) / Declarations.CurrencyValue) / CDbl(Label12.Text)
        End If

        '---перерасчет при правильном заполнении
        MyOrder = Label6.Text                              'Номер заказа

        cmd.ActiveConnection = Declarations.MyConn
        cmd.CommandText = "spp_SalesWorkplace4_EditOrder"
        cmd.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        cmd.CommandTimeout = 300

        MyParam = cmd.CreateParameter("@Order", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 10)
        cmd.Parameters.Append(MyParam)

        MyParam1 = cmd.CreateParameter("@Str", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
        cmd.Parameters.Append(MyParam1)
        MyParam2 = cmd.CreateParameter("@MyItemID", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 35)
        cmd.Parameters.Append(MyParam2)
        MyParam3 = cmd.CreateParameter("@MyName1", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
        cmd.Parameters.Append(MyParam3)
        MyParam4 = cmd.CreateParameter("@MyName2", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 25)
        cmd.Parameters.Append(MyParam4)
        MyParam5 = cmd.CreateParameter("@Cost", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam5)
        MyParam6 = cmd.CreateParameter("@CostIntr", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam6)
        MyParam7 = cmd.CreateParameter("@Qty", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam7)
        MyParam8 = cmd.CreateParameter("@Unit", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam8)
        MyParam9 = cmd.CreateParameter("@Discount", ADODB.DataTypeEnum.adBSTR, ADODB.ParameterDirectionEnum.adParamInput, 6)
        cmd.Parameters.Append(MyParam9)
        MyParam10 = cmd.CreateParameter("@EditOrRecalc", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput)
        cmd.Parameters.Append(MyParam10)

        If TextBox4.Text = "" And TextBox5.Text = "" Then
            '---Ничего не пересчитываем
            Exit Sub
        ElseIf TextBox4.Text <> "" And TextBox5.Text = "" Then
            '---Пересчитываем только скидку
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                MyStr = DataGridView2.Rows(i).Cells(1).Value.ToString

                MyParam.Value = MyOrder
                MyParam1.Value = MyStr
                MyParam2.Value = ""
                MyParam3.Value = ""
                MyParam4.Value = ""
                MyParam5.Value = DataGridView2.Rows(i).Cells(6).Value
                MyParam6.Value = 0
                MyParam7.Value = 0
                MyParam8.Value = 0
                MyParam9.Value = Replace(MyDisc, ",", ".")
                MyParam10.Value = 1
                Try
                    cmd.Execute()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Next i
        ElseIf TextBox4.Text = "" And TextBox5.Text <> "" Then
            '---Пересчитываем только маржу
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                MyStr = DataGridView2.Rows(i).Cells(1).Value.ToString
                MyFullDiscTRCost = CStr(Math.Round(CDbl(DataGridView2.Rows(i).Cells(8).Value) * 100 / (100 - MyMargin), 2))
                MyStrTRCost = CDbl(DataGridView2.Rows(i).Cells(8).Value) * MyPropCoeff
                MyFullDiscCost = MyFullDiscTRCost + MyStrTRCost
                MyDiscCost = MyFullDiscCost / CDbl(DataGridView2.Rows(i).Cells(7).Value)
                MyCost = Math.Round((MyDiscCost * 100) / (100 - CDbl(DataGridView2.Rows(i).Cells(10).Value)), 2)
                'MyFullCost = CStr(Math.Round(CDbl(MyFullDiscCost) * 100 / (100 - CDbl(DataGridView2.Rows(i).Cells(9).Value)), 2))
                'MyCost = CStr(Math.Round(CDbl(MyFullCost) / CDbl(DataGridView2.Rows(i).Cells(6).Value), 2)) 'Цена товара

                MyParam.Value = MyOrder
                MyParam1.Value = MyStr
                MyParam2.Value = ""
                MyParam3.Value = ""
                MyParam4.Value = ""
                MyParam5.Value = MyCost
                MyParam6.Value = 0
                MyParam7.Value = 0
                MyParam8.Value = 0
                MyParam9.Value = Replace(DataGridView2.Rows(i).Cells(10).Value, ",", ".")
                MyParam10.Value = 1
                Try
                    cmd.Execute()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Next i
        Else
            '---Пересчитываем скидку и маржу
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                MyStr = DataGridView2.Rows(i).Cells(1).Value.ToString
                MyFullDiscTRCost = CStr(Math.Round(CDbl(DataGridView2.Rows(i).Cells(8).Value) * 100 / (100 - MyMargin), 2))
                MyStrTRCost = CDbl(DataGridView2.Rows(i).Cells(8).Value) * MyPropCoeff
                MyFullDiscCost = MyFullDiscTRCost + MyStrTRCost
                MyDiscCost = MyFullDiscCost / CDbl(DataGridView2.Rows(i).Cells(7).Value)
                MyCost = Math.Round((MyDiscCost * 100) / (100 - MyDisc), 2)
                'MyFullCost = CStr(Math.Round(CDbl(MyFullDiscCost) * 100 / (100 - MyDisc), 2))
                'MyCost = CStr(Math.Round(CDbl(MyFullCost) / CDbl(DataGridView2.Rows(i).Cells(6).Value), 2)) 'цена товара

                MyParam.Value = MyOrder
                MyParam1.Value = MyStr
                MyParam2.Value = ""
                MyParam3.Value = ""
                MyParam4.Value = ""
                MyParam5.Value = MyCost
                MyParam6.Value = 0
                MyParam7.Value = 0
                MyParam8.Value = 0
                MyParam9.Value = Replace(MyDisc, ",", ".")
                MyParam10.Value = 1
                Try
                    cmd.Execute()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Next i
        End If
        OrderPreparation()
        TotalRecount()
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Скидка"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""скидка"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Маржа"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""маржа"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из редактирования строк заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        '---проверка маржи заказа
        If CDbl(Label28.Text) < Declarations.MinMarginLevelManager Then
            MyRez = MsgBox("Маржа вашего заказа меньше, чем разрешено для данного клиента. Отменить выход и переделать заказ?", vbYesNo, "Внимание!")
            If MyRez = vbYes Then
                Exit Sub
            End If
        End If

        Me.Close()
    End Sub

    Private Sub DataGridView2_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ChangeButtonsStatus()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния кнопок
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ChangeButtonsStatus()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого запаса по коду с начала строки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
        Else
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox2.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), Len(UCase(Trim(TextBox2.Text)))), 1) = 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    System.Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа в спецификацию
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            ExportSpecToLO()
        Else
            ExportSpecToExcel()
        End If
    End Sub

    Private Sub ExportSpecToExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа в Excel спецификацию
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""Скандика"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "Адрес:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "Спецификация поставки"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        '--заголовок таблицы
        MyWRKBook.ActiveSheet.Range("A10") = "N п/п"
        MyWRKBook.ActiveSheet.Range("B10") = "Код товара Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("D10") = "Наименование товара"
        MyWRKBook.ActiveSheet.Range("E10") = "Ед измерения"
        MyWRKBook.ActiveSheet.Range("F10") = "Кол-во"
        MyWRKBook.ActiveSheet.Range("G10") = "Цена без НДС"
        MyWRKBook.ActiveSheet.Range("H10") = "Сумма без НДС"
        MyWRKBook.ActiveSheet.Range("I10") = "Срок поставки (нед)"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:I10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:I10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:I10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:I10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12

        '--Вывод строк спецификации
        MySQLStr = "SELECT tbl_OR030300.OR03005 AS ItemCode, ISNULL(SC010300.SC01060, N'') AS SuppItemCode, LTRIM(RTRIM(LTRIM(RTRIM(tbl_OR030300.OR03006)) "
        'MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, ISNULL(View_1.txt, N'') AS UOM, tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'шт' THEN 'pcs(шт.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'кг' THEN 'kg (кг)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'л' THEN 'litre (литр)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'компл' THEN 'set (Компл.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)'  "
        MySQLStr = MySQLStr & "END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "ROUND(tbl_OR030300.OR03008 * (100.0 - tbl_OR030300.OR03018) / 100, 2) AS Price, "
        MySQLStr = MySQLStr & "tbl_OR030300.WeekQTY AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_OR030300.OR03010 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE     (tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_OR030300.OR03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("OR03011").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=ЕСЛИ(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---Выгрузка справочной информации
        MyWRKBook.ActiveSheet.Range("N10") = "Единица измерения"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(шт.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (м)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (кг)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (км)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (литр)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (Упак.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (Компл.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (пара)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportSpecToLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заказа в LibreOffice спецификацию
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---ширина колонок
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        '---защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""Скандика"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "Адрес:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "Спецификация поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        '--заголовок таблицы
        oSheet.getCellRangeByName("A10").String = "N п/п"
        oSheet.getCellRangeByName("B10").String = "Код товара Scala"
        oSheet.getCellRangeByName("C10").String = "Код товара поставщика"
        oSheet.getCellRangeByName("D10").String = "Наименование товара"
        oSheet.getCellRangeByName("E10").String = "Ед измерения"
        oSheet.getCellRangeByName("F10").String = "Кол-во"
        oSheet.getCellRangeByName("G10").String = "Цена без НДС"
        oSheet.getCellRangeByName("H10").String = "Сумма без НДС"
        oSheet.getCellRangeByName("I10").String = "Срок поставки (нед)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--Вывод строк спецификации
        MySQLStr = "SELECT tbl_OR030300.OR03005 AS ItemCode, ISNULL(SC010300.SC01060, ISNULL(tbl_OR030300.SuppItemCode, N'')) "
        MySQLStr = MySQLStr & "AS SuppItemCode, LTRIM(RTRIM(LTRIM(RTRIM(tbl_OR030300.OR03006)) "
        'MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, ISNULL(View_1.txt, N'') AS UOM, tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'шт' THEN 'pcs(шт.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'кг' THEN 'kg (кг)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'л' THEN 'litre (литр)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'компл' THEN 'set (Компл.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)'  "
        MySQLStr = MySQLStr & "END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "ROUND(tbl_OR030300.OR03008 * (100.0 - tbl_OR030300.OR03018) / 100, 2) AS Price, "
        MySQLStr = MySQLStr & "tbl_OR030300.WeekQTY AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_OR030300.OR03010 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE     (tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_OR030300.OR03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("OR03011").Value
                oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("Price").Value
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----Защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---Выгрузка справочной информации
        oSheet.getCellRangeByName("N10").String = "Единица измерения"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(шт.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (м)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (кг)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (км)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (литр)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (Упак.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (Компл.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (пара)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----в начало файла
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----закрытие паролем
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----видимость
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Ввод стоимости доставки
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyShipmentsCost = New ShipmentsCost
        MyShipmentsCost.ShowDialog()

        OrderPreparation()
        TotalRecount()
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Увеличение цены каждой строки КП на указанный %
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyPerCent As Double

        MyPerCent = 0
        '---проверка заполнения
        If TextBox6.Text <> "" Then
            MyPerCent = CDbl(TextBox6.Text)
        Else
            MsgBox("Поле 'Увеличить цену строк на' должно быть заполнено.", vbCritical, "Внимание!")
            TextBox6.Select()
            Exit Sub
        End If

        MySQLStr = "UPDATE tbl_OR030300 "
        MySQLStr = MySQLStr & "Set OR03008 = Round(OR03008 * (100 + " & MyPerCent & ") /100, 2) "
        MySQLStr = MySQLStr & "WHERE (OR03001 = N'" & Me.Label6.Text & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----Пересчет суммы заказа
        MySQLStr = "UPDATE tbl_OR010300 WITH (ROWLOCK) "
        MySQLStr = MySQLStr & "SET OR01024 = View_0.CorrSum "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT tbl_OR010300_1.OR01001, "
        MySQLStr = MySQLStr & "SUM(ROUND(ROUND((tbl_OR030300.OR03008 / tbl_OR030300.OR03022) * (100 - CONVERT(float, tbl_OR030300.OR03018) - CONVERT(float, tbl_OR030300.OR03017)) / 100, 2) * tbl_OR030300.OR03011, 2)) AS CorrSum "
        MySQLStr = MySQLStr & "FROM tbl_OR010300 AS tbl_OR010300_1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_OR030300 ON tbl_OR010300_1.OR01001 = tbl_OR030300.OR03001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR030300.OR03003 = N'000000') "
        MySQLStr = MySQLStr & "GROUP BY tbl_OR010300_1.OR01001 "
        MySQLStr = MySQLStr & "ORDER BY tbl_OR010300_1.OR01001) AS View_0 ON "
        MySQLStr = MySQLStr & "tbl_OR010300.OR01001 = View_0.OR01001 "
        MySQLStr = MySQLStr & "WHERE (tbl_OR010300.OR01001 = N'" & Me.Label6.Text & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        TextBox6.Text = ""
        OrderPreparation()
        TotalRecount()
    End Sub

    Private Sub TextBox6_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox6.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox6.Text) <> "" Then
            If InStr(TextBox6.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Увеличить цену строк на"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox6.Text
                Catch ex As Exception
                    MsgBox("В поле ""Увеличить цену строк на"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка Строк заказа без цены в форму спецификации
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If My.Settings.UseOffice = "LibreOffice" Then
            ExportNotPresentToLO()
        Else
            ExportNotPresentToExcel()
        End If
    End Sub

    Private Sub ExportNotPresentToExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка Строк заказа без цены в форму спецификации в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer

        MyObj = CreateObject("Excel.Application")
        Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")

        MyObj.SheetsInNewWorkbook = 1
        MyObj.ReferenceStyle = 1
        MyWRKBook = MyObj.Workbooks.Add

        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("A1") = SpecVersion
        With MyWRKBook.ActiveSheet.Range("A1").Font
            .Name = "Calibri"
            .Size = 9
            '.Color = -16776961
            .ColorIndex = 3
        End With

        MyWRKBook.ActiveSheet.Range("B2") = "Skandika"
        With MyWRKBook.ActiveSheet.Range("B2").Font
            .Name = "Calibri"
            .Size = 16
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A4") = "OOO ""Скандика"""
        With MyWRKBook.ActiveSheet.Range("A4").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("A5") = "Адрес:"
        With MyWRKBook.ActiveSheet.Range("A5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With

        MyWRKBook.ActiveSheet.Range("B5:H6").MergeCells = True
        MyWRKBook.ActiveSheet.Range("B5") = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        With MyWRKBook.ActiveSheet.Range("B5").Font
            .Name = "Tahoma"
            .Size = 9
            .Color = 0
            .Bold = True
        End With
        MyWRKBook.ActiveSheet.Range("B5:H6").WrapText = True
        MyWRKBook.ActiveSheet.Range("B5:H6").VerticalAlignment = -4160

        MyWRKBook.ActiveSheet.Range("D8") = "Спецификация поставки"
        With MyWRKBook.ActiveSheet.Range("D8").Font
            .Name = "Tahoma"
            .Size = 11.5
            .Color = 0
            .Bold = True
        End With

        '--заголовок таблицы
        MyWRKBook.ActiveSheet.Range("A10") = "N п/п"
        MyWRKBook.ActiveSheet.Range("B10") = "Код товара Scala"
        MyWRKBook.ActiveSheet.Range("C10") = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("D10") = "Наименование товара"
        MyWRKBook.ActiveSheet.Range("E10") = "Ед измерения"
        MyWRKBook.ActiveSheet.Range("F10") = "Кол-во"
        MyWRKBook.ActiveSheet.Range("G10") = "Цена без НДС"
        MyWRKBook.ActiveSheet.Range("H10") = "Сумма без НДС"
        MyWRKBook.ActiveSheet.Range("I10") = "Срок поставки (нед)"
        MyWRKBook.ActiveSheet.Range("A10:I10").Select()
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A10:I10").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A10:I10").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        MyWRKBook.ActiveSheet.Range("A10:I10").WrapText = True
        MyWRKBook.ActiveSheet.Range("A10:I10").VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A10:I10").HorizontalAlignment = -4108
        With MyWRKBook.ActiveSheet.Range("A10:I10").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 4
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 30
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 5
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 8
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 12
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 12

        '--Вывод строк спецификации
        MySQLStr = "SELECT tbl_OR030300.OR03005 AS ItemCode, ISNULL(SC010300.SC01060, ISNULL(tbl_OR030300.SuppItemCode, N'')) AS SuppItemCode, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(tbl_OR030300.OR03006)) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'шт' THEN 'pcs(шт.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'кг' THEN 'kg (кг)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'л' THEN 'litre (литр)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'компл' THEN 'set (Компл.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)'  "
        MySQLStr = MySQLStr & "END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "ROUND(tbl_OR030300.OR03008 * (100.0 - tbl_OR030300.OR03018) / 100, 2) AS Price, "
        MySQLStr = MySQLStr & "tbl_OR030300.WeekQTY AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_OR030300.OR03010 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE     ((tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') AND (SC010300.SC01001 IS NULL)) OR "
        MySQLStr = MySQLStr & "((tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') AND (SC010300.SC01055 = 0)) "
        MySQLStr = MySQLStr & "ORDER BY tbl_OR030300.OR03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("A" & CStr(i + 11)) = i + 1
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("B" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemCode").Value
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)).NumberFormat = "@"
                MyWRKBook.ActiveSheet.Range("C" & CStr(i + 11)) = Declarations.MyRec.Fields("SuppItemCode").Value
                MyWRKBook.ActiveSheet.Range("D" & CStr(i + 11)) = Declarations.MyRec.Fields("ItemName").Value
                MyWRKBook.ActiveSheet.Range("E" & CStr(i + 11)) = Declarations.MyRec.Fields("UOM").Value
                MyWRKBook.ActiveSheet.Range("F" & CStr(i + 11)) = Declarations.MyRec.Fields("OR03011").Value
                MyWRKBook.ActiveSheet.Range("G" & CStr(i + 11)) = Declarations.MyRec.Fields("Price").Value
                MyWRKBook.ActiveSheet.Range("I" & CStr(i + 11)) = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        '---

        MyWRKBook.ActiveSheet.Range("A11:I11").Select()
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A11:I11").Borders(6).LineStyle = -4142
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(7)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(8)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(9)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(10)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(11)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Borders(12)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A11:I1011").Font
            .Name = "Calibri"
            .Size = 9
            .Color = 0
            .Bold = False
        End With
        With MyWRKBook.ActiveSheet.Range("H11:H1011")
            '.FormulaR1C1 = "=ЕСЛИ(RC[-2]*RC[-1] = 0; """"; RC[-2]*RC[-1])"
            .FormulaR1C1 = "=IF(RC[-2]*RC[-1] = 0, """", RC[-2]*RC[-1])"
        End With
        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        '---Выгрузка справочной информации
        MyWRKBook.ActiveSheet.Range("N10") = "Единица измерения"
        MyWRKBook.ActiveSheet.Range("N11") = 0
        MyWRKBook.ActiveSheet.Range("O11") = "pcs(шт.)"
        MyWRKBook.ActiveSheet.Range("N12") = 1
        MyWRKBook.ActiveSheet.Range("O12") = "m (м)"
        MyWRKBook.ActiveSheet.Range("N13") = 2
        MyWRKBook.ActiveSheet.Range("O13") = "kg (кг)"
        MyWRKBook.ActiveSheet.Range("N14") = 3
        MyWRKBook.ActiveSheet.Range("O14") = "km (км)"
        MyWRKBook.ActiveSheet.Range("N15") = 4
        MyWRKBook.ActiveSheet.Range("O15") = "litre (литр)"
        MyWRKBook.ActiveSheet.Range("N16") = 5
        MyWRKBook.ActiveSheet.Range("O16") = "pack (Упак.)"
        MyWRKBook.ActiveSheet.Range("N17") = 6
        MyWRKBook.ActiveSheet.Range("O17") = "set (Компл.)"
        MyWRKBook.ActiveSheet.Range("N18") = 7
        MyWRKBook.ActiveSheet.Range("O18") = "pair (пара)"

        MyWRKBook.ActiveSheet.Range("N10:O18").Font.Color = 16777215
        'MyWRKBook.ActiveSheet.Range("N10:O18").Font.TintAndShade = 0
        MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=$O$11:$O$18")
        'MyWRKBook.ActiveSheet.Range("E11:E1011").Validation.Add(Type:=3, AlertStyle:=1, Operator:=1, Formula1:="=R11C15:R18C15")

        MyWRKBook.ActiveSheet.Cells.Locked = True
        MyWRKBook.ActiveSheet.Cells.FormulaHidden = True

        MyWRKBook.ActiveSheet.Range("A11:G1011").Locked = False
        MyWRKBook.ActiveSheet.Range("A11:G1011").FormulaHidden = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").Locked = False
        MyWRKBook.ActiveSheet.Range("I11:I1011").FormulaHidden = False

        MyWRKBook.ActiveSheet.Protect(Password:="!pass2009", DrawingObjects:=True, Contents:=True, Scenarios:=True)

        'MyWRKBook.Application.PrintCommunication = True
        'MyWRKBook.ActiveSheet.PageSetup.PrintArea = "$A$1:$I$1011"
        'MyWRKBook.Application.PrintCommunication = False
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesWide = 1
        'MyWRKBook.ActiveSheet.PageSetup.FitToPagesTall = 0
        'MyWRKBook.Application.PrintCommunication = True

        MyWRKBook.ActiveSheet.Range("A11").Select()
        MyObj.Application.Visible = True
        MyWRKBook = Nothing
        MyObj = Nothing
        oldCI = Nothing
    End Sub

    Private Sub ExportNotPresentToLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка Строк заказа без цены в форму спецификации в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim MySQLStr As String
        Dim SpecVersion As String               '--версия спецификации
        Dim i As Integer

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        '---ширина колонок
        oSheet.getColumns().getByName("A").Width = 1390
        oSheet.getColumns().getByName("B").Width = 2280
        oSheet.getColumns().getByName("C").Width = 2570
        oSheet.getColumns().getByName("D").Width = 5590
        oSheet.getColumns().getByName("E").Width = 1150
        oSheet.getColumns().getByName("F").Width = 1770
        oSheet.getColumns().getByName("G").Width = 2190
        oSheet.getColumns().getByName("H").Width = 2260
        oSheet.getColumns().getByName("I").Width = 2260
        '---защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "", True)
        '---заголовки
        '---версия спецификации
        MySQLStr = "SELECT Version "
        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Name = N'Спецификация предложения') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору.", vbCritical, "Внимание!")
            trycloseMyRec()
            Exit Sub
        Else
            SpecVersion = Trim(Declarations.MyRec.Fields("Version").Value)
            trycloseMyRec()
        End If
        oSheet.getCellRangeByName("A1").String = SpecVersion
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A1", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A1", 11)
        oSheet.getCellRangeByName("A1").CharColor = RGB(61, 65, 239) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий

        oSheet.getCellRangeByName("B2").String = "Skandika"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B2", "Calibri")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B2")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B2", 16)

        oSheet.getCellRangeByName("A4").String = "OOO ""Скандика"""
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A4", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A4")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A4", 11)

        oSheet.getCellRangeByName("A5").String = "Адрес:"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A5", 9)

        LOMergeCells(oServiceManager, oDispatcher, oFrame, "B5:H6")
        oSheet.getCellRangeByName("B5").String = "Россия,195027, Санкт Петербург, Шаумяна пр., д.4 Tel: +7 (812) 325 2040, Fax: +7 (812) 325 2039, www.skandikagroup.ru"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B5", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B5")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B5", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "B5:H6")

        oSheet.getCellRangeByName("D8").String = "Спецификация поставки"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "D8", "Tahoma")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "D8")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "D8", 11.5)

        '--заголовок таблицы
        oSheet.getCellRangeByName("A10").String = "N п/п"
        oSheet.getCellRangeByName("B10").String = "Код товара Scala"
        oSheet.getCellRangeByName("C10").String = "Код товара поставщика"
        oSheet.getCellRangeByName("D10").String = "Наименование товара"
        oSheet.getCellRangeByName("E10").String = "Ед измерения"
        oSheet.getCellRangeByName("F10").String = "Кол-во"
        oSheet.getCellRangeByName("G10").String = "Цена без НДС"
        oSheet.getCellRangeByName("H10").String = "Сумма без НДС"
        oSheet.getCellRangeByName("I10").String = "Срок поставки (нед)"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A10:I10", "Calibri")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A10:I10", 9)
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A10:I10")
        Dim LineFormat As Object
        LineFormat = oServiceManager.Bridge_GetStruct("com.sun.star.table.BorderLine2")
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 40
        oSheet.getCellRangeByName("A10:I10").TopBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").RightBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").LeftBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").BottomBorder = LineFormat
        oSheet.getCellRangeByName("A10:I10").VertJustify = 2
        oSheet.getCellRangeByName("A10:I10").HoriJustify = 2

        '--Вывод строк спецификации
        MySQLStr = "SELECT tbl_OR030300.OR03005 AS ItemCode, ISNULL(SC010300.SC01060, ISNULL(tbl_OR030300.SuppItemCode, N'')) AS SuppItemCode, "
        MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(tbl_OR030300.OR03006)) "
        MySQLStr = MySQLStr & "+ LTRIM(RTRIM(tbl_OR030300.OR03007)))) AS ItemName, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'шт' THEN 'pcs(шт.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'м' THEN 'm (м)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'кг' THEN 'kg (кг)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'км' THEN 'km (км)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'л' THEN 'litre (литр)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'упак' THEN 'pack (Упак.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'компл' THEN 'set (Компл.)' ELSE  "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(View_1.txt, N'') = 'пар' THEN 'pair (пара)'  "
        MySQLStr = MySQLStr & "END END END END END END END END AS UOM, "
        MySQLStr = MySQLStr & "tbl_OR030300.OR03011, "
        MySQLStr = MySQLStr & "ROUND(tbl_OR030300.OR03008 * (100.0 - tbl_OR030300.OR03018) / 100, 2) AS Price, "
        MySQLStr = MySQLStr & "tbl_OR030300.WeekQTY AS WeekQTY "
        MySQLStr = MySQLStr & "FROM tbl_OR030300 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT     0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM          SC090300 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE      (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_40 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_39 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_38 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_37 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_36 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_35 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_34 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_33 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_32 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_31 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_30 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS')"
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_29 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_28 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_27 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_26 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_25 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_24 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_23 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_22 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_21 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_20 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_19 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_18 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_17 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_16 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_15 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_14 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_13 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_12 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_11 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_10 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_9 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_8 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_7 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_6 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_5 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_4 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_3 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_2 WITH (NOLOCK) "
        MySQLStr = MySQLStr & "WHERE     (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT     40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM         SC090300 AS SC090300_1 WITH (NOLOCK)) AS View_1 ON tbl_OR030300.OR03010 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_OR030300.OR03005 = SC010300.SC01001 "
        MySQLStr = MySQLStr & "WHERE     ((tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') AND (SC010300.SC01001 IS NULL)) OR "
        MySQLStr = MySQLStr & "((tbl_OR030300.OR03001 = N'" & Trim(Label6.Text) & "') AND (SC010300.SC01055 = 0)) "
        MySQLStr = MySQLStr & "ORDER BY tbl_OR030300.OR03002 "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveFirst()
            i = 0
            While Declarations.MyRec.EOF = False
                oSheet.getCellRangeByName("A" & CStr(i + 11)).Value = i + 1
                oSheet.getCellRangeByName("B" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemCode").Value
                oSheet.getCellRangeByName("C" & CStr(i + 11)).String = Declarations.MyRec.Fields("SuppItemCode").Value
                oSheet.getCellRangeByName("D" & CStr(i + 11)).String = Declarations.MyRec.Fields("ItemName").Value
                oSheet.getCellRangeByName("E" & CStr(i + 11)).String = Declarations.MyRec.Fields("UOM").Value
                oSheet.getCellRangeByName("F" & CStr(i + 11)).Value = Declarations.MyRec.Fields("OR03011").Value
                oSheet.getCellRangeByName("G" & CStr(i + 11)).Value = Declarations.MyRec.Fields("Price").Value
                oSheet.getCellRangeByName("H" & CStr(i + 11)).FormulaLocal = "=IF(F" & CStr(i + 11) & "*G" & CStr(i + 11) & " = 0;"""";F" & CStr(i + 11) & " * G" & CStr(i + 11) & ") "
                oSheet.getCellRangeByName("I" & CStr(i + 11)).Value = Declarations.MyRec.Fields("WeekQTY").Value

                i = i + 1
                Declarations.MyRec.MoveNext()
            End While
            trycloseMyRec()
        End If
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A11:I" & CStr(11 + i - 1))
        LineFormat.LineStyle = 0
        LineFormat.LineWidth = 20
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).TopBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).RightBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).LeftBorder = LineFormat
        oSheet.getCellRangeByName("A11:I" & CStr(11 + i - 1)).BottomBorder = LineFormat
        '----Защита ячеек
        LOSetCellProtection(oServiceManager, oDispatcher, oFrame, "A11:I500", False)
        '---Выгрузка справочной информации
        oSheet.getCellRangeByName("N10").String = "Единица измерения"
        oSheet.getCellRangeByName("N11").Value = 0
        oSheet.getCellRangeByName("O11").String = "pcs(шт.)"
        oSheet.getCellRangeByName("N12").Value = 1
        oSheet.getCellRangeByName("O12").String = "m (м)"
        oSheet.getCellRangeByName("N13").Value = 2
        oSheet.getCellRangeByName("O13").String = "kg (кг)"
        oSheet.getCellRangeByName("N14").Value = 3
        oSheet.getCellRangeByName("O14").String = "km (км)"
        oSheet.getCellRangeByName("N15").Value = 4
        oSheet.getCellRangeByName("O15").String = "litre (литр)"
        oSheet.getCellRangeByName("N16").Value = 5
        oSheet.getCellRangeByName("O16").String = "pack (Упак.)"
        oSheet.getCellRangeByName("N17").Value = 6
        oSheet.getCellRangeByName("O17").String = "set (Компл.)"
        oSheet.getCellRangeByName("N18").Value = 7
        oSheet.getCellRangeByName("O18").String = "pair (пара)"
        oSheet.getCellRangeByName("N10:O18").CharColor = RGB(255, 255, 255) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOSetValidation(oSheet, "E11:E" & CStr(11 + i - 1), "=$O$11:$O$18")
        '----в начало файла
        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)
        '----закрытие паролем
        LOPasswordProtect(oSheet, "!pass2022", True)
        '----видимость
        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрос на поиск поставщика товаров
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If MySearchSupplier Is Nothing Then
            MySearchSupplier = New SearchSupplier
            MySearchSupplier.Show()
        Else
            'MySearchSupplier.BringToFront()
            MySearchSupplier.Close()
            MySearchSupplier = New SearchSupplier
            MySearchSupplier.Show()
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование заявки на поиск товаров (в рабочее место поисковика) 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyEditRequest = New EditRequest
        MyEditRequest.StartParam = "Create"
        MyEditRequest.WindowFrom = "OrderLines"
        MyEditRequest.ShowDialog()
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление строк заказа 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        OrderPreparation()
        ChangeButtonsStatus()
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Будут отображены только акционные товары
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If Button24.Text = "Только акции" Then '---оставляем только доступные к заказу продукты
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Button8.Text = "Все прод." Then
                DataPreparation(1, 1)
            Else
                DataPreparation(0, 1)
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button24.Text = "Все прод."
        Else                                             '---показываем все продукты
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            If Button8.Text = "Все прод." Then
                DataPreparation(1, 0)
            Else
                DataPreparation(0, 0)
            End If
            System.Windows.Forms.Cursor.Current = Cursors.Default
            Button24.Text = "Только акции"
        End If
    End Sub

    Private Sub Button8_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.MouseEnter
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отображение tooltip
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ToolTip1.Show("Только дост", Button8)
    End Sub

    Private Sub Button8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.MouseLeave
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Скрытие tooltip
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ToolTip1.Hide(Button8)
    End Sub

    Private Sub Button24_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button24.MouseEnter
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отображение tooltip
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ToolTip1.Show("Только акции", Button24)
    End Sub

    Private Sub Button24_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button24.MouseLeave
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Скрытие tooltip
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        ToolTip1.Hide(Button24)
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Печать КП
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MyPrintProposal = New PrintProposal
        MyPrintProposal.PropNum = Trim(Label6.Text)
        MyPrintProposal.Show()
    End Sub
End Class