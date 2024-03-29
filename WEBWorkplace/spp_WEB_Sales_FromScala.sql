USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Sales_FromScala]    Дата сценария: 05/29/2015 12:42:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Sales_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продажам (заголовки)                                       |
|    в промежуточную БД обмена с WEB из Scala                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

SET NOCOUNT ON

--=======================Информация по строкам заказов==================================
truncate table tbl_WEB_CardSalesDetails


-------------------------Строки продаж из истории---------------------------------------
INSERT INTO tbl_WEB_CardSalesDetails
SELECT ST030300.ST03001 AS CustomerCode, 
	View_2.MinOrderDate AS OrderDate, 
	ST030300.ST03009 AS OrderNum, 
	ST030300.ST03062 AS StrNum, 
    ST030300.ST03017 AS ItemCode, 
	ST030300.ST03020 AS OrderedQTY, 
	0 AS ReadyQTY, 
	ST030300.ST03020 AS ShippedQTY, 
    Round((ST030300.ST03021 * (100 - ST030300.ST03022) / 100) * (100 + CONVERT(numeric(28, 8), 
		CASE WHEN Ltrim(Rtrim(SY290300.SY29003)) = '' THEN '0' ELSE SY290300.SY29003 END)) / 100, 2) AS PriceRUR,
	Ltrim(Rtrim(Ltrim(Rtrim(SC010300.SC01002)) + ' ' + Ltrim(Rtrim(SC010300.SC01003)))) AS ItemName,
	1 AS IsClosed,
	9 AS OrderType
FROM ST030300 INNER JOIN
	tbl_WEB_Clients ON ST030300.ST03001 = tbl_WEB_Clients.Code INNER JOIN
    SC010300 ON ST030300.ST03017 = SC010300.SC01001 INNER JOIN
    SY290300 ON SC010300.SC01144 = SY290300.SY29001 INNER JOIN
    (SELECT OR20001, 
		MIN(OR20015) AS MinOrderDate
    FROM (SELECT OR20001, 
			OR20015
        FROM OR200300
        UNION
        SELECT OR01001, 
			OR01015
        FROM OR010300) AS View_1
    GROUP BY OR20001) AS View_2 ON 
	ST030300.ST03009 = View_2.OR20001
WHERE (tbl_WEB_Clients.WorkOverWEB = 1) 
	AND (ST030300.ST03063 = N'000000')

-------------------------Строки продаж из OR03-------------------------------------------
INSERT INTO tbl_WEB_CardSalesDetails
SELECT OR010300.OR01003 AS CustomerCode, 
	View_2.MinOrderDate AS OrderDate, 
	OR010300.OR01001 AS OrderNum, 
	OR030300.OR03002 AS StrNum, 
    OR030300.OR03005 AS ItemCode, 
	OR030300.OR03011 AS OrderedQTY, 
	OR030300.OR03012 AS ReadyQTY, 
	0 AS ShippedQTY, 
    ROUND(((OR030300.OR03008 / OR030300.OR03022 * SYCH0100.SYCH006) * (100 - CONVERT(float, OR030300.OR03017) - CONVERT(float, 
		OR030300.OR03018)) / 100) * (100 + CONVERT(numeric(28, 8), CASE WHEN Ltrim(Rtrim(SY290300.SY29003)) 
		= '' THEN '0' ELSE SY290300.SY29003 END)) / 100, 2) AS Price, 
	Ltrim(Rtrim(Ltrim(Rtrim(SC010300.SC01002)) + ' ' + Ltrim(Rtrim(SC010300.SC01003)))) AS ItemName,
	0 AS IsClosed, 
	OR010300.OR01002 AS OrderType
FROM OR010300 INNER JOIN
	tbl_WEB_Clients ON OR010300.OR01003 = tbl_WEB_Clients.Code INNER JOIN
    OR030300 ON OR010300.OR01001 = OR030300.OR03001 INNER JOIN
    (SELECT OR20001, 
		MIN(OR20015) AS MinOrderDate
    FROM (SELECT OR20001, 
			OR20015
        FROM OR200300
        UNION
        SELECT OR01001, 
			OR01015
        FROM OR010300 AS OR010300_1) AS View_1
    GROUP BY OR20001) AS View_2 ON 
	OR010300.OR01001 = View_2.OR20001 INNER JOIN
    SC010300 ON OR030300.OR03005 = SC010300.SC01001 INNER JOIN
    SY290300 ON SC010300.SC01144 = SY290300.SY29001 INNER JOIN
    SYCH0100 ON OR010300.OR01028 = SYCH0100.SYCH001
WHERE (tbl_WEB_Clients.WorkOverWEB = 1) 
	AND (OR030300.OR03003 = N'000000') 
	AND (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


-------------------------Обновление состояния заказа-------------------------------------
UPDATE tbl_WEB_CardSalesDetails
SET IsClosed = 0
FROM tbl_WEB_CardSalesDetails INNER JOIN
	(SELECT DISTINCT OrderNum
	FROM tbl_WEB_CardSalesDetails AS tbl_WEB_CardSalesDetails_1
    WHERE (IsClosed = 0)) AS View_6 ON 
	tbl_WEB_CardSalesDetails.OrderNum = View_6.OrderNum


-------------------------Обновление типа заказа (мин.)-----------------------------------
UPDATE tbl_WEB_CardSalesDetails
SET OrderType = View_7.MinOrderType
FROM tbl_WEB_CardSalesDetails INNER JOIN
	(SELECT OrderNum, 
		MIN(OrderType) AS MinOrderType
	FROM dbo.tbl_WEB_CardSalesDetails
	GROUP BY OrderNum) AS View_7 ON 
	tbl_WEB_CardSalesDetails.OrderNum = View_7.OrderNum



--=======================Информация по заголовкам заказов================================
-------------------------Новые заказы----------------------------------------------------
INSERT INTO tbl_WEB_CardSales
SELECT View_1.ClientCode, 
	View_1.OrderDate, 
	View_1.OrderNum, 
	View_1.Discount, 
	View_1.OrderSumm, 
	View_1.ShipmentState, 
    View_1.OrderState,
	View_1.WebOrderNum,
	1 AS RMStatus,
	1 AS WEBStatus
FROM tbl_WEB_CardSales RIGHT OUTER JOIN
	(SELECT TOP (100) PERCENT tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
        tbl_WEB_CardSalesDetails.OrderNum, 
		'' AS Discount, 
		Round(SUM(tbl_WEB_CardSalesDetails.OrderedQTY * tbl_WEB_CardSalesDetails.Price), 2) AS OrderSumm, 
		ISNULL(tbl_WEB_OrderNum.WebOrderNum, N'') AS WebOrderNum, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 
			THEN 2 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ShippedQTY) = 0 THEN 0 
			ELSE 1 END END AS ShipmentState, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 THEN -4 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) 
			> 0 THEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) / SUM(tbl_WEB_CardSalesDetails.OrderedQTY) 
            * 100 ELSE CASE WHEN tbl_WEB_CardSalesDetails.OrderType > 0 THEN - 1 ELSE - 2 END END END AS OrderState
    FROM tbl_WEB_CardSalesDetails LEFT OUTER JOIN
        tbl_WEB_OrderNum ON tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_OrderNum.ScaOrderNUm
    GROUP BY tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
		tbl_WEB_CardSalesDetails.OrderNum, 
        tbl_WEB_OrderNum.WebOrderNum, 
		tbl_WEB_CardSalesDetails.IsClosed, 
		tbl_WEB_CardSalesDetails.OrderType) AS View_1 ON 
	tbl_WEB_CardSales.ClientCode = View_1.ClientCode 
	AND tbl_WEB_CardSales.OrderNum = View_1.OrderNum
WHERE (tbl_WEB_CardSales.OrderNum IS NULL)


-------------------------Измененные заказы-----------------------------------------------
UPDATE tbl_WEB_CardSales
SET OrderDate = View_1.OrderDate, 
	OrderSumm = View_1.OrderSumm, 
	WEBOrderNum = View_1.WebOrderNum, 
	ShipmentState = View_1.ShipmentState, 
    OrderState = View_1.OrderState, 
	RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END, 
	WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END
FROM tbl_WEB_CardSales INNER JOIN
	(SELECT TOP (100) PERCENT tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
        tbl_WEB_CardSalesDetails.OrderNum, 
		'' AS Discount, 
		Round(SUM(tbl_WEB_CardSalesDetails.OrderedQTY * tbl_WEB_CardSalesDetails.Price), 2) AS OrderSumm, 
		ISNULL(tbl_WEB_OrderNum.WebOrderNum, N'') AS WebOrderNum, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 
			THEN 2 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ShippedQTY) = 0 THEN 0 
			ELSE 1 END END AS ShipmentState, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 THEN -4 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) 
			> 0 THEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) / SUM(tbl_WEB_CardSalesDetails.OrderedQTY) 
            * 100 ELSE CASE WHEN tbl_WEB_CardSalesDetails.OrderType > 0 THEN - 1 ELSE - 2 END END END AS OrderState
    FROM tbl_WEB_CardSalesDetails LEFT OUTER JOIN
        tbl_WEB_OrderNum ON tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_OrderNum.ScaOrderNUm
    GROUP BY tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
		tbl_WEB_CardSalesDetails.OrderNum, 
        tbl_WEB_OrderNum.WebOrderNum, 
		tbl_WEB_CardSalesDetails.IsClosed, 
		tbl_WEB_CardSalesDetails.OrderType) AS View_1 ON 
	tbl_WEB_CardSales.ClientCode = View_1.ClientCode 
	AND tbl_WEB_CardSales.OrderNum = View_1.OrderNum
WHERE (tbl_WEB_CardSales.OrderDate <> View_1.OrderDate) OR
	(tbl_WEB_CardSales.OrderSumm <> View_1.OrderSumm) OR
    (tbl_WEB_CardSales.WEBOrderNum <> View_1.WebOrderNum) OR
    (tbl_WEB_CardSales.ShipmentState <> View_1.ShipmentState) OR
    (tbl_WEB_CardSales.OrderState <> View_1.OrderState)


-------------------------Удаленные заказы------------------------------------------------
UPDATE tbl_WEB_CardSales
SET RMStatus = 2, 
	WEBStatus = 2
FROM tbl_WEB_CardSales LEFT OUTER JOIN
	(SELECT TOP (100) PERCENT tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
        tbl_WEB_CardSalesDetails.OrderNum, 
		'' AS Discount, 
		Round(SUM(tbl_WEB_CardSalesDetails.OrderedQTY * tbl_WEB_CardSalesDetails.Price), 2) AS OrderSumm, 
		ISNULL(tbl_WEB_OrderNum.WebOrderNum, N'') AS WebOrderNum, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 
			THEN 2 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ShippedQTY) = 0 THEN 0 
			ELSE 1 END END AS ShipmentState, 
        CASE WHEN tbl_WEB_CardSalesDetails.IsClosed = 1 THEN -4 ELSE CASE WHEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) 
			> 0 THEN SUM(tbl_WEB_CardSalesDetails.ReadyQTY) / SUM(tbl_WEB_CardSalesDetails.OrderedQTY) 
            * 100 ELSE CASE WHEN tbl_WEB_CardSalesDetails.OrderType > 0 THEN - 1 ELSE - 2 END END END AS OrderState
    FROM tbl_WEB_CardSalesDetails LEFT OUTER JOIN
        tbl_WEB_OrderNum ON tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_OrderNum.ScaOrderNUm
    GROUP BY tbl_WEB_CardSalesDetails.ClientCode, 
		tbl_WEB_CardSalesDetails.OrderDate, 
		tbl_WEB_CardSalesDetails.OrderNum, 
        tbl_WEB_OrderNum.WebOrderNum, 
		tbl_WEB_CardSalesDetails.IsClosed, 
		tbl_WEB_CardSalesDetails.OrderType) AS View_1 ON 
	tbl_WEB_CardSales.ClientCode = View_1.ClientCode 
	AND tbl_WEB_CardSales.OrderNum = View_1.OrderNum
WHERE (View_1.OrderNum IS NULL)