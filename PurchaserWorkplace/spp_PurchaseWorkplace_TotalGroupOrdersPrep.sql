USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_PurchaseWorkplace_TotalGroupOrdersPrep]    Дата сценария: 05/11/2012 08:43:07 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_PurchaseWorkplace_TotalGroupOrdersPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список сгруппированных заказов на закупку                                 |
|    разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@WH nvarchar(6),                                     -- номер склада
@SupplierCode nvarchar(10),                          -- код поставщика
@Actual int                                          -- 0 - все заказы, 1- только непринятые (активные)


WITH RECOMPILE
AS

-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyOrders') 
and xtype = N'U')
DROP TABLE #_MyOrders

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyOrders(
	[ID] [nvarchar](10) COLLATE Cyrillic_General_BIN,                 --номер консолидированного заказа на закупку
	[OrderDate] [datetime],                                           --дата консолидированного заказа на закупку
	[OrderSum] [float],                                               --Сумма консолидированного заказа на закупку
	[Purchaser] [nvarchar](255) COLLATE Cyrillic_General_BIN,         --имя закупщика
	[SupplierPlacedDate] [datetime],                                  --Дата размещения заказа у поставщика
	[ConfirmedDate] [datetime],                                       --Дата подтверждения заказа
	[SupplierOrderNumber] [nvarchar](50) COLLATE Cyrillic_General_BIN,--Номер заказа поставщика
	[IsActive] int,                                                   --0 - все заказы, 1- только непринятые (активные)
	[DelayedDate] datetime,                                           --Подтвержденная дата поставки (для задолженного заказа)
	[QTYInOrder] int                                                  --кол-во заказов в обобщенном
)

-----------------Список заказов
INSERT INTO #_MyOrders
SELECT tbl_PurchaseWorkplace_ConsolidatedOrders.ID, 
	tbl_PurchaseWorkplace_ConsolidatedOrders.OrderDate, 
	NULL AS OrderSum, 
    ISNULL(ScalaSystemDB.dbo.ScaUsers.FullName, 'Неизвестен') AS Purchaser, 
	tbl_PurchaseWorkplace_ConsolidatedOrders.SupplierPlacedDate, 
    tbl_PurchaseWorkplace_ConsolidatedOrders.ConfirmedDate, 
	tbl_PurchaseWorkplace_ConsolidatedOrders.SupplierOrderNumber, 
	NULL AS IsClosed, 
	NULL AS DelayedDate,
	NULL AS QTYInOrder
FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN
	ScalaSystemDB.dbo.ScaUsers ON tbl_PurchaseWorkplace_ConsolidatedOrders.UserID = ScalaSystemDB.dbo.ScaUsers.UserID
WHERE (tbl_PurchaseWorkplace_ConsolidatedOrders.SupplierCode = @SupplierCode) AND 
	(tbl_PurchaseWorkplace_ConsolidatedOrders.WH = @WH)

-----------------Сумма заказов
UPDATE #_MyOrders
SET OrderSum = Round(View_1.Sum,2)
FROM #_MyOrders INNER JOIN
	(SELECT PC010300.PC01052, 
		SUM(ROUND(PC030300.PC03008 / PC030300.PC03019 * PC030300.PC03010 * SYCH0100.SYCH006, 2)) AS Sum
	FROM PC030300 INNER JOIN
		PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN
        SYCH0100 ON PC010300.PC01022 = SYCH0100.SYCH001 AND 
		PC010300.PC01015 >= SYCH0100.SYCH004 AND 
        PC010300.PC01015 < SYCH0100.SYCH005
	GROUP BY PC010300.PC01052) AS View_1 ON 
	#_MyOrders.ID = View_1.PC01052

-----------------Непринятые заказы 
UPDATE #_MyOrders
SET IsActive = 1
FROM #_MyOrders INNER JOIN
	(SELECT PC010300.PC01052 AS OrderN
    FROM PC010300 INNER JOIN
		(SELECT PC03001
        FROM PC030300
        WHERE (PC03010 <> 0) OR
			(PC03011 <> 0)
        GROUP BY PC03001) AS View_2 ON 
		PC010300.PC01001 = View_2.PC03001
    WHERE (PC010300.PC01002 <> 0) AND 
		(PC010300.PC01023 = @WH) AND 
		(PC010300.PC01003 = @SupplierCode)
    GROUP BY PC010300.PC01052) AS View_1 ON 
	#_MyOrders.ID = View_1.OrderN

UPDATE #_MyOrders
SET IsActive = 1
FROM #_MyOrders INNER JOIN
	(SELECT tbl_PurchaseWorkplace_ConsolidatedOrders.ID, 
		COUNT(PC010300.PC01001) AS Expr1
    FROM PC010300 RIGHT OUTER JOIN
		tbl_PurchaseWorkplace_ConsolidatedOrders ON 
		PC010300.PC01052 = tbl_PurchaseWorkplace_ConsolidatedOrders.ID
    GROUP BY tbl_PurchaseWorkplace_ConsolidatedOrders.ID
    HAVING (COUNT(PC010300.PC01001) = 0)) AS View_1 ON 
	#_MyOrders.ID = View_1.ID

UPDATE #_MyOrders
SET IsActive = ISNULL(IsActive,0)


-----------------Подтвержденная дата поставки (для задолженного заказа) - минимальная из задолженных
UPDATE #_MyOrders
SET DelayedDate = View_1.Expr1
FROM #_MyOrders INNER JOIN
	(SELECT PC010300.PC01052, 
		MIN(PC030300.PC03031) AS Expr1
    FROM PC030300 INNER JOIN
		PC010300 ON PC030300.PC03001 = PC010300.PC01001
    WHERE (PC030300.PC03029 = N'1')
    GROUP BY PC010300.PC01052) AS View_1 ON 
	#_MyOrders.ID = View_1.PC01052

-----------------кол-во заказов в обобщенном
UPDATE #_MyOrders
SET QTYInOrder = View_1.Expr1
FROM #_MyOrders INNER JOIN
	(SELECT PC01052, 
		COUNT(PC01001) AS Expr1
    FROM PC010300
    GROUP BY PC01052) AS View_1 ON 
	#_MyOrders.ID = View_1.PC01052

UPDATE #_MyOrders
SET QTYInOrder = ISNULL(QTYInOrder,0) 


-----------------Выбор результатов
SELECT ID,
	OrderDate,
	OrderSum,
	Purchaser,
	SupplierPlacedDate,
	ConfirmedDate,
	SupplierOrderNumber,
	CASE WHEN IsActive = 0 THEN '' ELSE 'X' END AS IsActive,
	DelayedDate,
	QTYInOrder
FROM #_MyOrders
WHERE (IsActive >= @Actual)
Order By ID
