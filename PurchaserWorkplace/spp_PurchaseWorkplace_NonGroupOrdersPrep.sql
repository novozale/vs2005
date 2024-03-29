USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_PurchaseWorkplace_NonGroupOrdersPrep]    Дата сценария: 05/12/2012 10:16:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_PurchaseWorkplace_NonGroupOrdersPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список несгруппированных заказов на закупку ненулевого типа               |
|    разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@WH nvarchar(6),                                     -- номер склада
@SupplierCode nvarchar(10)                           -- код поставщика


WITH RECOMPILE
AS

-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyOrders') 
and xtype = N'U')
DROP TABLE #_MyOrders

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyOrders(
	[PC01001] [nvarchar](10) COLLATE Cyrillic_General_BIN,            --номер заказа на закупку
	[PC01015] [datetime],                                             --дата заказа
	[OrderSum] [float] NULL,                                          --Сумма заказа
	[Purchaser] [nvarchar](255) COLLATE Cyrillic_General_BIN,         --Код + имя закупщика
	[SalesOrder] [nvarchar](50) COLLATE Cyrillic_General_BIN,         --номер заказа на продажу
	[DestWH] [nvarchar](50) COLLATE Cyrillic_General_BIN,             --Номер склада назначения
	[Salesman] [nvarchar](255) COLLATE Cyrillic_General_BIN           --Продавец
)

------------------Список заказов
INSERT INTO #_MyOrders
SELECT PC010300.PC01001 AS OrderN, 
	PC010300.PC01015 AS OrderDate, 
	NULL AS OrderSum, 
	LTRIM(RTRIM(LTRIM(RTRIM(PC010300.PC01046)) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, N''))))) AS Purchaser, 
	LTRIM(RTRIM(PC010300.PC01060)) AS SalesOrderN, 
    PC010300.PC01006 AS DestWH,
	NULL
FROM PC010300 INNER JOIN
    (SELECT PC03001
    FROM PC030300
	WHERE (PC03010 <> 0) OR
		(PC03011 <> 0)
    GROUP BY PC03001) AS View_2 ON 
	PC010300.PC01001 = View_2.PC03001 LEFT OUTER JOIN
    (SELECT SYPD001, SYPD002, SYPD003
    FROM SYPD0300
    WHERE (SYPD002 = N'ENG')) AS View_1 ON 
	UPPER(PC010300.PC01046) = UPPER(View_1.SYPD001)
WHERE (PC010300.PC01002 <> 0) AND 
	(PC010300.PC01023 = @WH) AND 
	(PC010300.PC01001 NOT IN
		(SELECT PC010300_2.PC01001
		FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN
			PC010300 AS PC010300_2 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300_2.PC01052
		GROUP BY PC010300_2.PC01001)) AND 
	(PC010300.PC01003 = @SupplierCode)

------------------Обновление сумм заказов (руб)
UPDATE #_MyOrders
SET OrderSum = Round(View_1.Sum,2)
FROM #_MyOrders INNER JOIN
	(SELECT PC030300.PC03001, 
		SUM(Round(PC030300.PC03008 / PC030300.PC03019 * PC030300.PC03010 * SYCH0100.SYCH006,2)) AS Sum
    FROM PC030300 INNER JOIN
		PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN
        SYCH0100 ON PC010300.PC01022 = SYCH0100.SYCH001 AND 
		PC010300.PC01015 >= SYCH0100.SYCH004 AND 
        PC010300.PC01015 < SYCH0100.SYCH005
    GROUP BY PC030300.PC03001) AS View_1 ON 
	#_MyOrders.PC01001 = View_1.PC03001

-------------------Обновление информации по продавцу
UPDATE #_MyOrders
SET Salesman = View_2.Salesman
FROM #_MyOrders INNER JOIN
	(SELECT View_1.OR01001, 
		LTRIM(RTRIM(LTRIM(RTRIM(View_1.OR01019)) + ' ' + LTRIM(RTRIM(ISNULL(ST010300.ST01002, N''))))) AS Salesman
	FROM (SELECT OR01001, OR01019
		FROM OR010300
        UNION
        SELECT OR20001, OR20019
        FROM OR200300) AS View_1 LEFT OUTER JOIN
    ST010300 ON View_1.OR01019 = ST010300.ST01001) AS View_2 ON
	#_MyOrders.SalesOrder = View_2.OR01001

UPDATE #_MyOrders
SET Salesman = ISNULL(Salesman,'')

SELECT * 
FROM #_MyOrders
ORDER By PC01001