USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementWorkplace_GroupOrdersPrep]    Дата сценария: 10/17/2013 13:04:25 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_DisplacementWorkplace_GroupOrdersPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список заказов, входящих в выбранный сгруппированный заказ                |
|    разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@CoOrderNum nvarchar(10)                                              -- Номер консолидированного заказа


WITH RECOMPILE
AS


-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyOrders') 
and xtype = N'U')
DROP TABLE #_MyOrders

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyOrders(
	[SC7C001] [nvarchar](10),                                         --номер заказа на перемещение
	[SC7C007] [datetime],                                             --дата отгрузки
	[IsShipped] int,                                                  --отгружен
	[SC7C009] [datetime],                                             --дата приемки
	[IsReceived] int,                                                 --принят
	[IsActive] int,                                                   --активен или закрыт (тип 2)
	[SalesOrderNum] [nvarchar](10),                                   --номер заказа на продажу
	[Employee] [nvarchar](255)                                        --сотрудник созздавший заказ
)
                    

-----------------Список заказов
INSERT INTO #_MyOrders
SELECT tbl_DisplacementOrder_Shipment.OrderNumber, 
	SC7C0300.SC7C007, 
	NULL,
	SC7C0300.SC7C009,
	NULL,
	NULL, 
	ISNULL(tbl_DisplacementOrder.SalesOrderNumber, N'') AS SalesOrderNumber, 
	ISNULL(ScalaSystemDB.dbo.ScaUsers.FullName, N'') AS Employee
FROM ScalaSystemDB.dbo.ScaUsers RIGHT OUTER JOIN
	tbl_DisplacementOrder ON ScalaSystemDB.dbo.ScaUsers.UserID = tbl_DisplacementOrder.UserCode RIGHT OUTER JOIN
    tbl_DisplacementOrder_Shipment INNER JOIN
    SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001 ON 
    tbl_DisplacementOrder.OrderNumber = tbl_DisplacementOrder_Shipment.OrderNumber
WHERE (tbl_DisplacementOrder_Shipment.ShipmentsNumber = @CoOrderNum)


-----------------Информация по отгрузке
UPDATE #_MyOrders
SET IsShipped = View_1.IsShipped
FROM #_MyOrders INNER JOIN
	(SELECT SC7D001, 
		CASE WHEN SUM(SC7D004) = SUM(SC7D006) AND SUM(SC7D004) > 0 THEN 2 ELSE CASE WHEN SUM(SC7D006) = 0 THEN 0 ELSE 1 END END AS IsShipped
	FROM SC7D0300
	GROUP BY SC7D001) AS View_1 ON
	#_MyOrders.SC7C001 = View_1.SC7D001

-----------------Информация о приемке
UPDATE #_MyOrders
SET IsReceived = View_1.IsReceived
FROM #_MyOrders INNER JOIN
	(SELECT SC7D001, 
		CASE WHEN SUM(SC7D004) = SUM(SC7D007) AND SUM(SC7D004) > 0 THEN 2 ELSE CASE WHEN SUM(SC7D007) = 0 THEN 0 ELSE 1 END END AS IsReceived
	FROM SC7D0300
	GROUP BY SC7D001) AS View_1 ON
	#_MyOrders.SC7C001 = View_1.SC7D001

-----------------Активность заказа
UPDATE #_MyOrders
SET IsActive = View_1.IsActive
FROM #_MyOrders INNER JOIN
	(SELECT SC7C001, 
		CASE WHEN SC7C010 = 2 THEN 0 ELSE 1 END AS IsActive
	FROM SC7C0300) AS View_1 ON
	#_MyOrders.SC7C001 = View_1.SC7C001

-----------------Выставление 0 
UPDATE #_MyOrders
SET IsShipped = ISNULL(IsShipped,0),
	IsReceived = ISNULL(IsReceived,0),
	IsActive = ISNULL(IsActive,0),
	SalesOrderNum = ISNULL(SalesOrderNum,''),
	Employee = ISNULL(Employee,'')

----------------Выбор результатов
SELECT SC7C001,
	SC7C007,
	CASE WHEN IsShipped = 0 THEN 'не отгружен' ELSE CASE WHEN IsShipped = 1 THEN 'частично' ELSE 'отгружен' END END AS IsShipped,
	SC7C009,
	CASE WHEN IsReceived = 0 THEN 'не принят' ELSE CASE WHEN IsReceived = 1 THEN 'частично' ELSE 'принят' END END AS IsReceived,
	CASE WHEN IsActive = 0 THEN 'закрыт' ELSE 'активен' END AS IsActive,
	SalesOrderNum,
	Employee
FROM #_MyOrders
ORDER By SC7C001
/*
------------------Список заказов
INSERT INTO #_MyOrders
SELECT PC010300.PC01001 AS OrderN, 
	PC010300.PC01015 AS OrderDate, 
	NULL AS OrderSum, 
	LTRIM(RTRIM(LTRIM(RTRIM(PC010300.PC01046)) + ' ' + LTRIM(RTRIM(ISNULL(View_1.SYPD003, N''))))) AS Purchaser, 
	LTRIM(RTRIM(PC010300.PC01060)) AS SalesOrderN, 
    PC010300.PC01006 AS DestWH,
	NULL
FROM PC010300 WITH(NOLOCK) LEFT OUTER JOIN
    (SELECT SYPD001, SYPD002, SYPD003
    FROM SYPD0300 WITH(NOLOCK)
    WHERE (SYPD002 = N'ENG')) AS View_1 ON 
	UPPER(PC010300.PC01046) = UPPER(View_1.SYPD001)
WHERE (PC010300.PC01052 = @CoOrderNum) 

------------------Обновление сумм заказов (руб)
UPDATE #_MyOrders
SET OrderSum = Round(View_1.Sum,2)
FROM #_MyOrders INNER JOIN
	(SELECT PC030300.PC03001, 
		SUM(Round(PC030300.PC03008 / PC030300.PC03019 * PC030300.PC03010 * SYCH0100.SYCH006,2)) AS Sum
    FROM PC030300 WITH(NOLOCK) INNER JOIN
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
		FROM OR010300 WITH(NOLOCK)
        UNION
        SELECT OR20001, OR20019
        FROM OR200300 WITH(NOLOCK) ) AS View_1 LEFT OUTER JOIN
    ST010300 ON View_1.OR01019 = ST010300.ST01001) AS View_2 ON
	#_MyOrders.SalesOrder = View_2.OR01001

UPDATE #_MyOrders
SET Salesman = ISNULL(Salesman,'')

SELECT * 
FROM #_MyOrders
ORDER By PC01001
*/