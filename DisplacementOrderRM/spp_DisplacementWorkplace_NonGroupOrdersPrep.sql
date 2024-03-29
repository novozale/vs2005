USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementWorkplace_NonGroupOrdersPrep]    Дата сценария: 10/17/2013 13:04:55 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_DisplacementWorkplace_NonGroupOrdersPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список несгруппированных заказов на перемещение 1 и 2 типа                |
|    разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@WHFrom nvarchar(6),                                     -- номер склада отправки
@WHTo nvarchar(6)                                        -- номер склада получения


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
	[SalesOrderNum] [nvarchar](10),                                   --номер заказа на продажу
	[Employee] [nvarchar](255),                                       --сотрудник создавший заказ
	[Status] [nvarchar] (2)                                           --статус заказа 1-в работе 2-закрыт
)


-----------------Список заказов
INSERT INTO #_MyOrders
SELECT SC7C0300.SC7C001, 
	SC7C0300.SC7C007, 
	NULL,
	SC7C0300.SC7C009, 
	ISNULL(tbl_DisplacementOrder.SalesOrderNumber, N'') AS SalesOrderNumber, 
    ISNULL(ScalaSystemDB.dbo.ScaUsers.FullName, N'') AS Employee,
	SC7C0300.SC7C010
FROM SC7C0300 INNER JOIN
    tbl_DisplacementOrder ON SC7C0300.SC7C001 = tbl_DisplacementOrder.OrderNumber LEFT OUTER JOIN
    ScalaSystemDB.dbo.ScaUsers ON tbl_DisplacementOrder.UserCode = ScalaSystemDB.dbo.ScaUsers.UserID
WHERE (SC7C0300.SC7C010 = 1 OR SC7C0300.SC7C010 = 2) AND
	(SC7C0300.SC7C004 = @WHFrom) AND
	(SC7C0300.SC7C006 = @WHTo) AND
	(tbl_DisplacementOrder.ReadyFlag = 1) AND
	(SC7C0300.SC7C001 NOT IN (SELECT DISTINCT OrderNumber FROM tbl_DisplacementOrder_Shipment))

-----------------Информация по отгрузке
UPDATE #_MyOrders
SET IsShipped = View_1.IsShipped
FROM #_MyOrders INNER JOIN
	(SELECT SC7D001, 
		CASE WHEN SUM(SC7D004) = SUM(SC7D006) AND SUM(SC7D004) > 0 THEN 2 ELSE CASE WHEN SUM(SC7D006) = 0 THEN 0 ELSE 1 END END AS IsShipped
	FROM SC7D0300
	GROUP BY SC7D001) AS View_1 ON
	#_MyOrders.SC7C001 = View_1.SC7D001

-----------------Выставление 0 
UPDATE #_MyOrders
SET SalesOrderNum = ISNULL(SalesOrderNum,''),
	Employee = ISNULL(Employee,''),
	IsShipped = ISNULL(IsShipped,0)

----------------Выбор результатов
SELECT SC7C001,
	SC7C007,
	CASE WHEN IsShipped = 0 THEN 'не отгружен' ELSE CASE WHEN IsShipped = 1 THEN 'частично' ELSE 'отгружен' END END AS IsShipped,
	SC7C009,
	SalesOrderNum,
	Employee,
	[Status]
FROM #_MyOrders
ORDER By SC7C001
