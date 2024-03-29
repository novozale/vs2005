USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementWorkplace_WHToListPrep]    Дата сценария: 10/21/2013 09:39:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_DisplacementWorkplace_WHToListPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список складов на которые идет перемещение с выбранного склада            |
|    разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@ActiveOrNot int,                                    -- (0) Выводить всех поставщиков или (1) только активных 
@WH nvarchar(6)                                      -- номер склада, откуда идет перемещение


WITH RECOMPILE
AS


-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyWHTo') 
and xtype = N'U')
DROP TABLE #_MySuppliers

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyWHTo(
	[SC23001] [nvarchar](6) COLLATE Cyrillic_General_BIN,             --Код склада назначения
	[SC23002] [nvarchar](35) COLLATE Cyrillic_General_BIN,            --название склада назначения
	[NonGroupOrders] float NULL,                                      --кол - во несгруппированных заказов
	[ReadyGroupOrders] float NULL,                                    --кол-во сгруппированных заказов в работе (не до конца принятых)
	[NonShippedOrders] float NULL,                                    --кол - во неотправленных заказов
	[NonReceivedOrders] float NULL                                    --кол - во непринятых заказов
)


------Список складов
INSERT INTO #_MyWHTo
SELECT SC23001 AS WHCode,
	SC23002 AS WHName,
	NULL AS NonGroupOrders,
	NULL AS ReadyGroupOrders,
	NULL AS NonShippedOrders,
	NULL AS NonReceivedOrders
FROM SC230300 WITH(NOLOCK)
WHERE (SC23001 <> @WH)


------кол - во несгруппированных заказов
/*
UPDATE #_MyWHTo
SET NonGroupOrders = View_1.QTY
FROM #_MyWHTo LEFT OUTER JOIN
	(SELECT SC7C006, 
		COUNT(SC7C001) AS QTY
	FROM SC7C0300
	WHERE (SC7C010 <> N'2') AND 
		(SC7C010 <> N'0') AND
		(SC7C004 = @WH) AND
		(SC7C001 NOT IN (SELECT DISTINCT OrderNumber FROM tbl_DisplacementOrder_Shipment))
	GROUP BY SC7C006) AS View_1 ON 
	#_MyWHTo.SC23001 = View_1.SC7C006
*/
UPDATE #_MyWHTo
SET NonGroupOrders = View_2.QTY
FROM #_MyWHTo LEFT OUTER JOIN
	(SELECT SC7C006, 
		COUNT(SC7C001) AS QTY
	FROM (SELECT SC7C0300.SC7C006, 
			SC7C0300.SC7C001
        FROM SC7C0300 INNER JOIN
			tbl_DisplacementOrder ON SC7C0300.SC7C001 = tbl_DisplacementOrder.OrderNumber
        WHERE (SC7C0300.SC7C010 = 1 OR SC7C0300.SC7C010 = 2) AND 
			(SC7C0300.SC7C004 = @WH) AND 
			(tbl_DisplacementOrder.ReadyFlag = 1) AND 
            (SC7C0300.SC7C001 NOT IN (SELECT DISTINCT OrderNumber FROM tbl_DisplacementOrder_Shipment))) AS View_1
	GROUP BY SC7C006) AS View_2 ON
	#_MyWHTo.SC23001 = View_2.SC7C006

------Кол - во заказов в работе
UPDATE #_MyWHTo
SET ReadyGroupOrders = View_1.QTY
FROM #_MyWHTo LEFT OUTER JOIN
	(SELECT SC7C0300.SC7C006, 
		COUNT(DISTINCT tbl_DisplacementOrder_Shipment.ShipmentsNumber) AS QTY
	FROM tbl_DisplacementOrder_Shipment INNER JOIN
		SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001
	WHERE (SC7C0300.SC7C004 = @WH) AND 
		(SC7C0300.SC7C010 <> N'2') AND
		(SC7C0300.SC7C010 <> N'0')
	GROUP BY SC7C0300.SC7C006) AS View_1 ON 
	#_MyWHTo.SC23001 = View_1.SC7C006

------Кол - во неотправленных заказов
UPDATE #_MyWHTo
SET NonShippedOrders = View_1.QTY
FROM #_MyWHTo LEFT OUTER JOIN
	(SELECT View_10.SC7C006, 
		COUNT(DISTINCT tbl_DisplacementOrder_ShipmentInfo.ID) AS QTY
	FROM tbl_DisplacementOrder_ShipmentInfo INNER JOIN
		tbl_DisplacementOrder_Shipment ON tbl_DisplacementOrder_ShipmentInfo.ID = tbl_DisplacementOrder_Shipment.ShipmentsNumber INNER JOIN
    (SELECT SC7C0300.SC7C001, 
		SC7C0300.SC7C006, 
		SC7C0300.SC7C007
    FROM SC7C0300 INNER JOIN
		SC7D0300 ON SC7C0300.SC7C001 = SC7D0300.SC7D001
    WHERE (SC7D0300.SC7D006 < SC7D0300.SC7D004) AND 
		(SC7C0300.SC7C010 <> N'2') AND 
		(SC7C0300.SC7C010 <> N'0') AND
		(SC7C0300.SC7C004 = @WH)
    GROUP BY SC7C0300.SC7C001, 
		SC7C0300.SC7C006, 
		SC7C0300.SC7C007) AS View_10 ON 
    tbl_DisplacementOrder_Shipment.OrderNumber = View_10.SC7C001
WHERE (tbl_DisplacementOrder_ShipmentInfo.ShipmentDate < GETDATE()) OR
    (View_10.SC7C007 < GETDATE())
GROUP BY View_10.SC7C006) AS View_1 ON 
	#_MyWHTo.SC23001 = View_1.SC7C006

------Кол - во непринятых заказов
UPDATE #_MyWHTo
SET NonReceivedOrders = View_1.QTY
FROM #_MyWHTo LEFT OUTER JOIN
	(SELECT View_10.SC7C006, 
		COUNT(DISTINCT tbl_DisplacementOrder_ShipmentInfo.ID) AS QTY
	FROM tbl_DisplacementOrder_ShipmentInfo INNER JOIN
		tbl_DisplacementOrder_Shipment ON tbl_DisplacementOrder_ShipmentInfo.ID = tbl_DisplacementOrder_Shipment.ShipmentsNumber INNER JOIN
    (SELECT SC7C0300.SC7C001, 
		SC7C0300.SC7C006, 
		SC7C0300.SC7C009
    FROM SC7C0300 INNER JOIN
		SC7D0300 ON SC7C0300.SC7C001 = SC7D0300.SC7D001
    WHERE (SC7D0300.SC7D007 < SC7D0300.SC7D004) AND 
		(SC7C0300.SC7C010 <> N'2') AND 
		(SC7C0300.SC7C010 <> N'0') AND
		(SC7C0300.SC7C004 = @WH)
    GROUP BY SC7C0300.SC7C001, 
		SC7C0300.SC7C006, 
		SC7C0300.SC7C009) AS View_10 ON 
    tbl_DisplacementOrder_Shipment.OrderNumber = View_10.SC7C001
WHERE (tbl_DisplacementOrder_ShipmentInfo.ReceivingDate < GETDATE()) OR
    (View_10.SC7C009 < GETDATE())
GROUP BY View_10.SC7C006) AS View_1 ON 
	#_MyWHTo.SC23001 = View_1.SC7C006

------убираем NULL
UPDATE #_MyWHTo
SET NonGroupOrders = ISNULL(NonGroupOrders,0),
	ReadyGroupOrders = ISNULL(ReadyGroupOrders,0),
	NonShippedOrders = ISNULL(NonShippedOrders,0),
	NonReceivedOrders = ISNULL(NonReceivedOrders,0)

-----------------Вывод данных-------------------------------
if @ActiveOrNot = 0
BEGIN    ---все склады назначения
	SELECT * FROM #_MyWHTo
	Order BY SC23001
END
ELSE
BEGIN    ---только те склады назначения, которые активны (есть не принятые до конц заказы)
	SELECT * FROM #_MyWHTo
	WHERE (NonGroupOrders <> 0) OR
		(ReadyGroupOrders <> 0) OR
		(NonShippedOrders <> 0) OR
		(NonReceivedOrders <> 0)
	Order BY SC23001
END
