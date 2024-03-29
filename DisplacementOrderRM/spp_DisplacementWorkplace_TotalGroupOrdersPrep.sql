USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementWorkplace_TotalGroupOrdersPrep]    Дата сценария: 10/17/2013 13:05:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_DisplacementWorkplace_TotalGroupOrdersPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список сгруппированных заказов на перемещение                             |
|    разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@WHFrom nvarchar(6),                                     -- номер склада откуда
@WHTo nvarchar(6),                                       -- Номер склада куда
@Actual int                                              -- 0 - все обобщенные заказы (доставки), 1- только непринятые (активные)


WITH RECOMPILE
AS


-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyOrders') 
and xtype = N'U')
DROP TABLE #_MyOrders

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyOrders(
	[ID] float,                                                            --номер консолидированного заказа на перемещение
	[ShipmentDate] [datetime],                                             --дата отправки консолидированного заказа на перемещение
	[ShipmentDelay] int,                                                   --задержка с отправкой 0-нет 1-да
	[ReceivingDate] [datetime],                                            --дата приемки консолидированного заказа на перемещение
	[ReceivingDelay] int,                                                  --задержка с приемкой 0-нет 1-да
	[QTYInOrder] int,                                                      --кол-во заказов в обобщенном
	[UserName] [nvarchar](255) COLLATE Cyrillic_General_BIN,               --имя создавшего заказ
	[IsActive] int,                                                        -- 0-заказ закрыт (все тип 2) или 1-нет
	[TransportCompanyDocNum] [nvarchar](50) COLLATE Cyrillic_General_BIN,  -- номер документа перевозчика
	[Comments] [nvarchar](255) COLLATE Cyrillic_General_BIN,               -- комментарий
	[BlockedForEdits] int,                                                 --заблокирован для редактирования (1) или нет (0)
	[BlockedForAdd] int                                                    --заблокирован для добавления заказов (1) или нет (0)
)


-----------------Список заказов
INSERT INTO #_MyOrders
SELECT tbl_DisplacementOrder_ShipmentInfo.ID, 
	tbl_DisplacementOrder_ShipmentInfo.ShipmentDate, 
	NULL,
	tbl_DisplacementOrder_ShipmentInfo.ReceivingDate, 
	NULL,
	NULL,
    ScalaSystemDB.dbo.ScaUsers.FullName,
	NULL,
	tbl_DisplacementOrder_ShipmentInfo.TransportCompanyDocNum,
	tbl_DisplacementOrder_ShipmentInfo.Comments,
	NULL,
	NULL
FROM tbl_DisplacementOrder_ShipmentInfo INNER JOIN
	ScalaSystemDB.dbo.ScaUsers ON tbl_DisplacementOrder_ShipmentInfo.UserID = ScalaSystemDB.dbo.ScaUsers.UserID
WHERE (tbl_DisplacementOrder_ShipmentInfo.WHFrom = @WHFrom) AND 
	(tbl_DisplacementOrder_ShipmentInfo.WHTo = @WHTo)

------------------Задержка с отправкой
UPDATE #_MyOrders
SET ShipmentDelay = View_1.ShipmentDelay
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT tbl_DisplacementOrder_ShipmentInfo.ID, 
		CASE WHEN COUNT(DISTINCT tbl_DisplacementOrder_Shipment.OrderNumber) > 0 THEN 1 ELSE 0 END AS ShipmentDelay
	FROM tbl_DisplacementOrder_ShipmentInfo INNER JOIN
		tbl_DisplacementOrder_Shipment ON tbl_DisplacementOrder_ShipmentInfo.ID = tbl_DisplacementOrder_Shipment.ShipmentsNumber INNER JOIN
        SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001 AND 
        tbl_DisplacementOrder_ShipmentInfo.ShipmentDate > SC7C0300.SC7C007
	GROUP BY tbl_DisplacementOrder_ShipmentInfo.ID) AS View_1 ON
	#_MyOrders.ID = View_1.ID

-----------------Задержка с приемкой
UPDATE #_MyOrders
SET ReceivingDelay = View_1.ReceivingDelay
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT tbl_DisplacementOrder_ShipmentInfo.ID, 
		CASE WHEN COUNT(DISTINCT tbl_DisplacementOrder_Shipment.OrderNumber) > 0 THEN 1 ELSE 0 END AS ReceivingDelay
	FROM tbl_DisplacementOrder_ShipmentInfo INNER JOIN
		tbl_DisplacementOrder_Shipment ON tbl_DisplacementOrder_ShipmentInfo.ID = tbl_DisplacementOrder_Shipment.ShipmentsNumber INNER JOIN
        SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001 AND 
        tbl_DisplacementOrder_ShipmentInfo.ReceivingDate > SC7C0300.SC7C009
	GROUP BY tbl_DisplacementOrder_ShipmentInfo.ID) AS View_1 ON
	#_MyOrders.ID = View_1.ID


-----------------Количество включенных заказов на перемещение
UPDATE #_MyOrders
SET QTYInOrder = View_1.QTY
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT ShipmentsNumber, 
		COUNT(OrderNumber) AS QTY
	FROM tbl_DisplacementOrder_Shipment
	GROUP BY ShipmentsNumber) AS View_1 ON
	#_MyOrders.ID = View_1.ShipmentsNumber


-----------------Заказ закрыт или нет
UPDATE #_MyOrders
SET IsActive = View_1.IsActive
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT tbl_DisplacementOrder_Shipment.ShipmentsNumber, 
		2 - MIN(SC7C0300.SC7C010) AS IsActive
	FROM tbl_DisplacementOrder_Shipment INNER JOIN
		SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001
	GROUP BY tbl_DisplacementOrder_Shipment.ShipmentsNumber) AS View_1 ON
	#_MyOrders.ID = View_1.ShipmentsNumber

-----------------Заказ заблокирован для редактирования (как только была первая отгрузка)
UPDATE #_MyOrders
SET BlockedForEdits = View_1.BlockedForEdits
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT tbl_DisplacementOrder_Shipment.ShipmentsNumber, 
		CASE WHEN SUM(SC7D0300.SC7D006) > 0 THEN 1 ELSE 0 END AS BlockedForEdits
	FROM tbl_DisplacementOrder_Shipment INNER JOIN
		SC7D0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7D0300.SC7D001
	GROUP BY tbl_DisplacementOrder_Shipment.ShipmentsNumber) AS View_1 ON
	#_MyOrders.ID = View_1.ShipmentsNumber

-----------------Заказ заблокирован для добавления заказов (как только была первая приемка)
UPDATE #_MyOrders
SET BlockedForAdd = View_1.BlockedForAdd
FROM #_MyOrders WITH(NOLOCK) LEFT OUTER JOIN
	(SELECT tbl_DisplacementOrder_Shipment.ShipmentsNumber, 
		CASE WHEN SUM(SC7D0300.SC7D007) > 0 THEN 1 ELSE 0 END AS BlockedForAdd
	FROM tbl_DisplacementOrder_Shipment INNER JOIN
		SC7D0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7D0300.SC7D001
	GROUP BY tbl_DisplacementOrder_Shipment.ShipmentsNumber) AS View_1 ON
	#_MyOrders.ID = View_1.ShipmentsNumber

-----------------Выставление 0 
UPDATE #_MyOrders
SET ShipmentDelay = ISNULL(ShipmentDelay,0),
	ReceivingDelay = ISNULL(ReceivingDelay,0),
	QTYInOrder = ISNULL(QTYInOrder,0),
	IsActive = ISNULL(IsActive,1),
	TransportCompanyDocNum = ISNULL(TransportCompanyDocNum,''),
	Comments = ISNULL(Comments,'')

-----------------Выбор результатов
SELECT ID,
	ShipmentDate,
	CASE WHEN ShipmentDelay = 0 THEN '' ELSE 'X' END AS ShipmentDelay,
	ReceivingDate,
	CASE WHEN ReceivingDelay = 0 THEN '' ELSE 'X' END AS ReceivingDelay,
	QTYInOrder,
	UserName,
	CASE WHEN IsActive = 0 THEN '' ELSE 'X' END AS IsActive,
	TransportCompanyDocNum,
	Comments,
	BlockedForEdits,
	BlockedForAdd
FROM #_MyOrders
WHERE (IsActive >= @Actual)
Order By ID
