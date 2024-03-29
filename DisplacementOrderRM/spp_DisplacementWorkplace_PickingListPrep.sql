USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementWorkplace_PickingListPrep]    Дата сценария: 10/17/2013 13:05:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_DisplacementWorkplace_PickingListPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список запасов, входящих в выбранный сгруппированный заказ                |
|    разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@CoOrderNum nvarchar(10)                                              -- Номер консолидированного заказа


WITH RECOMPILE
AS

/*
-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyShipment') 
and xtype = N'U')
DROP TABLE #_MyShipment

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyShipment(
	[StockCode] [nvarchar](35),                                       --код запаса
	[StockName] [nvarchar](80) NULL,                                  --название запаса
	[WHFrom][nvarchar](6),                                            --номер склада отгрузки
	[QTYOrdered] float NULL,                                          --заказано к перемещению
	[QTYOrderedRest] float NULL,                                      --осталось отгрузить
	[BatchID] [nvarchar](20),                                         --номер партии
	[QTYOrderedBatch] float,                                          --заказано к перемещению из партии
	[QTYOrderedBatchRest] float,                                      --осталось отгрузить из партии
	[UOM] int,                                                        --код единицы измерения
	[UOMName][nvarchar](30) NULL,                                     --единица измерения
	[BINNumber] [nvarchar](6) NULL                                    --Номер ячейки по умолчанию
)
*/                    

-----------------Товары и партии с количеством к отгрузке
INSERT INTO #_MyShipment
SELECT LTRIM(RTRIM(SC7D0300.SC7D003)) AS StockCode,
	NULL, 
	SC7C0300.SC7C004 AS WHFrom,
	NULL,
	NULL, 
	SC7E0300.SC7E004 AS BatchID, 
	SUM(SC7E0300.SC7E013) AS QTYOrderedBatch, 
	SUM(SC7E0300.SC7E013 - SC7E0300.SC7E007) AS QTYOrderedBatchRest,
	CONVERT(int,SC7D0300.SC7D008) AS UOM,
	NULL,
	NULL
FROM tbl_DisplacementOrder_Shipment INNER JOIN
    SC7C0300 ON tbl_DisplacementOrder_Shipment.OrderNumber = SC7C0300.SC7C001 INNER JOIN
    SC7D0300 ON SC7C0300.SC7C001 = SC7D0300.SC7D001 INNER JOIN
    SC7E0300 ON SC7D0300.SC7D001 = SC7E0300.SC7E001 AND 
	SC7D0300.SC7D002 = SC7E0300.SC7E002
WHERE (tbl_DisplacementOrder_Shipment.ShipmentsNumber = @CoOrderNum)
GROUP BY SC7E0300.SC7E004, 
	SC7C0300.SC7C004, 
	LTRIM(RTRIM(SC7D0300.SC7D003)),
	SC7D0300.SC7D008

----------------Единица измерения
UPDATE #_MyShipment
SET UOMName = UOM.UOM_Description
FROM #_MyShipment INNER JOIN
	(SELECT     0 AS UOM_code, SC09002 AS UOM_Description
	FROM          SC090300 WITH (NOLOCK)
	WHERE      (SC09001 = 'RUS')
	UNION
	SELECT     1 AS Expr1, SC09003
	FROM         SC090300 AS SC090300_40 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     2 AS Expr1, SC09004
	FROM         SC090300 AS SC090300_39 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     3 AS Expr1, SC09005
	FROM         SC090300 AS SC090300_38 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     4 AS Expr1, SC09006
	FROM         SC090300 AS SC090300_37 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     5 AS Expr1, SC09007
	FROM         SC090300 AS SC090300_36 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     6 AS Expr1, SC09008
	FROM         SC090300 AS SC090300_35 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     7 AS Expr1, SC09009
	FROM         SC090300 AS SC090300_34 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     8 AS Expr1, SC09010
	FROM         SC090300 AS SC090300_33 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     9 AS Expr1, SC09011
	FROM         SC090300 AS SC090300_32 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     10 AS Expr1, SC09012
	FROM         SC090300 AS SC090300_31 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     11 AS Expr1, SC09013
	FROM         SC090300 AS SC090300_30 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     12 AS Expr1, SC09014
	FROM         SC090300 AS SC090300_29 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     13 AS Expr1, SC09015
	FROM         SC090300 AS SC090300_28 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     14 AS Expr1, SC09016
	FROM         SC090300 AS SC090300_27 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     15 AS Expr1, SC09017
	FROM         SC090300 AS SC090300_26 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     16 AS Expr1, SC09018
	FROM         SC090300 AS SC090300_25 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     17 AS Expr1, SC09019
	FROM         SC090300 AS SC090300_24 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     18 AS Expr1, SC09020
	FROM         SC090300 AS SC090300_23 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     19 AS Expr1, SC09021
	FROM         SC090300 AS SC090300_22 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     20 AS Expr1, SC09022
	FROM         SC090300 AS SC090300_21 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     21 AS Expr1, SC09023
	FROM         SC090300 AS SC090300_20 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     22 AS Expr1, SC09024
	FROM         SC090300 AS SC090300_19 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     23 AS Expr1, SC09025
	FROM         SC090300 AS SC090300_18 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     24 AS Expr1, SC09026
	FROM         SC090300 AS SC090300_17 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     25 AS Expr1, SC09027
	FROM         SC090300 AS SC090300_16 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     26 AS Expr1, SC09028
	FROM         SC090300 AS SC090300_15 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     27 AS Expr1, SC09029
	FROM         SC090300 AS SC090300_14 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     28 AS Expr1, SC09030
	FROM         SC090300 AS SC090300_13 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     29 AS Expr1, SC09031
	FROM         SC090300 AS SC090300_12 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     30 AS Expr1, SC09032
	FROM         SC090300 AS SC090300_11 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     31 AS Expr1, SC09033
	FROM         SC090300 AS SC090300_10 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     32 AS Expr1, SC09034
	FROM         SC090300 AS SC090300_9 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     33 AS Expr1, SC09035
	FROM         SC090300 AS SC090300_8 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     34 AS Expr1, SC09036
	FROM         SC090300 AS SC090300_7 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     35 AS Expr1, SC09037
	FROM         SC090300 AS SC090300_6 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     36 AS Expr1, SC09038
	FROM         SC090300 AS SC090300_5 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     37 AS Expr1, SC09039
	FROM         SC090300 AS SC090300_4 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     38 AS Expr1, SC09040
	FROM         SC090300 AS SC090300_3 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     39 AS Expr1, SC09041
	FROM         SC090300 AS SC090300_2 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')
	UNION
	SELECT     40 AS Expr1, SC09042
	FROM         SC090300 AS SC090300_1 WITH (NOLOCK)
	WHERE     (SC09001 = 'RUS')) AS UOM ON UOM.UOM_code = #_MyShipment.UOM

UPDATE #_MyShipment
SET UOMName = ISNULL(UOMName,'')


----------------Ячейка хранения по умолчанию
UPDATE #_MyShipment
SET BINNumber = SC030300.SC03013
FROM #_MyShipment INNER JOIN
	SC030300 ON #_MyShipment.StockCode = SC030300.SC03001 AND 
	#_MyShipment.WHFrom = SC030300.SC03002

UPDATE #_MyShipment
SET BINNumber = ISNULL(BINNumber,'')

----------------Количества сгруппированные потоварно
UPDATE #_MyShipment
SET QTYOrdered = View_1.QTYOrderedBatch, 
	QTYOrderedRest = View_1.QTYOrderedBatchRest
FROM #_MyShipment INNER JOIN
	(SELECT StockCode, 
		SUM(QTYOrderedBatch) AS QTYOrderedBatch, 
		SUM(QTYOrderedBatchRest) AS QTYOrderedBatchRest
    FROM #_MyShipment AS #_MyShipment_1
    GROUP BY StockCode) AS View_1 ON 
	#_MyShipment.StockCode = View_1.StockCode

----------------Названия запасов
UPDATE #_MyShipment
SET StockName = LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003))))
FROM #_MyShipment INNER JOIN
	SC010300 ON #_MyShipment.StockCode = SC010300.SC01001

----------------Выбор результатов
/*
SELECT * FROM #_MyShipment
ORDER BY StockCode,
	BatchID
*/