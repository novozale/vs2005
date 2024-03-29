USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_PurchaseWorkplace_PurchaseGroupOrderPreparation]    Дата сценария: 05/11/2012 08:42:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




CREATE Procedure [dbo].[spp_PurchaseWorkplace_PurchaseGroupOrderPreparation]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит данные по обобщенному заказу на закупку                                   |
|    любого типа                                                                       |
|    Разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@CommonOrderNumber [nvarchar](20)                     -- номер обобщенного заказа

WITH RECOMPILE
AS

Declare @MyCurrCode [int];                                --код валюты заказа
Declare @MyExchRate [numeric](28, 8);                     --Курс валюты заказа
declare @MyDate datetime;                                 --текущая дата
Declare @MyOrderSum [numeric](28, 8);                     --Сумма заказа
Declare @MyOrderLang [nvarchar](3);                       --язык заказа

Set @MyDate = CONVERT(datetime,CONVERT(nvarchar,DATEPART(dd,GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(mm,GETDATE())) + '/' + CONVERT(nvarchar,DATEPART(yyyy,GETDATE())),103)

----------------код валюты заказа и язык----------------------------------------------
SELECT @MyCurrCode = PL010300.PL01026,
	@MyOrderLang = PL010300.PL01027
FROM (SELECT PC01003
	FROM PC010300
    WHERE (PC01052 = @CommonOrderNumber)
    GROUP BY PC01003) AS View_1 INNER JOIN
    PL010300 ON View_1.PC01003 = PL010300.PL01001

----------------Курс валюты заказа-----------------------------------------------------
SELECT @MyExchRate = SYCH007
FROM SYCH0100
WHERE (SYCH001 = @MyCurrCode) AND 
	(SYCH004 <= @MyDate) AND 
	(SYCH005 > @MyDate)
/*
-----------------очистка временных таблиц----------------------------------------------
IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MyPCOrder') 
and xtype = N'U')
DROP TABLE #_MyPCOrder

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MyPCOrder(
	[PC03005] [nvarchar](35) COLLATE Cyrillic_General_BIN,            --код товара
	[SC01060] [nvarchar](35) COLLATE Cyrillic_General_BIN,            --код товара поставщика
	[PC03006] [nvarchar] (52) COLLATE Cyrillic_General_BIN,           --название товара
	[QTY] [numeric](20, 8),                                           --количество
	[PC03009] [int],                                                  --единица измерения
	[PC03009_Name][nvarchar](10),                                     --название единицы измерения (SC09)
	[Price] [numeric](28, 8),                                         --цена за 1
	[StrSum] [numeric](28, 8),                                        --Сумма строки
	[OrderSum] [numeric](28, 8)                                       --Сумма заказа
)
*/

-----------------список товаров в заказе----------------------------------------------
INSERT INTO #_MyPCOrder
SELECT PC030300.PC03005, 
	SC010300.SC01060,
	LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003)))) AS PC03006, 
	SUM(PC030300.PC03010) AS QTY,
	SC010300.SC01134,
	NULL,
	NULL,
	NULL,
	NULL
FROM PC030300 INNER JOIN
    PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN
    SC010300 ON PC030300.PC03005 = SC010300.SC01001
WHERE (PC010300.PC01052 = @CommonOrderNumber)
GROUP BY PC030300.PC03005, 
	SC010300.SC01060,
	LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003)))),
	SC010300.SC01134

--------------------Английские наименования продуктов---------------------------------
if Ltrim(Rtrim(@MyOrderLang)) <> 'RUS'
BEGIN
	UPDATE #_MyPCOrder
	SET PC03006 = View_1.SC04004
	FROM #_MyPCOrder INNER JOIN
		(SELECT SC04001, SC04004
		FROM SC040300
		WHERE (SC04002 = N'B_') AND 
			(SC04003 = N'ENG')) AS View_1 ON 
	#_MyPCOrder.PC03005 = View_1.SC04001
END

--------------------название единицы измерения----------------------------------------
UPDATE #_MyPCOrder
SET PC03009_Name = ISNULL(View_1.txt, 'Неизв.')
FROM #_MyPCOrder LEFT OUTER JOIN
	(SELECT 0 AS num, SC09002 AS txt
    FROM SC090300
    WHERE      (SC09001 = 'RUS')
    UNION
    SELECT     1 AS Expr1, SC09003
    FROM         SC090300 AS SC090300_40
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     2 AS Expr1, SC09004
    FROM         SC090300 AS SC090300_39
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     3 AS Expr1, SC09005
    FROM         SC090300 AS SC090300_38
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     4 AS Expr1, SC09006
    FROM         SC090300 AS SC090300_37
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     5 AS Expr1, SC09007
    FROM         SC090300 AS SC090300_36
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     6 AS Expr1, SC09008
    FROM         SC090300 AS SC090300_35
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     7 AS Expr1, SC09009
    FROM         SC090300 AS SC090300_34
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     8 AS Expr1, SC09010
    FROM         SC090300 AS SC090300_33
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     9 AS Expr1, SC09011
    FROM         SC090300 AS SC090300_32
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     10 AS Expr1, SC09012
    FROM         SC090300 AS SC090300_31
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     11 AS Expr1, SC09013
    FROM         SC090300 AS SC090300_30
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     12 AS Expr1, SC09014
    FROM         SC090300 AS SC090300_29
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     13 AS Expr1, SC09015
    FROM         SC090300 AS SC090300_28
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     14 AS Expr1, SC09016
    FROM         SC090300 AS SC090300_27
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     15 AS Expr1, SC09017
    FROM         SC090300 AS SC090300_26
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     16 AS Expr1, SC09018
    FROM         SC090300 AS SC090300_25
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     17 AS Expr1, SC09019
    FROM         SC090300 AS SC090300_24
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     18 AS Expr1, SC09020
    FROM         SC090300 AS SC090300_23
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     19 AS Expr1, SC09021
    FROM         SC090300 AS SC090300_22
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     20 AS Expr1, SC09022
    FROM         SC090300 AS SC090300_21
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     21 AS Expr1, SC09023
    FROM         SC090300 AS SC090300_20
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     22 AS Expr1, SC09024
    FROM         SC090300 AS SC090300_19
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     23 AS Expr1, SC09025
    FROM         SC090300 AS SC090300_18
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     24 AS Expr1, SC09026
    FROM         SC090300 AS SC090300_17
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     25 AS Expr1, SC09027
    FROM         SC090300 AS SC090300_16
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     26 AS Expr1, SC09028
    FROM         SC090300 AS SC090300_15
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     27 AS Expr1, SC09029
    FROM         SC090300 AS SC090300_14
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     28 AS Expr1, SC09030
    FROM         SC090300 AS SC090300_13
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     29 AS Expr1, SC09031
    FROM         SC090300 AS SC090300_12
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     30 AS Expr1, SC09032
    FROM         SC090300 AS SC090300_11
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     31 AS Expr1, SC09033
    FROM         SC090300 AS SC090300_10
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     32 AS Expr1, SC09034
    FROM         SC090300 AS SC090300_9
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     33 AS Expr1, SC09035
    FROM         SC090300 AS SC090300_8
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     34 AS Expr1, SC09036
    FROM         SC090300 AS SC090300_7
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     35 AS Expr1, SC09037
    FROM         SC090300 AS SC090300_6
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     36 AS Expr1, SC09038
    FROM         SC090300 AS SC090300_5
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     37 AS Expr1, SC09039
    FROM         SC090300 AS SC090300_4
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     38 AS Expr1, SC09040
    FROM         SC090300 AS SC090300_3
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     39 AS Expr1, SC09041
    FROM         SC090300 AS SC090300_2
    WHERE     (SC09001 = 'RUS')
    UNION
    SELECT     40 AS Expr1, SC09042
    FROM         SC090300 AS SC090300_1
    WHERE     (SC09001 = 'RUS')) AS View_1 ON #_MyPCOrder.PC03009 = View_1.num

--------------------Прайс товаров-----------------------------------------------
UPDATE #_MyPCOrder
SET Price = Round(View_1.Price,2)
FROM #_MyPCOrder Inner Join
	(SELECT dbo.SC010300.SC01001, 
		dbo.SC010300.SC01055 * dbo.SYCH0100.SYCH006 / @MyExchRate AS Price
	FROM dbo.SC010300 INNER JOIN
		dbo.SYCH0100 ON dbo.SC010300.SC01056 = dbo.SYCH0100.SYCH001
	WHERE (dbo.SYCH0100.SYCH004 <= @MyDate) AND 
		(dbo.SYCH0100.SYCH005 > @MyDate)) AS View_1 ON 
	#_MyPCOrder.PC03005 = View_1.SC01001

--------------------Сумма строки------------------------------------------------
UPDATE #_MyPCOrder
SET StrSum = Round(Price * QTY,2)

--------------------Сумма заказа------------------------------------------------
SELECT @MyOrderSum = Round(SUM(StrSum),2)
FROM #_MyPCOrder

UPDATE #_MyPCOrder
SET OrderSum = @MyOrderSum

/*
SELECT * 
FROM #_MyPCOrder
Order BY PC03005
*/
