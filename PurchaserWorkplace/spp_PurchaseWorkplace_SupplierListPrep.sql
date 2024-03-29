USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_PurchaseWorkplace_SupplierListPrep]    Дата сценария: 05/11/2012 08:42:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_PurchaseWorkplace_SupplierListPrep]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Готовит список поставщиков                                                        |
|    разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@ActiveOrNot int,                                    -- (0) Выводить всех поставщиков или (1) только активных 
@WH nvarchar(6),                                     -- номер склада
@WhereCond1 nvarchar(250),                           -- Условие поиска 1
@WhereCond2 nvarchar(250)                            -- Условие поиска 2


WITH RECOMPILE
AS

DECLARE @MyWhereCond1 nvarchar(250)                    -- обработанное условие поиска 1
DECLARE @MyWhereCond2 nvarchar(250)                    -- обработанное условие поиска 2
SELECT @MyWhereCond1 = '%' + UPPER(@WhereCond1) + '%'
SELECT @MyWhereCond2 = '%' + UPPER(@WhereCond2) + '%'

-----------------очистка временных таблиц----------------------------------------------

IF exists(select * from tempdb..sysobjects where 
id = object_id(N'tempdb..#_MySuppliers') 
and xtype = N'U')
DROP TABLE #_MySuppliers

-----------------создание временных таблиц---------------------------------------------

CREATE TABLE #_MySuppliers(
	[PL01001] [nvarchar](10) COLLATE Cyrillic_General_BIN,            --Код поставщика
	[PL01002] [nvarchar](35) COLLATE Cyrillic_General_BIN,            --название поставщика
	[PL01003] [nvarchar](150) COLLATE Cyrillic_General_BIN,           --адрес поставщика
	[NonGroupOrders] float NULL,                                      --кол - во несгруппированных заказов
	[NonPlacedOrders] float NULL,                                     --кол - во неразмещенных заказов
	[NonConfirmedOrders] float NULL,                                  --кол - во неподтвержденных заказов
	[DelayedOrders] float NULL                                        --кол - во задолженных заказов
)

------Список поставщиков
INSERT INTO #_MySuppliers
SELECT PL01001 AS SuppCode,
	PL01002 AS SuppName,
	Ltrim(Rtrim(Ltrim(Rtrim(PL01003)) + ' ' + Ltrim(Rtrim(PL01004)) + ' ' + Ltrim(Rtrim(PL01005)) + ' ' + Ltrim(Rtrim(PL01006)))) AS SuppAddress,
	NULL AS NonGroupOrders,
	NULL AS NonPlacedOrders,
	NULL AS NonConfirmedOrders,
	NULL AS DelayedOrders
FROM PL010300

------кол - во несгруппированных заказов
UPDATE #_MySuppliers
SET NonGroupOrders = View_1.Expr1
FROM #_MySuppliers INNER JOIN
	(SELECT PC010300.PC01003, 
		COUNT(DISTINCT PC030300.PC03001) AS Expr1
    FROM PC030300 INNER JOIN
		PC010300 ON PC030300.PC03001 = PC010300.PC01001
        WHERE ((PC030300.PC03010 <> 0) OR (PC030300.PC03011 <> 0)) AND
			(PC010300.PC01002 <> 0) AND 
			(PC010300.PC01023 = @WH) AND
			(PC010300.PC01001 NOT IN (SELECT PC010300_2.PC01001
				FROM tbl_PurchaseWorkplace_ConsolidatedOrders INNER JOIN
					PC010300 AS PC010300_2 ON tbl_PurchaseWorkplace_ConsolidatedOrders.ID = PC010300_2.PC01052
                GROUP BY PC010300_2.PC01001)) 
	GROUP BY PC010300.PC01003) AS View_1 ON 
	#_MySuppliers.PL01001 = View_1.PC01003

------кол - во неразмещенных заказов
UPDATE #_MySuppliers
SET NonPlacedOrders = View_1.Expr1
FROM #_MySuppliers INNER JOIN
	(SELECT SupplierCode, COUNT(ID) AS Expr1
    FROM tbl_PurchaseWorkplace_ConsolidatedOrders
    WHERE (WH = @WH) AND (SupplierPlacedDate IS NULL)
    GROUP BY SupplierCode) AS View_1 ON 
#_MySuppliers.PL01001 = View_1.SupplierCode

------кол - во неподтвержденных заказов
UPDATE #_MySuppliers
SET NonConfirmedOrders = View_1.Expr1
FROM #_MySuppliers INNER JOIN
	(SELECT SupplierCode, COUNT(ID) AS Expr1
    FROM tbl_PurchaseWorkplace_ConsolidatedOrders
    WHERE (WH = @WH) AND (ConfirmedDate IS NULL)
    GROUP BY SupplierCode) AS View_1 ON 
#_MySuppliers.PL01001 = View_1.SupplierCode

------кол - во задолженных заказов
UPDATE #_MySuppliers
SET DelayedOrders = View_1.Expr1
FROM #_MySuppliers INNER JOIN
	(SELECT PC010300.PC01003 AS SuppCode, 
		COUNT(DISTINCT PC010300.PC01001) AS Expr1
    FROM PC030300 INNER JOIN
		PC010300 ON PC030300.PC03001 = PC010300.PC01001
    WHERE ((PC030300.PC03010 <> 0) OR
		(PC030300.PC03011 <> 0)) AND
		(PC030300.PC03029 = N'1') AND 
		(PC010300.PC01023 = @WH) AND
		(PC030300.PC03031 < GETDATE())
    GROUP BY PC010300.PC01003) AS View_1 ON 
	#_MySuppliers.PL01001 = View_1.SuppCode

------убираем NULL
UPDATE #_MySuppliers
SET NonGroupOrders = ISNULL(NonGroupOrders,0),
	NonPlacedOrders = ISNULL(NonPlacedOrders,0),
	NonConfirmedOrders = ISNULL(NonConfirmedOrders,0),
	DelayedOrders = ISNULL(DelayedOrders,0)

-----------------Вывод данных-------------------------------
if @ActiveOrNot = 0
BEGIN
	if Ltrim(Rtrim(@WhereCond1)) = ''
	BEGIN
		if Ltrim(Rtrim(@WhereCond2)) = ''
		BEGIN
			-----Оба условия поиска пусты
			SELECT * FROM #_MySuppliers
			Order BY PL01001
		END
		ELSE
		BEGIN
			-----условие поиска введено в окно 2
			SELECT * FROM #_MySuppliers
			WHERE (Upper(PL01001) Like @MyWhereCond2) OR
				(Upper(PL01002) Like @MyWhereCond2) OR
				(Upper(PL01003) Like @MyWhereCond2)
			Order BY PL01001
		END
	END
	ELSE
	BEGIN
		if Ltrim(Rtrim(@WhereCond2)) = ''
		BEGIN
			-----условие поиска введено в окно 1
			SELECT * FROM #_MySuppliers
			WHERE (Upper(PL01001) Like @MyWhereCond1) OR
				(Upper(PL01002) Like @MyWhereCond1) OR
				(Upper(PL01003) Like @MyWhereCond1)
			Order BY PL01001
		END
		ELSE
		BEGIN
			-----условие поиска введено в оба окна
			SELECT * FROM #_MySuppliers
			WHERE ((Upper(PL01001) Like @MyWhereCond1) AND
				(Upper(PL01001) Like @MyWhereCond2)) OR
				((Upper(PL01002) Like @MyWhereCond1) AND
				(Upper(PL01002) Like @MyWhereCond2)) OR
				((Upper(PL01003) Like @MyWhereCond1) AND
				(Upper(PL01003) Like @MyWhereCond2))
			Order BY PL01001
		END
	END
END
ELSE
BEGIN
	if Ltrim(Rtrim(@WhereCond1)) = ''
	BEGIN
		if Ltrim(Rtrim(@WhereCond2)) = ''
		BEGIN
			-----Оба условия поиска пусты
			SELECT * FROM #_MySuppliers
			WHERE ((NonGroupOrders <> 0) OR
				(NonPlacedOrders <> 0) OR
				(NonConfirmedOrders <> 0) OR
				(DelayedOrders <> 0))
			Order BY PL01001
		END
		ELSE
		BEGIN
			-----условие поиска введено в окно 2
			SELECT * FROM #_MySuppliers
			WHERE ((Upper(PL01001) Like @MyWhereCond2) OR
				(Upper(PL01002) Like @MyWhereCond2) OR
				(Upper(PL01003) Like @MyWhereCond2)) AND
				((NonGroupOrders <> 0) OR
				(NonPlacedOrders <> 0) OR
				(NonConfirmedOrders <> 0) OR
				(DelayedOrders <> 0))
			Order BY PL01001
		END
	END
	ELSE
	BEGIN
		if Ltrim(Rtrim(@WhereCond2)) = ''
		BEGIN
			-----условие поиска введено в окно 1
			SELECT * FROM #_MySuppliers
			WHERE ((Upper(PL01001) Like @MyWhereCond1) OR
				(Upper(PL01002) Like @MyWhereCond1) OR
				(Upper(PL01003) Like @MyWhereCond1)) AND
				((NonGroupOrders <> 0) OR
				(NonPlacedOrders <> 0) OR
				(NonConfirmedOrders <> 0) OR
				(DelayedOrders <> 0))
			Order BY PL01001
		END
		ELSE
		BEGIN
			-----условие поиска введено в оба окна
			SELECT * FROM #_MySuppliers
			WHERE (((Upper(PL01001) Like @MyWhereCond1) AND
				(Upper(PL01001) Like @MyWhereCond2)) OR
				((Upper(PL01002) Like @MyWhereCond1) AND
				(Upper(PL01002) Like @MyWhereCond2)) OR
				((Upper(PL01003) Like @MyWhereCond1) AND
				(Upper(PL01003) Like @MyWhereCond2))) AND
				((NonGroupOrders <> 0) OR
				(NonPlacedOrders <> 0) OR
				(NonConfirmedOrders <> 0) OR
				(DelayedOrders <> 0))
			Order BY PL01001
		END
	END
END
