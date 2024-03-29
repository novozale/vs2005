USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_ShopAvailability_FromDB]    Дата сценария: 05/29/2015 12:43:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_ShopAvailability_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по доступности товаров на складах                             |
|    из промежуточной БД обмена с WEB в файлы                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

SET NOCOUNT ON

----------------- Очистка временных таблиц -----------------------------------------------------------------
IF exists(select * from tempdb..sysobjects where id = object_id(N'tempdb..#Result') and xtype = N'U')
DROP TABLE #Result

------------------Создание временных таблиц-----------------------------------------------------------------
CREATE TABLE #Result(
	[WH] [nvarchar](10),							        -- код Склада в Scala
	[Code] [nvarchar](35),									-- код товара в Scala
	[QTY] [numeric] (28,8) NULL							-- цена в рублях
)


--=================Загрузка запасов======================================================
INSERT INTO #Result
SELECT SC230300.SC23001, 
	tbl_WEB_Items.Code, 
	NULL AS QTY
FROM tbl_WEB_Items CROSS JOIN
	SC230300
WHERE (LTRIM(RTRIM(tbl_WEB_Items.SubGroupCode)) <> N'') 
	AND (SC230300.SC23006 = N'1')


--=================Количество на складах==================================================
UPDATE #Result
SET QTY = SC030300.SC03003 - SC030300.SC03004 - SC030300.SC03005
FROM #Result INNER JOIN
	SC030300 ON #Result.Code = SC030300.SC03001 
	AND #Result.WH = SC030300.SC03002

UPDATE #Result
SET QTY = ISNULL(QTY,0)


--=================Выгрузка результата====================================================
SELECT 
	'"' + Replace(Ltrim(Rtrim(WH)),'"','""') + '"',
	'"' + Replace(Ltrim(Rtrim(Code)),'"','""') + '"',
	QTY
FROM #Result
ORDER BY Code,
	WH
