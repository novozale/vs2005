USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Price_FromDB]    Дата сценария: 05/29/2015 12:42:09 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Price_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по прайс листам                                               |
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
	[CustomerCode] [nvarchar](35),							-- код клиента в Scala
	[Code] [nvarchar](35),									-- код товара в Scala
	[GroupCode] [nvarchar](50) NULL,						-- код группы продукта
	[SubGroupCode] [nvarchar](50) NULL,						-- код подгруппы продукта
	[DiscountType] [nvarchar] (100) NULL,					-- тип скидки или согласованный ассортимент
	[Discount] [numeric] (28,8)	NULL,						-- размер скидки (%)
	[Price] [numeric] (28,8) NULL							-- цена в рублях
)


--=================Загрузка запасов======================================================
INSERT INTO #Result
SELECT View_1.Code AS CustomerCode, 
	tbl_WEB_Items.Code, 
	tbl_WEB_Items.GroupCode, 
	tbl_WEB_Items.SubGroupCode, 
	'' AS DiscountType, 
    0 AS Discount, 
	NULL AS Price
FROM tbl_WEB_Items CROSS JOIN
	(SELECT Code
	FROM tbl_WEB_Clients
    WHERE (WorkOverWEB = 1)
    UNION
    SELECT '000000' AS Code) AS View_1
WHERE (Ltrim(Rtrim(tbl_WEB_Items.SubGroupCode)) <> N'')


--=================прайс с учетом общей скидки клиенту====================================
-------------------базовый прайс (00)-----------------------------------------------------
UPDATE #Result
SET Price = ROUND(SC390300.SC39005 * SYCH0100.SYCH006, 2),
	Discount = 0
FROM #Result INNER JOIN
	SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001
WHERE (SC390300.SC39002 = '00')
	AND (#Result.CustomerCode = '000000') 
	AND (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


-------------------базовые прайс листы покупателей----------------------------------------
UPDATE #Result
SET  Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - CONVERT(numeric(28, 8), 
		CASE WHEN tbl_WEB_Clients.Discount = '' THEN '0' ELSE tbl_WEB_Clients.Discount END)) / 100, 2),
	Discount = CONVERT(numeric(28, 8), CASE WHEN tbl_WEB_Clients.Discount = '' THEN '0' 
		ELSE tbl_WEB_Clients.Discount END),
	DiscountType = N'общая скидка клиенту'
FROM #Result INNER JOIN
	SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_Clients ON #Result.CustomerCode = tbl_WEB_Clients.Code 
	AND SC390300.SC39002 = tbl_WEB_Clients.BasePrice
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


--=================прайс с учетом скидок на группы========================================
UPDATE #Result
SET Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountGroup.Discount) / 100, 2),
	Discount = tbl_WEB_DiscountGroup.Discount,
	DiscountType = N'скидка на группу'
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountGroup ON #Result.GroupCode = tbl_WEB_DiscountGroup.GroupCode AND 
    #Result.CustomerCode = tbl_WEB_DiscountGroup.ClientCode INNER JOIN
    tbl_WEB_Clients ON #Result.CustomerCode = tbl_WEB_Clients.Code 
	AND SC390300.SC39002 = tbl_WEB_Clients.BasePrice
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


--=================прайс с учетом скидок на подгруппы=====================================
UPDATE #Result
SET Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountSubgroup.Discount) / 100, 2),
	Discount = tbl_WEB_DiscountSubgroup.Discount,
	DiscountType = N'скидка на подгруппу' 
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountSubgroup ON #Result.GroupCode = tbl_WEB_DiscountSubgroup.GroupCode AND 
    #Result.SubGroupCode = tbl_WEB_DiscountSubgroup.SubgroupCode 
	AND #Result.CustomerCode = tbl_WEB_DiscountSubgroup.ClientCode INNER JOIN
    tbl_WEB_Clients ON #Result.CustomerCode = tbl_WEB_Clients.Code 
	AND SC390300.SC39002 = tbl_WEB_Clients.BasePrice
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


--=================прайс с учетом скидок на товар=========================================
UPDATE #Result
SET  Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountItem.Discount) / 100, 2),
	Discount = tbl_WEB_DiscountItem.Discount,
	DiscountType = N'скидка на товар'
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountItem ON #Result.Code = tbl_WEB_DiscountItem.ItemCode 
	AND #Result.CustomerCode = tbl_WEB_DiscountItem.ClientCode INNER JOIN
	tbl_WEB_Clients ON #Result.CustomerCode = tbl_WEB_Clients.Code 
	AND SC390300.SC39002 = tbl_WEB_Clients.BasePrice
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


--=================прайс с учетом согласованного ассортимента=============================
UPDATE #Result
SET Price = ROUND(tbl_WEB_AgreedRange.AgreedPrice, 2),
	Discount = 0,
	DiscountType = N'согласованный ассортимент'
FROM #Result INNER JOIN
    tbl_WEB_AgreedRange ON #Result.Code = tbl_WEB_AgreedRange.ItemCode AND 
    #Result.CustomerCode = tbl_WEB_AgreedRange.ClientCode INNER JOIN
    SYCH0100 ON tbl_WEB_AgreedRange.CurrCode = SYCH0100.SYCH001
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE())


--=================прайс с учетом НДС=====================================================
UPDATE #Result
SET Price = Price * (100 + View_1.VAT) /100
FROM #Result INNER JOIN
	(SELECT SC010300.SC01001 AS Code, 
		CONVERT(numeric(28, 8), CASE WHEN Ltrim(Rtrim(SY290300.SY29003)) = '' THEN '0' ELSE SY290300.SY29003 END) AS VAT
    FROM SC010300 INNER JOIN
    SY290300 ON SC010300.SC01144 = CONVERT(integer, SY290300.SY29001)) AS View_1 ON 
	#Result.Code = View_1.Code



--=================Выгрузка результата====================================================
SELECT 
	'"' + Replace(Ltrim(Rtrim(CustomerCode)),'"','""') + '"',
	'"' + Replace(Ltrim(Rtrim(Code)),'"','""') + '"',
	Price,
	CASE WHEN DiscountType = 'согласованный ассортимент' THEN 1 ELSE 0 END AS AgreedRange
FROM #Result
ORDER BY CustomerCode,
	Code
