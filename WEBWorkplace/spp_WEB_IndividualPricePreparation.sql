USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_IndividualPricePreparation]    Дата сценария: 05/29/2015 12:40:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_IndividualPricePreparation]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    подготовка информации по индивидуальному прайс листу клиента                      |
|    из промежуточной БД обмена с WEB                                                  |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/
@ClientCode nvarchar(50),									-- код клиента
@SubgroupFlag bit											-- флаг выгрузки продуктов 0 - выгружать все
																-- 1 - выгружать только с подгруппами

WITH RECOMPILE
AS
SET NOCOUNT ON 

DECLARE @MyClientDiscount [numeric] (28,8);					-- Скидка клиенту в целом
DECLARE @MyBasePrice [nvarchar] (10);						-- базовый прайс, назначенный клиенту

----------------- Очистка временных таблиц -----------------------------------------------------------------
IF exists(select * from tempdb..sysobjects where id = object_id(N'tempdb..#Result') and xtype = N'U')
DROP TABLE #Result

------------------Создание временных таблиц-----------------------------------------------------------------
CREATE TABLE #Result(
	[Code] [nvarchar](35),									-- код товара в Scala
	[GroupCode] [nvarchar](50) NULL,						-- код группы продукта
	[GroupName] [nvarchar](100) NULL,						-- имя группы продукта
	[SubGroupCode] [nvarchar](50) NULL,						-- код подгруппы продукта
	[SubgroupName] [nvarchar](100) NULL,					-- имя подгруппы продукта
	[Name] [nvarchar](100) NULL,							-- имя товара
	[ManufacturerCode] [bigint] NULL,						-- код производителя
	[ManufacturerName] [nvarchar](200) NULL,				-- имя производителя
	[ManufacturerItemCode] [nvarchar](35) NULL,				-- код товара производителя
	[DiscountType] [nvarchar] (100) NULL,					-- тип скидки или согласованный ассортимент
	[Discount] [numeric] (28,8)	NULL,						-- размер скидки (%)
	[PriCost] [numeric] (28,8) NULL,						-- расчетная себестоимость по базовому прайсу (руб)
	[Price] [numeric] (28,8) NULL,							-- цена в рублях
	[Margin] [numeric] (28,8) NULL							-- расчетная маржа (%)
)


--=================Загрузка запасов======================================================
IF @SubgroupFlag = 0
BEGIN		----------выгрузка всех продуктов
	INSERT INTO #Result
	SELECT Code, 
		GroupCode, 
		'' AS GroupName, 
		SubGroupCode, 
		'' AS SubGroupName, 
		Name, 
		ManufacturerCode, 
		'' AS ManufacturerName, 
        ManufacturerItemCode, 
		'' AS DiscountType, 
		0 AS Discount, 
		NULL AS PriCost, 
		NULL AS Price, 
		NULL AS Margin
	FROM tbl_WEB_Items
	WHERE (WEBStatus <> 2)
END

ELSE
BEGIN		----------выгрузка только продуктов с подгруппами
	INSERT INTO #Result
	SELECT Code, 
		GroupCode, 
		'' AS GroupName, 
		SubGroupCode, 
		'' AS SubGroupName, 
		Name, 
		ManufacturerCode, 
		'' AS ManufacturerName, 
        ManufacturerItemCode, 
		'' AS DiscountType, 
		0 AS Discount, 
		NULL AS PriCost, 
		NULL AS Price, 
		NULL AS Margin
	FROM tbl_WEB_Items
	WHERE (Ltrim(Rtrim(SubGroupCode)) <> N'')
		AND (WEBStatus <> 2)
END


--=================Названия групп продуктов===============================================
UPDATE #Result
SET GroupName = tbl_WEB_ItemGroup.Name
FROM #Result INNER JOIN
	tbl_WEB_ItemGroup ON #Result.GroupCode = tbl_WEB_ItemGroup.Code


--=================Названия подгрупп продуктов============================================
UPDATE #Result
SET SubgroupName = tbl_WEB_ItemSubGroup.Name
FROM #Result INNER JOIN
	tbl_WEB_ItemSubGroup ON #Result.SubGroupCode = tbl_WEB_ItemSubGroup.SubgroupCode AND 
    #Result.GroupCode = tbl_WEB_ItemSubGroup.GroupCode


--=================Названия производителей================================================
UPDATE #Result
SET ManufacturerName = tbl_WEB_Manufacturers.Name
FROM #Result INNER JOIN
	tbl_WEB_Manufacturers ON #Result.ManufacturerCode = tbl_WEB_Manufacturers.ID



--=================расчетная себестоимость (руб)==========================================
UPDATE #Result
SET PriCost = Round(SC010300.SC01053, 2)
FROM #Result INNER JOIN
	SC010300 ON #Result.Code = SC010300.SC01001



--=================прайс с учетом общей скидки клиенту====================================
SELECT @MyClientDiscount = CASE WHEN Ltrim(Rtrim(Discount)) = '' THEN 0 ELSE CONVERT(numeric(28,8),Discount) END, 
	@MyBasePrice = BasePrice
FROM tbl_WEB_Clients
WHERE (Code = @ClientCode)

PRINT @MyClientDiscount

IF @MyClientDiscount = 0
BEGIN			--грузим базовый прайс в рублях без скидок
	UPDATE #Result
	SET Price = ROUND(SC390300.SC39005 * SYCH0100.SYCH006, 2)
	FROM #Result INNER JOIN
		SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
        SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001
	WHERE (SC390300.SC39002 = @MyBasePrice) 
		AND (SYCH0100.SYCH004 <= GETDATE()) 
		AND (SYCH0100.SYCH005 > GETDATE())
END
ELSE
BEGIN			--грузим базовый прайс в рублях со скидкой
	UPDATE #Result
	SET DiscountType = N'общая скидка клиенту', 
		Discount = @MyClientDiscount, 
		Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - @MyClientDiscount) / 100, 2)
	FROM #Result INNER JOIN
		SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
        SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001
	WHERE (SC390300.SC39002 = @MyBasePrice) 
		AND (SYCH0100.SYCH004 <= GETDATE()) 
		AND (SYCH0100.SYCH005 > GETDATE())
END


--=================прайс с учетом скидок на группы========================================
UPDATE #Result
SET DiscountType = N'скидка на группу', 
	Discount = tbl_WEB_DiscountGroup.Discount, 
	Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountGroup.Discount) / 100, 2)
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountGroup ON #Result.GroupCode = tbl_WEB_DiscountGroup.GroupCode
WHERE (SC390300.SC39002 = @MyBasePrice) 
	AND (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE()) 
	AND (tbl_WEB_DiscountGroup.ClientCode = @ClientCode)


--=================прайс с учетом скидок на подгруппы=====================================
UPDATE #Result
SET DiscountType = N'скидка на подгруппу', 
	Discount = tbl_WEB_DiscountSubgroup.Discount, 
    Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountSubgroup.Discount) / 100, 2)
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountSubgroup ON #Result.GroupCode = tbl_WEB_DiscountSubgroup.GroupCode AND 
    #Result.SubGroupCode = tbl_WEB_DiscountSubgroup.SubgroupCode
WHERE (SC390300.SC39002 = @MyBasePrice) 
	AND (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE()) 
	AND (tbl_WEB_DiscountSubgroup.ClientCode = @ClientCode)


--=================прайс с учетом скидок на товар=========================================
UPDATE #Result
SET DiscountType = N'скидка на товар', 
	Discount = tbl_WEB_DiscountItem.Discount, 
    Price = ROUND((SC390300.SC39005 * SYCH0100.SYCH006) * (100 - tbl_WEB_DiscountItem.Discount) / 100, 2)
FROM #Result INNER JOIN
    SC390300 ON #Result.Code = SC390300.SC39001 INNER JOIN
    SYCH0100 ON SC390300.SC39003 = SYCH0100.SYCH001 INNER JOIN
    tbl_WEB_DiscountItem ON #Result.Code = tbl_WEB_DiscountItem.ItemCode 
WHERE (SC390300.SC39002 = @MyBasePrice) 
	AND (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE()) 
	AND (tbl_WEB_DiscountItem.ClientCode = @ClientCode)


--=================прайс с учетом согласованного ассортимента=============================
UPDATE #Result
SET DiscountType = N'согласованный ассортимент', 
	Discount = 0, 
	Price = ROUND(tbl_WEB_AgreedRange.AgreedPrice, 2)
FROM #Result INNER JOIN
	tbl_WEB_AgreedRange ON #Result.Code = tbl_WEB_AgreedRange.ItemCode INNER JOIN
    SYCH0100 ON tbl_WEB_AgreedRange.CurrCode = SYCH0100.SYCH001
WHERE (SYCH0100.SYCH004 <= GETDATE()) 
	AND (SYCH0100.SYCH005 > GETDATE()) 
	AND (tbl_WEB_AgreedRange.ClientCode = @ClientCode)


--=================Маржа индивидуального прайс листа======================================
UPDATE #Result
SET Margin = CASE WHEN Price = 0 THEN 0 ELSE Round((Price - PriCost) / Price * 100, 2) END



--=================Выгрузка результата====================================================
SELECT Code,
	Name,
	ManufacturerCode,
	ManufacturerName,
	ManufacturerItemCode,
	DiscountType,
	Discount,
	PriCost,
	Price,
	Margin
FROM #Result
ORDER BY Code