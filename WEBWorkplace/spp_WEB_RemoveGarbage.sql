USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_RemoveGarbage]    Дата сценария: 05/29/2015 12:42:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_RemoveGarbage]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    удаление мусора в промежуточной БД                                                |
|                                                                                      |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS


-------------продавцы с несуществующим городом-------------------------------------------
UPDATE tbl_WEB_Salesmans
SET City = 0
WHERE (City NOT IN
	(SELECT DISTINCT ID
	FROM tbl_WEB_Cities)) 
	AND (City <> 0)


-------------Подгруппы товара без группы-------------------------------------------------
DELETE FROM tbl_WEB_ItemSubGroup
WHERE (Ltrim(Rtrim(GroupCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_ItemGroup)) 
	AND (RMStatus <> 2) 
	AND (WEBStatus <> 2)


-------------Товары без кода группы------------------------------------------------------
DELETE FROM tbl_WEB_Items
WHERE (Ltrim(Rtrim(GroupCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_ItemGroup)) 
	AND (RMStatus <> 2) 
	AND (WEBStatus <> 2)


-------------Товары с несуществующим кодом подгруппы-------------------------------------
UPDATE tbl_WEB_Items
SET SubGroupCode = N''
WHERE ((Ltrim(Rtrim(GroupCode)) + Ltrim(Rtrim(SubGroupCode))) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(SubgroupID)) AS SubgroupID
    FROM tbl_WEB_ItemSubGroup)) 
	AND (SubGroupCode <> '')


-------------Скидки на несуществующие группы---------------------------------------------
DELETE FROM tbl_WEB_DiscountGroup
WHERE (Ltrim(Rtrim(GroupCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_ItemGroup))


-------------Скидки на группы на несуществующих и неработающих через WEB клиентов--------
DELETE FROM tbl_WEB_DiscountGroup
WHERE (Ltrim(Rtrim(ClientCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
	FROM tbl_WEB_Clients
    WHERE (WorkOverWEB = 1)))


-------------Скидки на несуществующие подгруппы------------------------------------------
DELETE FROM tbl_WEB_DiscountSubgroup
WHERE ((Ltrim(Rtrim(GroupCode)) + Ltrim(Rtrim(SubgroupCode))) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(SubgroupID)) AS SubgroupID
    FROM tbl_WEB_ItemSubGroup))


-------------Скидки на подгруппы на несуществующих и неработающих через WEB клиентов-----
DELETE FROM tbl_WEB_DiscountSubgroup
WHERE (Ltrim(Rtrim(ClientCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code))
    FROM tbl_WEB_Clients
    WHERE (WorkOverWEB = 1)))


-------------Скидки на несуществующие товары или товары без подгрупп---------------------
DELETE FROM tbl_WEB_DiscountItem
WHERE (Ltrim(Rtrim(ItemCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_Items
    WHERE (SubGroupCode <> N'')))


-------------Скидки на товары на несуществующих и неработающих через WEB клиентов--------
DELETE FROM tbl_WEB_DiscountItem
WHERE (Ltrim(Rtrim(ClientCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_Clients
    WHERE (WorkOverWEB = 1)))


-------------Согласованный ассортимент на несуществующие товары или товары без подгрупп--
DELETE FROM tbl_WEB_AgreedRange
WHERE (Ltrim(Rtrim(ItemCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code
    FROM tbl_WEB_Items
    WHERE (SubGroupCode <> N'')))


-------------Согласованный ассортимент на несуществующих и неработающих через WEB клиентов
DELETE FROM tbl_WEB_AgreedRange
WHERE (Ltrim(Rtrim(ClientCode)) NOT IN
	(SELECT DISTINCT Ltrim(Rtrim(Code)) AS Code	
	FROM tbl_WEB_Clients
    WHERE (WorkOverWEB = 1)))