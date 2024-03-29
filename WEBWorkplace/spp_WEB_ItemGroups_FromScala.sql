USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_ItemGroups_FromScala]    Дата сценария: 05/29/2015 12:40:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_ItemGroups_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по группам товаров                                            |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

--------------------------Загрузка отсутствующих-----------------------------------------
---загружаются только производители товаров, присутствующих в прайс - листах-------------
INSERT INTO tbl_WEB_ItemGroup
SELECT tbl_RexelProductCategory.CategoryNum AS Code, 
	tbl_RexelProductCategory.CategoryName AS Name, 
	'' AS WEBName, 
	1 AS RMStatus, 
	1 AS WEBStatus
FROM tbl_RexelProductCategory LEFT OUTER JOIN
    tbl_WEB_ItemGroup AS tbl_WEB_ItemGroup_1 ON 
	tbl_RexelProductCategory.CategoryNum = tbl_WEB_ItemGroup_1.Code
WHERE (tbl_WEB_ItemGroup_1.Code IS NULL)


--------------------------Обновление существующих----------------------------------------
UPDATE tbl_WEB_ItemGroup
SET Name = tbl_RexelProductCategory.CategoryName, 
	RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END,
	WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END 
FROM tbl_RexelProductCategory INNER JOIN
    tbl_WEB_ItemGroup ON tbl_RexelProductCategory.CategoryNum = tbl_WEB_ItemGroup.Code 
	AND tbl_RexelProductCategory.CategoryName <> tbl_WEB_ItemGroup.Name


--------------------------Удаление отсутствующих-----------------------------------------
UPDATE tbl_WEB_ItemGroup
SET RMStatus = 2, 
	WEBStatus = 2
FROM tbl_RexelProductCategory RIGHT OUTER JOIN
    tbl_WEB_ItemGroup ON tbl_RexelProductCategory.CategoryNum = tbl_WEB_ItemGroup.Code
WHERE (tbl_RexelProductCategory.CategoryNum IS NULL)