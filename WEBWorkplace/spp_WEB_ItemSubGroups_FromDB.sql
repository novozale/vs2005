USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_ItemSubGroups_FromDB]    Дата сценария: 05/29/2015 12:41:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_ItemSubGroups_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по подгруппам продуктов                                       |
|    из промежуточной БД обмена с WEB в файлы                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/
@FullUploadFlag integer,								-- флаг 1 - полная выгрузка
														--		0 - частичная выгрузка
@MarkFlag integer										-- флаг 1 - помечать, что выгружено
														--		0 - не помечать

WITH RECOMPILE
AS


if @FullUploadFlag = 1
BEGIN					---------полная выгрузка
	SELECT DISTINCT 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.SubgroupID)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.GroupCode)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Description)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Rezerv1)),'"','""') + '"', 
		1 AS RMStatus
	FROM tbl_WEB_ItemSubGroup INNER JOIN
		tbl_WEB_Items ON tbl_WEB_ItemSubGroup.GroupCode = tbl_WEB_Items.GroupCode AND 
        tbl_WEB_ItemSubGroup.SubgroupCode = tbl_WEB_Items.SubGroupCode
	WHERE (tbl_WEB_ItemSubGroup.WEBStatus <> 2)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_ItemSubGroup
		SET RMStatus = 0 
		WHERE (WEBStatus <> 2)
	END
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT DISTINCT
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.SubgroupID)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.GroupCode)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Description)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_ItemSubGroup.Rezerv1)),'"','""') + '"', 
		tbl_WEB_ItemSubGroup.RMStatus
	FROM tbl_WEB_ItemSubGroup INNER JOIN
		tbl_WEB_Items ON tbl_WEB_ItemSubGroup.GroupCode = tbl_WEB_Items.GroupCode AND 
        tbl_WEB_ItemSubGroup.SubgroupCode = tbl_WEB_Items.SubGroupCode
	WHERE (tbl_WEB_ItemSubGroup.WEBStatus <> 0)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_ItemSubGroup
		SET RMStatus = 0 
		WHERE (WEBStatus <> 0)
	END
END