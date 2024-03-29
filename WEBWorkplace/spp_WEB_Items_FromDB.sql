USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Items_FromDB]    Дата сценария: 05/29/2015 12:41:19 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Items_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продуктам                                                  |
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
	SELECT 
		'"' + Replace(LTRIM(RTRIM(Code)),'"','""') + '"', 
		'"' + Replace(LTrim(Rtrim(ManufacturerItemCode)),'"','""') + '"', 
		'"' + Replace(LTRIM(RTRIM(GroupCode)),'"','""') + '"', 
		'"' + Replace(LTRIM(RTRIM(GroupCode)) + LTRIM(RTRIM(SubGroupCode)),'"','""') + '"', 
		'"' + Replace(CASE WHEN Ltrim(Rtrim(WEBName)) = '' THEN Ltrim(Rtrim(Name)) ELSE Ltrim(Rtrim(WEBName)) END,'"','""') + '"',
		'"' + Replace(Ltrim(Rtrim(Description)),'"','""') + '"',
		'"' + Replace(CONVERT(nvarchar(10), ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(CountryCode)), 4),'"','""') + '"', 
		CASE WHEN WHAssortiment = 0 THEN 2 ELSE WHAssortiment END, 
		'"' + Replace(LTRIM(RTRIM(UOM)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Rezerv)),'"','""') + '"', 
        1 
	FROM tbl_WEB_Items
	WHERE (SubGroupCode <> N'')
	ORDER BY Code	

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Items
		SET RMStatus = 0 
		WHERE (WEBStatus <> 2)
	END
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT 
		'"' + Replace(LTRIM(RTRIM(Code)),'"','""') + '"', 
		'"' + Replace(LTrim(Rtrim(ManufacturerItemCode)),'"','""') + '"', 
		'"' + Replace(LTRIM(RTRIM(GroupCode)),'"','""') + '"', 
		'"' + Replace(LTRIM(RTRIM(GroupCode)) + LTRIM(RTRIM(SubGroupCode)),'"','""') + '"', 
		'"' + Replace(CASE WHEN Ltrim(Rtrim(WEBName)) = '' THEN Ltrim(Rtrim(Name)) ELSE Ltrim(Rtrim(WEBName)) END,'"','""') + '"',
		'"' + Replace(Ltrim(Rtrim(Description)),'"','""') + '"',
		'"' + Replace(CONVERT(nvarchar(10), ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(CountryCode)), 4),'"','""') + '"', 
		CASE WHEN WHAssortiment = 0 THEN 2 ELSE WHAssortiment END, 
		'"' + Replace(LTRIM(RTRIM(UOM)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Rezerv)),'"','""') + '"', 
        RMStatus 
	FROM tbl_WEB_Items
	WHERE (SubGroupCode <> N'')
		AND (RMStatus <> 0)
	ORDER BY Code

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Items
		SET RMStatus = 0 
		WHERE (WEBStatus <> 0)
	END
END