USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Manufacturers_FromDB]    Дата сценария: 05/29/2015 12:41:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Manufacturers_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по производителям                                             |
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
		'"' + Replace(CONVERT(nvarchar(10), tbl_WEB_Items.ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(tbl_WEB_Items.CountryCode)), 4),'"','""') + '"', 
        '"' + Replace(Ltrim(Rtrim(tbl_WEB_Manufacturers.Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_Items.Country)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_Manufacturers.Rezerv1)),'"','""') + '"', 
		1 AS RMStatus
	FROM tbl_WEB_Items INNER JOIN
		tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID
	WHERE (tbl_WEB_Items.SubGroupCode <> N'')
		AND (tbl_WEB_Manufacturers.WEBStatus <> 2)
	ORDER BY '"' + Replace(CONVERT(nvarchar(10), tbl_WEB_Items.ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(tbl_WEB_Items.CountryCode)), 4),'"','""') + '"'

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Manufacturers
		SET RMStatus = 0 
		WHERE (WEBStatus <> 2)
	END
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT DISTINCT 
		'"' + Replace(CONVERT(nvarchar(10), tbl_WEB_Items.ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(tbl_WEB_Items.CountryCode)), 4),'"','""') + '"', 
        '"' + Replace(Ltrim(Rtrim(tbl_WEB_Manufacturers.Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_Items.Country)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_Manufacturers.Rezerv1)),'"','""') + '"', 
		tbl_WEB_Manufacturers.RMStatus
	FROM tbl_WEB_Items INNER JOIN
		tbl_WEB_Manufacturers ON tbl_WEB_Items.ManufacturerCode = tbl_WEB_Manufacturers.ID
	WHERE (tbl_WEB_Items.SubGroupCode <> N'')
		AND (tbl_WEB_Manufacturers.WEBStatus <> 0)
	ORDER BY '"' + Replace(CONVERT(nvarchar(10), tbl_WEB_Items.ManufacturerCode) + RIGHT('0000' + LTRIM(RTRIM(tbl_WEB_Items.CountryCode)), 4),'"','""') + '"'

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Manufacturers
		SET RMStatus = 0 
		WHERE (WEBStatus <> 0)
	END
END