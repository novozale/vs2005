USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Salesmans_FromDB]    Дата сценария: 05/29/2015 12:43:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Salesmans_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продавцам                                                  |
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
	SELECT '"' + Replace(Ltrim(Rtrim(Code)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Email)),'"','""') + '"', 
		'"' + Replace(CASE WHEN CONVERT(nvarchar(20),City) = '0' THEN '' ELSE CONVERT(nvarchar(20),City) END,'"','""') + '"' AS City, 
		OfficeLeader, 
		OnDuty, 
		'"' + Replace(Ltrim(Rtrim(Rezerv1)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Rezerv2)),'"','""') + '"', 
		1 AS RMStatus
	FROM tbl_WEB_Salesmans 
	WHERE (IsActive = 1)
		AND (WEBStatus <> 2)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Salesmans
		SET RMStatus = 0 
		WHERE (WEBStatus <> 2)
	END
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT '"' + Replace(Ltrim(Rtrim(Code)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Name)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Email)),'"','""') + '"', 
		'"' + Replace(CASE WHEN CONVERT(nvarchar(20),City) = '0' THEN '' ELSE CONVERT(nvarchar(20),City) END,'"','""') + '"' AS City, 
		OfficeLeader, 
		OnDuty, 
		'"' + Replace(Ltrim(Rtrim(Rezerv1)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Rezerv2)),'"','""') + '"', 
		RMStatus
	FROM tbl_WEB_Salesmans
	WHERE ((IsActive = 1) 
		AND (WEBStatus <> 0))
		OR (WEBStatus = 2)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_Salesmans
		SET RMStatus = 0 
		WHERE (WEBStatus <> 0)
	END
END