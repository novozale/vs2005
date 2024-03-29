USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_CardDiscounts_FromDB]    Дата сценария: 05/29/2015 12:39:33 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_CardDiscounts_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по общим скидкам клиентам, работающим через WEB               |
|    из промежуточной БД обмена с WEB в файлы                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/


WITH RECOMPILE
AS


SELECT 
	'"' + Replace(Ltrim(Rtrim(Code)),'"','""') + '"', 
	'"' + Replace(CASE WHEN Ltrim(Rtrim(Discount)) = '' THEN '0' ELSE Ltrim(Rtrim(Discount)) END,'"','""') + '"', 
	'"' + Replace(Ltrim(Rtrim(Name)),'"','""') + '"'
FROM tbl_WEB_Clients
WHERE (WorkOverWEB = 1)
ORDER BY Code
