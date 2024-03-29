USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_CardSalesDetails_FromDB]    Дата сценария: 05/29/2015 12:39:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_CardSalesDetails_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продажам (строки)                                          |
|    из промежуточной БД обмена с WEB в файлы                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/
@FullUploadFlag integer									-- флаг 1 - полная выгрузка
														--		0 - частичная выгрузка

WITH RECOMPILE
AS

SET NOCOUNT ON

if @FullUploadFlag = 1
BEGIN					---------полная выгрузка
	SELECT     
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ClientCode)),'"','""') + '"', 
		'"' + CONVERT(nvarchar(30),tbl_WEB_CardSalesDetails.OrderDate,104) + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.OrderNum)),'"','""') + '"', 
        '"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.StrNum)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ItemCode)),'"','""') + '"', 
		tbl_WEB_CardSalesDetails.OrderedQTY, 
        tbl_WEB_CardSalesDetails.ReadyQTY, 
		tbl_WEB_CardSalesDetails.ShippedQTY, 
		tbl_WEB_CardSalesDetails.Price,
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ItemName)),'"','""') + '"' AS ItemName
	FROM tbl_WEB_CardSalesDetails INNER JOIN
        tbl_WEB_CardSales ON 
		tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum
	WHERE (tbl_WEB_CardSales.WEBStatus <> 2)
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT     
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ClientCode)),'"','""') + '"', 
		'"' + CONVERT(nvarchar(30),tbl_WEB_CardSalesDetails.OrderDate,104) + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.OrderNum)),'"','""') + '"', 
        '"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.StrNum)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ItemCode)),'"','""') + '"', 
		tbl_WEB_CardSalesDetails.OrderedQTY, 
        tbl_WEB_CardSalesDetails.ReadyQTY, 
		tbl_WEB_CardSalesDetails.ShippedQTY, 
		tbl_WEB_CardSalesDetails.Price,
		'"' + Replace(Ltrim(Rtrim(tbl_WEB_CardSalesDetails.ItemName)),'"','""') + '"' AS ItemName
	FROM tbl_WEB_CardSalesDetails INNER JOIN
        tbl_WEB_CardSales ON 
		tbl_WEB_CardSalesDetails.OrderNum = tbl_WEB_CardSales.OrderNum
	WHERE (tbl_WEB_CardSales.WEBStatus <> 0)
END