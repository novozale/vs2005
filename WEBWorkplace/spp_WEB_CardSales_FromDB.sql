USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_CardSales_FromDB]    Дата сценария: 05/29/2015 12:40:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_CardSales_FromDB]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продажам (заголовки)                                       |
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

SET NOCOUNT ON

if @FullUploadFlag = 1
BEGIN					---------полная выгрузка
	SELECT 
		'"' + Replace(Ltrim(Rtrim(ClientCode)),'"','""') + '"', 
		'"' + CONVERT(nvarchar(30),OrderDate,104) + '"', 
		'"' + Replace(Ltrim(Rtrim(OrderNum)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Discount)),'"','""') + '"', 
		OrderSumm, 
		1 AS Status, 
		ShipmentState, 
		OrderState, 
		'"' + Replace(Ltrim(Rtrim(WEBOrderNum)),'"','""') + '"'
	FROM tbl_WEB_CardSales
	WHERE (WEBStatus <> 2)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_CardSales
		SET RMStatus = 0 
		WHERE (WEBStatus <> 2)
	END
END
ELSE					---------частичная выгрузка
BEGIN
	SELECT 
		'"' + Replace(Ltrim(Rtrim(ClientCode)),'"','""') + '"', 
		'"' + CONVERT(nvarchar(30),OrderDate,104) + '"', 
		'"' + Replace(Ltrim(Rtrim(OrderNum)),'"','""') + '"', 
		'"' + Replace(Ltrim(Rtrim(Discount)),'"','""') + '"', 
		OrderSumm, 
		RMStatus AS Status, 
		ShipmentState, 
		OrderState, 
		'"' + Replace(Ltrim(Rtrim(WEBOrderNum)),'"','""') + '"'
	FROM tbl_WEB_CardSales
	WHERE (WEBStatus <> 0)

	if @MarkFlag = 1
	BEGIN
		UPDATE tbl_WEB_CardSales
		SET RMStatus = 0 
		WHERE (WEBStatus <> 0)
	END
END