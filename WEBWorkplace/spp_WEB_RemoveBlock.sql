USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_RemoveBlock]    Дата сценария: 05/29/2015 12:42:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spp_WEB_RemoveBlock] 
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Снятие блокировки в таблице tbl_WEB_2Block                                        |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
WITH RECOMPILE
AS

UPDATE tbl_WEB_2Block
SET [User] = N'', 
	ActionDt = CONVERT(DATETIME, '1900-10-01 00:00:00', 102)