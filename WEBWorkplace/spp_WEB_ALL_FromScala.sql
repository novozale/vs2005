USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_ALL_FromScala]    Дата сценария: 05/29/2015 12:39:20 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_ALL_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации                                                               |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

exec spp_WEB_Manufacturers_FromScala

exec spp_WEB_Salesmans_FromScala

exec spp_WEB_ItemGroups_FromScala

exec spp_WEB_Items_FromScala

exec spp_WEB_Clients_FromScala

exec spp_WEB_RemoveGarbage