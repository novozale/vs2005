USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Salesmans_FromScala]    Дата сценария: 05/29/2015 12:43:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Salesmans_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по продавцам                                                  |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

--------------------------Загрузка отсутствующих-----------------------------------------
---загружаются только незаблокированные продавцы-----------------------------------------
INSERT INTO tbl_WEB_Salesmans
SELECT View_3.ST01001, 
	View_3.ST01002, 
	'' AS Email, 
	NULL AS City, 
	0 AS OfficeLeader, 
	0 AS OnDuty, 
	0 AS IsActive, 
	'' AS Rezerv1, 
	'' AS Rezerv2, 
	1 AS RMStatus, 
	1 AS WEBStatus,
	1 AS ScalaStatus
FROM tbl_WEB_Salesmans AS tbl_WEB_Salesmans_1 RIGHT OUTER JOIN
	(SELECT ST010300.ST01001, 
		ST010300.ST01002
    FROM ST010300 INNER JOIN
		ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName
    WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0)) AS View_3 ON 
	tbl_WEB_Salesmans_1.Code = View_3.ST01001
WHERE (tbl_WEB_Salesmans_1.Code IS NULL)


--------------------------Обновление существующих----------------------------------------
UPDATE tbl_WEB_Salesmans
SET Name = ST010300.ST01002, 
	RMStatus = CASE WHEN IsActive = 1 THEN 3 ELSE RMStatus END, 
	WEBStatus = CASE WHEN IsActive = 1 THEN 3 ELSE WEBStatus END
FROM tbl_WEB_Salesmans INNER JOIN
    ST010300 ON tbl_WEB_Salesmans.Code = ST010300.ST01001 
	AND tbl_WEB_Salesmans.Name <> ST010300.ST01002


--------------------------Удаление отсутствующих-----------------------------------------
UPDATE tbl_WEB_Salesmans
SET IsActive = 0, 
	RMStatus = 2, 
	WEBStatus = 2, 
	ScalaStatus = 2
FROM tbl_WEB_Salesmans LEFT OUTER JOIN
	(SELECT ST010300.ST01001, 
		ST010300.ST01002
    FROM ST010300 INNER JOIN
		ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName
    WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0)) AS View_3 ON 
	tbl_WEB_Salesmans.Code = View_3.ST01001
WHERE (View_3.ST01001 IS NULL)