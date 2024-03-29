USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Clients_FromScala]    Дата сценария: 05/29/2015 12:40:14 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Clients_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по клиентам                                                   |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

--------------------------Загрузка отсутствующих-----------------------------------------
INSERT INTO tbl_WEB_Clients
SELECT View_2.Code, 
	View_2.Name, 
	View_2.Address, 
	View_2.Discount, 
	0 AS WorkOverWEB,
	'00' AS BasePrice
FROM tbl_WEB_Clients AS tbl_WEB_Clients_1 RIGHT OUTER JOIN
	(SELECT SL010300.SL01001 AS Code, 
		SL010300.SL01002 AS Name, 
		LTRIM(RTRIM(LTRIM(RTRIM(SL010300.SL01003)) + ' ' + LTRIM(RTRIM(SL010300.SL01004)) + ' ' + LTRIM(RTRIM(SL010300.SL01005)))) AS Address, 
		ISNULL(View_1.SL14007, N'') AS Discount
    FROM SL010300 LEFT OUTER JOIN
	(SELECT SL14001, 
		SL14007
    FROM SL140300
    WHERE (SL14002 = N'00')) AS View_1 ON 
	SL010300.SL01001 = View_1.SL14001) AS View_2 ON 
    tbl_WEB_Clients_1.Code = View_2.Code
WHERE (tbl_WEB_Clients_1.Code IS NULL)


--------------------------Обновление существующих----------------------------------------
UPDATE tbl_WEB_Clients
SET Name = View_2.Name, 
	Address = View_2.Address, 
	Discount = View_2.Discount
FROM tbl_WEB_Clients INNER JOIN
    (SELECT SL010300.SL01001 AS Code, 
		SL010300.SL01002 AS Name, 
		LTRIM(RTRIM(LTRIM(RTRIM(SL010300.SL01003)) + ' ' + LTRIM(RTRIM(SL010300.SL01004)) + ' ' + LTRIM(RTRIM(SL010300.SL01005)))) AS Address, 
		ISNULL(View_1.SL14007, N'') AS Discount
    FROM SL010300 LEFT OUTER JOIN
	(SELECT SL14001, 
		SL14007
    FROM SL140300
    WHERE (SL14002 = N'00')) AS View_1 ON 
	SL010300.SL01001 = View_1.SL14001) AS View_2 ON 
    tbl_WEB_Clients.Code = View_2.Code


--------------------------Удаление отсутствующих-----------------------------------------
DELETE FROM tbl_WEB_Clients
FROM tbl_WEB_Clients LEFT OUTER JOIN
	(SELECT SL010300.SL01001 AS Code, 
		SL010300.SL01002 AS Name, LTRIM(RTRIM(LTRIM(RTRIM(SL010300.SL01003)) + ' ' + LTRIM(RTRIM(SL010300.SL01004)) + ' ' + LTRIM(RTRIM(SL010300.SL01005)))) AS Address, 
		ISNULL(View_1.SL14007, N'') AS Discount
    FROM SL010300 LEFT OUTER JOIN
    (SELECT SL14001, 
		SL14007
    FROM SL140300
    WHERE (SL14002 = N'00')) AS View_1 ON 
	SL010300.SL01001 = View_1.SL14001) AS View_2 ON 
    tbl_WEB_Clients.Code = View_2.Code
WHERE (View_2.Code IS NULL)