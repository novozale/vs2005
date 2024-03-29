USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Manufacturers_FromScala]    Дата сценария: 05/29/2015 12:41:56 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Manufacturers_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по производителям                                             |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

--------------------------Загрузка отсутствующих-----------------------------------------
---загружаются только производители товаров, присутствующих в прайс - листах-------------
INSERT INTO tbl_WEB_Manufacturers
SELECT View_2.ID, 
	View_2.Name, 
	'' AS WEBName, 
	'' AS Rezerv1, 
	1 AS RMStatus, 
	1 AS WEBStatus
FROM tbl_WEB_Manufacturers AS tbl_WEB_Manufacturers_1 RIGHT OUTER JOIN
	(SELECT tbl_Manufacturers.ID, 
		tbl_Manufacturers.Name
    FROM tbl_PurchasePriceHistory INNER JOIN
		tbl_ItemCard0300 ON tbl_PurchasePriceHistory.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN
        tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID
    WHERE (tbl_PurchasePriceHistory.DateTo = CONVERT(DATETIME, '31/12/9999', 103))
    GROUP BY tbl_Manufacturers.ID, 
		tbl_Manufacturers.Name) AS View_2 ON 
	tbl_WEB_Manufacturers_1.ID = View_2.ID
WHERE (tbl_WEB_Manufacturers_1.ID IS NULL)


--------------------------Обновление существующих----------------------------------------
UPDATE tbl_WEB_Manufacturers
SET Name = tbl_Manufacturers.Name, 
	RMStatus = CASE WHEN RMStatus = 1 THEN 1 ELSE 3 END,
	WEBStatus = CASE WHEN WEBStatus = 1 THEN 1 ELSE 3 END
FROM tbl_WEB_Manufacturers INNER JOIN
    tbl_Manufacturers ON tbl_WEB_Manufacturers.ID = tbl_Manufacturers.ID 
	AND tbl_WEB_Manufacturers.Name <> tbl_Manufacturers.Name


--------------------------Удаление отсутствующих-----------------------------------------
UPDATE tbl_WEB_Manufacturers
SET RMStatus = 2, 
	WEBStatus = 2
FROM tbl_WEB_Manufacturers LEFT OUTER JOIN
	(SELECT tbl_Manufacturers.ID, 
		tbl_Manufacturers.Name
    FROM tbl_PurchasePriceHistory INNER JOIN
		tbl_ItemCard0300 ON tbl_PurchasePriceHistory.SC01001 = tbl_ItemCard0300.SC01001 INNER JOIN
        tbl_Manufacturers ON tbl_ItemCard0300.Manufacturer = tbl_Manufacturers.ID
    WHERE (tbl_PurchasePriceHistory.DateTo = CONVERT(DATETIME, '31/12/9999', 103))
    GROUP BY tbl_Manufacturers.ID, 
		tbl_Manufacturers.Name) AS View_2 ON 
	tbl_WEB_Manufacturers.ID = View_2.ID
WHERE (View_2.ID IS NULL)