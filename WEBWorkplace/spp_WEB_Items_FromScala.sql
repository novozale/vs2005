USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_Items_FromScala]    Дата сценария: 05/29/2015 12:41:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE Procedure [dbo].[spp_WEB_Items_FromScala]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    выгрузка информации по товарам                                                    |
|    из Scala в промежуточную БД обмена с WEB                                          |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015г.                                                 |
--------------------------------------------------------------------------------------*/

WITH RECOMPILE
AS

--------------------------Загрузка отсутствующих-----------------------------------------
---загружаются только товары, присутствующих в прайс - листах-------------
INSERT INTO tbl_WEB_Items
SELECT View_10.*
FROM tbl_WEB_Items AS tbl_WEB_Items_1 RIGHT OUTER JOIN
	(SELECT SC010300.SC01001, 
		LTRIM(RTRIM(ISNULL(tbl_ItemCard0300.ManufacturerItemCode, N''))) AS ManufacturerItemCode, 
        tbl_ItemCard0300.Manufacturer AS ManufacturerCode, 
		ISNULL(View_6.SC33025, N'643') AS CountryCode, 
		ISNULL(View_6.SY24003, N'Россия') AS Country, 
		tbl_ItemCard0300.RexelProductCategory AS GroupCode, 
		'' AS SubGroupCode, 
		LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003)))) AS Name, 
		'' AS WEBName, 
		'' AS Description, 
		CASE WHEN CHARINDEX('1', SC010300.SC01128) > 0 THEN 1 ELSE 0 END AS WHAssortiment, 
		View_12.txt AS UOM, 
		'' AS Rezerv, 
		1 AS RMStatus, 
		1 AS WEBStatus
	FROM (SELECT SC01001, 
			PL01001
        FROM tbl_PurchasePriceHistory
        WHERE (DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) 
			AND (SC01001 <> N'')
        GROUP BY SC01001, PL01001) AS View_2 INNER JOIN
        SC010300 INNER JOIN
        tbl_Manufacturers INNER JOIN
        tbl_ItemCard0300 ON tbl_Manufacturers.ID = tbl_ItemCard0300.Manufacturer ON 
		SC010300.SC01001 = tbl_ItemCard0300.SC01001 ON 
        View_2.SC01001 = SC010300.SC01001 
		AND View_2.PL01001 = SC010300.SC01058 INNER JOIN
		(SELECT     0 AS num, SC09002 AS txt
        FROM          SC090300 WITH (NOLOCK)
        WHERE      (SC09001 = 'RUS')
        UNION
        SELECT     1 AS Expr1, SC09003
        FROM         SC090300 AS SC090300_40 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     2 AS Expr1, SC09004
        FROM         SC090300 AS SC090300_39 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     3 AS Expr1, SC09005
        FROM         SC090300 AS SC090300_38 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     4 AS Expr1, SC09006
        FROM         SC090300 AS SC090300_37 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     5 AS Expr1, SC09007
        FROM         SC090300 AS SC090300_36 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     6 AS Expr1, SC09008
        FROM         SC090300 AS SC090300_35 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     7 AS Expr1, SC09009
        FROM         SC090300 AS SC090300_34 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     8 AS Expr1, SC09010
        FROM         SC090300 AS SC090300_33 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     9 AS Expr1, SC09011
        FROM         SC090300 AS SC090300_32 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     10 AS Expr1, SC09012
        FROM         SC090300 AS SC090300_31 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     11 AS Expr1, SC09013
        FROM         SC090300 AS SC090300_30 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     12 AS Expr1, SC09014
        FROM         SC090300 AS SC090300_29 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     13 AS Expr1, SC09015
        FROM         SC090300 AS SC090300_28 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     14 AS Expr1, SC09016
        FROM         SC090300 AS SC090300_27 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     15 AS Expr1, SC09017
        FROM         SC090300 AS SC090300_26 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     16 AS Expr1, SC09018
        FROM         SC090300 AS SC090300_25 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     17 AS Expr1, SC09019
        FROM         SC090300 AS SC090300_24 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     18 AS Expr1, SC09020
        FROM         SC090300 AS SC090300_23 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     19 AS Expr1, SC09021
        FROM         SC090300 AS SC090300_22 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     20 AS Expr1, SC09022
        FROM         SC090300 AS SC090300_21 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     21 AS Expr1, SC09023
        FROM         SC090300 AS SC090300_20 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     22 AS Expr1, SC09024
        FROM         SC090300 AS SC090300_19 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     23 AS Expr1, SC09025
        FROM         SC090300 AS SC090300_18 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     24 AS Expr1, SC09026
        FROM         SC090300 AS SC090300_17 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     25 AS Expr1, SC09027
        FROM         SC090300 AS SC090300_16 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     26 AS Expr1, SC09028
        FROM         SC090300 AS SC090300_15 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     27 AS Expr1, SC09029
        FROM         SC090300 AS SC090300_14 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     28 AS Expr1, SC09030
        FROM         SC090300 AS SC090300_13 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     29 AS Expr1, SC09031
        FROM         SC090300 AS SC090300_12 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     30 AS Expr1, SC09032
        FROM         SC090300 AS SC090300_11 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     31 AS Expr1, SC09033
        FROM         SC090300 AS SC090300_10 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     32 AS Expr1, SC09034
        FROM         SC090300 AS SC090300_9 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     33 AS Expr1, SC09035
        FROM         SC090300 AS SC090300_8 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     34 AS Expr1, SC09036
        FROM         SC090300 AS SC090300_7 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     35 AS Expr1, SC09037
        FROM         SC090300 AS SC090300_6 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     36 AS Expr1, SC09038
        FROM         SC090300 AS SC090300_5 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     37 AS Expr1, SC09039
        FROM         SC090300 AS SC090300_4 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     38 AS Expr1, SC09040
        FROM         SC090300 AS SC090300_3 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     39 AS Expr1, SC09041
        FROM         SC090300 AS SC090300_2 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     40 AS Expr1, SC09042
        FROM         SC090300 AS SC090300_1 WITH (NOLOCK)
        WHERE (SC09001 = 'RUS')) AS View_12 ON 
		SC010300.SC01135 = View_12.num LEFT OUTER JOIN
        (SELECT View_4.SC33001, 
			View_4.SC33025, 
			View_5.SY24003
        FROM (SELECT SC330300.SC33001, 
				CASE WHEN SC330300.SC33025 = '' THEN '643' ELSE SC330300.SC33025 END AS SC33025
            FROM SC330300 INNER JOIN
				(SELECT SC33001, 
					MAX(SC33003) AS SC33003
                FROM SC330300 AS SC330300_1
                GROUP BY SC33001) AS View_3 ON 
			SC330300.SC33001 = View_3.SC33001 
			AND SC330300.SC33003 = View_3.SC33003) AS View_4 INNER JOIN
		(SELECT SY24002, 
			SY24003
        FROM SY240300
        WHERE (SY24001 = N'BM')) AS View_5 ON 
		View_4.SC33025 = View_5.SY24002) AS View_6 ON 
    SC010300.SC01001 = View_6.SC33001) AS View_10 ON 
	tbl_WEB_Items_1.Code = View_10.SC01001
WHERE (tbl_WEB_Items_1.Code IS NULL)


--------------------------Обновление существующих----------------------------------------
UPDATE tbl_WEB_Items
SET ManufacturerItemCode = Ltrim(Rtrim(View_10.ManufacturerItemCode)), 
	ManufacturerCode = View_10.ManufacturerCode, 
	CountryCode = View_10.CountryCode, 
    Country = View_10.Country, 
	GroupCode = View_10.GroupCode, 
	Name = View_10.Name, 
	WHAssortiment = View_10.WHAssortiment, 
	UOM = View_10.UOM,
	RMStatus = CASE WHEN tbl_WEB_Items.RMStatus = 1 THEN 1 ELSE 3 END,
	WEBStatus = CASE WHEN tbl_WEB_Items.WEBStatus = 1 THEN 1 ELSE 3 END
FROM tbl_WEB_Items INNER JOIN
    (SELECT SC010300.SC01001, 
		LTRIM(RTRIM(ISNULL(tbl_ItemCard0300.ManufacturerItemCode, N''))) AS ManufacturerItemCode, 
        tbl_ItemCard0300.Manufacturer AS ManufacturerCode, 
		ISNULL(View_6.SC33025, N'643') AS CountryCode, 
		ISNULL(View_6.SY24003, N'Россия') AS Country, 
		tbl_ItemCard0300.RexelProductCategory AS GroupCode, 
		'' AS SubGroupCode, 
		LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003)))) AS Name, 
		'' AS WEBName, 
		'' AS Description, 
		CASE WHEN CHARINDEX('1', SC010300.SC01128) > 0 THEN 1 ELSE 0 END AS WHAssortiment, 
		View_12.txt AS UOM, 
		'' AS Rezerv, 
		1 AS RMStatus, 
		1 AS WEBStatus
	FROM (SELECT SC01001, 
			PL01001
        FROM tbl_PurchasePriceHistory
        WHERE (DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) 
			AND (SC01001 <> N'')
        GROUP BY SC01001, PL01001) AS View_2 INNER JOIN
        SC010300 INNER JOIN
        tbl_Manufacturers INNER JOIN
        tbl_ItemCard0300 ON tbl_Manufacturers.ID = tbl_ItemCard0300.Manufacturer ON 
		SC010300.SC01001 = tbl_ItemCard0300.SC01001 ON 
        View_2.SC01001 = SC010300.SC01001 
		AND View_2.PL01001 = SC010300.SC01058 INNER JOIN
		(SELECT     0 AS num, SC09002 AS txt
        FROM          SC090300 WITH (NOLOCK)
        WHERE      (SC09001 = 'RUS')
        UNION
        SELECT     1 AS Expr1, SC09003
        FROM         SC090300 AS SC090300_40 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     2 AS Expr1, SC09004
        FROM         SC090300 AS SC090300_39 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     3 AS Expr1, SC09005
        FROM         SC090300 AS SC090300_38 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     4 AS Expr1, SC09006
        FROM         SC090300 AS SC090300_37 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     5 AS Expr1, SC09007
        FROM         SC090300 AS SC090300_36 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     6 AS Expr1, SC09008
        FROM         SC090300 AS SC090300_35 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     7 AS Expr1, SC09009
        FROM         SC090300 AS SC090300_34 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     8 AS Expr1, SC09010
        FROM         SC090300 AS SC090300_33 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     9 AS Expr1, SC09011
        FROM         SC090300 AS SC090300_32 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     10 AS Expr1, SC09012
        FROM         SC090300 AS SC090300_31 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     11 AS Expr1, SC09013
        FROM         SC090300 AS SC090300_30 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     12 AS Expr1, SC09014
        FROM         SC090300 AS SC090300_29 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     13 AS Expr1, SC09015
        FROM         SC090300 AS SC090300_28 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     14 AS Expr1, SC09016
        FROM         SC090300 AS SC090300_27 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     15 AS Expr1, SC09017
        FROM         SC090300 AS SC090300_26 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     16 AS Expr1, SC09018
        FROM         SC090300 AS SC090300_25 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     17 AS Expr1, SC09019
        FROM         SC090300 AS SC090300_24 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     18 AS Expr1, SC09020
        FROM         SC090300 AS SC090300_23 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     19 AS Expr1, SC09021
        FROM         SC090300 AS SC090300_22 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     20 AS Expr1, SC09022
        FROM         SC090300 AS SC090300_21 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     21 AS Expr1, SC09023
        FROM         SC090300 AS SC090300_20 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     22 AS Expr1, SC09024
        FROM         SC090300 AS SC090300_19 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     23 AS Expr1, SC09025
        FROM         SC090300 AS SC090300_18 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     24 AS Expr1, SC09026
        FROM         SC090300 AS SC090300_17 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     25 AS Expr1, SC09027
        FROM         SC090300 AS SC090300_16 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     26 AS Expr1, SC09028
        FROM         SC090300 AS SC090300_15 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     27 AS Expr1, SC09029
        FROM         SC090300 AS SC090300_14 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     28 AS Expr1, SC09030
        FROM         SC090300 AS SC090300_13 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     29 AS Expr1, SC09031
        FROM         SC090300 AS SC090300_12 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     30 AS Expr1, SC09032
        FROM         SC090300 AS SC090300_11 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     31 AS Expr1, SC09033
        FROM         SC090300 AS SC090300_10 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     32 AS Expr1, SC09034
        FROM         SC090300 AS SC090300_9 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     33 AS Expr1, SC09035
        FROM         SC090300 AS SC090300_8 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     34 AS Expr1, SC09036
        FROM         SC090300 AS SC090300_7 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     35 AS Expr1, SC09037
        FROM         SC090300 AS SC090300_6 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     36 AS Expr1, SC09038
        FROM         SC090300 AS SC090300_5 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     37 AS Expr1, SC09039
        FROM         SC090300 AS SC090300_4 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     38 AS Expr1, SC09040
        FROM         SC090300 AS SC090300_3 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     39 AS Expr1, SC09041
        FROM         SC090300 AS SC090300_2 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     40 AS Expr1, SC09042
        FROM         SC090300 AS SC090300_1 WITH (NOLOCK)
        WHERE (SC09001 = 'RUS')) AS View_12 ON 
		SC010300.SC01135 = View_12.num LEFT OUTER JOIN
        (SELECT View_4.SC33001, 
			View_4.SC33025, 
			View_5.SY24003
        FROM (SELECT SC330300.SC33001, 
				CASE WHEN SC330300.SC33025 = '' THEN '643' ELSE SC330300.SC33025 END AS SC33025
            FROM SC330300 INNER JOIN
				(SELECT SC33001, 
					MAX(SC33003) AS SC33003
                FROM SC330300 AS SC330300_1
                GROUP BY SC33001) AS View_3 ON 
			SC330300.SC33001 = View_3.SC33001 
			AND SC330300.SC33003 = View_3.SC33003) AS View_4 INNER JOIN
		(SELECT SY24002, 
			SY24003
        FROM SY240300
        WHERE (SY24001 = N'BM')) AS View_5 ON 
		View_4.SC33025 = View_5.SY24002) AS View_6 ON 
    SC010300.SC01001 = View_6.SC33001) AS View_10 ON 
		tbl_WEB_Items.Code = View_10.SC01001
WHERE (tbl_WEB_Items.ManufacturerItemCode <> Ltrim(Rtrim(View_10.ManufacturerItemCode))) OR
	(tbl_WEB_Items.ManufacturerCode <> View_10.ManufacturerCode) OR
    (tbl_WEB_Items.CountryCode <> View_10.CountryCode) OR
    (tbl_WEB_Items.Country <> View_10.Country) OR
    (tbl_WEB_Items.GroupCode <> View_10.GroupCode) OR
    (tbl_WEB_Items.Name <> View_10.Name) OR
    (tbl_WEB_Items.WHAssortiment <> View_10.WHAssortiment) OR
	(tbl_WEB_Items.UOM <> View_10.UOM)


--------------------------Удаление отсутствующих-----------------------------------------
UPDATE tbl_WEB_Items
SET RMStatus = 2, 
	WEBStatus = 2
FROM tbl_WEB_Items LEFT OUTER JOIN
	(SELECT SC010300.SC01001, 
		LTRIM(RTRIM(ISNULL(tbl_ItemCard0300.ManufacturerItemCode, N''))) AS ManufacturerItemCode, 
        tbl_ItemCard0300.Manufacturer AS ManufacturerCode, 
		ISNULL(View_6.SC33025, N'643') AS CountryCode, 
		ISNULL(View_6.SY24003, N'Россия') AS Country, 
		tbl_ItemCard0300.RexelProductCategory AS GroupCode, 
		'' AS SubGroupCode, 
		LTRIM(RTRIM(LTRIM(RTRIM(SC010300.SC01002)) + ' ' + LTRIM(RTRIM(SC010300.SC01003)))) AS Name, 
		'' AS WEBName, 
		'' AS Description, 
		CASE WHEN CHARINDEX('1', SC010300.SC01128) > 0 THEN 1 ELSE 0 END AS WHAssortiment, 
		View_12.txt AS UOM, 
		'' AS Rezerv, 
		1 AS RMStatus, 
		1 AS WEBStatus
	FROM (SELECT SC01001, 
			PL01001
        FROM tbl_PurchasePriceHistory
        WHERE (DateTo = CONVERT(DATETIME, '9999-12-31 00:00:00', 102)) 
			AND (SC01001 <> N'')
        GROUP BY SC01001, PL01001) AS View_2 INNER JOIN
        SC010300 INNER JOIN
        tbl_Manufacturers INNER JOIN
        tbl_ItemCard0300 ON tbl_Manufacturers.ID = tbl_ItemCard0300.Manufacturer ON 
		SC010300.SC01001 = tbl_ItemCard0300.SC01001 ON 
        View_2.SC01001 = SC010300.SC01001 
		AND View_2.PL01001 = SC010300.SC01058 INNER JOIN
		(SELECT     0 AS num, SC09002 AS txt
        FROM          SC090300 WITH (NOLOCK)
        WHERE      (SC09001 = 'RUS')
        UNION
        SELECT     1 AS Expr1, SC09003
        FROM         SC090300 AS SC090300_40 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     2 AS Expr1, SC09004
        FROM         SC090300 AS SC090300_39 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     3 AS Expr1, SC09005
        FROM         SC090300 AS SC090300_38 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     4 AS Expr1, SC09006
        FROM         SC090300 AS SC090300_37 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     5 AS Expr1, SC09007
        FROM         SC090300 AS SC090300_36 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     6 AS Expr1, SC09008
        FROM         SC090300 AS SC090300_35 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     7 AS Expr1, SC09009
        FROM         SC090300 AS SC090300_34 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     8 AS Expr1, SC09010
        FROM         SC090300 AS SC090300_33 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     9 AS Expr1, SC09011
        FROM         SC090300 AS SC090300_32 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     10 AS Expr1, SC09012
        FROM         SC090300 AS SC090300_31 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     11 AS Expr1, SC09013
        FROM         SC090300 AS SC090300_30 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     12 AS Expr1, SC09014
        FROM         SC090300 AS SC090300_29 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     13 AS Expr1, SC09015
        FROM         SC090300 AS SC090300_28 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     14 AS Expr1, SC09016
        FROM         SC090300 AS SC090300_27 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     15 AS Expr1, SC09017
        FROM         SC090300 AS SC090300_26 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     16 AS Expr1, SC09018
        FROM         SC090300 AS SC090300_25 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     17 AS Expr1, SC09019
        FROM         SC090300 AS SC090300_24 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     18 AS Expr1, SC09020
        FROM         SC090300 AS SC090300_23 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     19 AS Expr1, SC09021
        FROM         SC090300 AS SC090300_22 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     20 AS Expr1, SC09022
        FROM         SC090300 AS SC090300_21 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     21 AS Expr1, SC09023
        FROM         SC090300 AS SC090300_20 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     22 AS Expr1, SC09024
        FROM         SC090300 AS SC090300_19 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     23 AS Expr1, SC09025
        FROM         SC090300 AS SC090300_18 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     24 AS Expr1, SC09026
        FROM         SC090300 AS SC090300_17 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     25 AS Expr1, SC09027
        FROM         SC090300 AS SC090300_16 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     26 AS Expr1, SC09028
        FROM         SC090300 AS SC090300_15 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     27 AS Expr1, SC09029
        FROM         SC090300 AS SC090300_14 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     28 AS Expr1, SC09030
        FROM         SC090300 AS SC090300_13 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     29 AS Expr1, SC09031
        FROM         SC090300 AS SC090300_12 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     30 AS Expr1, SC09032
        FROM         SC090300 AS SC090300_11 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     31 AS Expr1, SC09033
        FROM         SC090300 AS SC090300_10 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     32 AS Expr1, SC09034
        FROM         SC090300 AS SC090300_9 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     33 AS Expr1, SC09035
        FROM         SC090300 AS SC090300_8 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     34 AS Expr1, SC09036
        FROM         SC090300 AS SC090300_7 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     35 AS Expr1, SC09037
        FROM         SC090300 AS SC090300_6 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     36 AS Expr1, SC09038
        FROM         SC090300 AS SC090300_5 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     37 AS Expr1, SC09039
        FROM         SC090300 AS SC090300_4 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     38 AS Expr1, SC09040
        FROM         SC090300 AS SC090300_3 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     39 AS Expr1, SC09041
        FROM         SC090300 AS SC090300_2 WITH (NOLOCK)
        WHERE     (SC09001 = 'RUS')
        UNION
        SELECT     40 AS Expr1, SC09042
        FROM         SC090300 AS SC090300_1 WITH (NOLOCK)
        WHERE (SC09001 = 'RUS')) AS View_12 ON 
		SC010300.SC01135 = View_12.num LEFT OUTER JOIN
        (SELECT View_4.SC33001, 
			View_4.SC33025, 
			View_5.SY24003
        FROM (SELECT SC330300.SC33001, 
				CASE WHEN SC330300.SC33025 = '' THEN '643' ELSE SC330300.SC33025 END AS SC33025
            FROM SC330300 INNER JOIN
				(SELECT SC33001, 
					MAX(SC33003) AS SC33003
                FROM SC330300 AS SC330300_1
                GROUP BY SC33001) AS View_3 ON 
			SC330300.SC33001 = View_3.SC33001 
			AND SC330300.SC33003 = View_3.SC33003) AS View_4 INNER JOIN
		(SELECT SY24002, 
			SY24003
        FROM SY240300
        WHERE (SY24001 = N'BM')) AS View_5 ON 
		View_4.SC33025 = View_5.SY24002) AS View_6 ON 
    SC010300.SC01001 = View_6.SC33001) AS View_10 ON 
		tbl_WEB_Items.Code = View_10.SC01001
WHERE (View_10.SC01001 IS NULL)
