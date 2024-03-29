USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_ImportWHMark]    Дата сценария: 07/11/2012 11:42:55 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE Procedure [dbo].[spp_ImportWHMark]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    импорт признака складского ассортимента                                           |
|    разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyScalaCode [nvarchar](35),                             -- Код запаса в Scala
@MyWHMark nvarchar(10),                                  -- признак складской
@My09Update int,                                         -- Обновляем ли ручные ROP,МЖЗ и страх запас на 09 складе 
@My09ROP float,                                          -- ручной ROP на 09 складе
@My09MGZ float,                                          -- ручной МЖЗ на 09 складе
@My09InsLevel float,                                     -- ручной страховой запас на 09 складе
@My07Update int,                                         -- Обновляем ли ручные ROP,МЖЗ и страх запас на 07 складе 
@My07ROP float,                                          -- ручной ROP на 07 складе
@My07MGZ float,                                          -- ручной МЖЗ на 07 складе
@My07InsLevel float,                                     -- ручной страховой запас на 07 складе
@My03Update int,                                         -- Обновляем ли ручные ROP,МЖЗ и страх запас на 03 складе 
@My03ROP float,                                          -- ручной ROP на 03 складе
@My03MGZ float,                                          -- ручной МЖЗ на 03 складе
@My03InsLevel float,                                     -- ручной страховой запас на 03 складе
@My01Update int,                                         -- Обновляем ли ручные ROP,МЖЗ и страх запас на 01 складе 
@My01ROP float,                                          -- ручной ROP на 01 складе
@My01MGZ float,                                          -- ручной МЖЗ на 01 складе
@My01InsLevel float,                                     -- ручной страховой запас на 01 складе
@MyUser nvarchar(255)                                    -- пользователь Scala

WITH RECOMPILE
AS

declare @MyCC as float;                                  -- для кол - ва
SELECT @MyScalaCode = LTrim(Rtrim(@MyScalaCode))
SELECT @MyWHMark = Ltrim(Rtrim(@MyWHMark))

-------Обновление признака складской
IF @MyWHMark = '0000'
BEGIN
	UPDATE SC010300
	SET SC01128 = ''
	WHERE (SC01001 = @MyScalaCode)
END
ELSE
BEGIN
	UPDATE SC010300
	SET SC01128 = @MyWHMark
	WHERE (SC01001 = @MyScalaCode)
END

-------при необходимости - удаление из консигнационных запасов
DELETE FROM tbl_ItemInfo0300
FROM (SELECT SC010300.SC01001, 
		View_2.WH, 
		1 AS IsWH
    FROM SC010300 CROSS JOIN
		(SELECT SC23001 AS WH, 
			CHARINDEX('1', SC23007) AS WHPos
		FROM SC230300
		WHERE (SC23006 = N'1')) AS View_2
	WHERE (SC010300.SC01001 = @MyScalaCode) AND 
		(SUBSTRING(@MyWHMark, View_2.WHPos, 1) = N'1')) AS View_1 INNER JOIN
    tbl_ItemInfo0300 ON View_1.SC01001 = tbl_ItemInfo0300.SC03001 AND 
	View_1.WH = tbl_ItemInfo0300.SC03002

SELECT @MyCC = COUNT(*) 
FROM tbl_ItemInfo0300
WHERE (SC03001 = @MyScalaCode)

IF @MyCC = 0
BEGIN
	UPDATE tbl_ItemCard0300
	SET IsConsignment = N'0'
	WHERE (SC01001 = @MyScalaCode)
END


-------Обновление ручных значений ROP, МЖЗ и страхового запаса
--09 склад
IF @My09Update = -1
BEGIN
	IF @My09ROP = 0 AND @My09MGZ = 0 AND @My09InsLevel = 0
	BEGIN
		---Новые значения не выставляем, закрываем старые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'09')
            AND (Code = @MyScalaCode)
		
		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'09')
	END
	ELSE
	BEGIN
		---Закрываем старые значения и выставляем новые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'09')
            AND (Code = @MyScalaCode)

		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'09')

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL)
        VALUES (NEWID(), 
        @MyScalaCode, 
        '09',
        @My09MGZ, 
        @My09ROP,
        @My09InsLevel) 

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) 
        VALUES (NEWID(), 
        @MyScalaCode, 
        '09',
        @My09MGZ, 
        @My09ROP,
        @My09InsLevel, 
        @MyUser,
        GETDATE(), 
		Convert(datetime,'31/12/9999',103))
	END
END

--07 склад
IF @My07Update = -1
BEGIN
	IF @My07ROP = 0 AND @My07MGZ = 0 AND @My07InsLevel = 0
	BEGIN
		---Новые значения не выставляем, закрываем старые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'07')
            AND (Code = @MyScalaCode)
		
		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'07')
	END
	ELSE
	BEGIN
		---Закрываем старые значения и выставляем новые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'07')
            AND (Code = @MyScalaCode)

		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'07')

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL)
        VALUES (NEWID(), 
        @MyScalaCode, 
        '07',
        @My07MGZ, 
        @My07ROP,
        @My07InsLevel) 

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) 
        VALUES (NEWID(), 
        @MyScalaCode, 
        '07',
        @My07MGZ, 
        @My07ROP,
        @My07InsLevel, 
        @MyUser,
        GETDATE(), 
		Convert(datetime,'31/12/9999',103))
	END
END

--03 склад
IF @My03Update = -1
BEGIN
	IF @My03ROP = 0 AND @My03MGZ = 0 AND @My03InsLevel = 0
	BEGIN
		---Новые значения не выставляем, закрываем старые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'03')
            AND (Code = @MyScalaCode)
		
		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'03')
	END
	ELSE
	BEGIN
		---Закрываем старые значения и выставляем новые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'03')
            AND (Code = @MyScalaCode)

		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'03')

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL)
        VALUES (NEWID(), 
        @MyScalaCode, 
        '03',
        @My03MGZ, 
        @My03ROP,
        @My03InsLevel) 

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) 
        VALUES (NEWID(), 
        @MyScalaCode, 
        '03',
        @My03MGZ, 
        @My03ROP,
        @My03InsLevel, 
        @MyUser,
        GETDATE(), 
		Convert(datetime,'31/12/9999',103))
	END
END

--01 склад
IF @My01Update = -1
BEGIN
	IF @My01ROP = 0 AND @My01MGZ = 0 AND @My01InsLevel = 0
	BEGIN
		---Новые значения не выставляем, закрываем старые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'01')
            AND (Code = @MyScalaCode)
		
		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'01')
	END
	ELSE
	BEGIN
		---Закрываем старые значения и выставляем новые
		Update tbl_ForecastOrderR2_CustomMGZROPINS_History 
        SET DateTo = GETDATE()
        WHERE (DateTo = Convert(datetime,'31/12/9999',103)) 
            AND (WH = N'01')
            AND (Code = @MyScalaCode)

		DELETE FROM tbl_ForecastOrderR2_CustomMGZROPINS
		WHERE (Code = @MyScalaCode) AND 
			(WH = N'01')

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL)
        VALUES (NEWID(), 
        @MyScalaCode, 
        '01',
        @My01MGZ, 
        @My01ROP,
        @My01InsLevel) 

		INSERT INTO tbl_ForecastOrderR2_CustomMGZROPINS_History 
        (ID, Code, WH, MGZ, ROP, IshuranceLVL, UserID, DateFrom, DateTo) 
        VALUES (NEWID(), 
        @MyScalaCode, 
        '01',
        @My01MGZ, 
        @My01ROP,
        @My01InsLevel, 
        @MyUser,
        GETDATE(), 
		Convert(datetime,'31/12/9999',103))
	END
END
 