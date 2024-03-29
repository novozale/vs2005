USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementOrderCreationFromExcel]    Дата сценария: 09/24/2013 10:01:14 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[spp_DisplacementOrderCreationFromExcel]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Процедура формирования заказа на перемещение из Excel                             |
|    Разработчик Новожилов А.Н. 2013                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@SrcWarNo nvarchar(6),                                              -- исходный склад
@DestWarNo nvarchar(6),                                             -- склад назначения
@MyOtherWHFlag integer,                                             -- Включать в заказ на перемещение запасы для заказов на продажу на других складах
@MyOrderDate datetime,                                              -- дата предполагаемой отгрузки
@MyShipDate datetime,                                               -- дата доставки (предполагаемая)
@MyRezStr nvarchar(4000) output,                                    -- результирующая строка
@MyRelocOrderNum nvarchar(30) output                                -- номер созданного заказа на перемещение

AS

DECLARE @MySrcWarNo nvarchar(6);                                    -- исходный склад
DECLARE @MyDestWarNo nvarchar(6);                                   -- склад назначения
DECLARE @MyZeroDate datetime                                        -- 1 января 1900 года
DECLARE @MyUserID int                                               -- код пользователя
DECLARE @MyUserName nvarchar(35)                                    -- имя пользователя
Declare @MyCount1 int;                                              -- счетчик
DECLARE @MyItemCode nvarchar(35)                                    -- код запаса
DECLARE @MyUOM int                                                  -- единица измерения (SC01135 - продажи)
DECLARE @MyQTYStr float                                             -- кол - во к перемещению в строке
DECLARE @MyQTYStrRest float                                         -- оставшееся кол - во к перемещению в строке
DECLARE @MyStrNum nvarchar(6)                                       -- номер строки в заказе
DECLARE @MySubStrNum nvarchar(6)                                    -- номер подстроки в заказе
DECLARE @MyBatchNum nvarchar(12)                                    -- номер партии запаса
DECLARE @MyBatchQTY float                                           -- доступное кол - во в партии

SELECT @MySrcWarNo = Ltrim(Rtrim(@SrcWarNo))
SELECT @MyDestWarNo = Ltrim(Rtrim(@DestWarNo))
SELECT @MyZeroDate = CONVERT(datetime,'01/01/1900',103)
SELECT @MyRezStr = ''
SELECT @MyRelocOrderNum = ''

---------------------Пользователь--------------------------------------------------------
SELECT @MyUserName = CASE WHEN dbo.GetUserByPCANDThread(host_name, host_process_id) = '' 
	THEN 'Admin'
	ELSE Right(dbo.GetUserByPCANDThread(host_name, host_process_id),LEN(dbo.GetUserByPCANDThread(host_name, host_process_id)) - LEN('EskRU\')) END 
FROM sys.dm_exec_sessions
WHERE (session_id = @@SPID)

SELECT @MyUserID = UserID
FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK)
WHERE (UPPER(UserName) = UPPER(@MyUserName))

----номер заказа и увеличение счетчика
SELECT @MyRelocOrderNum = RIGHT('0000000000' + CONVERT(nvarchar,SY68002),10)
FROM SY6803XX
WHERE (SY68001 = N'SC7C')

UPDATE SY6803XX WITH (ROWLOCK)
SET SY68002 = CONVERT(float,@MyRelocOrderNum) + 1
WHERE (SY68001 = N'SC7C')

----Заголовок заказа на перемещение
INSERT INTO SC7C0300
	(SC7C001, SC7C002, SC7C003, SC7C004, SC7C005, SC7C006, SC7C007, SC7C008, SC7C009, SC7C010)
VALUES (@MyRelocOrderNum,      --номер заказа
	@MyOrderDate,              --дата заказа
	'', 
	@MySrcWarNo,               --исходный склад
	'', 
	@MyDestWarNo,              --склад назначения
	@MyOrderDate,              --дата заказа
	@MyZeroDate,               --01.01.1900
	@MyShipDate,               --предполагаемая дата доставки
	N'1')                      --флаг - 1 тип

----Данные в доп таблицу tbl_DisplacementOrder
DELETE FROM tbl_DisplacementOrder
WHERE (OrderNumber = @MyRelocOrderNum)

INSERT INTO tbl_DisplacementOrder
	(ID, OrderNumber, UserCode, SalesOrderNumber, ReadyFlag)
VALUES (NEWID(), 
	@MyRelocOrderNum, 
	@MyUserID, 
	N'', 
	1)

----курсор по запасам - для формирования строк заказа
IF @MyOtherWHFlag = 0         -- 0 -- в заказ на перемещение запасы для заказов на продажу на других складах не включаем
BEGIN
	DECLARE String_Cursor CURSOR FOR
/*
	SELECT Ltrim(Rtrim(SC330300.SC33001)) AS ItemCode, 
		CASE WHEN SUM(SC330300.SC33005 - SC330300.SC33006) >= #_MyOrder.QTY THEN #_MyOrder.QTY ELSE SUM(SC330300.SC33005 - SC330300.SC33006) END AS QTY,
		ISNULL(SC010300.SC01135, 0) AS UOM
	FROM SC330300 WITH (NOLOCK) INNER JOIN
		PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
        PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
		PC190300.PC19002 = PC030300.PC03002 INNER JOIN
        PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN
        #_MyOrder ON SC330300.SC33001 = #_MyOrder.ItemCode LEFT OUTER JOIN
        SC010300 ON SC330300.SC33001 = SC010300.SC01001
	WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
		(SC330300.SC33002 = @MySrcWarNo) AND 
		(PC010300.PC01060 NOT IN (SELECT OR01001
			FROM OR010300
			WHERE (OR01050 <> @MySrcWarNo)))
	GROUP BY SC330300.SC33001, 
		ISNULL(SC010300.SC01135, 0), 
		#_MyOrder.QTY
	ORDER BY ItemCode
*/
	SELECT View_1.SC33001, 
		CASE WHEN SUM(View_1.QTY) >= #_MyOrder.QTY THEN #_MyOrder.QTY ELSE SUM(View_1.QTY) END AS QTY, 
		ISNULL(SC010300.SC01135, 0) AS UOM
	FROM (SELECT SC330300.SC33001, 
			SC330300.SC33003, 
			SC330300.SC33005 - SC330300.SC33006 AS QTY, 
			ISNULL(PC010300.PC01060, N'') AS PC01060
        FROM PC030300 RIGHT OUTER JOIN
			PC190300 ON PC030300.PC03001 = PC190300.PC19001 AND 
			PC030300.PC03002 = PC190300.PC19002 LEFT OUTER JOIN
            PC010300 ON PC030300.PC03001 = PC010300.PC01001 RIGHT OUTER JOIN
            SC330300 WITH (NOLOCK) ON PC190300.PC19005 = SC330300.SC33009
        WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
			(SC330300.SC33002 = @MySrcWarNo)
        GROUP BY SC330300.SC33001, 
			SC330300.SC33005 - SC330300.SC33006, 
			SC330300.SC33003, PC010300.PC01060) AS View_1 INNER JOIN
        #_MyOrder ON View_1.SC33001 = #_MyOrder.ItemCode LEFT OUTER JOIN
        SC010300 ON View_1.SC33001 = SC010300.SC01001
	WHERE (View_1.PC01060 NOT IN
		(SELECT OR01001
        FROM OR010300
        WHERE (OR01050 <> @MySrcWarNo)))
	GROUP BY View_1.SC33001, 
		ISNULL(SC010300.SC01135, 0), 
		#_MyOrder.QTY
	ORDER BY View_1.SC33001
END
ELSE                          -- 1 -- в заказ на перемещение запасы для заказов на продажу на других складах включаем
BEGIN
	DECLARE String_Cursor CURSOR FOR
/*
	SELECT Ltrim(Rtrim(SC330300.SC33001)) AS ItemCode, 
		CASE WHEN SUM(SC330300.SC33005 - SC330300.SC33006) >= #_MyOrder.QTY THEN #_MyOrder.QTY ELSE SUM(SC330300.SC33005 - SC330300.SC33006) END AS QTY,
		ISNULL(SC010300.SC01135, 0) AS UOM
	FROM SC330300 WITH (NOLOCK) INNER JOIN
		PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
        PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
		PC190300.PC19002 = PC030300.PC03002 INNER JOIN
        PC010300 ON PC030300.PC03001 = PC010300.PC01001 INNER JOIN
        #_MyOrder ON SC330300.SC33001 = #_MyOrder.ItemCode LEFT OUTER JOIN
        SC010300 ON SC330300.SC33001 = SC010300.SC01001
	WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
		(SC330300.SC33002 = @MySrcWarNo)
	GROUP BY SC330300.SC33001, 
		ISNULL(SC010300.SC01135, 0), 
		#_MyOrder.QTY
	ORDER BY ItemCode
*/
	SELECT LTRIM(RTRIM(SC330300.SC33001)) AS ItemCode, 
		CASE WHEN SUM(SC330300.SC33005 - SC330300.SC33006) >= #_MyOrder.QTY THEN #_MyOrder.QTY ELSE SUM(SC330300.SC33005 - SC330300.SC33006) END AS QTY, 
		ISNULL(SC010300.SC01135, 0) AS UOM
	FROM SC330300 WITH (NOLOCK) INNER JOIN
        #_MyOrder ON SC330300.SC33001 = #_MyOrder.ItemCode LEFT OUTER JOIN
		SC010300 ON SC330300.SC33001 = SC010300.SC01001
	WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
		(SC330300.SC33002 = @MySrcWarNo)
	GROUP BY SC330300.SC33001, 
		ISNULL(SC010300.SC01135, 0), 
		#_MyOrder.QTY
	ORDER BY ItemCode
END

SELECT @MyStrNum = '000010'

OPEN String_Cursor
FETCH NEXT FROM String_Cursor INTO @MyItemCode, @MyQTYStr, @MyUOM
WHILE @@FETCH_STATUS = 0
BEGIN
	IF @MyQTYStr > 0
	BEGIN
		--SELECT @MyRezStr = @MyRezStr + 'Код запаса: ' + @MyItemCode + ' Кол - во: ' + CONVERT(nvarchar(30),@MyQTYStr) + ' Единица измерения: ' + CONVERT(nvarchar(30),@MyUOM) + CHAR(13) 
		SELECT @MyQTYStrRest = @MyQTYStr
		INSERT INTO SC7D0300
			(SC7D001, SC7D002, SC7D003, SC7D004, SC7D005, SC7D006, SC7D007, SC7D008, SC7D009, SC7D010, SC7D011, SC7D012)
		VALUES (@MyRelocOrderNum, 
			@MyStrNum, 
			@MyItemCode, 
			@MyQTYStr, 
			@MyQTYStr, 
			0, 
			0, 
			@MyUOM, 
			0, 
			N'', 
			N'', 
			N'')

		-----SC110300
		-----уход
		INSERT INTO SC110300
			(SC11001, SC11002, SC11003, SC11004, SC11005, SC11006, SC11007, SC11008, SC11009, SC11010, SC11011, SC11012, SC11013, SC11014, 
			SC11015, SC11016, SC11017, SC11018, SC11019, SC11020, SC11021, SC11022, SC11023)
		VALUES (@MyItemCode,                          ---код запаса
			@MyOrderDate,                             ---дата з-за на перемещение
			@MyRelocOrderNum,                         ---номер заказа на перемещение
			@MySrcWarNo,                              ---номер склада откуда идет перемещение
			- @MyQTYStr,                              ---кол - во
			N'', 
			N'*STO*', 
			@MyShipDate,                              ---предполагаемая дата приемки
			0, 
			N'', 
			0, 
			@MyStrNum,                                ---номер строки з-за на перемещение
			N'', 
			N'', 
			N'T', 
			0x1, 
			N'', 
			0, 
			N'0', 
			N'1',                                     ---признак исходящий 1 входящий 0
			@MyZeroDate, 
			@MyZeroDate, 
			N'')
		-----приход
		INSERT INTO SC110300
			(SC11001, SC11002, SC11003, SC11004, SC11005, SC11006, SC11007, SC11008, SC11009, SC11010, SC11011, SC11012, SC11013, SC11014, 
			SC11015, SC11016, SC11017, SC11018, SC11019, SC11020, SC11021, SC11022, SC11023)
		VALUES (@MyItemCode,                          ---код запаса
			@MyOrderDate,                             ---дата з-за на перемещение
			@MyRelocOrderNum,                         ---номер заказа на перемещение
			@MyDestWarNo,                             ---номер склада назначения
			@MyQTYStr,                                ---кол - во
			N'', 
			N'*STO*', 
			@MyShipDate,                              ---предполагаемая дата приемки
			0, 
			N'', 
			0, 
			@MyStrNum,                                ---номер строки з-за на перемещение
			N'', 
			N'', 
			N'T', 
			0x1, 
			N'', 
			0, 
			N'0', 
			N'0',                                     ---признак исходящий 1 входящий 0
			@MyZeroDate, 
			@MyZeroDate, 
			N'')


		SELECT @MySubStrNum = '000010'
		---курсор по подстрокам и партиям
		IF @MyOtherWHFlag = 0         -- 0 -- в заказ на перемещение запасы для заказов на продажу на других складах не включаем
		BEGIN
			DECLARE SubString_Cursor CURSOR FOR
/*
			SELECT SC330300.SC33003, 
				SC330300.SC33005 - SC330300.SC33006 AS QTY
			FROM SC330300 WITH (NOLOCK) INNER JOIN
				PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
				PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
				PC190300.PC19002 = PC030300.PC03002 INNER JOIN
				PC010300 ON PC030300.PC03001 = PC010300.PC01001
			WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
				(SC330300.SC33002 = @MySrcWarNo) AND 
				(SC330300.SC33001 = @MyItemCode) AND 
				(PC010300.PC01060 NOT IN
					(SELECT OR01001
					FROM OR010300
					WHERE (OR01050 <> @MySrcWarNo)))
			GROUP BY SC330300.SC33005 - SC330300.SC33006, 
				SC330300.SC33003
			ORDER BY SC330300.SC33003
*/
			SELECT SC33003, 
				QTY
			FROM (SELECT SC330300.SC33003, 
					SC330300.SC33005 - SC330300.SC33006 AS QTY, 
					ISNULL(PC010300.PC01060, N'') AS PC01060
				FROM PC030300 RIGHT OUTER JOIN
					PC190300 ON PC030300.PC03001 = PC190300.PC19001 AND 
					PC030300.PC03002 = PC190300.PC19002 LEFT OUTER JOIN
					PC010300 ON PC030300.PC03001 = PC010300.PC01001 RIGHT OUTER JOIN
                    SC330300 WITH (NOLOCK) ON PC190300.PC19005 = SC330300.SC33009
                WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
					(SC330300.SC33002 = @MySrcWarNo) AND 
					(SC330300.SC33001 = @MyItemCode)
                GROUP BY SC330300.SC33005 - SC330300.SC33006, 
					SC330300.SC33003, 
					PC010300.PC01060) AS View_1
			WHERE (PC01060 NOT IN
				(SELECT OR01001
                FROM OR010300
                WHERE (OR01050 <> @MySrcWarNo)))
			ORDER BY SC33003
		END
		ELSE                         -- 1 -- в заказ на перемещение запасы для заказов на продажу на других складах включаем
		BEGIN
			DECLARE SubString_Cursor CURSOR FOR
/*
			SELECT SC330300.SC33003, 
				SC330300.SC33005 - SC330300.SC33006 AS QTY
			FROM SC330300 WITH (NOLOCK) INNER JOIN
				PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
				PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
				PC190300.PC19002 = PC030300.PC03002 INNER JOIN
				PC010300 ON PC030300.PC03001 = PC010300.PC01001
			WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
				(SC330300.SC33002 = @MySrcWarNo) AND 
				(SC330300.SC33001 = @MyItemCode)
			GROUP BY SC330300.SC33005 - SC330300.SC33006, 
				SC330300.SC33003
			ORDER BY SC330300.SC33003
*/
			SELECT SC33003, 
				SC33005 - SC33006 AS QTY
			FROM SC330300 WITH (NOLOCK)
			WHERE (SC33005 - SC33006 > 0) AND 
				(SC33002 = @MySrcWarNo) AND 
				(SC33001 = @MyItemCode)
			GROUP BY SC33005 - SC33006, 
				SC33003
			ORDER BY SC33003
		END

		OPEN SubString_Cursor
		FETCH NEXT FROM SubString_Cursor INTO @MyBatchNum, @MyBatchQTY
		WHILE @@FETCH_STATUS = 0
		BEGIN
			IF @MyQTYStrRest < @MyBatchQTY
			BEGIN
				SELECT @MyBatchQTY = @MyQTYStrRest
			END

			IF @MyBatchQTY > 0 --отгрузка по партиям только если количество больше нуля
			BEGIN
				INSERT INTO SC7E0300
					(SC7E001, SC7E002, SC7E003, SC7E004, SC7E005, SC7E006, SC7E007, SC7E008, SC7E009, SC7E010, SC7E011, SC7E012, SC7E013, SC7E014, SC7E015)
				VALUES (@MyRelocOrderNum, 
					@MyStrNum, 
					@MySubStrNum, 
					@MyBatchNum, 
					N'', 
					N'', 
					0, 
					@MyOrderDate, 
					N'0000000000',
					0, 
					@MyShipDate, 
					N'0000000000', 
					@MyBatchQTY, 
					@MyOrderDate, 
					N'0000000000')

				---обновление данных в SC33
				UPDATE SC330300 WITH (ROWLOCK)
				SET SC33006 = SC33006 + @MyBatchQTY
				WHERE (SC33001 = @MyItemCode) AND 
					(SC33002 = @MySrcWarNo) AND
					(SC33003 = @MyBatchNum)

				SELECT @MyQTYStrRest = @MyQTYStrRest - @MyBatchQTY
				SELECT @MySubStrNum = RIGHT('000000' + CONVERT(nvarchar,(Convert(int,@MySubStrNum) / 10 + 1) * 10),6)
			END
			FETCH NEXT FROM SubString_Cursor INTO @MyBatchNum, @MyBatchQTY
		END
		CLOSE SubString_Cursor
		DEALLOCATE SubString_Cursor
		---конец курсора по подстрокам и партиям

		---обновление данных в SC03
		UPDATE SC030300 WITH (ROWLOCK)
		SET SC03004 = SC03004 + @MyQTYStr,
			SC03016 = SC03016 + @MyQTYStr
		WHERE (SC03001 = @MyItemCode) AND
			(SC03002 = @MySrcWarNo)

		---обновление данных в SC01
		UPDATE SC010300 WITH (ROWLOCK)
		SET  SC01043 = SC01043 + @MyQTYStr,
			SC01183 = SC01183 + @MyQTYStr
		WHERE (SC01001 = @MyItemCode)

		---обновление данных во временной таблице
		UPDATE #_MyOrder
		SET RestQTY = RestQTY - @MyQTYStr
		WHERE (ItemCode = @MyItemCode)

		SELECT @MyStrNum = RIGHT('000000' + CONVERT(nvarchar,(Convert(int,@MyStrNum) / 10 + 1) * 10),6)
	END
	FETCH NEXT FROM String_Cursor INTO @MyItemCode, @MyQTYStr, @MyUOM
END
CLOSE String_Cursor
DEALLOCATE String_Cursor


---------------------Проверяем - есть ли строки в з-зе на перемещение. если нет - удаляем заголовок----
SELECT @MyCount1 = COUNT(*)
FROM SC7D0300
WHERE (SC7D001 = @MyRelocOrderNum)

IF @MyCount1 = 0
BEGIN
	DELETE FROM SC7C0300
	WHERE (SC7C001 = @MyRelocOrderNum)

	DELETE FROM tbl_DisplacementOrder
	WHERE (OrderNumber = @MyRelocOrderNum)

	SELECT @MyRelocOrderNum = ''
END

