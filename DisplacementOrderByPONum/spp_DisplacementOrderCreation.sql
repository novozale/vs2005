USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_DisplacementOrderCreation]    Дата сценария: 11/22/2012 12:34:06 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spp_DisplacementOrderCreation]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Процедура формирования заказа на перемещение товаров из заказа на закупку         |
|    Разработчик Новожилов А.Н. 2012                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyPurchaseOrderNum nvarchar(15),                                   -- номер заказа на закупку
@MyDestWarNo nvarchar(6),                                           -- склад назначения
@MyOrderDate datetime,                                              -- дата предполагаемой отгрузки
@MyShipDate datetime,                                               -- дата доставки (предполагаемая)
@MyRezStr nvarchar(4000) output                                     -- результирующая строка

AS

DECLARE @MyWarNo nvarchar(6)                                        -- исходный склад
--DECLARE @MyOrderDate datetime                                       -- дата создания заказа
--DECLARE @MyShipDate datetime                                        -- дата доставки (предполагаемая)
DECLARE @MyZeroDate datetime                                        -- 1 января 1900 года
DECLARE @MyUserID int                                               -- код пользователя
DECLARE @MyUserName nvarchar(35)                                    -- имя пользователя
DECLARE @MySalesOrderNum nvarchar(10)                               -- Номер заказа на продажу
DECLARE @MyRowNum float                                             -- кол - во строк
DECLARE @MyDisplacementOrderN nvarchar(12)                          -- номер заказа на перемещение
DECLARE @MyStrNum nvarchar(6)                                       -- номер строки в заказе
DECLARE @MyItemCode nvarchar(35)                                    -- код запаса
DECLARE @MyUOM int                                                  -- единица измерения (SC01135 - продажи)
DECLARE @MyQTYStr float                                             -- кол - во к перемещению в строке
DECLARE @MySubStrNum nvarchar(6)                                    -- номер подстроки в заказе
DECLARE @MyBatchNum nvarchar(12)                                    -- номер партии запаса
DECLARE @MyBatchQTY float                                           -- доступное кол - во в партии

SELECT @MyRezStr = ''
SELECT @MyPurchaseOrderNum = Ltrim(Rtrim(@MyPurchaseOrderNum))
SELECT @MyDestWarNo = Ltrim(Rtrim(@MyDestWarNo))

--SELECT @MyRezStr = @MyRezStr + '@MyPurchaseOrderNum ' + ISNULL(@MyPurchaseOrderNum,'NULL') + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@MyDestWarNo ' + ISNULL(@MyDestWarNo,'NULL') + CHAR(13)
---------------------Определение исходного склада----------------------------------------
SELECT @MyWarNo = PC01023 
FROM PC010300 WITH(NOLOCK)
WHERE (PC01001 = @MyPurchaseOrderNum)

--SELECT @MyRezStr = @MyRezStr + '@MyWarNo ' + ISNULL(@MyWarNo,'NULL') + CHAR(13)

---------------------Даты----------------------------------------------------------------
--SELECT @MyOrderDate = dateadd( day, datediff(day, 0, GETDATE()), 0)
--SELECT @MyShipDate = DATEADD(day, 1, @MyOrderDate)
SELECT @MyZeroDate = CONVERT(datetime,'01/01/1900',103)

--SELECT @MyRezStr = @MyRezStr + '@MyOrderDate ' + CONVERT(nvarchar(30),ISNULL(@MyOrderDate,'NULL'),103) + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@MyShipDate ' + CONVERT(nvarchar(30),ISNULL(@MyShipDate,'NULL'),103) + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@MyZeroDate ' + CONVERT(nvarchar(30),ISNULL(@MyZeroDate,'NULL'),103) + CHAR(13)

---------------------Пользователь--------------------------------------------------------
SELECT @MyUserName = CASE WHEN dbo.GetUserByPCANDThread(host_name, host_process_id) = '' 
			THEN 'Admin'
			ELSE Right(dbo.GetUserByPCANDThread(host_name, host_process_id),LEN(dbo.GetUserByPCANDThread(host_name, host_process_id)) - LEN('EskRU\')) END 
		FROM sys.dm_exec_sessions
        WHERE (session_id = @@SPID)
--PRINT @MyUserName
--SELECT @MyRezStr = @MyRezStr + '@MyUserName ' + ISNULL(@MyUserName,'NULL') + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@@SPID ' + CONVERT(nvarchar(30),ISNULL(@@SPID,'NULL')) + CHAR(13)


SELECT @MyUserID = UserID
	FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK)
	WHERE (UPPER(UserName) = UPPER(@MyUserName))
--PRINT CONVERT(nvarchar(35),@MyUserID)
--SELECT @MyRezStr = @MyRezStr + '@MyUserID ' + CONVERT(nvarchar(30),ISNULL(@MyUserID,'NULL')) + CHAR(13)

---------------------Номер Заказа на продажу---------------------------------------------
SELECT @MySalesOrderNum = Ltrim(Rtrim(OR010300.OR01001))
FROM PC010300 WITH(NOLOCK) INNER JOIN
	OR010300 ON PC010300.PC01060 = OR010300.OR01001
WHERE (PC010300.PC01001 = @MyPurchaseOrderNum)

IF @MySalesOrderNum IS NULL
BEGIN
	SELECT @MySalesOrderNum = ''
END

--SELECT @MyRezStr = @MyRezStr + '@MySalesOrderNum ' + ISNULL(@MySalesOrderNum,'NULL') + CHAR(13)

---------------------кол - во строк (запасов) из которых можно сформировать заказ--------
SELECT @MyRowNum = COUNT(PC010300.PC01001)
FROM SC330300 WITH(NOLOCK) INNER JOIN
	PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
    PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
	PC190300.PC19002 = PC030300.PC03002 INNER JOIN
    PC010300 ON PC030300.PC03001 = PC010300.PC01001
WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
	(PC010300.PC01001 = @MyPurchaseOrderNum) AND 
	(SC330300.SC33002 = @MyWarNo)

--SELECT @MyRezStr = @MyRezStr + '@MyRowNum ' + CONVERT(nvarchar(30),ISNULL(@MyRowNum,'NULL')) + CHAR(13)

IF @MyRowNum > 0
BEGIN  --если есть из чего - формируем заказ на перемещение
	----номер заказа и увеличение счетчика
	SELECT @MyDisplacementOrderN = RIGHT('0000000000' + CONVERT(nvarchar,SY68002),10)
	FROM SY6803XX
	WHERE (SY68001 = N'SC7C')

	UPDATE SY6803XX
	SET SY68002 = CONVERT(float,@MyDisplacementOrderN) + 1
	WHERE (SY68001 = N'SC7C')

--SELECT @MyRezStr = @MyRezStr + '@MyDisplacementOrderN ' + @MyDisplacementOrderN + CHAR(13)

	----Заголовок заказа на перемещение
	INSERT INTO SC7C0300
		(SC7C001, SC7C002, SC7C003, SC7C004, SC7C005, SC7C006, SC7C007, SC7C008, SC7C009, SC7C010)
	VALUES (@MyDisplacementOrderN,  --номер заказа
		@MyOrderDate,              --дата заказа
		'', 
		@MyWarNo,                  --исходный склад
		'', 
		@MyDestWarNo,              --склад назначения
		@MyOrderDate,              --дата заказа
		@MyZeroDate,               --01.01.1900
		@MyShipDate,               --предполагаемая дата доставки
		N'1')                      --флаг - 1 тип

	----Данные в доп таблицу tbl_DispllacementOrder
	DELETE FROM tbl_DisplacementOrder
	WHERE (OrderNumber = @MyDisplacementOrderN)


--SELECT @MyRezStr = @MyRezStr + '-----' + '@MyDisplacementOrderN ' + ISNULL(@MyDisplacementOrderN,'NULL') + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@MyUserID ' + CONVERT(nvarchar(50),ISNULL(@MyUserID,'NULL')) + CHAR(13)
--SELECT @MyRezStr = @MyRezStr + '@MySalesOrderNum ' + ISNULL(@MySalesOrderNum,'NULL') + CHAR(13)

	INSERT INTO tbl_DisplacementOrder
    (ID, OrderNumber, UserCode, SalesOrderNumber, ReadyFlag)
	VALUES (NEWID(), 
		@MyDisplacementOrderN, 
		@MyUserID, 
		@MySalesOrderNum, 
		1)

	SELECT @MyStrNum = '000010'
	----курсор по запасам - для формирования строк заказа
	DECLARE String_Cursor CURSOR FOR
	SELECT SC330300.SC33001 AS ItemCode, 
		SUM(SC330300.SC33005 - SC330300.SC33006) AS QTY, 
		ISNULL(SC010300.SC01135, 0) AS UOM
	FROM SC330300 WITH(NOLOCK) INNER JOIN
		PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
        PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
		PC190300.PC19002 = PC030300.PC03002 INNER JOIN
        PC010300 ON PC030300.PC03001 = PC010300.PC01001 LEFT OUTER JOIN
        SC010300 ON SC330300.SC33001 = SC010300.SC01001
	WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
		(PC010300.PC01001 = @MyPurchaseOrderNum) AND 
		(SC330300.SC33002 = @MyWarNo)
	GROUP BY SC330300.SC33001, 
		ISNULL(SC010300.SC01135, 0)
	ORDER BY ItemCode

	OPEN String_Cursor
	FETCH NEXT FROM String_Cursor INTO @MyItemCode, @MyQTYStr, @MyUOM
	WHILE @@FETCH_STATUS = 0
	BEGIN
		INSERT INTO SC7D0300
			(SC7D001, SC7D002, SC7D003, SC7D004, SC7D005, SC7D006, SC7D007, SC7D008, SC7D009, SC7D010, SC7D011, SC7D012)
		VALUES (@MyDisplacementOrderN, 
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
			@MyDisplacementOrderN,                    ---номер заказа на перемещение
			@MyWarNo,                                 ---номер склада
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
			@MyDisplacementOrderN,                    ---номер заказа на перемещение
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
		DECLARE SubString_Cursor CURSOR FOR
		SELECT SC330300.SC33003, 
			SC330300.SC33005 - SC330300.SC33006 AS QTY
		FROM SC330300 WITH(NOLOCK) INNER JOIN
			PC190300 ON SC330300.SC33009 = PC190300.PC19005 INNER JOIN
            PC030300 ON PC190300.PC19001 = PC030300.PC03001 AND 
			PC190300.PC19002 = PC030300.PC03002 INNER JOIN
            PC010300 ON PC030300.PC03001 = PC010300.PC01001
		WHERE (SC330300.SC33005 - SC330300.SC33006 > 0) AND 
			(SC330300.SC33001 = @MyItemCode) AND 
			(PC010300.PC01001 = @MyPurchaseOrderNum) AND
			(SC330300.SC33002 = @MyWarNo)
		GROUP BY SC330300.SC33005 - SC330300.SC33006, 
			SC330300.SC33003

		OPEN SubString_Cursor
		FETCH NEXT FROM SubString_Cursor INTO @MyBatchNum, @MyBatchQTY
		WHILE @@FETCH_STATUS = 0
		BEGIN
			INSERT INTO SC7E0300
				(SC7E001, SC7E002, SC7E003, SC7E004, SC7E005, SC7E006, SC7E007, SC7E008, SC7E009, SC7E010, SC7E011, SC7E012, SC7E013, SC7E014, SC7E015)
			VALUES (@MyDisplacementOrderN, 
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
				UPDATE SC330300
				SET SC33006 = SC33006 + @MyBatchQTY
				WHERE (SC33001 = @MyItemCode) AND 
					(SC33002 = @MyWarNo) AND
					(SC33003 = @MyBatchNum)

			SELECT @MySubStrNum = RIGHT('000000' + CONVERT(nvarchar,(Convert(int,@MySubStrNum) / 10 + 1) * 10),6)
			FETCH NEXT FROM SubString_Cursor INTO @MyBatchNum, @MyBatchQTY
		END
		CLOSE SubString_Cursor
		DEALLOCATE SubString_Cursor
		

		---конец курсора по подстрокам и партиям
		---обновление данных в SC03
		UPDATE SC030300
		SET SC03004 = SC03004 + @MyQTYStr,
			SC03016 = SC03016 + @MyQTYStr
		WHERE (SC03001 = @MyItemCode) AND
			(SC03002 = @MyWarNo)

		---обновление данных в SC01
		UPDATE SC010300
		SET  SC01043 = SC01043 + @MyQTYStr,
			SC01183 = SC01183 + @MyQTYStr
		WHERE (SC01001 = @MyItemCode)

		SELECT @MyStrNum = RIGHT('000000' + CONVERT(nvarchar,(Convert(int,@MyStrNum) / 10 + 1) * 10),6)
		FETCH NEXT FROM String_Cursor INTO @MyItemCode, @MyQTYStr, @MyUOM
	END
	CLOSE String_Cursor
	DEALLOCATE String_Cursor

	SELECT @MyRezStr = @MyRezStr + 'номер заказа на перемещение ' + ISNULL(@MyDisplacementOrderN,'')
END
ELSE
BEGIN  --Если не из чего - сообщаем об этом
	SELECT @MyRezStr = @MyRezStr + 'По заказу на закупку с номером ' + ISNULL(@MyPurchaseOrderNum,'') + ' перемещать нечего. Свободных к перемещению запасов нет. Или заказ еще не принят, или все принятое по этому заказу зарезервировано, или уже продано и отгружено. '
END
