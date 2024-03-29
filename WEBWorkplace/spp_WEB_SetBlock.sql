USE [ScaDataDB]
GO
/****** Объект:  StoredProcedure [dbo].[spp_WEB_SetBlock]    Дата сценария: 05/29/2015 12:43:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[spp_WEB_SetBlock] 
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Установка блокировки в таблице tbl_WEB_2Block                                     |
|                                                                                      |
|    Разработчик Новожилов А.Н. 2015                                                   |
|                                                                                      |
--------------------------------------------------------------------------------------*/
@MyUser nvarchar(255)

WITH RECOMPILE
AS

Declare @BlockUser nvarchar(255);            --пользователь, который выставил блокировку
Declare @BlockTime datetime;                 --дата и время выставления блокировки
Declare @WaitingTime integer;                --установленное максимальное время ожидания
declare @MyCounter integer;                  --счетчик
declare @MyDateNew datetime;                 --Новое значение даты / времени (текущее)
Declare @MyTimeStr nvarchar(10);             --строка для задания промежутка времени

SELECT @MyCounter = COUNT(*)
FROM tbl_WEB_2Block WITH(NOLOCK)

if @MyCounter = 0 --записи в таблице нет - создаем по умолчанию
BEGIN
	INSERT INTO tbl_WEB_2Block
    ([User], ActionDt, WaitingTime)
	VALUES (@MyUser, GETDATE(), 900)
END
ELSE
BEGIN
	SELECT @BlockUser = [User], 
		@BlockTime = ActionDt, 
		@WaitingTime = WaitingTime
	FROM tbl_WEB_2Block WITH(NOLOCK)

	IF Ltrim(Rtrim(@BlockUser)) = '' --никто не блокирует
	BEGIN
		Update tbl_WEB_2Block 
        Set [User] = @MyUser, 
        ActionDt = GETDATE() 
	END
	ELSE
	BEGIN  --кто - то заблокировал
		SELECT @MyDateNew = GETDATE()
		WHILE Ltrim(Rtrim(@BlockUser)) != '' AND Datediff(s, @BlockTime, @MyDateNew) < @WaitingTime
		BEGIN
			SELECT @MyTimeStr = '00:00:' + Right('00' + CONVERT(nvarchar(10),Round(3*RAND(),0)),2)
			WAITFOR DELAY @MyTimeStr

			SELECT @BlockUser = [User], 
			@BlockTime = ActionDt, 
			@WaitingTime = WaitingTime,
			@MyDateNew = GETDATE()
			FROM tbl_WEB_2Block WITH(NOLOCK)
		END
		IF Ltrim(Rtrim(@BlockUser)) != ''
		BEGIN
			INSERT INTO tbl_WEB_2UnblockInfo 
            (ID, [User], ActionDt, Descr)
            VALUES (NEWID(), 
            @BlockUser,
            GETDATE(),
            N'Removed')
		END
		
		Update tbl_WEB_2Block 
        Set [User] = @MyUser, 
        ActionDt = GETDATE()
	END
END
