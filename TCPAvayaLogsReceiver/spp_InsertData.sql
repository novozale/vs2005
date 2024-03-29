USE [AvayaLogs]
GO
/****** Объект:  StoredProcedure [dbo].[spp_InsertData]    Дата сценария: 08/31/2013 19:37:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

 CREATE Procedure [dbo].[spp_InsertData]
/*-------------------------------------------------------------------------------------
|                                                                                      |
|    Вставляет данные по 1 событию из телефонной станции                               |
|    автор - Новожилов А.Н. 01.09.2013                                                 |
--------------------------------------------------------------------------------------*/
@AvayaIP [nvarchar](20),
@CALL_TIME [nvarchar](40),
@DURATION [nvarchar](25),
@DURATION_S [integer],
@CALLER_PHONE [nvarchar](40),
@DIRECTION [nvarchar](4),
@DIALED_PHONE [nvarchar](40),
@Field7 [nvarchar](40),
@ACC [nvarchar](40),
@INTERNAL [nvarchar](4),
@CallID [nvarchar](40),
@ANOTHER_RECORDS [nvarchar](4),
@ABONENT_DEVICE1 [nvarchar](40),
@ABONENT_NAME1 [nvarchar](40),
@ABONENT_DEVICE2 [nvarchar](40),
@ABONENT_NAME2 [nvarchar](40),
@RETENTION_TIME [integer],
@PARKING_TIME [integer],
@AUTH_VALID [nvarchar](4),
@AUTH_CODE [nvarchar](40),
@USER_PAYMENTS [nvarchar](40),
@COST [nvarchar](40), 
@CURRENCY [nvarchar](40),
@AOC_QTY [nvarchar](40),
@UNITS_QTY [nvarchar](40),
@AOC_UNITS_QTY [nvarchar](40),
@COST_PER_UNIT [nvarchar](40),
@ALLOWANCE [nvarchar](40),
@OUT_CALL_REASON [nvarchar](40),
@OUT_ABONENT_ID [nvarchar](40),
@OUT_CALL_NUMBER [nvarchar](40)

WITH RECOMPILE
AS


--------Объявление переменных
DECLARE @MyAvayaIP [nvarchar](20);
DECLARE @MyCALL_TIME [datetime];            --
DECLARE @MyDURATION [time](7);              --
DECLARE @MyCALLER_PHONE [nvarchar](40);
DECLARE @MyDIRECTION [nvarchar](4);
DECLARE @MyDIALED_PHONE [nvarchar](40);
DECLARE @MyField7 [nvarchar](40);
DECLARE @MyACC [nvarchar](40);
DECLARE @MyINTERNAL [nvarchar](4);
DECLARE @MyCallID [nvarchar](40);
DECLARE @MyANOTHER_RECORDS [nvarchar](4);
DECLARE @MyABONENT_DEVICE1 [nvarchar](40);
DECLARE @MyABONENT_NAME1 [nvarchar](40);
DECLARE @MyABONENT_DEVICE2 [nvarchar](40);
DECLARE @MyABONENT_NAME2 [nvarchar](40);
DECLARE @MyAUTH_VALID [nvarchar](4);
DECLARE @MyAUTH_CODE [nvarchar](40);
DECLARE @MyUSER_PAYMENTS [nvarchar](40);
DECLARE @MyCOST [float];                     --
DECLARE @MyCURRENCY [nvarchar](40);
DECLARE @MyAOC_QTY [float];                  --
DECLARE @MyUNITS_QTY [float];                --
DECLARE @MyAOC_UNITS_QTY [float];            --
DECLARE @MyCOST_PER_UNIT [float];            --
DECLARE @MyALLOWANCE [float];                --
DECLARE @MyOUT_CALL_REASON [nvarchar](40);
DECLARE @MyOUT_ABONENT_ID [nvarchar](40);
DECLARE @MyOUT_CALL_NUMBER [nvarchar](40);

---------Преобразования
SELECT @MyAvayaIP = LTRIM(Rtrim(@AvayaIP))
SELECT @MyCALL_TIME = CONVERT(datetime,@CALL_TIME,120) 
SELECT @MyDURATION = CONVERT(time(7),@DURATION) 
SELECT @MyCALLER_PHONE = LTRIM(Rtrim(@CALLER_PHONE))
SELECT @MyDIRECTION = LTRIM(Rtrim(@DIRECTION))
SELECT @MyDIALED_PHONE = LTRIM(Rtrim(@DIALED_PHONE))
SELECT @MyField7 = LTRIM(Rtrim(@Field7))
SELECT @MyACC = LTRIM(Rtrim(@ACC))
SELECT @MyINTERNAL = LTRIM(Rtrim(@INTERNAL))
SELECT @MyCallID = LTRIM(Rtrim(@CallID))
SELECT @MyANOTHER_RECORDS = LTRIM(Rtrim(@ANOTHER_RECORDS))
SELECT @MyABONENT_DEVICE1 = LTRIM(Rtrim(@ABONENT_DEVICE1))
SELECT @MyABONENT_NAME1 = LTRIM(Rtrim(@ABONENT_NAME1))
SELECT @MyABONENT_DEVICE2 = LTRIM(Rtrim(@ABONENT_DEVICE2))
SELECT @MyABONENT_NAME2 = LTRIM(Rtrim(@ABONENT_NAME2))
SELECT @MyAUTH_VALID = LTRIM(Rtrim(@AUTH_VALID))
SELECT @MyAUTH_CODE = LTRIM(Rtrim(@AUTH_CODE))
SELECT @MyUSER_PAYMENTS = LTRIM(Rtrim(@USER_PAYMENTS))
SELECT @MyCOST = CONVERT(float,@COST)
SELECT @MyCURRENCY = LTRIM(Rtrim(@CURRENCY))
SELECT @MyAOC_QTY = CONVERT(float,@AOC_QTY)
SELECT @MyUNITS_QTY = CONVERT(float,@UNITS_QTY)
SELECT @MyAOC_UNITS_QTY = CONVERT(float,@AOC_UNITS_QTY)
SELECT @MyCOST_PER_UNIT = CONVERT(float,@COST_PER_UNIT)
SELECT @MyALLOWANCE = CONVERT(float,@ALLOWANCE)
SELECT @MyOUT_CALL_REASON = LTRIM(Rtrim(@OUT_CALL_REASON))
SELECT @MyOUT_ABONENT_ID = LTRIM(Rtrim(@OUT_ABONENT_ID))
SELECT @MyOUT_CALL_NUMBER = LTRIM(Rtrim(@OUT_CALL_NUMBER))




INSERT INTO tbl_Data
	(ID,      AvayaIP,    CALL_TIME,    DURATION,    DURATION_S,  CALLER_PHONE,    DIRECTION,    DIALED_PHONE,    Field7,    ACC,    INTERNAL,    CallID,    ANOTHER_RECORDS,    ABONENT_DEVICE1,    ABONENT_NAME1,    ABONENT_DEVICE2,    ABONENT_NAME2,    RETENTION_TIME,  PARKING_TIME,  AUTH_VALID,    AUTH_CODE,    USER_PAYMENTS,    COST,    CURRENCY,    AOC_QTY,    UNITS_QTY,    AOC_UNITS_QTY,    COST_PER_UNIT,    ALLOWANCE,    OUT_CALL_REASON,    OUT_ABONENT_ID,    OUT_CALL_NUMBER)
VALUES     
	(NEWID(), @MyAvayaIP, @MyCALL_TIME, @MyDURATION, @DURATION_S, @MyCALLER_PHONE, @MyDIRECTION, @MyDIALED_PHONE, @MyField7, @MyACC, @MyINTERNAL, @MyCallID, @MyANOTHER_RECORDS, @MyABONENT_DEVICE1, @MyABONENT_NAME1, @MyABONENT_DEVICE2, @MyABONENT_NAME2, @RETENTION_TIME, @PARKING_TIME, @MyAUTH_VALID, @MyAUTH_CODE, @MyUSER_PAYMENTS, @MyCOST, @MyCURRENCY, @MyAOC_QTY, @MyUNITS_QTY, @MyAOC_UNITS_QTY, @MyCOST_PER_UNIT, @MyALLOWANCE, @MyOUT_CALL_REASON, @MyOUT_ABONENT_ID, @MyOUT_CALL_NUMBER)