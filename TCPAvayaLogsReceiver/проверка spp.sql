USE [AvayaLogs]
GO

DECLARE	@return_value int

EXEC	@return_value = [dbo].[spp_InsertData]
		@AvayaIP = N'192.168.10.177',
		@CALL_TIME = N'2013/08/31 15:00:59',
		@DURATION = N'00:00:06',
		@DURATION_S = 0,
		@CALLER_PHONE = N' ',
		@DIRECTION = N'o',
		@DIALED_PHONE = N'89219615419',
		@Field7 = N'89219615419',
		@ACC = N' ',
		@INTERNAL = N'0',
		@CallID = N'1000370',
		@ANOTHER_RECORDS = N'0',
		@ABONENT_DEVICE1 = N'E121',
		@ABONENT_NAME1 = N'Danilov',
		@ABONENT_DEVICE2 = N'T9013',
		@ABONENT_NAME2 = N'Line 13.29',
		@RETENTION_TIME = 0,
		@PARKING_TIME = 0,
		@AUTH_VALID = N' ',
		@AUTH_CODE = N' ',
		@USER_PAYMENTS = N' ',
		@COST = N'0000.00',
		@CURRENCY = N' ',
		@AOC_QTY = N'0000.00',
		@UNITS_QTY = N'0',
		@AOC_UNITS_QTY = N'0',
		@COST_PER_UNIT = N'618',
		@ALLOWANCE = N'0.01',
		@OUT_CALL_REASON = N'U MT',
		@OUT_ABONENT_ID = N'Danilov',
		@OUT_CALL_NUMBER = N'89219615419'

SELECT	'Return Value' = @return_value

GO