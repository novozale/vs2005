USE [AvayaLogs]
GO
/****** Объект:  Table [dbo].[tbl_Data]    Дата сценария: 08/31/2013 19:36:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_Data](
	[ID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_tbl_Data_ID]  DEFAULT (newid()),
	[AvayaIP] [nvarchar](20) NULL,
	[CALL_TIME] [datetime] NULL,
	[DURATION] [time] NULL,
	[DURATION_S] [int] NULL,
	[CALLER_PHONE] [nvarchar](40) NULL,
	[DIRECTION] [nvarchar](4) NULL,
	[DIALED_PHONE] [nvarchar](40) NULL,
	[Field7] [nvarchar](40) NULL,
	[ACC] [nvarchar](40) NULL,
	[INTERNAL] [nvarchar](4) NULL,
	[CallID] [nvarchar](40) NULL,
	[ANOTHER_RECORDS] [nvarchar](4) NULL,
	[ABONENT_DEVICE1] [nvarchar](40) NULL,
	[ABONENT_NAME1] [nvarchar](40) NULL,
	[ABONENT_DEVICE2] [nvarchar](40) NULL,
	[ABONENT_NAME2] [nvarchar](40) NULL,
	[RETENTION_TIME] [int] NULL,
	[PARKING_TIME] [int] NULL,
	[AUTH_VALID] [nvarchar](4) NULL,
	[AUTH_CODE] [nvarchar](40) NULL,
	[USER_PAYMENTS] [nvarchar](40) NULL,
	[COST] [float] NULL,
	[CURRENCY] [nvarchar](40) NULL,
	[AOC_QTY] [float] NULL,
	[UNITS_QTY] [float] NULL,
	[AOC_UNITS_QTY] [float] NULL,
	[COST_PER_UNIT] [float] NULL,
	[ALLOWANCE] [float] NULL,
	[OUT_CALL_REASON] [nvarchar](40) NULL,
	[OUT_ABONENT_ID] [nvarchar](40) NULL,
	[OUT_CALL_NUMBER] [nvarchar](40) NULL,
 CONSTRAINT [PK_tbl_Data] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
