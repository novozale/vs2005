USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_CardSales]    Дата сценария: 05/29/2015 12:36:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_CardSales](
	[ClientCode] [nvarchar](50) NOT NULL,
	[OrderDate] [datetime] NOT NULL,
	[OrderNum] [nvarchar](15) NOT NULL,
	[Discount] [nvarchar](50) NOT NULL,
	[OrderSumm] [numeric](28, 8) NOT NULL,
	[ShipmentState] [int] NOT NULL,
	[OrderState] [int] NOT NULL,
	[WEBOrderNum] [nvarchar](50) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL
) ON [PRIMARY]
