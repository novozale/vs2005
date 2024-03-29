USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_CardSalesDetails]    Дата сценария: 05/29/2015 12:36:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_CardSalesDetails](
	[ClientCode] [nvarchar](50) NOT NULL,
	[OrderDate] [datetime] NOT NULL,
	[OrderNum] [nvarchar](15) NOT NULL,
	[StrNum] [nvarchar](50) NOT NULL,
	[ItemCode] [nvarchar](35) NOT NULL,
	[OrderedQTY] [numeric](28, 8) NOT NULL,
	[ReadyQTY] [numeric](28, 8) NOT NULL,
	[ShippedQTY] [numeric](28, 8) NOT NULL,
	[Price] [numeric](28, 8) NOT NULL,
	[ItemName] [nvarchar](100) NOT NULL,
	[IsClosed] [int] NOT NULL,
	[OrderType] [int] NOT NULL
) ON [PRIMARY]
