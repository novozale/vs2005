USE [ScaDataDB]
GO
/****** Object:  Table [dbo].[tbl_ShipmentsCost_Price]    Script Date: 03/17/2011 14:31:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_ShipmentsCost_Price](
	[ID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_tbl_ShipmentsCost_Price_ID]  DEFAULT (newid()),
	[WHNum] [nvarchar](6) NOT NULL,
	[ShipmentsType] [int] NOT NULL,
	[CostType] [int] NOT NULL,
	[Destination] [nvarchar](255) NOT NULL,
	[PriceType] [int] NOT NULL,
	[PriceFrom] [float] NOT NULL,
	[PriceTo] [float] NOT NULL,
	[PriceVal] [float] NOT NULL
) ON [PRIMARY]
