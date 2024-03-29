USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_ShipmentsCost_Fact]    Дата сценария: 07/04/2012 10:08:08 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_ShipmentsCost_Fact](
	[ID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_tbl_ShipmentsCost_Fact_ID]  DEFAULT (newid()),
	[PL01001] [nvarchar](10) NOT NULL,
	[PL03002] [nvarchar](15) NOT NULL,
	[DocDate] [datetime] NOT NULL,
	[DocSumm] [float] NOT NULL,
	[ShipmentsType] [int] NOT NULL
) ON [PRIMARY]
