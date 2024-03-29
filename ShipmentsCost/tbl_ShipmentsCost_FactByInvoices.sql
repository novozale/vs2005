USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_ShipmentsCost_FactByInvoices]    Дата сценария: 07/04/2012 10:07:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_ShipmentsCost_FactByInvoices](
	[ID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_tbl_ShipmentCost_FactByInvoices_ID]  DEFAULT (newid()),
	[DocID] [uniqueidentifier] NOT NULL,
	[SL03002] [nvarchar](15) NOT NULL,
	[InvoiceSumm] [float] NOT NULL,
	[ShipmentCost] [float] NOT NULL
) ON [PRIMARY]
