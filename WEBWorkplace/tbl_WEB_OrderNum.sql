USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_OrderNum]    Дата сценария: 05/29/2015 12:38:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_OrderNum](
	[ID] [uniqueidentifier] NOT NULL,
	[ScaOrderNUm] [nvarchar](10) NOT NULL,
	[WebOrderNum] [nvarchar](15) NOT NULL
) ON [PRIMARY]
