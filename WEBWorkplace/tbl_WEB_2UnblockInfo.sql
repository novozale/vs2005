USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_2UnblockInfo]    Дата сценария: 05/29/2015 12:35:35 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_2UnblockInfo](
	[ID] [uniqueidentifier] ROWGUIDCOL  NOT NULL CONSTRAINT [DF_tbl_WEB_2UnblockInfo_ID]  DEFAULT (newid()),
	[User] [nvarchar](50) NOT NULL,
	[ActionDt] [datetime] NOT NULL,
	[Descr] [nvarchar](50) NOT NULL
) ON [PRIMARY]
