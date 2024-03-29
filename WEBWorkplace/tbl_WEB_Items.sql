USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_Items]    Дата сценария: 05/29/2015 12:38:15 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_Items](
	[Code] [nvarchar](35) NOT NULL,
	[ManufacturerItemCode] [nvarchar](100) NOT NULL,
	[ManufacturerCode] [bigint] NULL,
	[CountryCode] [nvarchar](10) NULL,
	[Country] [nvarchar](50) NULL,
	[GroupCode] [nvarchar](50) NULL,
	[SubGroupCode] [nvarchar](50) NULL,
	[Name] [nvarchar](100) NOT NULL,
	[WEBName] [nvarchar](100) NULL,
	[Description] [nvarchar](max) NULL,
	[WHAssortiment] [int] NULL,
	[UOM] [nvarchar](100) NULL,
	[Rezerv] [nvarchar](max) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL
) ON [PRIMARY]
