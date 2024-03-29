USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_ItemGroup]    Дата сценария: 05/29/2015 12:37:49 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_ItemGroup](
	[Code] [nvarchar](50) NOT NULL,
	[Name] [nvarchar](100) NOT NULL,
	[WEBName] [nvarchar](255) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL,
 CONSTRAINT [PK_tbl_WEB_ItemGroup] PRIMARY KEY CLUSTERED 
(
	[Code] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
