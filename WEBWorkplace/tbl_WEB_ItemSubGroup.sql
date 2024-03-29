USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_ItemSubGroup]    Дата сценария: 05/29/2015 12:38:02 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_ItemSubGroup](
	[SubgroupID] [nvarchar](20) NOT NULL,
	[SubgroupCode] [nvarchar](10) NOT NULL,
	[GroupCode] [nvarchar](10) NOT NULL,
	[Name] [nvarchar](100) NOT NULL,
	[Description] [nvarchar](max) NULL,
	[Rezerv1] [nvarchar](100) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL,
 CONSTRAINT [PK_tbl_WEB_ItemSubGroup] PRIMARY KEY CLUSTERED 
(
	[SubgroupID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
