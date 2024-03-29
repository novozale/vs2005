USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_Manufacturers]    Дата сценария: 05/29/2015 12:38:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_Manufacturers](
	[ID] [bigint] NOT NULL,
	[Name] [nvarchar](200) NOT NULL,
	[WEBName] [nvarchar](200) NULL,
	[Rezerv1] [nvarchar](100) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL,
 CONSTRAINT [PK_tbl_WEB_Manufacturers] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
