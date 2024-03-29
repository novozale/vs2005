USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_Salesmans]    Дата сценария: 05/29/2015 12:39:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tbl_WEB_Salesmans](
	[Code] [nvarchar](10) NOT NULL,
	[Name] [nvarchar](50) NOT NULL,
	[Email] [nvarchar](50) NULL,
	[City] [bigint] NULL,
	[OfficeLeader] [char](1) NULL,
	[OnDuty] [char](1) NULL,
	[IsActive] [bit] NULL,
	[Rezerv1] [nvarchar](50) NULL,
	[Rezerv2] [nvarchar](50) NULL,
	[RMStatus] [int] NOT NULL,
	[WEBStatus] [int] NOT NULL,
	[ScalaStatus] [int] NOT NULL,
 CONSTRAINT [PK_tbl_WEB_Salesmans] PRIMARY KEY CLUSTERED 
(
	[Code] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF