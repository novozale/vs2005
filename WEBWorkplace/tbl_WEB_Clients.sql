USE [ScaDataDB]
GO
/****** Объект:  Table [dbo].[tbl_WEB_Clients]    Дата сценария: 05/29/2015 12:36:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tbl_WEB_Clients](
	[Code] [nvarchar](50) NOT NULL,
	[Name] [nvarchar](80) NOT NULL,
	[Address] [nvarchar](200) NOT NULL,
	[Discount] [nvarchar](20) NOT NULL,
	[WorkOverWEB] [bit] NOT NULL,
	[BasePrice] [nvarchar](3) NOT NULL,
 CONSTRAINT [PK_tbl_WEB_Clients] PRIMARY KEY CLUSTERED 
(
	[Code] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
